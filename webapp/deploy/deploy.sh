#!/usr/bin/env bash
# ============================================================
# 超声报告数据纠错系统 — Linux 一键部署脚本
#
# 用法（在服务器上以 root 执行）：
#   bash deploy.sh                 # 部署到 /opt/qc-web，端口 8000
#   bash deploy.sh --port 9000     # 指定端口
#   bash deploy.sh --nginx         # 同时安装 nginx 反向代理（80 -> 本服务）
#   bash deploy.sh --no-systemd    # 仅安装文件与依赖，不注册 systemd 服务
#
# 可选环境变量：
#   APP_DIR       安装目录（默认 /opt/qc-web）
#   PORT          服务端口（默认 8000）
#   SERVICE_USER  运行用户（默认 qcweb，不存在则自动创建）
#   PIP_INDEX_URL pip 镜像源（内网离线部署可指向本地 pypi 镜像）
# ============================================================
set -euo pipefail

APP_DIR="${APP_DIR:-/opt/qc-web}"
PORT="${PORT:-8000}"
SERVICE_USER="${SERVICE_USER:-qcweb}"
INSTALL_NGINX=0
USE_SYSTEMD=1

while [ $# -gt 0 ]; do
  case "$1" in
    --app-dir)  APP_DIR="$2"; shift 2 ;;
    --port)     PORT="$2"; shift 2 ;;
    --user)     SERVICE_USER="$2"; shift 2 ;;
    --nginx)    INSTALL_NGINX=1; shift ;;
    --no-systemd) USE_SYSTEMD=0; shift ;;
    -h|--help)
      grep '^#' "$0" | sed 's/^# \{0,1\}//' | head -20
      exit 0 ;;
    *) echo "未知参数: $1（使用 --help 查看用法）"; exit 1 ;;
  esac
done

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
WEBAPP_DIR="$(dirname "$SCRIPT_DIR")"
REPO_DIR="$(dirname "$WEBAPP_DIR")"

echo "==> [1/7] 环境检查"
if [ "$(id -u)" -ne 0 ]; then
  echo "请以 root 运行（或 sudo bash deploy.sh）"; exit 1
fi
PYTHON_BIN="$(command -v python3 || true)"
if [ -z "$PYTHON_BIN" ]; then
  echo "未找到 python3，请先安装 Python 3.8+（如 apt install python3 python3-venv）"; exit 1
fi
"$PYTHON_BIN" - <<'EOF'
import sys
assert sys.version_info >= (3, 8), f"需要 Python 3.8+，当前 {sys.version.split()[0]}"
print(f"    Python: {sys.version.split()[0]}")
EOF

# 引擎与规则库：优先同仓库根目录
ENGINE_SRC=""
for cand in "$REPO_DIR/qc_engine.py" "$WEBAPP_DIR/qc_engine.py" "$SCRIPT_DIR/qc_engine.py"; do
  [ -f "$cand" ] && ENGINE_SRC="$cand" && break
done
if [ -z "$ENGINE_SRC" ]; then
  echo "未找到 qc_engine.py，请将其放在 webapp 同级目录"; exit 1
fi
RULES_SRC=""
for cand in "$REPO_DIR/rules.json" "$WEBAPP_DIR/rules.json" "$SCRIPT_DIR/rules.json"; do
  [ -f "$cand" ] && RULES_SRC="$cand" && break
done

echo "==> [2/7] 创建运行用户与安装目录"
if ! id "$SERVICE_USER" >/dev/null 2>&1; then
  useradd --system --home "$APP_DIR" --shell /usr/sbin/nologin "$SERVICE_USER"
  echo "    已创建系统用户: $SERVICE_USER"
fi
mkdir -p "$APP_DIR"/{logs,backend}

echo "==> [3/7] 复制程序文件"
cp -r "$WEBAPP_DIR/backend" "$APP_DIR/"
cp -r "$WEBAPP_DIR/frontend" "$APP_DIR/"
cp -r "$WEBAPP_DIR/samples" "$APP_DIR/"
cp "$WEBAPP_DIR/requirements.txt" "$APP_DIR/"
cp "$ENGINE_SRC" "$APP_DIR/qc_engine.py"
if [ -n "$RULES_SRC" ]; then
  if [ -f "$APP_DIR/rules.json" ]; then
    echo "    rules.json 已存在，保留现场配置（不覆盖）"
  else
    cp "$RULES_SRC" "$APP_DIR/rules.json"
  fi
else
  echo "    未找到 rules.json，将使用引擎内置默认规则"
fi

echo "==> [4/7] 创建虚拟环境并安装依赖"
if [ ! -d "$APP_DIR/venv" ]; then
  "$PYTHON_BIN" -m venv "$APP_DIR/venv"
fi
PIP_ARGS=()
if [ -d "$SCRIPT_DIR/wheels" ]; then
  PIP_ARGS+=(--no-index --find-links "$SCRIPT_DIR/wheels")
  echo "    检测到离线 wheels 目录，使用离线安装"
fi
"$APP_DIR/venv/bin/pip" install --quiet --upgrade pip
"$APP_DIR/venv/bin/pip" install --quiet "${PIP_ARGS[@]}" -r "$APP_DIR/requirements.txt"

echo "==> [5/7] 写入生产配置"
cat > "$APP_DIR/.env" <<EOF
# 超声报告数据纠错系统 — 运行配置（修改后重启服务生效）
QC_PORT=$PORT
QC_CORS_ORIGINS=*
QC_MAX_UPLOAD_BYTES=20971520
QC_MAX_REPORTS_PER_BATCH=5000
# QC_RULES_FILE=$APP_DIR/rules.json
EOF
chown -R "$SERVICE_USER":"$SERVICE_USER" "$APP_DIR"

if [ "$USE_SYSTEMD" = "1" ]; then
  echo "==> [6/7] 注册 systemd 服务"
  sed -e "s|@APP_DIR@|$APP_DIR|g" \
      -e "s|@PORT@|$PORT|g" \
      -e "s|@SERVICE_USER@|$SERVICE_USER|g" \
      "$SCRIPT_DIR/qc-web.service" > /etc/systemd/system/qc-web.service
  systemctl daemon-reload
  systemctl enable --now qc-web >/dev/null 2>&1
  sleep 2
  systemctl --no-pager --lines=0 status qc-web || true
else
  echo "==> [6/7] 跳过 systemd（--no-systemd），手动启动命令："
  echo "    cd $APP_DIR && venv/bin/uvicorn backend.main:app --host 0.0.0.0 --port $PORT"
fi

if [ "$INSTALL_NGINX" = "1" ]; then
  echo "==> [7/7] 配置 nginx 反向代理"
  if ! command -v nginx >/dev/null 2>&1; then
    echo "    未安装 nginx，请先安装后重新执行（apt install nginx / yum install nginx）"
  else
    sed -e "s|@PORT@|$PORT|g" "$SCRIPT_DIR/nginx.conf" > /etc/nginx/conf.d/qc-web.conf
    nginx -t && systemctl reload nginx
    echo "    nginx 已配置：http://<服务器IP>/  →  127.0.0.1:$PORT"
  fi
else
  echo "==> [7/7] 跳过 nginx（需要反代时加 --nginx 参数）"
fi

echo ""
echo "============================================================"
echo "部署完成！验证方式："
echo "  curl http://127.0.0.1:$PORT/api/health"
echo "  浏览器访问 http://<服务器IP>:$PORT/"
echo "常用运维命令："
echo "  systemctl status qc-web     # 查看状态"
echo "  systemctl restart qc-web    # 重启（修改 rules.json 后执行）"
echo "  journalctl -u qc-web -f     # 查看实时日志"
echo "============================================================"

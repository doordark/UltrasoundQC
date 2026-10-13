# 超声报告数据纠错系统（Web 端）

基于**离线超声报告质控引擎**（`qc_engine.py`，纯 Python 标准库、零外网依赖）的
Web 端数据纠错系统：**前后端分离**架构，支持常见格式文件上传解析、一键触发纠错、
实时状态提示、问题逐条定位与修改建议，并预置多组标准测试数据集，**免上传即可演示**纠错效果。

| 特性           | 说明                                                              |
| -------------- | ----------------------------------------------------------------- |
| 前后端分离     | 后端 REST API（FastAPI）+ 原生 HTML/CSS/JS 前端，接口完全适配     |
| 多格式上传     | `.txt / .md / .json / .csv / .tsv / .xlsx` 自动解析               |
| 一键纠错       | 上传 → 解析 → 纠错全流程可视化，步骤条实时状态提示                |
| 预置测试数据集 | 5 组标准数据集覆盖 A~H 全部规则，点击即检                         |
| 结果清晰直观   | 汇总统计卡片 + 规则分布 + 问题卡片（原文片段/修改建议）+ 报告明细 |
| 结果导出       | 一键导出纠错结果为 Excel（质控概览 + 错误明细两个 Sheet）         |
| 生产就绪       | 一键部署脚本、systemd 守护、nginx 反代、环境变量配置              |
| 轻量稳定       | 引擎纯标准库实现，单份报告处理约 0.1~0.3ms，无 GPU/外网依赖       |

---

## 1. 系统架构

```text
┌───────────────────────────┐        ┌──────────────────────────────┐
│  前端（浏览器）            │  HTTP  │  后端（FastAPI / uvicorn）    │
│  frontend/                │ ─────► │  backend/main.py              │
│  · 上传 / 拖拽文件        │  JSON  │  · REST API（/api/*）          │
│  · 一键纠错 + 状态步骤条   │        │  · 文件解析（parsers.py）      │
│  · 结果展示 / 数据集演示   │        │  · 引擎适配（engine_adapter.py）│
└───────────────────────────┘        └──────────────┬───────────────┘
                                                    │ 调用
                                     ┌──────────────▼───────────────┐
                                     │  质控引擎 qc_engine.py        │
                                     │  规则库 rules.json（外置可改） │
                                     │  规则 A~H · 毫秒级 · 零依赖   │
                                     └──────────────────────────────┘
```

- 前后端通过 JSON REST 接口通信，**同源部署无跨域问题**；分离部署时后端已开启 CORS，
  通过 `QC_CORS_ORIGINS` 配置允许的前端来源。
- 引擎与规则库为独立模块：修改 `rules.json` 后重启服务即生效，无需改动 Web 代码。

## 2. 目录结构

```text
webapp/
├── README.md               # 本文档
├── requirements.txt        # 后端依赖（fastapi / uvicorn / python-multipart）
├── backend/                # 后端服务
│   ├── main.py             # FastAPI 应用与 REST 接口
│   ├── engine_adapter.py   # 质控引擎加载与批量纠错封装
│   ├── parsers.py          # 上传文件解析（txt/json/csv/tsv/xlsx，零第三方依赖）
│   ├── datasets.py         # 预置测试数据集加载
│   ├── config.py           # 运行配置（环境变量）
│   └── datasets/           # 预置测试数据集（JSON，可自行增删）
├── frontend/               # 前端静态页面（原生 JS，无构建步骤）
│   ├── index.html
│   └── assets/             # styles.css / app.js
├── samples/                # 上传演示样例文件（json/csv/txt/xlsx 各 5 份报告）
└── deploy/                 # Linux 生产部署
    ├── deploy.sh           # 一键部署脚本（venv + 依赖 + systemd + nginx）
    ├── qc-web.service      # systemd 服务单元模板
    └── nginx.conf          # nginx 反向代理模板
```

运行时还需两个文件（位于 `webapp/` 同级目录，部署脚本会自动带上）：

- `qc_engine.py` — 质控引擎
- `rules.json` — 规则库（可选；缺失时使用引擎内置默认规则）

## 3. 本地启动

### 3.1 环境要求

- Python **3.8+**（推荐 3.10+）
- 依赖仅 3 个纯 Python 包：`fastapi`、`uvicorn`、`python-multipart`

### 3.2 Windows（PowerShell）

```powershell
cd webapp
python -m venv venv
venv\Scripts\pip install -r requirements.txt
venv\Scripts\python -m uvicorn backend.main:app --host 127.0.0.1 --port 8000
```

### 3.3 Linux / macOS

```bash
cd webapp
python3 -m venv venv
venv/bin/pip install -r requirements.txt
venv/bin/uvicorn backend.main:app --host 0.0.0.0 --port 8000
```

启动后浏览器访问 **http://127.0.0.1:8000/** 即可使用；
`http://127.0.0.1:8000/docs` 为自动生成的交互式 API 文档。

> 首次启动自检：`curl http://127.0.0.1:8000/api/health` 应返回
> `"status": "ok"`、`engine_version` 与 `rules_source`。

## 4. 测试用例使用方法

### 4.1 预置测试数据集（免上传，直接演示）

切换到「**测试数据集**」页签 → 点击「**一键纠错**」即可；「查看数据」可先预览报告原文。

| 数据集               | 内容                               | 预期结果                     |
| -------------------- | ---------------------------------- | ---------------------------- |
| 正常报告集（对照组） | 4 份规范阴性报告                   | **0 条问题**（验证零误报）   |
| 典型错误集           | 规则 A/B/C/D 各 1 份               | 4 条问题（A、B、C、D 各 1）  |
| 完整性规范集         | 规则 E×2、F×2                      | 4 条问题（E、F 各 2）        |
| 性别与部位集         | 规则 G×2 份、H×2 份                | 5 条问题（G 3 条、H 2 条）   |
| 混合工作列表         | 10 份混合正常/错误（模拟批量送检） | 8 条问题（含单份多问题 G×2） |

数据集为 `backend/datasets/*.json`，格式如下，**新增数据集只需添加 JSON 文件并重启服务**：

```json
{
  "id": "my-dataset",
  "name": "数据集名称",
  "description": "用途说明",
  "category": "分类",
  "tags": ["A", "B"],
  "expectation": "预期检出结果说明",
  "reports": [
    {
      "report_id": "X001",
      "exam_item": "甲状腺彩超",
      "gender": "女",
      "age": "44",
      "finding": "超声所见……",
      "conclusion": "超声提示……",
      "expect_rules": ["A"]
    }
  ]
}
```

### 4.2 上传样例文件测试

「上传纠错」页签顶部提供 4 种格式的样例文件下载（各含 5 份报告）：

| 样例文件              | 格式特点                               | 预期结果                        |
| --------------------- | -------------------------------------- | ------------------------------- |
| `sample_reports.json` | `{"reports": [...]}` 包装的标准 JSON   | 解析 5 份，检出 4 条（A/B/C/H） |
| `sample_reports.csv`  | 表头行 + 每行一份报告                  | 同上（四种格式结果一致）        |
| `sample_reports.xlsx` | Excel 表格，首行表头                   | 同上                            |
| `sample_reports.txt`  | 标签行 + `------` 分隔行的多份文本报告 | 同上                            |

操作步骤：上传文件 → 查看解析预览 → 点击「**开始纠错**」→ 查看结果卡片与报告明细。

### 4.3 常用自测命令（curl）

```bash
# 健康探针
curl http://127.0.0.1:8000/api/health

# 预置数据集一键纠错
curl -X POST http://127.0.0.1:8000/api/datasets/typical-errors/check

# 上传文件一键纠错
curl -X POST -F "file=@webapp/samples/sample_reports.csv" \
     http://127.0.0.1:8000/api/upload/check
```

## 5. 上传文件格式说明

| 格式        | 说明                                                                                                    |
| ----------- | ------------------------------------------------------------------------------------------------------- |
| **JSON**    | 单个报告对象、对象数组或 `{"reports": [...]}` / `{"data": [...]}` 包装                                  |
| **CSV/TSV** | 首行为表头，列名使用标准字段名或中文键名自动映射；每行一份报告                                          |
| **XLSX**    | Excel/WPS 工作表，首行为表头（列名同 CSV），仅解析第一个工作表；无需安装 openpyxl（内置轻量解析）       |
| **TXT/MD**  | 多份报告按分隔行（`------` 等）或空行分隔；每份内使用「检查项目：」「超声所见：」「超声提示：」等标签行 |

**标准字段**（兼容中文键名与常见大小写变体）：

| 字段         | 中文键名                       | 必填                             |
| ------------ | ------------------------------ | -------------------------------- |
| `exam_item`  | 检查项目 / 检查部位 / 项目名称 | 否                               |
| `gender`     | 性别                           | 否                               |
| `age`        | 年龄                           | 否                               |
| `finding`    | 超声所见                       | **是**（与 conclusion 至少一项） |
| `conclusion` | 超声提示 / 超声结论            | **是**（与 finding 至少一项）    |

## 6. 接口说明

Base URL：`http://<服务器>:<端口>`；请求/响应均为 JSON（上传接口为 `multipart/form-data`）。

### 6.1 接口总览

| 方法 | 路径                       | 说明                                        |
| ---- | -------------------------- | ------------------------------------------- |
| GET  | `/api/health`              | 服务状态、引擎版本、规则来源探针            |
| GET  | `/api/rules`               | 当前生效的规则库                            |
| GET  | `/api/datasets`            | 预置测试数据集列表                          |
| GET  | `/api/datasets/{id}`       | 数据集详情（含报告列表）                    |
| POST | `/api/datasets/{id}/check` | 对预置数据集一键纠错                        |
| POST | `/api/check`               | 单份报告纠错                                |
| POST | `/api/check/batch`         | 批量纠错                                    |
| POST | `/api/export`              | 导出纠错结果为 Excel（质控概览 + 错误明细） |
| POST | `/api/upload`              | 上传并解析文件（不纠错，供预览）            |
| POST | `/api/upload/check`        | 上传 → 解析 → 纠错，单请求返回完整结果      |

### 6.2 POST /api/check — 单份报告纠错

请求体为标准 JSON 报告对象（兼容中文键名）：

```json
{
  "exam_item": "甲状腺彩超",
  "gender": "女",
  "age": "44",
  "finding": "甲状腺右叶见一大小约1.2cm×0.9cm低回声结节，边界清，形态规则。",
  "conclusion": "甲状腺未见明显异常。"
}
```

响应（200）：

```json
{
  "ok": true,
  "engine_version": "1.1.0",
  "report": { "exam_item": "甲状腺彩超", "gender": "女", "age": "44" },
  "parsed": {
    "finding": "……",
    "conclusion": "……",
    "parse_notes": ["标准 JSON 输入"]
  },
  "issues": [
    {
      "rule_id": "A",
      "rule": "结论遗漏病变",
      "severity": "high",
      "field": "conclusion",
      "fragment": "甲状腺未见明显异常。",
      "message": "所见描述阳性病变（结节等），但结论为阴性表述“未见明显异常”，疑似遗漏病变",
      "suggestion": "请核对所见与结论是否矛盾：若确有病灶，结论应逐项描述病变部位、大小与性质"
    }
  ],
  "issue_count": 1,
  "elapsed_ms": 0.12
}
```

### 6.3 POST /api/check/batch — 批量纠错

```json
{ "reports": [ { "exam_item": "……", "finding": "……", "conclusion": "……" }, … ] }
```

响应（200）：

```json
{
  "ok": true,
  "engine_version": "1.1.0",
  "total": 2,
  "summary": {
    "total": 2,
    "reports_with_issues": 1,
    "issue_total": 1,
    "elapsed_ms": 0.35,
    "by_rule": { "A": 1 },
    "by_severity": { "high": 1 }
  },
  "results": [
    {
      "index": 0,
      "report_id": "R001",
      "report": { "exam_item": "……", "gender": "女", "age": "44", "finding": "……", "conclusion": "……" },
      "parse_notes": ["标准 JSON 输入"],
      "issues": [ … ],
      "issue_count": 1,
      "elapsed_ms": 0.12
    }
  ]
}
```

### 6.4 POST /api/upload / /api/upload/check — 文件上传

```bash
curl -X POST -F "file=@sample_reports.csv" http://127.0.0.1:8000/api/upload/check
```

- `/api/upload` 返回 `{ ok, filename, format, report_count, reports, warnings }`（仅解析）；
- `/api/upload/check` 返回批量纠错结果 + `source: {type:"upload", filename, format}` 与 `warnings`。

### 6.5 POST /api/export — 导出纠错结果（Excel）

请求体与 `/api/check/batch` 相同，可选 `source` 作为导出文件名前缀标识：

```json
{
  "source": "文件：sample_reports.json",
  "reports": [ { "exam_item": "……", "finding": "……", "conclusion": "……" }, … ]
}
```

响应为 Excel 文件下载（`Content-Disposition: attachment`），包含两个 Sheet：

- **质控概览**：生成时间、引擎版本、规则来源、报告文件数、存在问题报告数、问题总数、按规则统计；
- **错误明细**：文件名/报告ID、是否存在问题、规则、严重程度、位置、原文片段、问题描述、修改建议。

文件名格式：`<source>_<yyyyMMdd_HHmmss>.xlsx`（`source` 中的非法文件名字符自动过滤）。

```bash
curl -X POST -H "Content-Type: application/json" --data-binary "@body.json" \
     -o result.xlsx http://127.0.0.1:8000/api/export
```

前端「导出纠错结果（Excel）」按钮即调用本接口：纠错完成后点击即可下载，导出期间按钮置灰并提示进度。

### 6.6 错误响应

所有接口出错返回 `{"detail": "错误原因"}`，HTTP 状态码：`400`（参数/格式错误）、
`404`（数据集不存在）、`413`（文件超限）。

## 7. Linux 服务器部署指南

### 7.1 一键部署（推荐）

将整个项目（含 `webapp/`、`qc_engine.py`、`rules.json`）上传到服务器，例如 `/tmp/qc`，然后：

```bash
cd /tmp/qc/webapp/deploy
sudo bash deploy.sh                 # 部署到 /opt/qc-web，端口 8000，注册 systemd 服务
sudo bash deploy.sh --port 9000     # 自定义端口
sudo bash deploy.sh --nginx         # 同时配置 nginx 反向代理（80 端口访问）
```

脚本自动完成：环境检查 → 创建运行用户 `qcweb` → 复制程序/引擎/规则库 → 创建 venv 并安装依赖
→ 写入生产配置 `.env` → 注册并启动 systemd 服务 →（可选）配置 nginx。

部署完成后：

```bash
curl http://127.0.0.1:8000/api/health   # 验证服务
# 浏览器访问 http://<服务器IP>:8000/    （或 nginx 模式下 http://<服务器IP>/）
```

### 7.2 手动部署（无 systemd 环境）

```bash
mkdir -p /opt/qc-web && cd /opt/qc-web
# 复制 webapp/{backend,frontend,samples,requirements.txt}、qc_engine.py、rules.json 到当前目录
python3 -m venv venv
venv/bin/pip install -r requirements.txt
venv/bin/uvicorn backend.main:app --host 0.0.0.0 --port 8000
```

### 7.3 离线 / 内网部署

1. 在有外网的机器上下载依赖包：

   ```bash
   pip download -r requirements.txt -d wheels
   ```

2. 将 `wheels/` 目录放到服务器的 `webapp/deploy/` 下再执行 `deploy.sh`，
   脚本检测到 `wheels/` 后自动使用 `--no-index --find-links wheels` 离线安装。

### 7.4 nginx 反向代理（可选）

`deploy.sh --nginx` 自动生成 `/etc/nginx/conf.d/qc-web.conf`（模板见 `deploy/nginx.conf`）。
前端与 API 同源，无需任何跨域配置。若前端与后端分机部署：

- 前端静态文件（`frontend/`、`samples/`）由任意 Web 服务器托管；
- 修改前端 `frontend/assets/app.js` 顶部 `const API = ""` 为后端地址；
- 后端通过 `QC_CORS_ORIGINS` 允许前端来源域名（逗号分隔多个）。

### 7.5 防火墙放行

```bash
firewall-cmd --permanent --add-port=8000/tcp && firewall-cmd --reload   # firewalld
# 或
ufw allow 8000/tcp                                                       # ufw
```

## 8. 生产环境配置

配置文件：`/opt/qc-web/.env`（修改后 `systemctl restart qc-web` 生效）。

| 环境变量                   | 默认值     | 说明                                           |
| -------------------------- | ---------- | ---------------------------------------------- |
| `QC_PORT`                  | `8000`     | 服务端口（同时需改 systemd ExecStart）         |
| `QC_CORS_ORIGINS`          | `*`        | 允许的跨域来源，逗号分隔（如 `https://a.com`） |
| `QC_MAX_UPLOAD_BYTES`      | `20971520` | 单文件上传大小上限（字节，默认 20MB）          |
| `QC_MAX_REPORTS_PER_BATCH` | `5000`     | 单批纠错报告数上限                             |
| `QC_RULES_FILE`            | 自动查找   | 规则库路径（默认用安装目录下 `rules.json`）    |
| `QC_ENGINE_FILE`           | 自动查找   | 引擎源码路径（默认安装目录下 `qc_engine.py`）  |

**规则调优（零重编译）**：直接修改 `/opt/qc-web/rules.json`（漏报补词、误报精细化正则），
执行 `systemctl restart qc-web` 即生效。重新部署时脚本**不会覆盖**已存在的 `rules.json`，
现场规则修改可安全保留。

## 9. 运维常用命令

```bash
systemctl status qc-web        # 查看服务状态
systemctl restart qc-web       # 重启服务（改 rules.json 后执行）
systemctl stop qc-web          # 停止服务
journalctl -u qc-web -f        # 实时日志
tail -f /opt/qc-web/logs/qc-web.log
```

## 10. 质控规则速览（A~H）

| 规则 | 名称               | 检查逻辑简述                                          | 严重度 |
| ---- | ------------------ | ----------------------------------------------------- | ------ |
| A    | 结论遗漏病变       | 所见含强病灶词（结节/囊肿/占位等），结论为纯阴性表述  | high   |
| B    | 左右侧不符         | 所见与结论中同一器官左右侧冲突                        | high   |
| C    | 用词勘误           | 命中纠错映射表（如 回音→回声），给出规范替换建议      | medium |
| D    | 分级冲突           | 分级号不一致 / 风险征象分级偏低 / 非标准分级写法      | high   |
| E    | 结论未涵盖所见     | 所见病灶在结论中无任何等价呼应                        | high   |
| F    | 完整性与书写规范   | 建议缺时间数值、数值粘连、检查项目与内容不符等        | medium |
| G    | 性别与检查项目冲突 | 性别与检查项目/所见部位矛盾（男见子宫、女见前列腺等） | high   |
| H    | 所见遗漏检查部位   | 多部位检查项目仅描述部分部位，报出缺失部位            | high   |

规则完整逻辑与词表见仓库根目录 `README.md` 与 `rules.json`。

## 11. 常见问题

**Q：启动报 `未找到 qc_engine.py`？**
A：引擎需位于 `webapp/` 同级目录（或通过 `QC_ENGINE_FILE` 指定路径）；`deploy.sh` 会自动复制。

**Q：上传报 400「不支持的文件类型」？**
A：仅支持 `.txt/.md/.json/.csv/.tsv/.xlsx`；请检查文件扩展名是否正确。

**Q：修改了 rules.json 为什么不生效？**
A：规则在启动时加载，修改后需 `systemctl restart qc-web` 重启服务。

**Q：前端能打开但接口报跨域？**
A：同源部署（推荐）无跨域问题；分离部署时设置 `QC_CORS_ORIGINS` 为前端域名并重启。

**Q：如何备份现场配置？**
A：备份 `/opt/qc-web/rules.json` 与 `/opt/qc-web/.env` 两个文件即可，重新部署不会覆盖它们。

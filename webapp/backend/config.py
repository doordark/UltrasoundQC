"""运行配置：全部支持环境变量覆盖（生产部署通过 .env 注入）。"""
from __future__ import annotations

import os
from pathlib import Path

BACKEND_DIR = Path(__file__).resolve().parent
WEBAPP_ROOT = BACKEND_DIR.parent
FRONTEND_DIR = WEBAPP_ROOT / "frontend"
DATASETS_DIR = BACKEND_DIR / "datasets"

SERVICE_NAME = "超声报告数据纠错系统"
SERVICE_VERSION = "1.0.0"

# 监听地址（uvicorn 命令行参数优先，此处供 /api/health 展示）
HOST = os.environ.get("QC_HOST", "0.0.0.0")
PORT = int(os.environ.get("QC_PORT", "8000"))

# 跨域：前后端分离开发时前端 dev server 的来源；"*" 表示不限制
CORS_ORIGINS = [
    o.strip()
    for o in os.environ.get("QC_CORS_ORIGINS", "*").split(",")
    if o.strip()
]

# 上传限制（字节）
MAX_UPLOAD_BYTES = int(os.environ.get("QC_MAX_UPLOAD_BYTES", str(20 * 1024 * 1024)))
MAX_REPORTS_PER_BATCH = int(os.environ.get("QC_MAX_REPORTS_PER_BATCH", "5000"))

SUPPORTED_EXTS = {".txt", ".md", ".json", ".csv", ".tsv", ".xlsx"}

"""超声报告数据纠错系统 — 后端服务（FastAPI）。

前后端分离：本服务提供 REST API（/api/*），前端静态页面由 "/" 挂载，
亦可由 nginx 托管前端、反向代理 /api（见 deploy/nginx.conf）。

启动：uvicorn backend.main:app --host 0.0.0.0 --port 8000
"""
from __future__ import annotations

import os
import platform
import tempfile
import time
from datetime import datetime

from fastapi import Body, FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from . import config, datasets, parsers
from .engine_adapter import (
    ENGINE_FILE,
    ENGINE_VERSION,
    RULES_FILE,
    RULES_SOURCE,
    check_one,
    check_text,
    engine,
    normalize_report,
    rules_public,
    run_batch,
)

app = FastAPI(
    title=config.SERVICE_NAME,
    version=config.SERVICE_VERSION,
    description="离线超声报告质控引擎（A~J 十类规则）的 Web 端数据纠错系统",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=config.CORS_ORIGINS,
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============================ 基础 ============================

@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "service": config.SERVICE_NAME,
        "service_version": config.SERVICE_VERSION,
        "engine_version": ENGINE_VERSION,
        "rules_source": RULES_SOURCE,
        "rules_file": RULES_FILE,
        "dataset_count": len(datasets.list_datasets()),
        "supported_formats": sorted(config.SUPPORTED_EXTS),
        "server_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "python": platform.python_version(),
    }


@app.get("/api/rules")
def get_rules():
    return {
        "ok": True,
        "rules_source": RULES_SOURCE,
        "rules_file": RULES_FILE,
        "engine_version": ENGINE_VERSION,
        "rules": rules_public(),
    }


# ============================ 测试数据集 ============================

@app.get("/api/datasets")
def list_datasets():
    return {"ok": True, "datasets": datasets.list_datasets()}


@app.get("/api/datasets/{dataset_id}")
def get_dataset(dataset_id: str):
    ds = datasets.get_dataset(dataset_id)
    if ds is None:
        raise HTTPException(status_code=404, detail=f"测试数据集不存在：{dataset_id}")
    return {"ok": True, **ds}


@app.post("/api/datasets/{dataset_id}/check")
def check_dataset(dataset_id: str):
    """免上传一键纠错：直接对预置测试数据集执行质控。"""
    ds = datasets.get_dataset(dataset_id)
    if ds is None:
        raise HTTPException(status_code=404, detail=f"测试数据集不存在：{dataset_id}")
    batch = run_batch(ds["reports"])
    return {"ok": True, "source": {"type": "dataset", **{k: ds[k] for k in ("id", "name", "description")}}, **batch}


# ============================ 纠错接口 ============================

@app.post("/api/check")
def check_single(payload: dict = Body(...)):
    """单份报告质检：标准 JSON 报告对象，或 {"report": {...}} / {"text": "..."}。"""
    obj = payload.get("report") if isinstance(payload.get("report"), dict) else payload
    if isinstance(payload.get("text"), str):
        return {"ok": True, **check_text(payload["text"])}
    if not isinstance(obj, dict):
        raise HTTPException(status_code=400, detail="请求体必须是 JSON 对象")
    result = check_one(obj)
    return {"ok": True, **result}


def _normalize_reports(payload: dict) -> tuple[list, list]:
    """批量请求体 → 标准报告列表：支持 {"reports": [...]} 与 {"texts": [...]}。"""
    reports_in = payload.get("reports")
    texts_in = payload.get("texts")
    if isinstance(texts_in, list) and not isinstance(reports_in, list):
        reports_in = [engine.parse_report(str(t)) for t in texts_in]
    if not isinstance(reports_in, list) or not reports_in:
        raise HTTPException(status_code=400, detail="缺少 reports 数组（或 texts 数组）")
    if len(reports_in) > config.MAX_REPORTS_PER_BATCH:
        raise HTTPException(
            status_code=400,
            detail=f"单批最多 {config.MAX_REPORTS_PER_BATCH} 份报告，请分批提交",
        )

    reports, warnings = [], []
    for i, item in enumerate(reports_in):
        if isinstance(item, str):
            parsed = engine.parse_report(item)
            rep = {"report_id": f"R{i + 1:03d}", "exam_item": parsed.get("exam_item", ""),
                   "gender": parsed.get("gender", ""), "age": parsed.get("age", ""),
                   "finding": parsed["finding"], "conclusion": parsed["conclusion"]}
        else:
            rep = normalize_report(item)
            if rep is None:
                warnings.append(f"第 {i + 1} 条缺少超声所见/超声提示字段，已跳过")
                continue
            rep["report_id"] = str(item.get("report_id") or item.get("id") or f"R{i + 1:03d}")
        reports.append(rep)

    if not reports:
        raise HTTPException(status_code=400, detail="没有任何可处理的报告（均缺少超声所见/超声提示）")
    return reports, warnings


@app.post("/api/check/batch")
def check_batch(payload: dict = Body(...)):
    """批量质检：{"reports": [标准报告对象...]} 或 {"texts": ["报告全文"...]}。"""
    reports, warnings = _normalize_reports(payload)
    batch = run_batch(reports)
    return {"ok": True, "source": {"type": "inline"}, "warnings": warnings, **batch}


@app.post("/api/export")
def export_batch(payload: dict = Body(...)):
    """导出纠错结果为 Excel（两个 Sheet：质控概览 + 错误明细）。

    请求体同 /api/check/batch，可选 "source" 作为文件名前缀标识。
    """
    reports, _warnings = _normalize_reports(payload)
    results = [{"file": f"{rep.get('source_file')}/{rep['report_id']}" if rep.get("source_file")
                else rep["report_id"],
                "result": check_one(rep)} for rep in reports]

    fd, tmp_path = tempfile.mkstemp(prefix="qc_export_", suffix=".xlsx")
    os.close(fd)
    try:
        engine.export_excel(results, tmp_path, RULES_SOURCE)
    except Exception as exc:  # 清理临时文件后抛出
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise HTTPException(status_code=500, detail=f"结果导出失败：{exc}") from exc

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    source = str(payload.get("source") or "纠错结果").strip()
    safe = "".join(c for c in source if c not in '\\/:*?"<>|\n\t')[:40] or "纠错结果"
    return FileResponse(
        tmp_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"{safe}_{stamp}.xlsx",
        background=BackgroundTask(os.remove, tmp_path),
    )


# ============================ 文件上传 ============================

MAX_UPLOAD_FILES = 20
MAX_UPLOAD_TOTAL_BYTES = 100 * 1024 * 1024  # 单次上传总量上限 100MB


async def _read_upload(file: UploadFile) -> bytes:
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail=f"上传文件“{file.filename or ''}”为空")
    if len(data) > config.MAX_UPLOAD_BYTES:
        raise HTTPException(
            status_code=413,
            detail=f"文件“{file.filename or ''}”超过大小限制（{config.MAX_UPLOAD_BYTES // (1024 * 1024)}MB）",
        )
    return data


async def _read_uploads(files: list[UploadFile]) -> list[tuple[str, bytes]]:
    """读取一批上传文件（1~MAX_UPLOAD_FILES 个），校验单文件与总量限制。"""
    if not files:
        raise HTTPException(status_code=400, detail="未收到上传文件")
    if len(files) > MAX_UPLOAD_FILES:
        raise HTTPException(status_code=400, detail=f"单次最多上传 {MAX_UPLOAD_FILES} 个文件")
    items: list[tuple[str, bytes]] = []
    total = 0
    for f in files:
        data = await _read_upload(f)
        total += len(data)
        if total > MAX_UPLOAD_TOTAL_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"上传文件总量超过限制（{MAX_UPLOAD_TOTAL_BYTES // (1024 * 1024)}MB）",
            )
        items.append((f.filename or "", data))
    return items


def _merge_parsed(items: list[tuple[str, bytes]]) -> dict:
    """逐文件解析并合并报告：report_id 全局统一编号，报告标注 source_file 来源文件。"""
    reports: list[dict] = []
    warnings: list[str] = []
    formats: list[str] = []
    names: list[str] = []
    for fname, data in items:
        try:
            parsed = parsers.parse_upload(fname, data)
        except ValueError as exc:
            warnings.append(f"「{fname}」解析失败：{exc}")
            continue
        formats.append(parsed["format"])
        names.append(fname)
        for rep in parsed["reports"]:
            rep["source_file"] = fname
            reports.append(rep)
        warnings.extend(f"「{fname}」{w}" for w in parsed["warnings"])
    if not reports:
        raise HTTPException(status_code=400, detail="未能从上传文件中解析出任何报告，请检查文件内容与格式")
    for i, rep in enumerate(reports):
        rep["report_id"] = f"R{i + 1:03d}"
    fmt = formats[0] if len(set(formats)) == 1 else "mixed"
    filename = names[0] if len(names) == 1 else f"{len(names)} 个文件"
    return {"format": fmt, "filename": filename, "files": names,
            "reports": reports, "warnings": warnings}


@app.post("/api/upload")
async def upload_parse(file: list[UploadFile] = File(...)):
    """上传并解析一个或多个文件（不纠错），返回合并的标准报告列表供前端预览。

    同名表单字段 `file` 重复出现即为多文件。
    """
    items = await _read_uploads(file)
    merged = _merge_parsed(items)
    return {
        "ok": True,
        "filename": merged["filename"],
        "files": merged["files"],
        "file_count": len(items),
        "format": merged["format"],
        "report_count": len(merged["reports"]),
        "reports": merged["reports"],
        "warnings": merged["warnings"],
    }


@app.post("/api/upload/check")
async def upload_check(file: list[UploadFile] = File(...)):
    """一键纠错：上传（可多个文件）→ 解析合并 → 批量质控，单请求返回完整结果。"""
    items = await _read_uploads(file)
    merged = _merge_parsed(items)
    batch = run_batch(merged["reports"])
    return {
        "ok": True,
        "source": {"type": "upload", "filename": merged["filename"],
                   "files": merged["files"], "format": merged["format"]},
        "warnings": merged["warnings"],
        **batch,
    }


# ============================ 静态资源 ============================

SAMPLES_DIR = config.WEBAPP_ROOT / "samples"
if SAMPLES_DIR.is_dir():
    app.mount("/samples", StaticFiles(directory=str(SAMPLES_DIR)), name="samples")

if config.FRONTEND_DIR.is_dir():
    app.mount("/", StaticFiles(directory=str(config.FRONTEND_DIR), html=True), name="frontend")

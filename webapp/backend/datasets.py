"""预置测试数据集：读取 backend/datasets/*.json，供前端免上传直接演示纠错。"""
from __future__ import annotations

import json
from pathlib import Path

from . import config

_CACHE: dict | None = None


def _load_all() -> dict:
    global _CACHE
    if _CACHE is not None:
        return _CACHE

    datasets = {}
    for path in sorted(config.DATASETS_DIR.glob("*.json")):
        try:
            meta = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        ds_id = str(meta.get("id") or path.stem)
        reports = []
        for i, rep in enumerate(meta.get("reports", [])):
            reports.append({
                "report_id": str(rep.get("report_id") or f"R{i + 1:03d}"),
                "exam_item": str(rep.get("exam_item", "")),
                "gender": str(rep.get("gender", "")),
                "age": str(rep.get("age", "")),
                "finding": str(rep.get("finding", "")),
                "conclusion": str(rep.get("conclusion", "")),
                "expect_rules": rep.get("expect_rules", []),
            })
        datasets[ds_id] = {
            "id": ds_id,
            "name": meta.get("name", ds_id),
            "description": meta.get("description", ""),
            "category": meta.get("category", "通用"),
            "tags": meta.get("tags", []),
            "expectation": meta.get("expectation", ""),
            "report_count": len(reports),
            "reports": reports,
        }
    _CACHE = datasets
    return datasets


def list_datasets() -> list[dict]:
    out = []
    for ds in _load_all().values():
        item = {k: v for k, v in ds.items() if k != "reports"}
        out.append(item)
    return out


def get_dataset(ds_id: str) -> dict | None:
    return _load_all().get(ds_id)

"""质控引擎适配层：加载 qc_engine 模块与规则库，封装批量纠错入口。

引擎为纯 Python 标准库实现（仓库根目录 qc_engine.py），本模块负责：
1. 定位并导入 qc_engine（支持 QC_ENGINE_FILE 环境变量指定路径）；
2. 加载规则库 rules.json（支持 QC_RULES_FILE 环境变量指定路径，缺失时用内置规则）；
3. 提供 run_batch()：对标准 JSON 报告列表批量执行 A~H 质控规则并汇总统计。
"""
from __future__ import annotations

import importlib.util
import os
import sys
import time
from pathlib import Path

from . import config


def _candidate_engine_files() -> list[Path]:
    cands = []
    env = os.environ.get("QC_ENGINE_FILE")
    if env:
        cands.append(Path(env))
    cands += [
        config.WEBAPP_ROOT / "qc_engine.py",
        config.WEBAPP_ROOT.parent / "qc_engine.py",
        config.BACKEND_DIR / "qc_engine.py",
    ]
    return cands


def _candidate_rules_files() -> list[Path]:
    cands = []
    env = os.environ.get("QC_RULES_FILE")
    if env:
        cands.append(Path(env))
    cands += [
        config.WEBAPP_ROOT / "rules.json",
        config.WEBAPP_ROOT.parent / "rules.json",
        config.BACKEND_DIR / "rules.json",
    ]
    return cands


def _load_engine_module():
    for cand in _candidate_engine_files():
        if cand.exists():
            spec = importlib.util.spec_from_file_location("qc_engine", cand)
            if spec is None or spec.loader is None:
                continue
            mod = importlib.util.module_from_spec(spec)
            sys.modules["qc_engine"] = mod
            spec.loader.exec_module(mod)
            return mod, cand.resolve()
    raise FileNotFoundError(
        "未找到 qc_engine.py，请将其放在 webapp 同级目录，或通过 QC_ENGINE_FILE 指定路径"
    )


engine, ENGINE_FILE = _load_engine_module()

# 规则加载：显式路径优先，否则交给引擎默认查找（与引擎同目录的 rules.json）
_rules_path = next((p for p in _candidate_rules_files() if p.exists()), None)
RULES, RULES_SOURCE = engine.load_rules(str(_rules_path) if _rules_path else None)
RULES = engine.compile_rules(RULES)
RULES_FILE = str(_rules_path.resolve()) if _rules_path else "（内置默认规则）"
ENGINE_VERSION = engine.ENGINE_VERSION


def normalize_report(obj) -> dict | None:
    """把任意来源的报告对象归一化为标准五字段结构（含中文键名别名兼容）。"""
    if not isinstance(obj, dict):
        return None
    parsed = engine.parse_json_report(obj)
    report = {
        "exam_item": parsed.get("exam_item", ""),
        "gender": parsed.get("gender", ""),
        "age": parsed.get("age", ""),
        "finding": parsed.get("finding", ""),
        "conclusion": parsed.get("conclusion", ""),
    }
    if not report["finding"] and not report["conclusion"]:
        return None
    return report


def check_one(report: dict) -> dict:
    """单份标准 JSON 报告质检，返回与引擎 /api/check 一致的结构化结果。"""
    return engine.check_json(report, RULES)


def check_text(text: str) -> dict:
    """单份纯文本报告质检（自适应解析）。"""
    return engine.check_report(text, RULES)


def run_batch(reports: list[dict], start_index: int = 0) -> dict:
    """批量质检并汇总：reports 为标准五字段结构列表。"""
    t0 = time.perf_counter()
    results = []
    by_rule: dict = {}
    by_severity: dict = {}
    issue_total = 0
    reports_with_issues = 0

    for offset, rep in enumerate(reports):
        res = engine.check_json(rep, RULES)
        issues = res["issues"]
        issue_total += len(issues)
        if issues:
            reports_with_issues += 1
        for it in issues:
            by_rule[it["rule_id"]] = by_rule.get(it["rule_id"], 0) + 1
            by_severity[it["severity"]] = by_severity.get(it["severity"], 0) + 1
        results.append({
            "index": start_index + offset,
            "report_id": rep.get("report_id") or f"R{start_index + offset + 1:03d}",
            "report": {k: rep.get(k, "") for k in ("exam_item", "gender", "age", "finding", "conclusion")},
            "parse_notes": res["parsed"].get("parse_notes", []),
            "issues": issues,
            "issue_count": len(issues),
            "elapsed_ms": res["elapsed_ms"],
        })

    elapsed = round((time.perf_counter() - t0) * 1000, 2)
    return {
        "engine_version": ENGINE_VERSION,
        "total": len(reports),
        "summary": {
            "total": len(reports),
            "reports_with_issues": reports_with_issues,
            "issue_total": issue_total,
            "elapsed_ms": elapsed,
            "by_rule": dict(sorted(by_rule.items())),
            "by_severity": by_severity,
        },
        "results": results,
    }


def rules_public() -> dict:
    """对外暴露的规则视图（剔除预编译私有字段）。"""
    return {k: v for k, v in RULES.items() if not k.startswith("_")}

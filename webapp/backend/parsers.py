"""上传文件解析：支持 .txt/.md/.json/.csv/.tsv/.xlsx，统一产出标准五字段报告列表。

零第三方依赖：xlsx 为 zip+xml 手工解析（兼容 Excel/WPS 导出的 sharedStrings 与内联字符串）。
"""
from __future__ import annotations

import csv
import io
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from . import config
from .engine_adapter import engine, normalize_report

_SSML_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
# 报告分隔行：整行由 - = _ * ~ # 等符号组成（4 个以上）
_DELIM_RE = re.compile(r"^\s*[-=_*~#]{4,}\s*$")


# ============================ JSON ============================

def parse_json_bytes(data: bytes, warnings: list) -> list[dict]:
    text = engine.read_text_any(data)
    try:
        obj = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"JSON 格式错误：{exc}") from exc

    items = []
    if isinstance(obj, list):
        items = obj
    elif isinstance(obj, dict):
        if isinstance(obj.get("reports"), list):
            items = obj["reports"]
        elif isinstance(obj.get("data"), list):
            items = obj["data"]
        else:
            items = [obj]
    else:
        raise ValueError("JSON 顶层必须是对象或数组")

    reports = []
    for i, item in enumerate(items):
        if isinstance(item, str):
            rep = _text_to_report(item)
            if rep:
                reports.append(rep)
            continue
        rep = normalize_report(item)
        if rep is None:
            warnings.append(f"第 {i + 1} 条记录缺少超声所见/超声提示字段，已跳过")
            continue
        rep["report_id"] = str(item.get("report_id") or item.get("id") or f"R{i + 1:03d}")
        reports.append(rep)
    return reports


# ============================ CSV / TSV ============================

def parse_csv_bytes(data: bytes, delimiter: str, warnings: list) -> list[dict]:
    text = engine.read_text_any(data)
    rows = [r for r in csv.reader(io.StringIO(text), delimiter=delimiter) if any(c.strip() for c in r)]
    if not rows:
        raise ValueError("文件内容为空")
    return _rows_to_reports(rows, warnings)


# ============================ XLSX ============================

def _xlsx_shared_strings(z: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall(f"{_SSML_NS}si"):
        out.append("".join(t.text or "" for t in si.iter(f"{_SSML_NS}t")))
    return out


def _xlsx_first_sheet_rows(data: bytes, warnings: list) -> list[list[str]]:
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise ValueError("不是有效的 xlsx 文件（zip 结构损坏）") from exc

    sheet_names = sorted(
        (n for n in z.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)),
        key=lambda n: int(re.search(r"\d+", Path(n).stem).group(0)),
    )
    if not sheet_names:
        raise ValueError("xlsx 中未找到工作表")
    if len(sheet_names) > 1:
        warnings.append(f"文件含 {len(sheet_names)} 个工作表，仅解析第一个")

    shared = _xlsx_shared_strings(z)
    root = ET.fromstring(z.read(sheet_names[0]))
    rows = []
    for row_el in root.iter(f"{_SSML_NS}row"):
        cells: dict[int, str] = {}
        for c in row_el.iter(f"{_SSML_NS}c"):
            ref = c.get("r", "")
            col = 0
            for ch in re.match(r"[A-Z]+", ref).group(0) if re.match(r"[A-Z]+", ref) else "":
                col = col * 26 + (ord(ch) - 64)
            col = max(col - 1, 0)
            ctype = c.get("t", "")
            if ctype == "inlineStr":
                val = "".join(t.text or "" for t in c.iter(f"{_SSML_NS}t"))
            else:
                v = c.find(f"{_SSML_NS}v")
                raw = (v.text or "").strip() if v is not None else ""
                if ctype == "s" and raw.isdigit() and int(raw) < len(shared):
                    val = shared[int(raw)]
                else:
                    val = raw
            if val:
                cells[col] = val.strip()
        if cells:
            width = max(cells) + 1
            rows.append([cells.get(i, "") for i in range(width)])
    return rows


# ============================ TXT / MD ============================

def parse_txt_bytes(data: bytes, warnings: list) -> list[dict]:
    text = engine.read_text_any(data)
    lines = text.splitlines()
    has_delim = any(_DELIM_RE.match(ln) for ln in lines)
    if has_delim:
        blocks, cur = [], []
        for ln in lines:
            if _DELIM_RE.match(ln):
                if cur:
                    blocks.append("\n".join(cur))
                    cur = []
            else:
                cur.append(ln)
        if cur:
            blocks.append("\n".join(cur))
    else:
        blocks = re.split(r"\n{2,}", text)

    reports = []
    for i, block in enumerate(blocks):
        block = block.strip()
        if not block:
            continue
        parsed = engine.parse_report(block)
        if not parsed["finding"] and not parsed["conclusion"]:
            warnings.append(f"第 {i + 1} 段未识别到超声所见/超声提示，已跳过")
            continue
        reports.append({
            "report_id": f"R{len(reports) + 1:03d}",
            "exam_item": parsed.get("exam_item", ""),
            "gender": parsed.get("gender", ""),
            "age": parsed.get("age", ""),
            "finding": parsed["finding"],
            "conclusion": parsed["conclusion"],
        })
    return reports


def _text_to_report(text: str) -> dict | None:
    parsed = engine.parse_report(text)
    if not parsed["finding"] and not parsed["conclusion"]:
        return None
    return {
        "exam_item": parsed.get("exam_item", ""),
        "gender": parsed.get("gender", ""),
        "age": parsed.get("age", ""),
        "finding": parsed["finding"],
        "conclusion": parsed["conclusion"],
    }


# ============================ 表格行 → 报告 ============================

def _rows_to_reports(rows: list[list[str]], warnings: list) -> list[dict]:
    """首行为表头且含已知字段名 → 按列映射；否则整表按纯文本列处理。"""
    from .engine_adapter import engine as _eng  # 复用字段别名表

    header = [h.strip().lower() for h in rows[0]]
    alias_hit = any(a.lower() in header for al in _eng.JSON_FIELD_ALIASES.values() for a in al)

    reports = []
    if alias_hit:
        for i, row in enumerate(rows[1:], start=2):
            obj = {rows[0][j].strip(): (row[j] if j < len(row) else "") for j in range(len(rows[0]))}
            rep = normalize_report(obj)
            if rep is None:
                warnings.append(f"第 {i} 行缺少超声所见/超声提示，已跳过")
                continue
            rid = ""
            for key in ("report_id", "id", "报告id", "序号"):
                for k, v in obj.items():
                    if k.strip().lower() == key and str(v).strip():
                        rid = str(v).strip()
            rep["report_id"] = rid or f"R{len(reports) + 1:03d}"
            reports.append(rep)
    else:
        for i, row in enumerate(rows, start=1):
            text = "\n".join(c for c in row if c.strip())
            rep = _text_to_report(text)
            if rep is None:
                warnings.append(f"第 {i} 行未识别到超声所见/超声提示，已跳过")
                continue
            rep["report_id"] = f"R{len(reports) + 1:03d}"
            reports.append(rep)
    return reports


# ============================ 统一入口 ============================

def parse_upload(filename: str, data: bytes) -> dict:
    """按扩展名分派解析器，返回 {format, reports, warnings}。"""
    if len(data) > config.MAX_UPLOAD_BYTES:
        raise ValueError(f"文件超过大小限制（{config.MAX_UPLOAD_BYTES // (1024 * 1024)}MB）")
    ext = Path(filename or "").suffix.lower()
    if ext not in config.SUPPORTED_EXTS:
        raise ValueError(f"不支持的文件类型“{ext or '(无扩展名)'}”，支持：{sorted(config.SUPPORTED_EXTS)}")

    warnings: list[str] = []
    if ext == ".json":
        reports = parse_json_bytes(data, warnings)
        fmt = "json"
    elif ext in (".csv", ".tsv"):
        reports = parse_csv_bytes(data, "\t" if ext == ".tsv" else ",", warnings)
        fmt = ext.lstrip(".")
    elif ext == ".xlsx":
        rows = _xlsx_first_sheet_rows(data, warnings)
        if not rows:
            raise ValueError("xlsx 工作表为空")
        reports = _rows_to_reports(rows, warnings)
        fmt = "xlsx"
    else:
        reports = parse_txt_bytes(data, warnings)
        fmt = "txt"

    if not reports:
        raise ValueError("未能从文件中解析出任何报告，请检查文件内容与格式说明")
    return {"format": fmt, "reports": reports, "warnings": warnings}

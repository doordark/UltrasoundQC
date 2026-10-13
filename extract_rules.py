#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
医学质控规则离线提炼与导出脚本
================================
从大模型批量质控产出的 Excel 错误明细表中，提炼高频错误模式、医学敏感词、
错别字映射与逻辑规则，导出为轻量级 rules.json，供本地离线质控引擎读取。

用法:
    python extract_rules.py
    python extract_rules.py --excel ultrasound_qc_report.xlsx --output rules.json
"""

import argparse
import difflib
import json
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd

# ============================ 种子词典（基础库，可被动态扩展） ============================

# 阳性病变敏感词种子表
SEED_POSITIVE_KEYWORDS = [
    "结节", "囊肿", "占位", "积液", "强回声", "低回声", "无回声", "混合回声",
    "欠均匀", "肿大", "暗区", "条索状", "结石", "钙化", "息肉", "肿块",
    "包块", "肿瘤", "增厚", "增大", "扩张", "脂肪肝", "血流信号", "边界不清",
    "斑块", "团块", "肌瘤",
]

# 方位/左右侧对应词表（固定）
SIDE_KEYWORDS = ["左侧", "右侧", "左", "右", "双侧"]

# 高频错别字/不规范用词映射种子表
SEED_CORRECTION_MAPPING = {
    "回音": "回声",
    "左侧乳房": "左乳",
}

# 阴性结论表述种子表（与阳性发现矛盾判定用）
SEED_NEGATIVE_CONCLUSION_PHRASES = [
    "未见明显异常", "结构未见明显异常", "未见异常", "无明显异常",
    "未见明确异常", "双侧未见异常", "未见占位",
]

# ---- 挖掘用辅助配置 ----
# 回声类描述的修饰字（用于动态发现 "稍强回声"、"混合性回声" 等新词）
ECHO_MODIFIER_CHARS = set("低强无等稍弱混合成杂粗细点片网实密性液")
ECHO_MODIFIER_BLACKLIST = {"正常", "未见", "周边", "内部", "实质", "均匀", "欠均", "整体", "回声"}

# 动态发现回声/暗区变体的正则
RE_ECHO = re.compile(r"([\u4e00-\u9fa5]{1,3}回声)")
RE_DARK = re.compile(r"([\u4e00-\u9fa5]{0,2}暗区)")

# 错别字/用词类错误的分类关键词
TYPO_CATEGORY_RE = re.compile(r"错别?字|错字|用词|书写|不规范|笔误|表述|规范|简写")

# 逻辑/结论矛盾类错误的分类关键词
LOGIC_CATEGORY_RE = re.compile(r"逻辑|矛盾|不一致|前后|相符|对应")

# 从修改建议/错误原因中提取纠错对 ("A"应为"B" / A→B / 将A改为B / A误写为B / A系B之误)
RE_ARROW = re.compile(
    r"([\u4e00-\u9fa5A-Za-z0-9]{1,8})\s*(?:→|->|—>|－>)\s*([\u4e00-\u9fa5A-Za-z0-9]{1,8})")
RE_SHOULD = re.compile(
    r"[“\"'『「〔【\[]?([\u4e00-\u9fa5A-Za-z0-9]{1,8})[”\"'』」〕】\]]?应为[“\"'『「〔【\[]?"
    r"([\u4e00-\u9fa5A-Za-z0-9]{1,8})[”\"'』」〕】\]]?")
RE_CHANGE = re.compile(
    r"(?:将|把)?[“\"'『「〔【\[]?([\u4e00-\u9fa5A-Za-z0-9]{1,8})[”\"'』」〕】\]]?"
    r"(?:改为|改成|更正为|修正为|规范为|写成)[“\"'『「〔【\[]?([\u4e00-\u9fa5A-Za-z0-9]{1,8})[”\"'』」〕】\]]?")
RE_MISTYPE = re.compile(
    r"([\u4e00-\u9fa5A-Za-z0-9]{1,8})[”\"'』」〕】\]]?(?:误写为|误写成了|误用为)"
    r"[“\"'『「〔【\[]?([\u4e00-\u9fa5A-Za-z0-9]{1,8})")
RE_MISCALL = re.compile(
    r"([\u4e00-\u9fa5A-Za-z0-9]{1,8})[”\"'』」〕】\]]?系[“\"'『「〔【\[]?"
    r"([\u4e00-\u9fa5A-Za-z0-9]{1,8})[”\"'』」〕】\]]?之误")

# 阴性结论短语挖掘
RE_NEG_PHRASE = re.compile(r"(?:未见|无明显|无)[\u4e00-\u9fa5]{0,6}异常")

# 纠错对过滤：元话语词（建议/描述等指令性文字不属于词形纠错对）
META_PHRASE_STOPWORDS = {"建议", "考虑", "必要时", "根据", "结合", "临床", "描述", "结论",
                         "报告", "检查", "超声", "提示", "规范", "用语", "未见", "分期", "分级"}
META_CHAR_STOPWORDS = {"如", "若", "并", "或", "的", "了", "应", "需", "与", "及", "见", "检", "上", "中", "下"}


def _clean_token(token: str) -> str:
    """去掉匹配片段两端的引号、括号、标点等噪声字符。"""
    return token.strip(" \t\r\n“”\"'‘’「」『』〔〕【】()（）<>《》.,，。:：;；!！?？、·-—")


def _is_valid_pair(wrong: str, correct: str) -> bool:
    """校验纠错对是否为可信的词形错误：长度、元话语词、公共字三重过滤。"""
    if not wrong or not correct or wrong == correct:
        return False
    if not (1 <= len(wrong) <= 6 and 1 <= len(correct) <= 6):
        return False
    for token in (wrong, correct):
        if any(sw in token for sw in META_PHRASE_STOPWORDS) or token in META_CHAR_STOPWORDS:
            return False
    # 两个多字词须至少共享一个汉字（错别字/词序错误的普遍特征），过滤跨句差分噪声
    if len(wrong) >= 2 and len(correct) >= 2 and not set(wrong) & set(correct):
        return False
    # 单字对只保留同部首/形近难判定，直接要求同字长且非常用虚词（上面已过滤）
    if len(wrong) == 1 or len(correct) == 1:
        return False
    return True

# 已知规则类型 → 分类关键词映射（未命中的错误分类进入 extra_rules 动态扩展）
KNOWN_CATEGORY_MAP = {
    "逻辑矛盾/前后不一致": LOGIC_CATEGORY_RE,
    "左右侧不匹配": re.compile(r"左|右|侧"),
    "错别字/不规范用词": TYPO_CATEGORY_RE,
    "阳性病变漏提示": re.compile(r"阳性|漏|未提示|未描述|遗漏"),
}


# ============================ 数据读取 ============================

def load_error_rows(excel_path: str, sheet_name: str) -> tuple[pd.DataFrame, int]:
    """读取 Excel 错误明细表，返回错误行子集与总行数。"""
    df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    required = ["是否存在错误", "错误分类", "原文语句", "错误原因分析", "修改建议"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise SystemExit(f"[错误] 表 {sheet_name} 缺少必需列: {missing}，实际列: {list(df.columns)}")

    for col in required:
        df[col] = df[col].fillna("").astype(str).str.strip()

    err_df = df[df["是否存在错误"] == "是"].copy()
    return err_df, len(df)


# ============================ 敏感词提取 ============================

def extract_positive_keywords(err_df: pd.DataFrame, min_count: int) -> list[dict]:
    """从错误文本中统计阳性病变敏感词频次，并动态发现回声/暗区类新变体。"""
    corpus = "\n".join(err_df["原文语句"] + "\n" + err_df["错误原因分析"])

    counts: Counter = Counter()
    # 1) 种子词频次统计
    for word in SEED_POSITIVE_KEYWORDS:
        n = corpus.count(word)
        if n > 0:
            counts[word] = n

    # 2) 动态发现回声类变体（如 稍强回声、混合性回声）与暗区类变体（如 液性暗区）
    for match in RE_ECHO.findall(corpus):
        modifier = match[:-2]
        if modifier and all(ch in ECHO_MODIFIER_CHARS for ch in modifier) \
                and modifier not in ECHO_MODIFIER_BLACKLIST:
            counts[match] += 1
    for match in RE_DARK.findall(corpus):
        if "未见" not in match:
            counts[match] += 1

    # 3) 过滤低频词，种子词不受 min_count 限制（保持基础词表稳定）
    result = []
    for word, n in sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])):
        is_seed = word in SEED_POSITIVE_KEYWORDS
        if n >= min_count or is_seed:
            result.append({"word": word, "count": n, "source": "seed" if is_seed else "excel"})
    return result


# ============================ 纠错映射提取 ============================

def _pairs_from_patterns(text: str) -> list[tuple[str, str]]:
    """用多种正则从建议/原因文本中抽取 (误写, 正确) 纠错对。"""
    pairs = []
    for regex, wrong_idx, correct_idx in (
        (RE_ARROW, 1, 2), (RE_SHOULD, 1, 2), (RE_CHANGE, 1, 2),
        (RE_MISTYPE, 1, 2), (RE_MISCALL, 1, 2),
    ):
        for m in regex.finditer(text):
            wrong, correct = _clean_token(m.group(wrong_idx)), _clean_token(m.group(correct_idx))
            if wrong and correct:
                pairs.append((wrong, correct))
    return pairs


def _pairs_from_diff(original: str, suggestion: str) -> list[tuple[str, str]]:
    """当修改建议是整句改写时，用序列对齐提取被替换的词语对。"""
    if not (0.5 <= len(suggestion) / max(len(original), 1) <= 1.5):
        return []
    pairs = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, original, suggestion).get_opcodes():
        if tag != "replace":
            continue
        wrong, correct = original[i1:i2], suggestion[j1:j2]
        if not (1 <= len(wrong) <= 6 and 1 <= len(correct) <= 6):
            continue
        if not all("\u4e00" <= ch <= "\u9fa5" for ch in wrong + correct):
            continue
        pairs.append((wrong, correct))
    return pairs


def extract_correction_pairs(err_df: pd.DataFrame) -> list[dict]:
    """提取高频错别字/不规范用词映射：种子表 + 正则解析 + 序列对齐，均经原文校验。"""
    orig_corpus = "\n".join(err_df["原文语句"])
    pair_counter: Counter = Counter()

    for _, row in err_df.iterrows():
        category, original = row["错误分类"], row["原文语句"]
        suggestion, reason = row["修改建议"], row["错误原因分析"]

        candidates: list[tuple[str, str]] = []
        # 1) 分类属于错别字/用词类时，解析建议文本中的显式纠错表述
        if TYPO_CATEGORY_RE.search(category):
            candidates += _pairs_from_patterns(f"{suggestion}\n{reason}")
        # 2) 建议近似整句改写时，做差分对齐
        if TYPO_CATEGORY_RE.search(category) and original and suggestion:
            candidates += _pairs_from_diff(original, suggestion)

        # 3) 三重校验：词形合法性 + 误写词确实出现在原文语句中
        for wrong, correct in candidates:
            if _is_valid_pair(wrong, correct) and wrong in orig_corpus:
                pair_counter[(wrong, correct)] += 1

    # 合并种子映射（种子词不再重复计数）
    details, mapping = [], {}
    for (wrong, correct), n in sorted(pair_counter.items(), key=lambda kv: (-kv[1], kv[0])):
        mapping[wrong] = correct
        details.append({"wrong": wrong, "correct": correct, "count": n, "source": "excel"})
    for wrong, correct in SEED_CORRECTION_MAPPING.items():
        if wrong not in mapping:
            mapping[wrong] = correct
            details.append({"wrong": wrong, "correct": correct, "count": 0, "source": "seed"})
    return details


# ============================ 阴性结论短语与逻辑规则 ============================

def extract_negative_phrases(err_df: pd.DataFrame) -> list[str]:
    """从逻辑/结论矛盾类错误文本中挖掘阴性结论表述，并与种子表合并。"""
    text = err_df["原文语句"] + "\n" + err_df["错误原因分析"]
    mask = err_df["错误分类"].apply(lambda c: bool(LOGIC_CATEGORY_RE.search(c)))
    corpus = "\n".join(text[mask].astype(str))
    counts = Counter(RE_NEG_PHRASE.findall(corpus))
    merged = {p: counts[p] for p in counts}
    for phrase in SEED_NEGATIVE_CONCLUSION_PHRASES:
        merged.setdefault(phrase, merged.get(phrase, 0))
    return [p for p, _ in sorted(merged.items(), key=lambda kv: (-kv[1], kv[0]))]


def build_logical_rules(negative_phrases: list[str]) -> list[dict]:
    """导出供离线引擎执行的结构化逻辑规则模板。"""
    return [
        {
            "id": "LR001_finding_vs_conclusion",
            "name": "阳性发现与阴性结论矛盾",
            "type": "cross_section_contradiction",
            "description": "所见/描述部分包含阳性病变敏感词，但结论部分出现阴性表述",
            "applies_to": {"source": "finding", "target": "conclusion"},
            "condition": {"source_any": "@positive_keywords", "target_any": "@negative_conclusion_phrases"},
            "negative_conclusion_phrases": negative_phrases,
            "severity": "high",
        },
        {
            "id": "LR002_side_mismatch",
            "name": "左右侧不匹配",
            "type": "cross_section_contradiction",
            "description": "所见与结论之间的左右侧表述冲突（如左侧病变结论写成右侧）",
            "applies_to": {"source": "finding", "target": "conclusion"},
            "condition": {"side_conflict": "@side_keywords"},
            "severity": "high",
        },
        {
            "id": "LR003_nonstandard_wording",
            "name": "错别字/不规范用词",
            "type": "wording_check",
            "description": "全文命中纠错映射中的误写词，应按映射替换为规范用词",
            "applies_to": {"source": "full_text"},
            "condition": {"correction_mapping": "@correction_mapping"},
            "severity": "medium",
        },
    ]


# ============================ 动态扩展规则类型 ============================

def build_extra_rules(err_df: pd.DataFrame) -> list[dict]:
    """发现 Excel 中出现的新错误类型（未被已知规则覆盖），作为动态扩展规则追加。"""
    counter = Counter(cat.strip() for c in err_df["错误分类"] for cat in c.split("、") if cat.strip())
    extra = []
    for category, count in sorted(counter.items(), key=lambda kv: (-kv[1], kv[0])):
        if any(re_.search(category) for re_ in KNOWN_CATEGORY_MAP.values()):
            continue
        examples = err_df[err_df["错误分类"].str.contains(re.escape(category), regex=True)]
        extra.append({
            "category": category,
            "count": count,
            "examples": [
                {"original_text": r["原文语句"], "suggestion": r["修改建议"]}
                for _, r in examples.head(3).iterrows()
            ],
        })
    return extra


# ============================ 与已有规则库合并 ============================

def merge_existing(rules: dict, out_path: Path) -> dict:
    """若 rules.json 已存在，则加载旧库进行增量合并（旧词保留、新词追加、新值优先）。"""
    if not out_path.exists():
        return rules
    try:
        old = json.loads(out_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        print("[警告] 旧 rules.json 解析失败，将直接覆盖")
        return rules

    # 词表类：按 word 合并去重，count 以本次统计为准，旧独有词保留
    for key in ("positive_keywords", "side_keywords"):
        old_map = {e["word"]: e.get("count", 0) for e in old.get(key, []) if isinstance(e, dict) and "word" in e}
        merged = {**old_map, **{e["word"]: e["count"] for e in rules[key]}}
        rules[key] = [{"word": w, "count": c} for w, c in sorted(merged.items(), key=lambda kv: -kv[1])]

    # 阴性短语：并集保序
    rules["negative_conclusion_phrases"] = list(dict.fromkeys(
        [*old.get("negative_conclusion_phrases", []), *rules["negative_conclusion_phrases"]]))

    # 纠错映射：旧为底、新覆盖；明细按 (wrong, correct) 去重
    rules["correction_mapping"] = {**old.get("correction_mapping", {}), **rules["correction_mapping"]}
    detail_map = {(d.get("wrong"), d.get("correct")): d for d in old.get("correction_details", [])}
    for d in rules["correction_details"]:
        detail_map[(d["wrong"], d["correct"])] = d
    rules["correction_details"] = sorted(detail_map.values(), key=lambda d: -d.get("count", 0))

    # 逻辑规则 / 扩展规则：按 id / category 合并，新值优先
    for key, id_field in (("logical_rules", "id"), ("extra_rules", "category")):
        merged = {r.get(id_field): r for r in old.get(key, []) if isinstance(r, dict) and r.get(id_field)}
        for r in rules[key]:
            merged[r.get(id_field)] = r
        rules[key] = list(merged.values())
    return rules


# ============================ 主流程 ============================

def build_rules(excel_path: str, sheet_name: str, min_count: int) -> dict:
    err_df, total_rows = load_error_rows(excel_path, sheet_name)
    if err_df.empty:
        raise SystemExit(f"[错误] 表 {sheet_name} 中没有'是否存在错误=是'的记录")

    positive_keywords = extract_positive_keywords(err_df, min_count)
    side_counts = sorted(
        ((w, ("\n".join(err_df["原文语句"])).count(w)) for w in SIDE_KEYWORDS),
        key=lambda kv: -kv[1])
    correction_details = extract_correction_pairs(err_df)
    negative_phrases = extract_negative_phrases(err_df)
    logical_rules = build_logical_rules(negative_phrases)
    extra_rules = build_extra_rules(err_df)

    type_counter = Counter(cat.strip() for c in err_df["错误分类"] for cat in c.split("、") if cat.strip())

    return {
        "meta": {
            "version": "1.0",
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source_file": Path(excel_path).name,
            "source_sheet": sheet_name,
            "total_rows": total_rows,
            "error_rows": int(len(err_df)),
            "description": "超声报告离线质控规则库：由大模型批量质控结果提炼，供本地无大模型依赖的质控引擎使用",
        },
        "positive_keywords": positive_keywords,
        "side_keywords": [{"word": w, "count": c} for w, c in side_counts],
        "correction_mapping": {d["wrong"]: d["correct"] for d in correction_details},
        "correction_details": correction_details,
        "negative_conclusion_phrases": negative_phrases,
        "logical_rules": logical_rules,
        "error_type_stats": [{"category": k, "count": v} for k, v in sorted(type_counter.items(), key=lambda kv: -kv[1])],
        "extra_rules": extra_rules,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="医学质控规则离线提炼与导出")
    parser.add_argument("--excel", default="ultrasound_qc_report.xlsx", help="质控结果 Excel 文件")
    parser.add_argument("--sheet", default="错误明细表", help="错误明细表 Sheet 名称")
    parser.add_argument("--output", default="rules.json", help="输出 JSON 路径")
    parser.add_argument("--min-count", type=int, default=1, help="动态新词的最低出现次数阈值")
    args = parser.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if not Path(args.excel).exists():
        raise SystemExit(f"[错误] 找不到 Excel 文件: {args.excel}")

    rules = build_rules(args.excel, args.sheet, args.min_count)
    out_path = Path(args.output)
    rules = merge_existing(rules, out_path)

    out_path.write_text(json.dumps(rules, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    m = rules["meta"]
    print("=" * 56)
    print("医学质控规则提炼完成")
    print("=" * 56)
    print(f"数据来源      : {m['source_file']} / {m['source_sheet']}")
    print(f"错误明细行数  : {m['error_rows']} / {m['total_rows']}")
    print(f"阳性敏感词    : {len(rules['positive_keywords'])} 条")
    print(f"方位/侧别词   : {len(rules['side_keywords'])} 条")
    print(f"纠错映射      : {len(rules['correction_mapping'])} 条 "
          f"(其中 Excel 新增 {sum(1 for d in rules['correction_details'] if d.get('source') == 'excel')} 条)")
    print(f"阴性结论短语  : {len(rules['negative_conclusion_phrases'])} 条")
    print(f"逻辑规则      : {len(rules['logical_rules'])} 条")
    print(f"动态扩展规则  : {len(rules['extra_rules'])} 类")
    print(f"输出文件      : {out_path.resolve()}")


if __name__ == "__main__":
    main()

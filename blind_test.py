#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
blind_test.py — 本地引擎 vs 大模型 盲测对比工具
================================================
流程:
    python blind_test.py prepare                 # 从 init_report 确定性抽样 100 份 → blind_sample/
    dist\\UltrasoundQCEngine.exe check blind_sample -o blind_engine_result.xlsx
    python blind_test.py compare                 # 对比大模型结果 → 控制台报告 + blind_compare_report.xlsx

对比口径:
    - 报告级: 以"该份报告是否被判定有问题"为口径，计算混淆矩阵与查准/查全
    - 类别映射: 错别字→C  左右→B  分级→D  逻辑/不一致→A/B/D(语义类)  其余→范围外
    - FN 明细: 大模型报错而引擎漏报 → 逐条给出错误分类与原文，供补充 rules.json
    - FP 明细: 引擎报错而大模型认为无错 → 逐条给出引擎命中片段，供精细化正则
"""

import argparse
import random
import re
import shutil
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

from qc_engine import MiniExcel  # 复用免依赖 Excel 写入器

HERE = Path(__file__).resolve().parent
SAMPLE_DIR = HERE / "blind_sample"
SAMPLE_N = 100
SEED = 2026  # 固定随机种子，保证每次抽样/复测同一批报告

# 大模型错误分类 → 引擎规则字母
def llm_bucket(cat: str) -> str:
    if re.search(r"错别?字|用词|书写|笔误", cat):
        return "C"
    if re.search(r"左|右", cat) and re.search(r"侧|匹配|不符|错", cat):
        return "B"
    if re.search(r"分级|BI-?RADS", cat, re.IGNORECASE):
        return "D"
    if re.search(r"逻辑|一致|矛盾|结论|前后|相符|遗漏", cat):
        return "A/B/D"
    return "范围外"


def _norm_name(name: str) -> str:
    return Path(str(name).strip()).stem  # 去掉 .txt 等后缀并统一


def load_llm(llm_xlsx: str) -> dict:
    """返回 {文件名: [错误行dict]}，仅含大模型判定有错的行。"""
    df = pd.read_excel(llm_xlsx, sheet_name="错误明细表", dtype=str).fillna("")
    files = defaultdict(list)
    for _, r in df.iterrows():
        if str(r["是否存在错误"]).strip() == "是":
            files[_norm_name(r["文件名/报告ID"])].append({
                "category": str(r["错误分类"]).strip(),
                "original": str(r["原文语句"]).strip(),
                "suggestion": str(r["修改建议"]).strip(),
                "reason": str(r["错误原因分析"]).strip(),
            })
    return files


def load_engine(engine_xlsx: str) -> dict:
    """返回 {文件名: [问题行dict]}，仅含引擎判定有问题的行。"""
    df = pd.read_excel(engine_xlsx, sheet_name="错误明细", dtype=str).fillna("")
    files = defaultdict(list)
    for _, r in df.iterrows():
        if str(r["是否存在问题"]).strip() == "是":
            files[_norm_name(r["文件名/报告ID"])].append({
                "rule": str(r["规则"]).strip(),
                "fragment": str(r["原文片段"]).strip(),
                "message": str(r["问题描述"]).strip(),
                "suggestion": str(r["修改建议"]).strip(),
            })
    return files


# ============================ 抽样 ============================

def cmd_prepare(_args) -> None:
    src = HERE / "init_report"
    all_files = sorted(src.glob("*.txt"))
    if len(all_files) < SAMPLE_N:
        raise SystemExit(f"[错误] init_report 仅 {len(all_files)} 份报告")
    if SAMPLE_DIR.exists() and len(list(SAMPLE_DIR.glob("*.txt"))) == SAMPLE_N:
        print(f"[prepare] 已存在同批样本，跳过（{SAMPLE_DIR}，种子={SEED}）")
        return
    if SAMPLE_DIR.exists():
        shutil.rmtree(SAMPLE_DIR)
    SAMPLE_DIR.mkdir()
    picked = random.Random(SEED).sample(all_files, SAMPLE_N)
    for f in picked:
        shutil.copy2(f, SAMPLE_DIR / f.name)
    print(f"[prepare] 已抽样 {SAMPLE_N} 份 → {SAMPLE_DIR}（random.seed({SEED})，可复现）")


# ============================ 对比 ============================

def cmd_compare(args) -> None:
    llm = load_llm(args.llm)
    eng = load_engine(args.engine)
    sample_files = {_norm_name(f.name) for f in SAMPLE_DIR.glob("*.txt")}
    if len(sample_files) != SAMPLE_N:
        raise SystemExit(f"[错误] {SAMPLE_DIR} 样本数为 {len(sample_files)}，请先运行 prepare")

    llm_bad = {f for f in sample_files if f in llm}
    eng_bad = {f for f in sample_files if f in eng}
    tp, fn, fp = llm_bad & eng_bad, llm_bad - eng_bad, eng_bad - llm_bad
    tn = sample_files - llm_bad - eng_bad

    precision = len(tp) / (len(tp) + len(fp)) if tp or fp else 0.0
    recall = len(tp) / (len(tp) + len(fn)) if tp or fn else 0.0
    f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0

    print("=" * 60)
    print(f"盲测对比（样本 {SAMPLE_N} 份，随机种子 {SEED}）")
    print("=" * 60)
    print(f"大模型判定有错: {len(llm_bad)} 份 | 引擎判定有问题: {len(eng_bad)} 份")
    print(f"双方一致(TP)  : {len(tp)} 份")
    print(f"引擎漏报(FN)  : {len(fn)} 份  ← 大模型发现、引擎未报")
    print(f"引擎误报(FP)  : {len(fp)} 份  ← 引擎报了、大模型认为无错")
    print(f"双方一致(TN)  : {len(tn)} 份")
    print(f"报告级 查准率 P = {precision:.1%} | 查全率 R = {recall:.1%} | F1 = {f1:.1%}")

    fn_rows, fp_rows, fn_cat_stat = [], [], Counter()
    for f in sorted(fn):
        for e in llm[f]:
            bucket = llm_bucket(e["category"])
            fn_cat_stat[bucket] += 1
            fn_rows.append([f, e["category"], bucket, e["original"][:80], e["suggestion"][:80]])
    for f in sorted(fp):
        for e in eng[f]:
            fp_rows.append([f, e["rule"], e["fragment"][:60], e["message"][:80]])

    print("\n--- FN 错误分类分布（引擎漏报的大模型错误）---")
    for cat, n in fn_cat_stat.most_common():
        print(f"  {cat}: {n} 条")
    print("\n--- FN 明细（前 12 条）---")
    for r in fn_rows[:12]:
        print(f"  [{r[1]}|{r[2]}] {r[3]}")
        print(f"      建议: {r[4]}")
    print("\n--- FP 明细（引擎误报，前 12 条）---")
    for r in fp_rows[:12]:
        print(f"  {r[0][:20]} | {r[1]} | {r[3]}")
        print(f"      片段: {r[2]}")

    MiniExcel.write(str(HERE / "blind_compare_report.xlsx"), [
        ("对比总览", ["指标", "数值"], [
            ["样本数", SAMPLE_N], ["随机种子", SEED],
            ["大模型有错", len(llm_bad)], ["引擎有问题", len(eng_bad)],
            ["TP(一致)", len(tp)], ["FN(漏报)", len(fn)], ["FP(误报)", len(fp)], ["TN(一致)", len(tn)],
            ["查准率P", f"{precision:.1%}"], ["查全率R", f"{recall:.1%}"], ["F1", f"{f1:.1%}"],
        ]),
        ("FN漏报明细", ["文件", "大模型分类", "对应引擎规则", "原文语句", "修改建议"], fn_rows),
        ("FP误报明细", ["文件", "引擎规则", "命中片段", "问题描述"], fp_rows),
    ])
    print(f"\n[输出] 明细已导出: {HERE / 'blind_compare_report.xlsx'}")
    print("[说明] FN 中『C/错别字』条目可直接向 rules.json 补纠错对；『范围外』为引擎设计边界(数据单位/模板残留等)")


def main() -> None:
    ap = argparse.ArgumentParser(description="本地引擎 vs 大模型 盲测对比")
    sub = ap.add_subparsers(dest="mode", required=True)
    sub.add_parser("prepare", help=f"抽样 {SAMPLE_N} 份报告到 blind_sample/")
    p_cmp = sub.add_parser("compare", help="对比引擎结果与大模型结果")
    p_cmp.add_argument("--engine", default="blind_engine_result.xlsx", help="引擎输出 Excel")
    p_cmp.add_argument("--llm", default="ultrasound_qc_report.xlsx", help="大模型质控 Excel")
    args = ap.parse_args()
    if args.mode == "prepare":
        cmd_prepare(args)
    else:
        cmd_compare(args)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ultrasound QC Engine — 离线超声报告自动勘误与质控引擎
======================================================
纯 Python 标准库实现（零第三方依赖 / 零外网 / 零 GPU），专为医院老旧
Windows 环境（Win7/Win10, 32/64 位）设计，可由 PyInstaller 打包为单文件 exe。

三种用法:
    qc_engine.exe selftest                         # 内置自检（验证全部规则）
    qc_engine.exe check <文件或目录...> -o 结果.xlsx  # 命令行批量检测，导出 Excel
    qc_engine.exe serve --host 0.0.0.0 --port 8976  # HTTP API 常驻服务（工作站实时对接）

输入格式（工作站集成规范）:
    标准 JSON 对象，核心字段:
        {
          "exam_item": "检查项目",   "gender": "性别",   "age": "年龄",
          "finding":  "超声所见",    "conclusion": "超声提示"
        }
    亦支持中文键名（检查项目/性别/年龄/超声所见/超声提示）；命令行 .txt 批量
    检测仍走正则自适应文本解析（兼容旧版 PACS 导出）。

质控规则:
    A 结论遗漏病变: 所见含阳性病变词，结论却为纯阴性表述（带否定词识别，防误报）
    B 左右侧不符  : 所见与结论中同一器官的左右侧表述冲突（器官级配对 + 全局兜底）
    C 用词勘误    : 命中纠错映射表（如 回音->回声），给出替换建议
    D 分级冲突    : BI-RADS/TI-RADS 分级与所见恶性征象不匹配（如微钙化却给 2 类）
    E 结论未涵盖所见: 所见病灶/征象在结论中无任何呼应（产前报告豁免）
    F 完整性与书写规范: 检查项目缺失 / 建议缺时间数值 / 数值粘连 / 项目与内容不符
    G 性别与检查项目冲突: 患者性别与检查项目/所见部位矛盾（男见子宫卵巢、女见前列腺）
    H 所见遗漏检查部位: 检查项目要求多个部位（如肝胆胰脾），所见缺少部分器官描述
    I 数据与单位异常  : 测量值单位错误/缺失、未删测量值占位符、径线量纲异常（疑似小数点错）
    J 语句重复        : 叠词重复（如"大小大小"）/ 整句模板重复冗余未删除
"""

from __future__ import annotations

import argparse
import difflib
import json
import os
import re
import sys
import time
import zipfile
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.sax.saxutils import escape

ENGINE_VERSION = "1.2.3"
MAX_BODY_BYTES = 8 * 1024 * 1024  # HTTP 请求体上限 8MB

# ============================ 质控规则（9 类错误类型，名称不可改动） ============================
# 统一口径：9 类错误类型即 9 条规则本体，rule_id/rule/error_type 三者同名；
# 原 A~J 字母仅是引擎内部检测子项的实现标记（check 字段），不再作为对外分类。
ERR_DESC_MISMATCH = "描述与结论不一致"
ERR_GRADING = "分级不规范"
ERR_WORDING = "语病或表述不规范"
ERR_DATA_UNIT = "数据或单位错误"
ERR_TYPO_FIELD = "错别字或字段缺失"
ERR_TEMPLATE_LEFT = "模板文字未删除"
ERR_MISSING_PART = "遗漏检查部位"
ERR_OTHER = "其他"
ERR_GENDER_AGE = "性别年龄冲突"

# 9 类错误类型（展示顺序与质控 Excel 统计表一致）
ERROR_TYPE_ORDER = [
    ERR_DESC_MISMATCH, ERR_GRADING, ERR_WORDING, ERR_DATA_UNIT, ERR_TYPO_FIELD,
    ERR_TEMPLATE_LEFT, ERR_MISSING_PART, ERR_OTHER, ERR_GENDER_AGE,
]
# 引擎内部检测子项 → 错误类型（=规则）映射；"check" 标记见各规则实现
CHECK_ERROR_TYPE = {
    "A": ERR_DESC_MISMATCH, "B": ERR_DESC_MISMATCH, "E": ERR_DESC_MISMATCH,
    "C": ERR_TYPO_FIELD, "C2": ERR_TYPO_FIELD,
    "D": ERR_GRADING, "D1": ERR_GRADING, "D2": ERR_GRADING, "D3": ERR_GRADING, "D4": ERR_GRADING,
    "F2": ERR_WORDING, "F3": ERR_DATA_UNIT, "F4": ERR_OTHER,
    "F5": ERR_WORDING, "F6": ERR_WORDING, "F7": ERR_WORDING,
    "G": ERR_GENDER_AGE, "G2": ERR_GENDER_AGE,
    "H": ERR_MISSING_PART,
    "I": ERR_DATA_UNIT, "I1": ERR_DATA_UNIT, "I2": ERR_DATA_UNIT, "I3": ERR_DATA_UNIT,
    "I4": ERR_DATA_UNIT, "I5": ERR_DATA_UNIT, "I6": ERR_DATA_UNIT,
    "J1": ERR_WORDING, "J2": ERR_TEMPLATE_LEFT, "J3": ERR_TEMPLATE_LEFT,
}

# ============================ 内置默认规则（rules.json 缺失时兜底） ============================

DEFAULT_RULES = {
    "positive_keywords": [
        "结节", "囊肿", "占位", "积液", "强回声", "低回声", "无回声", "混合回声",
        "欠均匀", "肿大", "暗区", "条索状", "结石", "钙化", "息肉", "肿块",
        "包块", "肿瘤", "增厚", "增大", "扩张", "脂肪肝", "血流信号", "边界不清",
        "斑块", "团块", "肌瘤",
    ],
    "side_keywords": ["左侧", "右侧", "左", "右", "双侧"],
    "correction_mapping": {
        "回音": "回声",
        "左侧乳房": "左乳",
    },
    "negative_conclusion_phrases": [
        "未见明显异常", "结构未见明显异常", "未见异常", "无明显异常",
        "未见明确异常", "双侧未见异常", "未见占位", "无异常",
    ],
    "grading_rules": {
        "context_keywords": ["乳腺", "乳房", "副乳", "甲状腺"],
        "min_grade_features": {
            "微钙化": 4, "钙化": 4, "边界不清": 4, "边界欠清": 4,
            "形态不规则": 4, "纵横比失调": 4, "毛刺": 5, "蟹足样": 5,
            "血流丰富": 4, "血流较丰富": 4, "簇状钙化": 4,
        },
    },
    # 规则 G：性别与检查项目/所见部位冲突（词表可配置，医院现场可增删）
    # 男 + female_organs / 女 + male_organs → 冲突；乳腺属可配置项（男性乳腺检查存在但少见）
    "gender_conflict": {
        "male_organs": ["前列腺", "精囊", "睾丸", "附睾", "阴囊", "阴茎", "输精管"],
        "female_organs": ["子宫", "卵巢", "附件", "输卵管", "宫颈", "阴道", "子宫内膜",
                          "胎盘", "羊水", "孕囊", "胎儿", "乳腺", "乳房", "副乳", "前庭大腺"],
        "exempt_keywords": ["男性乳腺", "男乳", "男性乳房", "两性畸形", "假两性", "变性"],
    },
    # 规则 G2：年龄与检查内容逻辑冲突（如 79 岁女性产科胎儿检查，几乎不可能自然妊娠）
    "age_rules": {
        "pregnancy_max_age": 55,          # 超过该年龄出现妊娠相关描述即报年龄逻辑冲突
        "pregnancy_keywords": ["胎儿", "孕囊", "羊水", "胎盘", "脐动脉", "脐带", "宫内孕",
                               "早孕", "中孕", "晚孕", "单胎", "双胎", "胎心", "胎位",
                               "胎头", "胎腹", "股骨长", "双顶径", "头围", "腹围", "胎儿颈"],
    },
    # 规则 H：检查项目 → 应查部位映射（organs 中任一器官在所见中无描述即报漏项；
    # 仅当部分器官已描述而其余缺失时触发，单部位项目/整段未描述不误报）
    "exam_organ_map": {
        "organ_aliases": {
            "肝": ["肝", "肝脏"], "胆": ["胆", "胆囊", "胆管"], "胰": ["胰", "胰腺"],
            "脾": ["脾", "脾脏"], "肾": ["肾", "肾脏"], "输尿管": ["输尿管"],
            "膀胱": ["膀胱"], "甲状腺": ["甲状腺"], "乳腺": ["乳腺", "乳房", "乳"],
            "子宫": ["子宫"], "卵巢": ["卵巢"], "附件": ["附件", "输卵管"],
            "前列腺": ["前列腺"], "心脏": ["心脏", "心房", "心室", "心肌", "左室", "右室", "左房", "右房"],
        },
        "exam_items": [
            {"exam": ["肝胆胰脾", "肝胆脾胰"], "organs": ["肝", "胆", "胰", "脾"]},
            {"exam": ["肝胆脾", "肝脾胆"], "organs": ["肝", "胆", "脾"]},
            {"exam": ["肝胆胰", "肝胰胆"], "organs": ["肝", "胆", "胰"]},
            {"exam": ["胰脾", "脾胰"], "organs": ["胰", "脾"]},
            {"exam": ["肝胆"], "organs": ["肝", "胆"]},
            {"exam": ["肾输尿管膀胱", "泌尿系", "泌尿系统", "双肾输尿管"], "organs": ["肾", "输尿管", "膀胱"]},
            {"exam": ["肾输尿管", "双肾"], "organs": ["肾", "输尿管"]},
            {"exam": ["子宫附件", "子宫双附件", "双侧附件", "子宫卵巢"], "organs": ["子宫", "附件"]},
        ],
    },
    # 规则 I：测量值与单位异常（阈值可按院内口径调整）
    "measurement_rules": {
        "suspicious_units": ["m"],        # 径线写作“m”基本为单位误录（应为 mm/cm）
        "ratio_threshold": 6,             # 径线长短轴比例阈值（超过视为量纲异常）
        "min_largest_cm": 5,              # 量纲检查仅对最大径≥该值的测量值生效
        "min_smallest_cm": 1,             # 且短轴≥该值（扁平斑块长短轴悬殊属正常，排除防误报）
    },
    # 规则 J：语句重复/模板残留（阈值可按院内口径调整）
    "redundancy_rules": {
        "similarity_threshold": 0.85,     # 语句相似度阈值（归一化后 SequenceMatcher ratio）
        "min_sentence_len": 10,           # 参与重复比对的最短语句长度
    },
    # 规则 J3：模板残留碎片（孤立括号残片/无关字词，如"(不完整，)""未了。"）
    "stale_fragments": ["未了", "不完整", "未填", "未填写", "待补", "待填", "待查",
                        "此处填写", "见上", "同上", "？？", "??", "xx", "XX"],
    "stale_max_len": 8,                   # 括号内超过该长度的视为正常括注，不判残留
}

# 规则 A 否定识别：否定词须与阳性词同句（子句级判定，避免跨句误判）
NEG_GUARD_RE = re.compile(r"(未见|未探|未及|未显示|没有|排除|否认|无)")
POS_SUFFIX_NEG_RE = re.compile(r"(未见|无明显|正常|减少|减低|不丰富|稀疏|充盈|透声好|透声可|状|探查|扫查)")
CLAUSE_PUNCT = "。；;，,！？!?\n：:【】\r丨"

# 规则 A 强病灶词特征（实体性病变词）；仅弱描述词（增厚/血流信号/回声类）不触发 A
A_STRONG_MORPHEMES = ("结节", "囊肿", "占位", "积液", "肿块", "包块", "肿瘤", "结石",
                      "钙化", "息肉", "肿大", "暗区", "条索", "脂肪肝", "狭窄", "斑块",
                      "血栓", "积血", "分隔", "突出")

# 规则 E 结论未涵盖所见：病变词在结论中的等价表述（命中任一即视为已涵盖）
LESION_NOUNS = ["结节", "占位", "团块", "包块", "肿块", "肿物", "肿瘤", "肌瘤", "病灶", "病变", "息肉"]
# 实体病变的回声性指代（如结论写"左侧乳腺低回声"指代所见的"结节"）
EQUIV_ECHO = ["低回声", "稍低回声", "稍强回声", "高回声", "强回声", "混合回声"]
LESION_EQUIV = LESION_NOUNS + EQUIV_ECHO
COVERAGE_ALIASES = {
    "结节": LESION_EQUIV, "占位": LESION_EQUIV, "包块": LESION_EQUIV, "肿块": LESION_EQUIV,
    "团块": LESION_EQUIV, "肌瘤": LESION_EQUIV, "肿瘤": LESION_EQUIV, "息肉": LESION_EQUIV,
    "囊肿": ["囊肿", "囊性", "囊"], "囊性": ["囊肿", "囊性", "囊"],
    "积液": ["积液", "暗区", "积血"], "暗区": ["暗区", "积液", "囊", "包块", "团块", "低回声"],
    "积血": ["积血", "积液"], "结石": ["结石", "强回声"], "钙化": ["钙化", "光点", "光斑", "强回声"],
    "斑块": ["斑块", "硬化", "硬斑", "软斑", "混合斑"], "血栓": ["血栓"], "狭窄": ["狭窄", "闭塞"], "肿大": ["肿大", "增大"],
    "条索": ["条索"], "脂肪肝": ["脂肪肝", "肝回声"],
}
# 有测量值的回声性病灶（如"大小约11*4mm低回声团"）→ 应在结论中体现
MEASURED_ECHO_RE = re.compile(
    r"\d+(?:\.\d+)?\s*[*×xX]\s*\d+(?:\.\d+)?(?:\s*[*×xX]\s*\d+(?:\.\d+)?)?\s*(?:mm|cm)\s*的?\s*"
    r"((?:低|无|稍低|稍高|稍强|略低|略强|高|强|混合性?|等|不均质|欠均匀)回声(?:团|区|斑|块){0,2})")
# 特征词伴随病灶窗：病灶词之后 12 字内出现的伴随特征（如"结节，边界不清，内见微钙化"）
# 不单独要求结论覆盖，防止把病灶的修饰特征误判为独立遗漏
FEATURE_PREFIX_RE = re.compile(
    "(" + "|".join(LESION_NOUNS + ["淋巴结", "囊肿", "积液", "结石"]) + r")[\s\S]{0,10}$")
# 规则 E 特殊征象：所见出现而结论未呼应 → 提示
# （"显示不清"不列入：肠气干扰致某部位显示不清而结论省略属书写惯例，误报多；
#   充盈缺损等价物含"斑块"——颈动脉报告中"局部血流充盈缺损"即斑块的血流表现，
#   结论写"斑块形成"视为已呼应，LLM 盲测/全量整体口径亦如此）
SPECIAL_SIGNS = [
    ("充盈缺损", ["充盈缺损", "缺损", "斑块", "闭塞", "血栓"], "血流充盈缺损"),
]
# 左室舒张功能减低依据（E峰<A峰 / E/A<1 / 文字式"A峰大于E峰"）
DIASTOLIC_RE = re.compile(
    r"E峰\s*[<＜]\s*A峰|E/A\s*[:：]?\s*0\.|E/A\s*[<＜]\s*1|[AaＡ]峰\s*(?:大于|＞|>)\s*[EeＥ]峰")

# 产前/胎儿报告结论惯例不同，规则 E 豁免；规则 F 用于检查项目相符性
FETUS_CONTEXT_RE = re.compile(r"胎儿|羊水|胎盘|双顶径|股骨|CRL|NT|孕周|孕囊|卵黄囊|脐带|早孕|中孕|晚孕|产前|产科")
FETUS_CONTENT_RE = re.compile(r"胎儿|胎心|双顶径|股骨|羊水|胎盘|CRL|孕囊|卵黄囊|顶臀长|脐带|孕周")
EXAM_FETUS_RE = re.compile(r"胎儿|产科|产前|NT|早孕|中孕|晚孕|宫内孕")
POSTPARTUM_RE = re.compile(r"产后|人流|药流|流产|清宫|引产")  # 产后/流产后检查无胎儿属正常

# 规则 F 完整性与书写规范
MISSING_NUM_SUGGEST_RE = re.compile(r"建议[周月天年]")   # "建议周后复查" 缺时间数值
NUM_GLUE_RE = re.compile(r"(?<!\d)\d\.\d{2,}\.\d")       # "1.01.27" 数值粘连（排除 2024.07.23 类日期）
NUM_MEASURE_NO_UNIT_RE = re.compile(  # 测量值未写单位（"大小约2.5*1.6。"，末尾无 cm/mm）
    r"(?:大小|径|深径|厚|长|宽|高|约)\s*[约]?\s*\d+(?:\.\d+)?(?![\d.])\s*[*x×]\s*\d+(?:\.\d+)?(?![\d.a-zA-Z%万微])")
NUM_UNIT_LOST_RE = re.compile(  # 疑似丢失乘号的小数粘连（"1.512cm" 应为 1.5*1.2cm，排除 0.25/1.75 等规范小数）
    r"(?<![\d.\-])\d\.\d{3,}\s*(?:cm|mm)(?![a-zA-Z])")
PLACENTA_GRADE_RE = re.compile(r"(?:成熟度|功能)\s*[ⅠⅡⅢ0-3]\s*[+＋]\s*级")  # "Ⅱ+级"非Grannum写法

# 规则 F5-F7 基础语病（标点/括号/句末）
PUNCT_REPEAT_RE = re.compile(r"([。，；：、！])\1")          # 标点重复（"。。"、"，，"）
EMPTY_BRACKET_RE = re.compile(r"[（(][\s]*[）)]")            # 空括号"（）"未填写内容
SENT_END_OK_RE = re.compile(r"[。．.！？!?；;：:）)】\]】》>”\"\s]$")  # 正常句末标点/收尾
# 规则 I6 测量范围倒置（"2.5-1.6cm" 疑似起止值颠倒，应为 1.6-2.5cm）
RANGE_INVERT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*[-‐–—~至]\s*(\d+(?:\.\d+)?)\s*(cm|mm)")

# 规则 I 数据与单位异常
# 测量值（径线）：1~3 段数值以 * × x 连接，可带 mm/cm/m 单位
DIM_MEASURE_RE = re.compile(
    r"(?<![\d.])(\d+(?:\.\d+)?)(?:\s*[*×xX]\s*(\d+(?:\.\d+)?)){1,2}\s*(mm|cm|m)?(?![a-zA-Z\d])",
    re.IGNORECASE)
# 单位误录："m" 作径线单位（排除 m/s 血流速度等）
BAD_UNIT_RE = re.compile(r"\d\s*[mM](?!\s*/|eter|[a-zA-Z])")
# 未删除的测量值占位符（"大小约*cm" / "约xmm"）
PLACEHOLDER_MEASURE_RE = re.compile(
    r"(?:约|大小)\s*[*x×X]\s*(?:cm|mm|m)?(?![a-zA-Z\d])|(?:约|大小)\s*[xX]\s*(?:cm|mm)(?![a-zA-Z])",
    re.IGNORECASE)

# 规则 J 语句重复
DUP_WORD_RE = re.compile(r"([一-龥]{2,4})\1")            # 叠词重复（"大小大小"）
DUP_WORD_ALLOW = {"常常", "渐渐", "看看", "谢谢", "宝宝", "娃娃", "星星", "每每",
                  "时时", "处处", "往往", "爷爷", "奶奶", "哥哥", "姐姐", "弟弟", "妹妹"}
SENT_SPLIT_RE = re.compile(r"[。；;！？!?\n]+")
SENT_NORM_RE = re.compile(r"[\s，,、：:（）()【】\[\]“”\"'·\-—_]+")
# 部位主语短语（供 J2 矛盾重复判定部位一致性）：如"双侧腋窝""左侧颈总动脉""右乳"
SITE_PHRASE_RE = re.compile(
    r"((?:双侧|左侧|右侧|左|右)?[一-龥]{1,4}(?:动脉|静脉|窝|叶|腺|肌|管|部|区|支|段|乳))")
SITE_SIDE_PREFIX_RE = re.compile(r"^(双侧|左侧|右侧|左|右)")  # 比较部位时去除侧别前缀
# 规则 J3 模板残留碎片：孤立括号残片（"(不完整，)"）与残缺标记提示
STALE_PAREN_RE = re.compile(r"[（(]([^（）()\n]{1,12})[）)]")
STALE_FRAG_HINT_RE = re.compile(r"(未|待|不|缺|补|填|写|\?|？|…|\.{2,}|[，,。；;：:]$)")

# 规则 B 侧别-器官提取：侧别字(容许夹空格) + 后续 0~8 个汉字（后续再裁剪出器官名）
SIDE_ORGAN_RE = re.compile(r"([左右双])\s*侧?\s*([\u4e00-\u9fa5]{0,8})")
ORGAN_CUT_CHARS = set("见探有未可内上下的与及各低强无囊实结占肿区回液网密点右左")
ORGAN_SUFFIX_STRIP = ("头部", "体部", "尾部")
ORGAN_ALIAS = {"乳": "乳腺", "乳房": "乳腺"}

# 规则 D 分级提取（BI-RADS / TI-RADS / 裸"N类|N级"）
GRADE_RE = re.compile(
    r"(?:BI|TI)[-‐–—]?RADS[^0-9\n]{0,8}?([1-6][abc]?)|(?<![\d.\-~至])([1-6][abc]?)[ \t]*(?:类|级)(?!\d)",
    re.IGNORECASE)
GRADE_SYSTEM_RE = re.compile(  # 仅系统名（供 D3 分级写法检查定位）
    r"(?:BI|TI)[-‐–—]?RADS|C[-‐–—]?TIRADS?", re.IGNORECASE)

# 报告字段标签（正则按"标签+冒号"或"行首标签"两种模式匹配，自适应中英文混乱格式）
# 中文标签 + 英文标签（老旧 PACS 导出常见 examItem/finding/conclusion 风格）
FIELD_PATTERNS = {
    "exam_item": r"检\s*查\s*(?:项\s*目|名\s*称|部\s*位)|申\s*请\s*项\s*目|项目名称|\bexam\s*item\b",
    "name": r"姓\s*名|\bpatient\s*name\b|\bname\b",
    "gender": r"性\s*别|\bgender\b|\bsex\b",
    "age": r"年\s*龄|\bage\b",
    "finding": r"(?:超\s*声|检\s*查)?\s*所\s*见|声像图(?:\s*表\s*现)?|(?:超\s*声)?\s*表\s*现|\bfindings?\b",
    "conclusion": r"超\s*声\s*(?:诊\s*断|提\s*示|结\s*论)|诊\s*断\s*意\s*见"
                  r"|(?:超\s*声)?\s*(?:诊\s*断|提\s*示|结\s*论|印\s*象|意\s*见)"
                  r"|\bconclusions?\b|\bimpression\b|\bdiagnosis\b",
}
# 注意：字段正则含多分支 |，拼接前后缀时必须整体加 (?:...) 分组，
# 否则前缀(^锚定/行首标签)与后缀(冒号)只绑定到最后一个分支，
# 导致"所见/结论"等词在正文中间被误判为字段标签（截断所见段）
COLON_MODE_RES = {f: re.compile(r"(?:" + p + r")[ \t]*[:：]", re.IGNORECASE) for f, p in FIELD_PATTERNS.items()}
LINE_MODE_RES = {f: re.compile(r"^[ \t]*(?:【|◆|●)?(?:\d+[.、]\s*)?(?:" + p +
                               r")(?:】)?[ \t]*(?:[:：]|(?=\n)|$)", re.MULTILINE | re.IGNORECASE)
                 for f, p in FIELD_PATTERNS.items()}
META_FIELDS = ("name", "gender", "age", "exam_item")

# 标准 JSON 输入字段别名（工作站集成规范）：核心字段 exam_item/gender/age/finding/conclusion，
# 兼容中文键名与常见大小写变体
JSON_FIELD_ALIASES = {
    "exam_item": ["exam_item", "examitem", "exam_item_name", "检查项目", "检查部位", "项目名称", "申请项目"],
    "name": ["name", "patient_name", "姓名", "患者姓名"],
    "gender": ["gender", "sex", "性别"],
    "age": ["age", "年龄"],
    "finding": ["finding", "findings", "ultrasound_finding", "超声所见", "所见", "声像图表现", "超声表现"],
    "conclusion": ["conclusion", "conclusions", "impression", "diagnosis", "超声提示",
                   "超声诊断", "超声结论", "提示", "结论", "印象", "诊断意见"],
}
JSON_KEY_SET = {a.strip().lower() for aliases in JSON_FIELD_ALIASES.values() for a in aliases}

# 声明/知情同意等尾段截断（防止其中"未见明显异常"等模板句干扰跨段规则）
DISCLAIMER_RE = re.compile(r"知情同意|告知|声明|注意事项|温馨提示|本报告仅反映|报告医师|审核医师|影像医师")

_XML_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


# ============================ 基础工具 ============================

def _exe_dir() -> Path:
    """打包为 exe 后返回 exe 所在目录，否则返回脚本目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _ensure_stdio() -> None:
    """无控制台 exe（--noconsole）下 sys.stdout/stderr 为 None，重定向到日志文件。"""
    if sys.stdout is not None and sys.stderr is not None:
        return
    try:
        f = open(_exe_dir() / "qc_engine.log", "a", encoding="utf-8", buffering=1)
    except OSError:
        class _Null:
            def write(self, *_a): pass
            def flush(self): pass
        f = _Null()
    if sys.stdout is None:
        sys.stdout = f
    if sys.stderr is None:
        sys.stderr = f


def read_text_any(data: bytes) -> str:
    """多编码自适应解码：BOM → UTF-8 → GB18030 → UTF-16 → 宽松兜底。"""
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        try:
            return data.decode("utf-16")
        except (UnicodeDecodeError, UnicodeError):
            pass
    for enc in ("utf-8-sig", "gb18030"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return data.decode("gb18030", errors="replace")


def clip(text: str, limit: int = 60) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    return text if len(text) <= limit else text[:limit] + "…"


# ============================ 轻量 Excel(.xlsx) 写入器（免 openpyxl/pandas） ============================

class MiniExcel:
    """手写最小化 xlsx 生成器：仅依赖 zipfile，支持多 Sheet 与内联字符串。"""

    @staticmethod
    def _col_letter(idx: int) -> str:
        letters = ""
        idx += 1
        while idx:
            idx, rem = divmod(idx - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    @classmethod
    def _sheet_xml(cls, headers: list, rows: list) -> str:
        all_rows = [list(headers)] + [list(r) + [""] * (len(headers) - len(r)) for r in rows]
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
        for r_idx, row in enumerate(all_rows, 1):
            cells = []
            for c_idx, val in enumerate(row):
                ref = f"{cls._col_letter(c_idx)}{r_idx}"
                if val is None or val == "":
                    continue
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    cells.append(f'<c r="{ref}"><v>{val}</v></c>')
                else:
                    text = escape(_XML_CTRL_RE.sub("", str(val)))
                    cells.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>')
            parts.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
        parts.append("</sheetData></worksheet>")
        return "".join(parts)

    @classmethod
    def write(cls, path: str, sheets: list) -> None:
        """sheets: [(名称, 表头列表, 行列表), ...]"""
        ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>'
              '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
        for i in range(len(sheets)):
            ct.append(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
                      f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        ct.append("</Types>")

        wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>')
        for i, (name, _, _) in enumerate(sheets):
            wb += f'<sheet name="{escape(name[:31])}" sheetId="{i + 1}" r:id="rId{i + 1}"/>'
        wb += "</sheets></workbook>"

        wb_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for i in range(len(sheets)):
            wb_rels.append(f'<Relationship Id="rId{i + 1}" '
                           f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                           f'Target="worksheets/sheet{i + 1}.xml"/>')
        wb_rels.append("</Relationships>")

        root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                     '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                     'relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')

        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", "".join(ct))
            z.writestr("_rels/.rels", root_rels)
            z.writestr("xl/workbook.xml", wb)
            z.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
            for i, (_, headers, rows) in enumerate(sheets):
                z.writestr(f"xl/worksheets/sheet{i + 1}.xml", cls._sheet_xml(headers, rows))


# ============================ 规则加载 ============================

def _normalize_list(raw_items) -> list:
    """兼容 rules.json 中 "词" 与 {"word": "词", "count": n} 两种条目格式。"""
    out = []
    for it in raw_items or []:
        if isinstance(it, dict) and it.get("word"):
            out.append(str(it["word"]))
        elif isinstance(it, str) and it.strip():
            out.append(it.strip())
    return list(dict.fromkeys(out))


def _normalize_mapping(raw) -> dict:
    if isinstance(raw, dict):
        return {str(k): str(v) for k, v in raw.items()}
    if isinstance(raw, list):  # 兼容 correction_details: [{"wrong","correct"}]
        return {str(d["wrong"]): str(d["correct"]) for d in raw
                if isinstance(d, dict) and d.get("wrong") and d.get("correct")}
    return {}


def _normalize_phrases(raw, rules_raw: dict) -> list:
    phrases = _normalize_list(raw)
    if not phrases:  # 兜底：从 logical_rules 模板中提取
        for lr in rules_raw.get("logical_rules", []):
            if isinstance(lr, dict) and lr.get("negative_conclusion_phrases"):
                phrases = _normalize_list(lr["negative_conclusion_phrases"])
                break
    return phrases


def load_rules(rules_path) -> tuple[dict, str]:
    """加载规则库：优先外部 rules.json，失败/缺失时使用内置默认规则。"""
    raw = None
    if rules_path is None:
        candidate = _exe_dir() / "rules.json"
        if candidate.exists():
            rules_path = candidate
    else:
        rules_path = Path(rules_path)

    if rules_path is not None and rules_path.exists():
        try:
            raw = json.loads(rules_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            print(f"[警告] rules.json 解析失败({exc})，改用内置默认规则")
    if raw is None:
        rules = json.loads(json.dumps(DEFAULT_RULES))  # 深拷贝
        return rules, "内置默认规则"

    rules = {
        "positive_keywords": _normalize_list(raw.get("positive_keywords")) or DEFAULT_RULES["positive_keywords"],
        "side_keywords": _normalize_list(raw.get("side_keywords")) or DEFAULT_RULES["side_keywords"],
        "correction_mapping": _normalize_mapping(raw.get("correction_mapping")) or DEFAULT_RULES["correction_mapping"],
        "negative_conclusion_phrases": _normalize_phrases(raw.get("negative_conclusion_phrases"), raw)
                                       or DEFAULT_RULES["negative_conclusion_phrases"],
        "grading_rules": raw.get("grading_rules") or DEFAULT_RULES["grading_rules"],
        "gender_conflict": {**DEFAULT_RULES["gender_conflict"], **(raw.get("gender_conflict") or {})},
        "age_rules": {**DEFAULT_RULES["age_rules"], **(raw.get("age_rules") or {})},
        "exam_organ_map": {**DEFAULT_RULES["exam_organ_map"], **(raw.get("exam_organ_map") or {})},
        "measurement_rules": {**DEFAULT_RULES["measurement_rules"], **(raw.get("measurement_rules") or {})},
        "redundancy_rules": {**DEFAULT_RULES["redundancy_rules"], **(raw.get("redundancy_rules") or {})},
        "stale_fragments": raw.get("stale_fragments") or DEFAULT_RULES["stale_fragments"],
        "stale_max_len": raw.get("stale_max_len") or DEFAULT_RULES["stale_max_len"],
    }
    return rules, f"外部规则库 {Path(rules_path).name}"


def compile_rules(rules: dict) -> dict:
    """预编译正则/派生数据（启动一次，检测毫秒级）。"""
    kws = sorted(rules["positive_keywords"], key=len, reverse=True)
    rules["_positive_re"] = re.compile("|".join(re.escape(k) for k in kws)) if kws else re.compile(r"(?!x)x")
    rules["_negative_phrases"] = sorted(set(rules["negative_conclusion_phrases"]), key=len, reverse=True)
    return rules


# ============================ 格式自适应解析（Parser Adapter） ============================

def parse_report(text: str) -> dict:
    """正则自适应抽取 exam_item/gender/age/finding/conclusion 字段。"""
    candidates = []  # (start, end, field, priority) 冒号模式优先
    for field in FIELD_PATTERNS:
        for m in COLON_MODE_RES[field].finditer(text):
            candidates.append((m.start(), m.end(), field, 2))
        for m in LINE_MODE_RES[field].finditer(text):
            candidates.append((m.start(), m.end(), field, 1))
    candidates.sort(key=lambda t: (t[0], -t[3], -(t[1] - t[0])))

    labels = []
    for s, e, f, _p in candidates:
        if labels and s < labels[-1][1]:  # 与上一个已选标签重叠 → 丢弃
            continue
        labels.append((s, e, f))
    labels.sort(key=lambda t: t[0])

    fields: dict = {}
    for idx, (s, e, f) in enumerate(labels):
        nxt = labels[idx + 1][0] if idx + 1 < len(labels) else len(text)
        if f in META_FIELDS:
            nl = text.find("\n", e)
            vend = len(text) if nl == -1 else nl
        else:
            vend = nxt
        value = text[e:vend].strip(" \t：:，,；;")
        if f not in META_FIELDS:  # 截断声明/知情同意尾段
            dm = DISCLAIMER_RE.search(value)
            if dm:
                value = value[:dm.start()].strip(" \t：:，,；;\n")
        if fields.get(f, "") == "":
            fields[f] = value

    finding = fields.get("finding", "")
    conclusion = fields.get("conclusion", "")
    notes = []
    if not finding and not conclusion:  # 完全无标签 → 全文按所见处理
        finding, conclusion = text.strip(), ""
        notes.append("未识别到字段标签，全文按【所见】处理")
    elif not conclusion:
        notes.append("未识别到【结论】段，跨段规则(A/B/D)已跳过")

    gender = re.search(r"[男女]", fields.get("gender", ""))
    age = re.search(r"\d{1,3}", fields.get("age", ""))
    return {
        "exam_item": fields.get("exam_item", ""),
        "name": fields.get("name", ""),
        "gender": gender.group(0) if gender else "",
        "age": age.group(0) if age else "",
        "finding": finding,
        "conclusion": conclusion,
        "parse_notes": notes,
    }


def _normalize_gender(raw: str) -> str:
    """性别归一化：男/女/M/F/male/female/男性/女性 → '男'/'女'（无法识别返回空）。"""
    s = str(raw or "").strip()
    if not s:
        return ""
    m = re.search(r"[男女]", s)
    if m:
        return m.group(0)
    low = s.lower()
    if low in ("m", "male", "man", "boy"):
        return "男"
    if low in ("f", "female", "woman", "girl"):
        return "女"
    return ""


def parse_json_report(obj: dict) -> dict:
    """标准 JSON 输入解析（工作站集成规范）：按字段别名抽取各字段，无需正则猜格式。

    核心字段: exam_item/gender/age/finding/conclusion
    （兼容中文键名 检查项目/性别/年龄/超声所见/超声提示 及常见大小写变体）
    """
    lower_map = {str(k).strip().lower(): v for k, v in obj.items() if v is not None}
    fields: dict = {}
    for field, aliases in JSON_FIELD_ALIASES.items():
        for alias in aliases:
            val = lower_map.get(alias.lower())
            if val is not None and str(val).strip():
                fields[field] = str(val).strip()
                break

    notes = ["标准 JSON 输入"]
    if not fields.get("finding") and not fields.get("conclusion"):
        # 兼容旧版 {"text": 全文} 输入：按文本自适应解析，不判字段缺失
        raw_text = str(obj.get("text") or obj.get("content") or "").strip()
        if raw_text:
            parsed = parse_report(raw_text)
            parsed["parse_notes"] = ["标准 JSON 输入（text 全文兼容路径）"] + (parsed.get("parse_notes") or [])
            parsed["input_format"] = "text"  # 内容为文本报告，供字段缺失(C2)等文本规则生效
            return parsed
        notes.append("JSON 中未提供超声所见/超声提示字段")
    elif not fields.get("conclusion"):
        notes.append("未提供【超声提示】字段，跨段规则(A/B/D)已跳过")

    age_raw = str(fields.get("age", ""))
    age_m = re.search(r"\d{1,3}", age_raw)
    return {
        "exam_item": fields.get("exam_item", ""),
        "name": fields.get("name", ""),
        "gender": _normalize_gender(fields.get("gender", "")),
        "age": age_m.group(0) if age_m else "",
        "finding": fields.get("finding", ""),
        "conclusion": fields.get("conclusion", ""),
        "parse_notes": notes,
        "input_format": "json",
    }


# ============================ 质控规则 A-D ============================

def _clause_text(text: str, start: int, end: int) -> tuple[str, str]:
    """截取命中词所在子句的前/后片段（以标点为界，最长 14 字）。"""
    i = start - 1
    while i >= 0 and (start - i) <= 14 and text[i] not in CLAUSE_PUNCT:
        i -= 1
    j = end
    while j < len(text) and (j - end) <= 14 and text[j] not in CLAUSE_PUNCT:
        j += 1
    return text[i + 1:start], text[end:j]


def _non_negated_positive_words(text: str, rules: dict) -> list:
    """提取未带否定的阳性病变词 [(word, start)]：否定词须出现在同一子句（前段/后段）。"""
    found = []
    for m in rules["_positive_re"].finditer(text):
        before, after = _clause_text(text, m.start(), m.end())
        if NEG_GUARD_RE.search(before) or before.endswith("不"):
            continue
        if POS_SUFFIX_NEG_RE.search(after):
            continue
        if all(w != m.group(0) for w, _s in found):
            found.append((m.group(0), m.start()))
    return found


def rule_a_logical(finding: str, conclusion: str, rules: dict) -> list:
    """规则 A【结论遗漏病变】：所见含强病灶词，结论却是纯阴性表述。"""
    if not finding or not conclusion:
        return []
    strong_pairs = [(w, s) for w, s in _non_negated_positive_words(finding, rules)
                    if any(mo in w for mo in A_STRONG_MORPHEMES)]
    if not strong_pairs:
        return []
    strong_hits = [w for w, _s in strong_pairs]
    # 等价抑制：所见强词均为"结节"且两段均已提及"淋巴结"（反应性淋巴结描述习惯）→ 不算矛盾
    if all("结节" in w for w in strong_hits) and "淋巴结" in finding and "淋巴结" in conclusion:
        return []
    if _non_negated_positive_words(conclusion, rules):
        return []  # 结论已描述病变（如"斑块形成，余未见异常"）→ 不算矛盾
    # 结论以等价表述涵盖病灶（如"动脉粥样硬化性改变"涵盖"斑块"）→ 不算矛盾
    if any(a in conclusion for w, s in strong_pairs
           for a in _coverage_aliases(finding[max(0, s - 4):s + len(w)])):
        return []
    neg_hit = next((p for p in rules["_negative_phrases"] if p in conclusion), None)
    if not neg_hit:
        return []
    return [{
        "rule_id": "A", "rule": "结论遗漏病变", "severity": "high",
        "field": "conclusion", "fragment": clip(conclusion),
        "message": f"所见描述阳性病变（{('、'.join(strong_hits[:4]))}等），但结论为阴性表述“{neg_hit}”，疑似遗漏病变",
        "suggestion": "请核对所见与结论是否矛盾：若确有病灶，结论应逐项描述病变部位、大小与性质",
    }]


def _side_organ_map(text: str) -> tuple[dict, dict]:
    """提取文本中 (器官→侧别集合) 与 (器官→原始片段)，供规则 B 配对。"""
    organ_sides: dict = {}
    organ_example: dict = {}
    for m in SIDE_ORGAN_RE.finditer(text):
        side_char, raw_organ = m.group(1), m.group(2)
        cut = min((raw_organ.index(c) for c in ORGAN_CUT_CHARS if c in raw_organ), default=len(raw_organ))
        organ = raw_organ[:cut]
        for suf in ORGAN_SUFFIX_STRIP:  # "附睾头部"→"附睾"
            if organ.endswith(suf):
                organ = organ[:-len(suf)]
                break
        if organ.startswith("叶"):  # "左侧叶甲状腺"→"甲状腺"
            organ = organ[1:]
        organ = organ.replace("淋巴结", "").replace("淋巴", "")  # "腋窝淋巴"→"腋窝"（结节约断残留）
        organ = ORGAN_ALIAS.get(organ, organ)
        sides = {"L"} if side_char == "左" else {"R"} if side_char == "右" else {"L", "R"}
        key = organ or "*"  # "*" 表示未识别出器官的泛侧别提及
        organ_sides.setdefault(key, set()).update(sides)
        organ_example.setdefault(key, m.group(0))
    return organ_sides, organ_example


_SIDE_NAME = {"L": "左", "R": "右"}


def rule_b_side(finding: str, conclusion: str, rules: dict) -> list:
    """规则 B【左右侧不符】：所见与结论中同一器官左右侧冲突。"""
    if not finding or not conclusion:
        return []
    f_map, f_ex = _side_organ_map(finding)
    c_map, c_ex = _side_organ_map(conclusion)

    # 1) 器官级配对：同一器官在所见与结论中侧别互斥（左 vs 右，且任一方非双侧）
    for organ in sorted(set(f_map) & set(c_map)):
        if organ == "*":
            continue
        fs, cs = f_map[organ], c_map[organ]
        conflict = (fs == {"L"} and cs == {"R"}) or (fs == {"R"} and cs == {"L"})
        if conflict:
            return [{
                "rule_id": "B", "rule": "左右侧不符", "severity": "high",
                "field": "conclusion", "fragment": clip(c_ex.get(organ, organ)),
                "message": f"器官“{organ}”侧别冲突：所见为{('、'.join(sorted(_SIDE_NAME[s] for s in fs)))}侧，"
                           f"结论为{('、'.join(sorted(_SIDE_NAME[s] for s in cs)))}侧",
                "suggestion": f"请核对{organ}病变的左右侧别，确保所见与结论一致",
            }]
    # 2) 全局兜底：两段各自只出现单一侧别且互斥
    f_all, c_all = set().union(*f_map.values()) if f_map else set(), set().union(*c_map.values()) if c_map else set()
    if (f_all == {"L"} and c_all == {"R"}) or (f_all == {"R"} and c_all == {"L"}):
        return [{
            "rule_id": "B", "rule": "左右侧不符", "severity": "high",
            "field": "conclusion", "fragment": clip(conclusion),
            "message": f"所见仅提及{(_SIDE_NAME[next(iter(f_all))])}侧，结论仅提及{(_SIDE_NAME[next(iter(c_all))])}侧，左右侧疑似不符",
            "suggestion": "请核对报告整体的左右侧方向是否一致",
        }]
    return []


def rule_c_typos(text: str, rules: dict) -> list:
    """规则 C【用词勘误】：命中纠错映射表，给出规范替换建议。"""
    issues = []
    for wrong, correct in rules["correction_mapping"].items():
        if not wrong or wrong == correct:
            continue
        start = 0
        count = 0
        while count < 3:  # 每个误写词最多报告 3 处
            pos = text.find(wrong, start)
            if pos == -1:
                break
            issues.append({
                "rule_id": "C", "rule": "用词勘误", "severity": "medium",
                "field": "full_text", "fragment": clip(text[max(0, pos - 12):pos + len(wrong) + 12]),
                "message": f"检出不规范用词“{wrong}”",
                "suggestion": f"建议将“{wrong}”替换为“{correct}”",
            })
            start = pos + len(wrong)
            count += 1
    return issues


def rule_c_field_missing(parsed: dict) -> list:
    """规则 C2【字段缺失】：报告头部完整（有性别/年龄/检查项目）但核心字段（所见/结论）缺失。

    保守口径：仅对文本格式、且报告头完整的报告生效，避免无标签片段/JSON 输入误报。
    """
    if parsed.get("input_format") not in (None, "text"):  # 仅文本输入判字段缺失（parse_report 不含该键时按文本处理）
        return []
    notes = parsed.get("parse_notes") or []
    if any("全文按【所见】处理" in n for n in notes):
        return []  # 完全无标签片段，不判字段缺失
    has_header = bool(parsed.get("gender") or parsed.get("age") or parsed.get("exam_item"))
    if not has_header:
        return []
    finding, conclusion = parsed["finding"], parsed["conclusion"]
    issues = []
    if finding and not conclusion:
        issues.append({
            "rule_id": "C", "check": "C2", "rule": "字段缺失", "severity": "high",
            "field": "conclusion", "fragment": clip(finding[-24:]),
            "message": "报告缺少【超声提示/结论】字段（仅有所见描述）",
            "suggestion": "请补充超声提示（诊断结论）字段",
        })
    elif conclusion and not finding:
        issues.append({
            "rule_id": "C", "check": "C2", "rule": "字段缺失", "severity": "high",
            "field": "finding", "fragment": clip(conclusion[-24:]),
            "message": "报告缺少【超声所见】字段（仅有结论）",
            "suggestion": "请补充超声所见（声像图表现）字段",
        })
    return issues


def _extract_grades(text: str) -> list:
    """提取文本中的分级号（BI-RADS/TI-RADS/N类|N级）。"""
    grades = []
    for m in GRADE_RE.finditer(text):
        g = m.group(1) or m.group(2)
        if g and g not in grades:
            grades.append(g)
    return grades


def rule_d_grading(parsed: dict, rules: dict) -> list:
    """规则 D【分级冲突】：所见恶性征象与 BI-RADS 分级不匹配 / 前后分级不一致 / 胎盘分级写法。"""
    finding, conclusion = parsed["finding"], parsed["conclusion"]
    issues = []

    # D3: 分级写法不规范（BI-RADS/TI-RADS 分级号后应写"N类"而非"N级"）
    for field, txt in (("finding", finding), ("conclusion", conclusion)):
        if not txt:
            continue
        for m in GRADE_SYSTEM_RE.finditer(txt):
            tail = txt[m.end():m.end() + 3]
            gm = re.match(r"[ \t]*(\d+)[ \t]*级", tail)
            if gm:
                issues.append({
                    "rule_id": "D", "check": "D3", "rule": "分级写法不规范", "severity": "medium",
                    "field": field, "fragment": clip(txt[m.start():m.end() + len(gm.group(0))]),
                    "message": f"分级写法“{gm.group(1)}级”不规范，{m.group(0)} 分级号后应写“{gm.group(1)}类”",
                    "suggestion": f"请改为“{m.group(0)} {gm.group(1)}类”的规范写法",
                })

    # D4: 胎盘成熟度非 Grannum 标准写法（标准为 0~Ⅲ 级，不带 +；不依赖结论段）
    pm = PLACENTA_GRADE_RE.search(finding)
    if pm and ("胎盘" in finding or "成熟度" in pm.group(0)):
        issues.append({
            "rule_id": "D", "check": "D4", "rule": "分级写法不规范", "severity": "medium",
            "field": "finding", "fragment": clip(pm.group(0)),
            "message": f"胎盘成熟度写法“{pm.group(0)}”非标准 Grannum 分级（应为 0~Ⅲ 级，不带“+”）",
            "suggestion": "请按 Grannum 分级规范书写胎盘成熟度（0、Ⅰ、Ⅱ、Ⅲ 级）",
        })

    if not conclusion:
        return issues
    grades_conc, grades_find = _extract_grades(conclusion), _extract_grades(finding)

    # D1b: 同一病灶结论内多个分级并存且不一致（"右叶结节（3类）…（4a类）" 疑似分级未统一或模板残留）；
    # 多病灶各自不同分级（左叶3类、右叶4a类）属正常，按侧别语境区分
    if len(grades_conc) >= 2:
        gctx = [(m.group(1) or m.group(2), conclusion[max(0, m.start() - 12):m.start()])
                for m in GRADE_RE.finditer(conclusion)]
        uniq = {g for g, _c in gctx}
        sides = [("左" if "左" in c else ("右" if "右" in c else "")) for _g, c in gctx]
        has_side = any(sides)
        same_site = ({s for s in sides if s} != {"左", "右"}) if has_side else True
        if len(uniq) >= 2 and same_site:
            issues.append({
                "rule_id": "D", "check": "D1", "rule": "分级冲突", "severity": "high",
                "field": "conclusion", "fragment": clip(conclusion),
                "message": f"同一部位结论内出现多个分级（{'类、'.join(sorted(uniq))}类）且不一致，疑似分级未统一或模板残留",
                "suggestion": "请核对分级结论，统一为一个分级结果并删除残留的其他分级描述",
            })

    # D1: 所见与结论分级号不一致
    if len(grades_conc) == 1 and len(grades_find) == 1 and grades_conc != grades_find:
        issues.append({
            "rule_id": "D", "rule": "分级冲突", "severity": "high",
            "field": "conclusion", "fragment": clip(conclusion),
            "message": f"分级前后不一致：所见为{grades_find[0]}类，结论为{grades_conc[0]}类",
            "suggestion": "请核对分级依据，统一所见与结论的分级结果",
        })

    # D2: 所见含恶性征象但分级偏低（如 微钙化/边界不清 却给 2 类）
    grading = rules["grading_rules"]
    features_map = grading.get("min_grade_features", {})
    if len(grades_conc) == 1:
        ctx = (parsed.get("exam_item", "") + finding).lower()
        # 分级语境：乳腺/甲状腺等上下文关键词，或结论含分级系统名（BI-RADS/TI-RADS/C-TIRADS）
        organ_ctx = any(k in ctx for k in grading.get("context_keywords", [])) \
            or re.search(r"(BI|TI|C-TI)[-‐–—]?RADS", conclusion, re.IGNORECASE)
        if organ_ctx:
            grade = int(re.match(r"\d+", grades_conc[0]).group())  # "4a"→4
            features = [f for f in features_map if f in finding]
            if features:
                need = max(int(features_map[f]) for f in features)
                if grade < need:
                    issues.append({
                        "rule_id": "D", "check": "D2", "rule": "分级冲突", "severity": "medium",
                        "field": "conclusion", "fragment": clip(conclusion),
                        "message": f"所见含风险征象（{('、'.join(features[:4]))}），通常对应 {need} 类及以上，"
                                   f"但结论为 {grade} 类，存在分级偏低风险",
                        "suggestion": f"请复核分级：{('、'.join(features[:4]))} 等征象建议评估为 {need} 类或请上级医师复核",
                    })
    return issues


# ============================ 质控规则 E/F ============================

def _coverage_aliases(word: str) -> list:
    """病变词在结论中的等价表述：word 含 COVERAGE_ALIASES 键时并入其别名。"""
    aliases = [word]
    for key, vals in COVERAGE_ALIASES.items():
        if key in word:
            aliases.extend(v for v in vals if v not in aliases)
    return aliases


def _echo_aliases(phrase: str) -> list:
    """带测量值回声病灶的结论等价表述（按回声性质归类 + phrase 内病变词等价物）。"""
    if "强回声" in phrase or "高回声" in phrase:
        base = ["强回声", "高回声", "钙化", "光点", "光斑", "结石", "斑块"]
    elif "无回声" in phrase:
        base = ["无回声", "囊肿", "囊性", "囊", "暗区", "积液"]
    else:
        base = [phrase]
    aliases = list(base)
    for key, vals in COVERAGE_ALIASES.items():
        if key in phrase:
            aliases.extend(v for v in vals if v not in aliases)
    aliases.extend(v for v in LESION_NOUNS if v not in aliases)
    return aliases


def rule_e_coverage(parsed: dict, rules: dict) -> list:
    """规则 E【结论未涵盖所见】：所见病灶/征象在结论中无任何呼应（产前报告豁免）。"""
    finding, conclusion = parsed["finding"], parsed["conclusion"]
    if not finding or not conclusion:
        return []
    if FETUS_CONTEXT_RE.search(finding):
        return []  # 产前/胎儿报告结论惯例不同，豁免
    uncovered = []
    for w, pos in _non_negated_positive_words(finding, rules):
        if not any(mo in w for mo in A_STRONG_MORPHEMES):
            continue  # 仅实体性病灶词要求结论覆盖
        if "条索" in w or "分隔" in w:
            continue  # 条索/分隔属结构描述，结论惯例不逐一呼应
        if FEATURE_PREFIX_RE.search(finding[max(0, pos - 12):pos]):
            continue  # 病灶伴随特征（如"结节…内见微钙化"），不单独要求覆盖
        if w == "结节" and "淋巴结" in finding and "淋巴结" in conclusion:
            continue  # 淋巴结反应性描述惯例，同规则 A 等价抑制
        # 词前后语境扩展（"囊性结节"按囊肿等价、"不均质回声斑块"按斑块等价处理）
        ext = finding[max(0, pos - 4):pos + len(w)]
        if not any(a in conclusion for a in _coverage_aliases(ext)):
            uncovered.append(w)
    for m in MEASURED_ECHO_RE.finditer(finding):
        phrase = m.group(1)
        # 分类仍按 phrase 本身；测量值紧后的病变语境词（如"强回声扁平斑块"）
        # 仅用于合并 COVERAGE_ALIASES 等价族，避免混合回声词翻转分类分支
        aliases = _echo_aliases(phrase)
        for key, vals in COVERAGE_ALIASES.items():
            if key in finding[m.end():m.end() + 6]:
                aliases.extend(v for v in vals if v not in aliases)
        if not any(a in conclusion for a in aliases):
            uncovered.append(f"有测量值的{phrase}")
    for pat, aliases, name in SPECIAL_SIGNS:
        if re.search(pat, finding) and not any(a in conclusion for a in aliases):
            uncovered.append(name)
    # 舒张功能依据也可能出现在表格版式的结论段（如"E峰<A峰"）→ 检查所见+结论全文；
    # 豁免只看结论段：结论已另行含 E/A 或 E峰 数值（表格引用测量值）时视为已呼应；
    # 需剔除命中词本身（"A峰大于E峰"含"E峰"，不可自我豁免）；所见中其他 E峰 提及
    # （如"三尖瓣E峰"）与豁免无关，不得纳入
    dia_m = DIASTOLIC_RE.search(finding + "\n" + conclusion)
    if dia_m:
        conc_rest = conclusion.replace(dia_m.group(0), "\n", 1)
        if not any(a in conclusion for a in ("舒张功能", "松弛", "顺应性")) \
                and not re.search(r"E/A|E峰", conc_rest):
            uncovered.append("舒张功能减低依据（E峰<A峰）")
    if not uncovered:
        return []
    dedup = []  # "斑块" 与 "有测量值的不均质回声斑块" 等互相包含的条目合并
    for u in uncovered:
        core = u.replace("有测量值的", "")
        if any(core in d or d in core for d in dedup):
            continue
        dedup.append(u)
    uncovered = dedup
    items = "、".join(f"“{u}”" for u in uncovered[:4])
    return [{
        "rule_id": "E", "rule": "结论未涵盖所见", "severity": "high",
        "field": "conclusion", "fragment": clip(conclusion),
        "message": f"所见中的{items}未在结论中体现，存在结论未涵盖所见风险",
        "suggestion": "结论应逐项呼应超声所见中的阳性发现（部位+大小+性质），请核对是否遗漏",
    }]


def rule_f_completeness(parsed: dict, rules: dict) -> list:
    """规则 F【完整性与书写规范】：字段缺失 / 建议缺数值 / 数值粘连 / 项目内容不符。"""
    issues = []
    exam, finding, conclusion = parsed.get("exam_item", ""), parsed["finding"], parsed["conclusion"]

    # F1 检查项目缺失已移除：相当一部分医院导出格式本无该字段，盲测显示
    # 53 条全为口径性误报（LLM 亦不判错），仅保留 F4 的项目/内容相符性检查

    # F2: 复查建议缺少时间数值（"建议周后复查"→"建议3周后复查"）
    sm = MISSING_NUM_SUGGEST_RE.search(conclusion) if conclusion else None
    if sm:
        issues.append({
            "rule_id": "F", "check": "F2", "rule": "书写规范", "severity": "medium",
            "field": "conclusion", "fragment": clip(conclusion),
            "message": f"建议语“{sm.group(0)}后…”缺少具体时间数值",
            "suggestion": "复查建议应写明具体间隔时间（如“建议 3 周后复查”），便于随访管理",
        })

    # F3: 数值粘连（"1.01.27m/s" 疑似丢失分隔符）
    for field, txt in (("finding", finding), ("conclusion", conclusion)):
        if not txt:
            continue
        gm = NUM_GLUE_RE.search(txt)
        if gm:
            issues.append({
                "rule_id": "F", "check": "F3", "rule": "书写规范", "severity": "medium",
                "field": field, "fragment": clip(gm.group(0)),
                "message": f"检测到数值粘连“{gm.group(0)}”，疑似两个数值丢失分隔符",
                "suggestion": "请核对测量数值并补充分隔（如 1.0 与 1.27 m/s）",
            })

    # F5-F7 基础语病（标点重复 / 空括号 / 句末缺标点）
    for field, txt in (("finding", finding), ("conclusion", conclusion)):
        if not txt:
            continue
        for m in PUNCT_REPEAT_RE.finditer(txt):
            issues.append({
                "rule_id": "F", "check": "F5", "rule": "书写规范", "severity": "low",
                "field": field, "fragment": clip(txt[max(0, m.start() - 8):m.end() + 8]),
                "message": f"标点重复“{m.group(0)}”",
                "suggestion": f"请删除重复标点（保留一个“{m.group(1)}”）",
            })
        for m in EMPTY_BRACKET_RE.finditer(txt):
            issues.append({
                "rule_id": "F", "check": "F6", "rule": "书写规范", "severity": "low",
                "field": field, "fragment": clip(txt[max(0, m.start() - 8):m.end() + 8]),
                "message": f"空括号“{m.group(0)}”未填写内容",
                "suggestion": "请填写括号内容或删除空括号",
            })
        body = txt.rstrip()
        if body and not SENT_END_OK_RE.search(body):
            issues.append({
                "rule_id": "F", "check": "F7", "rule": "书写规范", "severity": "low",
                "field": field, "fragment": clip(body[-16:]),
                "message": f"{('超声所见' if field == 'finding' else '超声提示')}末尾缺少句末标点",
                "suggestion": "请在段落末尾补全句号等标点",
            })

    # F4: 检查项目与内容不符（产科项目无胎儿内容 / 妇科项目出现胎儿内容）
    if exam and finding:
        is_fetus_exam = bool(EXAM_FETUS_RE.search(exam))
        has_fetus_content = bool(FETUS_CONTENT_RE.search(finding))
        postpartum = bool(POSTPARTUM_RE.search(exam))
        if is_fetus_exam and not postpartum and not has_fetus_content:
            issues.append({
                "rule_id": "F", "check": "F4", "rule": "项目内容不符", "severity": "high",
                "field": "exam_item", "fragment": clip(exam),
                "message": "检查项目含胎儿/产科内容，但超声所见中无任何胎儿相关描述，项目与内容不符",
                "suggestion": "请核实检查项目与实际检查内容是否一致",
            })
        elif not is_fetus_exam and not postpartum \
                and re.search(r"妇科", exam) and has_fetus_content:
            issues.append({
                "rule_id": "F", "check": "F4", "rule": "项目内容不符", "severity": "high",
                "field": "exam_item", "fragment": clip(exam),
                "message": "检查项目为妇科超声，但所见包含胎儿/产科内容，项目与内容不符",
                "suggestion": "请核实检查项目名称（应改为产科/产前超声检查）",
            })
    return issues


# ============================ 质控规则 G/H ============================

def rule_g_gender_conflict(parsed: dict, rules: dict) -> list:
    """规则 G【性别与检查项目冲突】：患者性别与检查项目/所见部位矛盾。

    例：性别男但检查项目/所见含子宫、卵巢、附件；性别女但检查前列腺。
    器官词表在 rules.json 的 gender_conflict 段可配置（如乳腺可按院内要求增删）。
    """
    gender = parsed.get("gender", "")
    if gender not in ("男", "女"):
        return []  # 性别缺失/无法识别时跳过，不猜测
    exam, finding = parsed.get("exam_item", ""), parsed.get("finding", "")
    conclusion = parsed.get("conclusion", "")
    text = exam + "\n" + finding + "\n" + conclusion  # 结论段亦须覆盖（"结论：子宫肌瘤"对男性即冲突）
    if not text.strip():
        return []
    gc = rules.get("gender_conflict") or DEFAULT_RULES["gender_conflict"]
    if any(k and k in text for k in gc.get("exempt_keywords", [])):
        return []  # 男性乳腺/两性畸形等特殊场景豁免
    wrong_organs = gc.get("male_organs", []) if gender == "女" else gc.get("female_organs", [])
    hits = []
    for organ in wrong_organs:
        if not organ or organ not in text:
            continue
        where = "、".join(w for w, t in (("检查项目", exam), ("超声所见", finding),
                                        ("超声提示", conclusion)) if organ in t)
        hits.append((organ, where))
    issues = []
    for organ, where in hits[:3]:
        issues.append({
            "rule_id": "G", "rule": "性别与检查项目冲突", "severity": "high",
            "field": "exam_item" if organ in exam else ("conclusion" if organ in conclusion else "finding"),
            "fragment": clip(exam or finding or conclusion),
            "message": f"患者性别为{gender}，但{where}中出现{'女性' if gender == '男' else '男性'}"
                       f"特有部位“{organ}”，性别与检查项目冲突",
            "suggestion": "请核对患者性别与检查项目/所见是否对应，排除性别录入错误或报告套用模板错误",
        })
    return issues


def rule_g_age_conflict(parsed: dict, rules: dict) -> list:
    """规则 G2【年龄与检查内容逻辑冲突】：年龄超出合理生理范围仍出现该类检查内容。

    例：79 岁女性报告含胎儿/孕囊/羊水等妊娠描述（几乎不可能自然妊娠），
    疑似年龄录入错误或报告套用他人模板。归入「性别年龄冲突」错误类型。
    """
    age_text = str(parsed.get("age", "") or "")
    m = re.search(r"\d{1,3}", age_text)
    if not m:
        return []  # 年龄缺失/无法识别时跳过，不猜测
    age = int(m.group())
    if age <= 0 or age > 120:
        return []
    ar = rules.get("age_rules") or DEFAULT_RULES["age_rules"]
    max_age = int(ar.get("pregnancy_max_age", 55))
    if age < max_age:
        return []
    text = (parsed.get("exam_item", "") + "\n" + parsed.get("finding", "")
            + "\n" + parsed.get("conclusion", ""))
    hits = [k for k in ar.get("pregnancy_keywords", []) if k and k in text]
    if not hits:
        return []
    return [{
        "rule_id": "G", "check": "G2", "rule": "性别年龄冲突", "severity": "high",
        "field": "age",
        "fragment": clip(f"age: {age_text}｜{hits[0]}…"),
        "message": f"患者年龄 {age} 岁，但报告含妊娠相关描述（{('、'.join(hits[:4]))}），"
                   f"年龄与检查内容逻辑冲突，疑似年龄录入错误或套用他人模板",
        "suggestion": f"请核对患者年龄与检查内容是否对应：{age} 岁不应出现妊娠/胎儿类检查描述，"
                      "排除年龄录入错误或报告套用模板错误",
    }]


def rule_h_exam_coverage(parsed: dict, rules: dict) -> list:
    """规则 H【所见遗漏检查部位】：检查项目要求多个部位，所见缺少部分器官描述。

    例：检查项目"肝胆胰脾"但超声所见只描述肝脏 → 遗漏胆、胰、脾。
    器官别名与项目-部位映射在 rules.json 的 exam_organ_map 段可配置。
    仅当"部分器官已描述、部分缺失"时触发（单部位项目或整段未描述不误报）。
    """
    exam, finding = parsed.get("exam_item", ""), parsed.get("finding", "")
    if not exam.strip() or not finding.strip():
        return []
    emap = rules.get("exam_organ_map") or DEFAULT_RULES["exam_organ_map"]
    # 项目名匹配：取最长匹配的项目定义（"肝胆胰脾"优先于"肝胆"）
    best = None
    for item in emap.get("exam_items", []):
        for kw in item.get("exam", []):
            if kw and kw in exam and (best is None or len(kw) > len(best[0])):
                best = (kw, item.get("organs", []))
    if not best:
        return []
    kw, organs = best
    aliases = emap.get("organ_aliases", {})
    mentioned = [o for o in organs if any(a and a in finding for a in aliases.get(o, [o]))]
    missing = [o for o in organs if o not in mentioned]
    if not missing or not mentioned:
        return []  # 全部已描述，或整段未描述任一要求部位（后者不属"漏项"口径）
    return [{
        "rule_id": "H", "rule": "所见遗漏检查部位", "severity": "high",
        "field": "finding", "fragment": clip(finding),
        "message": f"检查项目“{kw}”应检查部位（{('、'.join(organs))}），"
                   f"但超声所见中缺少{('、'.join(missing))}的描述，所见遗漏检查部位",
        "suggestion": f"请补充{('、'.join(missing))}的超声所见描述，确保检查项目要求的部位逐一记录",
    }]


# ============================ 质控规则 I/J ============================

def rule_i_measurement(parsed: dict, rules: dict) -> list:
    """规则 I【数据与单位异常】：测量值单位错误/缺失、未删占位符、径线量纲异常。

    例："大小约2.5*1.6m" 单位应为 mm/cm；"1.5*12cm" 径线比例 8 倍，疑似小数点错（1.5*1.2cm）；
    "内见大小约*cm" 占位符未填。
    """
    mr = rules.get("measurement_rules") or DEFAULT_RULES["measurement_rules"]
    bad_units = set(mr.get("suspicious_units", ["m"]))
    ratio_th = float(mr.get("ratio_threshold", 6))
    min_largest = float(mr.get("min_largest_cm", 5))
    issues = []
    seen_bad = set()

    for field in ("finding", "conclusion"):
        text = parsed.get(field, "")
        if not text:
            continue

        # I1 单位误录（径线写作 "m"）
        for m in BAD_UNIT_RE.finditer(text):
            frag = text[max(0, m.start() - 6):m.end()]
            if m.group(0).strip() not in seen_bad:
                seen_bad.add(m.group(0).strip())
                issues.append({
                    "rule_id": "I", "check": "I1", "rule": "数据与单位异常", "severity": "high",
                    "field": field, "fragment": clip(frag),
                    "message": f"测量值单位“m”可疑（原文“{clip(frag)}”），径线单位应为 mm 或 cm",
                    "suggestion": "请核对测量值单位，径线通常以 mm/cm 记录",
                })

        # I2 测量值占位符未填写（"大小约*cm"）
        for m in PLACEHOLDER_MEASURE_RE.finditer(text):
            issues.append({
                "rule_id": "I", "check": "I2", "rule": "数据与单位异常", "severity": "medium",
                "field": field, "fragment": clip(text[max(0, m.start() - 6):m.end() + 6]),
                "message": f"存在未填写的测量值占位符“{m.group(0)}”",
                "suggestion": "请补充完整测量值或删除未填写的占位文本",
            })

        # I3 径线量纲异常：换算 cm 后长短轴比例过大（如 1.5*12cm，疑似小数点错）
        for m in DIM_MEASURE_RE.finditer(text):
            unit = (m.group(3) or "").lower()
            if unit in bad_units:
                continue  # 已由 I1 报出，不重复
            num_part = m.group(0)[:-len(unit)] if unit else m.group(0)
            vals = [float(v) for v in re.findall(r"\d+(?:\.\d+)?", num_part)]
            if len(vals) < 2:
                continue
            # 全部值换算到 cm 比较量纲
            scale = {"cm": 1.0, "mm": 0.1, "m": 100.0}.get(unit, 1.0)
            cm_vals = [v * scale for v in vals]
            largest, smallest = max(cm_vals), min(cm_vals)
            if smallest <= 0 or smallest < float(mr.get("min_smallest_cm", 1)):
                continue  # 短轴过小多为扁平斑块/细长结构，长短轴悬殊属正常
            if largest >= min_largest and largest / smallest >= ratio_th:
                issues.append({
                    "rule_id": "I", "check": "I3", "rule": "数据与单位异常", "severity": "high",
                    "field": field, "fragment": clip(m.group(0)),
                    "message": f"测量值“{m.group(0)}”径线比例约 {largest / smallest:.1f} 倍，"
                               f"量纲异常，疑似小数点或单位录入错误",
                    "suggestion": "请核对测量数值与单位（如 1.5*12cm 疑似应为 1.5*1.2cm）",
                })

        # I4 测量值未写单位（"大小约2.5*1.6。" 数值后无 cm/mm）
        for m in NUM_MEASURE_NO_UNIT_RE.finditer(text):
            issues.append({
                "rule_id": "I", "check": "I4", "rule": "数据与单位异常", "severity": "medium",
                "field": field, "fragment": clip(text[max(0, m.start() - 6):m.end() + 6]),
                "message": f"测量值“{m.group(0)}”缺少单位（应注明 mm 或 cm）",
                "suggestion": "请为测量值补充单位（径线通常以 mm/cm 记录）",
            })

        # I5 疑似丢失乘号的小数粘连（"1.512cm" 应为 1.5*1.2cm，排除 0.25/1.75 等规范小数）
        for m in NUM_UNIT_LOST_RE.finditer(text):
            issues.append({
                "rule_id": "I", "check": "I5", "rule": "数据与单位异常", "severity": "medium",
                "field": field, "fragment": clip(text[max(0, m.start() - 6):m.end() + 6]),
                "message": f"测量值“{m.group(0)}”疑似两个数值丢失乘号或分隔符",
                "suggestion": f"请核对测量值“{m.group(0)}”，疑似应为两个径线数值（如 1.5*1.2{m.group(0)[-2:] if m.group(0)[-2:] in ('cm', 'mm') else 'cm'}）",
            })

        # I6 测量范围倒置（"2.5-1.6cm" 起止值颠倒，应为 1.6-2.5cm）
        for m in RANGE_INVERT_RE.finditer(text):
            if float(m.group(1)) <= float(m.group(2)):
                continue  # 正常范围不报
            issues.append({
                "rule_id": "I", "check": "I6", "rule": "数据与单位异常", "severity": "medium",
                "field": field, "fragment": clip(text[max(0, m.start() - 6):m.end() + 6]),
                "message": f"测量范围“{m.group(0)}”起止值颠倒（{m.group(1)} 大于 {m.group(2)}）",
                "suggestion": f"请按从小到大顺序书写测量范围（如 {m.group(2)}-{m.group(1)}{m.group(3)}）",
            })
    return issues


def rule_j_redundancy(parsed: dict, rules: dict) -> list:
    """规则 J【语句重复】：叠词重复（"大小大小"）/ 整句模板重复冗余。"""
    rr = rules.get("redundancy_rules") or DEFAULT_RULES["redundancy_rules"]
    sim_th = float(rr.get("similarity_threshold", 0.78))
    min_len = int(rr.get("min_sentence_len", 10))
    issues = []

    for field in ("finding", "conclusion"):
        text = parsed.get(field, "")
        if not text:
            continue

        # J1 叠词重复（录入重复，"大小大小约" → "大小约"）
        for m in DUP_WORD_RE.finditer(text):
            if m.group(1) in DUP_WORD_ALLOW:
                continue
            issues.append({
                "rule_id": "J", "check": "J1", "rule": "语句重复", "severity": "medium",
                "field": field, "fragment": clip(text[max(0, m.start() - 6):m.end() + 6]),
                "message": f"检测到词语重复“{m.group(0)}”，疑似录入重复",
                "suggestion": f"请删除重复词语（将“{m.group(0)}”改为“{m.group(1)}”）",
            })

        # J2 模板文字未删除：整句重复 / 同一病灶描述短语重复出现
        # 2a 归一化完全重复的语句（同句出现≥2 次即报；同归一化 key 的记录数即出现次数）
        sents = [s.strip() for s in SENT_SPLIT_RE.split(text) if len(s.strip()) >= min_len]
        norm_map = {}
        for s in sents:
            key = SENT_NORM_RE.sub("", s)
            if len(key) < min_len:
                continue
            entry = norm_map.setdefault(key, {"sample": s, "nums": tuple(re.findall(r"\d+(?:\.\d+)?", s)), "count": 0})
            entry["count"] += 1
        for key, ent in norm_map.items():
            if ent["count"] < 2:
                continue
            issues.append({
                "rule_id": "J", "check": "J2", "rule": "语句重复", "severity": "medium",
                "field": field, "fragment": clip(ent["sample"]),
                "message": f"语句“{clip(ent['sample'])}”重复出现 {ent['count']} 次，疑似模板文字未删除",
                "suggestion": "请删除重复语句，保留一处准确描述",
            })
        # 2b 病灶描述重复的两种强特征（弱特征"短语前缀重复"误报过高已移除——
        #    多发斑块/两侧器官逐一描述模板句属正常书写，不判残留）：
        #    ① 肯定/否定矛盾（"探及肿大淋巴结"与"未见肿大淋巴结"并存，典型模板残留）
        #    ② 含未填测量值占位符的病灶描述（"见一约xmm不均质低回声团块"）
        lesion_any = LESION_NOUNS + ["淋巴结", "囊肿", "囊性", "积液", "结石", "斑块", "狭窄", "血栓"]
        occurs = []
        for w in sorted(set(lesion_any), key=len, reverse=True):
            for m in re.finditer(re.escape(w), text):
                if any(abs(m.start() - p) < 4 for p, _x in occurs):
                    continue  # 同一短语被多个病灶词命中，取一处
                occurs.append((m.start(), w))
        occurs.sort()
        for idx, (p1, w1) in enumerate(occurs):
            for p2, w2 in occurs[idx + 1:]:
                if p2 - p1 < 4 or w1 != w2:
                    continue  # 仅比对同一病灶词（跨词如"斑块/狭窄"同句并存属正常）
                seg1 = SENT_NORM_RE.sub("", text[p1:p1 + 32])
                seg2 = SENT_NORM_RE.sub("", text[p2:p2 + 32])
                ctx1, ctx2 = text[max(0, p1 - 8):p1], text[max(0, p2 - 8):p2]
                side1 = {s for s in "左右" if s in ctx1}
                side2 = {s for s in "左右" if s in ctx2}
                if side1 and side2 and not (side1 & side2):
                    continue  # 左右对称描述（各病灶一条）豁免
                neg1 = bool(re.search(r"未见|未探及|未", ctx1))
                neg2 = bool(re.search(r"未见|未探及|未", ctx2))
                # ① 肯定/否定矛盾：部位主语一致（如均为"双侧腋窝"）时才判模板残留，
                #    不同解剖部位的阳/阴性描述（颈总动脉有斑块 vs 椎动脉无斑块）属正常
                strong = False
                if neg1 != neg2:
                    site1 = SITE_PHRASE_RE.search(ctx1)
                    site2 = SITE_PHRASE_RE.search(ctx2)
                    if site1 and site2:
                        core1 = SITE_SIDE_PREFIX_RE.sub("", site1.group(1))
                        core2 = SITE_SIDE_PREFIX_RE.sub("", site2.group(1))
                        if core1 == core2 and core1:
                            win1 = SENT_NORM_RE.sub("", text[max(0, p1 - 10):p1 + 24])
                            win2 = SENT_NORM_RE.sub("", text[max(0, p2 - 10):p2 + 24])
                            mt = difflib.SequenceMatcher(None, win1, win2).find_longest_match(
                                0, len(win1), 0, len(win2))
                            strong = mt.size >= 5
                # ② 占位符残留（"约xmm"/"大小约*cm"）
                if not strong:
                    ph1 = bool(PLACEHOLDER_MEASURE_RE.search(text[max(0, p1 - 10):p1 + 32]))
                    ph2 = bool(PLACEHOLDER_MEASURE_RE.search(text[max(0, p2 - 10):p2 + 32]))
                    strong = ph1 or ph2
                if not strong:
                    continue
                issues.append({
                    "rule_id": "J", "check": "J2", "rule": "语句重复", "severity": "medium",
                    "field": field, "fragment": clip(text[p2:p2 + 24]),
                    "message": f"病灶描述“{seg2[:8]}”与前文“{clip(text[max(0, p1 - 6):p1 + 24])}”重复，"
                               f"疑似模板文字未删除",
                    "suggestion": "请合并重复的病灶描述，删除模板残留（如未填写的测量值占位描述）",
                })
                break

        # J3 模板残留碎片：孤立括号残片 / 与正文无关的残留字词（"(不完整，)"、"未了。"）
        max_len = int(rules.get("stale_max_len", 8))
        stale_words = [w for w in (rules.get("stale_fragments") or []) if w]
        for m in STALE_PAREN_RE.finditer(text):
            inner = m.group(1).strip()
            if len(inner) > max_len:
                continue  # 较长括注为正常补充说明
            if not inner or not (STALE_FRAG_HINT_RE.search(inner)
                                or any(w in inner for w in stale_words)):
                continue  # 仅当括号内为残片/未填标记时才判残留
            issues.append({
                "rule_id": "J", "check": "J3", "rule": "语句重复", "severity": "medium",
                "field": field, "fragment": clip(text[max(0, m.start() - 10):m.end() + 6]),
                "message": f"检出模板残留碎片“{m.group(0)}”，括号内容未填写/与正文无关",
                "suggestion": f"请删除模板残留片段“{m.group(0)}”，补全或删除未完成内容",
            })
        for w in stale_words:
            for m in re.finditer(re.escape(w), text):
                ctx = text[max(0, m.start() - 4):m.end() + 4]
                if re.search(r"[（(][^）)]*" + re.escape(w), text[max(0, m.start() - 10):m.end() + 4]):
                    continue  # 位于括号内（如"(不完整，)"），交给括号残留规则报一条
                issues.append({
                    "rule_id": "J", "check": "J3", "rule": "语句重复", "severity": "medium",
                    "field": field, "fragment": clip(ctx),
                    "message": f"检出疑似模板残留字词“{w}”，与正文无关或未填写完成",
                    "suggestion": f"请核对并删除残留字词“{w}”（模板未删干净或未填写完整）",
                })
                break
    return issues


# ============================ 检测入口 ============================

def check_parsed(parsed: dict, rules: dict) -> dict:
    """对已解析的报告字段执行全部规则（A-J），返回结构化结果。"""
    t0 = time.perf_counter()
    issues = []
    issues += rule_a_logical(parsed["finding"], parsed["conclusion"], rules)
    a_fired = any(i["rule_id"] == "A" for i in issues)
    issues += rule_b_side(parsed["finding"], parsed["conclusion"], rules)
    issues += rule_c_typos(parsed.get("exam_item", "") + "\n" + parsed["finding"] + "\n" + parsed["conclusion"], rules)
    issues += rule_c_field_missing(parsed)
    issues += rule_d_grading(parsed, rules)
    if not a_fired:  # A 与 E 同向（结论遗漏），已报 A 则不重复提示
        issues += rule_e_coverage(parsed, rules)
    issues += rule_f_completeness(parsed, rules)
    issues += rule_g_gender_conflict(parsed, rules)
    issues += rule_g_age_conflict(parsed, rules)
    issues += rule_h_exam_coverage(parsed, rules)
    issues += rule_i_measurement(parsed, rules)
    issues += rule_j_redundancy(parsed, rules)
    # 统一归口：规则 ID 即错误类型名（9 类，名称不可改）；check 保留为内部子项标记
    for iss in issues:
        et = CHECK_ERROR_TYPE.get(iss.get("check") or iss["rule_id"], ERR_OTHER)
        iss["error_type"] = et
        iss["rule_id"] = et
        iss["rule"] = et
    elapsed = round((time.perf_counter() - t0) * 1000, 2)
    return {
        "engine_version": ENGINE_VERSION,
        "input_format": parsed.get("input_format", "text"),
        "report": {k: parsed[k] for k in ("exam_item", "name", "gender", "age")},
        "parsed": {"finding": parsed["finding"], "conclusion": parsed["conclusion"],
                   "parse_notes": parsed["parse_notes"]},
        "issues": issues,
        "issue_count": len(issues),
        "elapsed_ms": elapsed,
    }


def check_report(text: str, rules: dict) -> dict:
    """对单份报告文本执行全部规则，返回结构化结果（文本自适应解析）。"""
    return check_parsed(parse_report(text), rules)


def check_json(payload: dict, rules: dict) -> dict:
    """对标准 JSON 报告对象执行全部规则，返回结构化结果（工作站集成入口）。"""
    return check_parsed(parse_json_report(payload), rules)


# ============================ CLI: 批量检测 + Excel 导出 ============================

DETAIL_HEADERS = ["文件名/报告ID", "是否存在问题", "错误类型", "位置", "原文片段", "问题描述", "修改建议"]


def collect_input_files(inputs: list) -> list:
    files = []
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            files += sorted(f for f in p.iterdir() if f.is_file() and f.suffix.lower() == ".txt")
        elif p.is_file():
            files.append(p)
        else:
            print(f"[警告] 路径不存在，已跳过: {item}")
    return list(dict.fromkeys(files))


def export_excel(results: list, out_path: str, rules_source: str) -> None:
    detail_rows, type_counter = [], {}
    ok_count = 0
    for r in results:
        name = r["file"]
        if r["result"]["issue_count"] == 0:
            ok_count += 1
            detail_rows.append([name, "否"] + [""] * 5)
            continue
        for iss in r["result"]["issues"]:
            et = iss.get("error_type") or ERR_OTHER
            type_counter[et] = type_counter.get(et, 0) + 1
            detail_rows.append([name, "是", et,
                                iss["field"], iss["fragment"], iss["message"], iss["suggestion"]])
    summary = [
        ["指标", "数值"],
        ["生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["引擎版本", ENGINE_VERSION],
        ["规则来源", rules_source],
        ["报告文件数", len(results)],
        ["存在问题报告数", len(results) - ok_count],
        ["问题总数", sum(type_counter.values())],
    ] + [["错误类型", "错误数量（条）"]] + [[et, type_counter.get(et, 0)]
                                        for et in ERROR_TYPE_ORDER if type_counter.get(et, 0)]
    MiniExcel.write(out_path, [("质控概览", summary[0], summary[1:]), ("错误明细", DETAIL_HEADERS, detail_rows)])


def cmd_check(args) -> int:
    rules, source = load_rules(args.rules)
    compile_rules(rules)
    print(f"[引擎] 规则来源: {source}")
    files = collect_input_files(args.inputs)
    if not files:
        print("[错误] 未找到任何 .txt 报告文件")
        return 2
    results = []
    for f in files:
        try:
            text = read_text_any(f.read_bytes())
            res = check_report(text, rules)
        except OSError as exc:
            print(f"[警告] 读取失败 {f.name}: {exc}")
            res = {"report": {}, "parsed": {}, "issues": [], "issue_count": 0, "elapsed_ms": 0}
            res = {"file": f.name, "result": res, "error": str(exc)}
            results.append(res)
            continue
        entry = {"file": f.name, "result": res}
        results.append(entry)
        if args.json:
            print(json.dumps(entry, ensure_ascii=False, indent=2))
    export_excel(results, args.output, source)

    bad = sum(1 for r in results if r["result"]["issue_count"] > 0)
    total_issues = sum(r["result"]["issue_count"] for r in results)
    avg_ms = round(sum(r["result"]["elapsed_ms"] for r in results) / max(len(results), 1), 2)
    print("=" * 56)
    print(f"离线质控完成: 报告 {len(results)} 份 | 有问题 {bad} 份 | 问题 {total_issues} 条")
    print(f"单份平均耗时 {avg_ms} ms | 结果已导出: {Path(args.output).resolve()}")
    return 0


# ============================ HTTP API 模式 ============================

class QCRequestHandler(BaseHTTPRequestHandler):
    server_version = f"UltrasoundQCEngine/{ENGINE_VERSION}"
    protocol_version = "HTTP/1.1"

    def _send_json(self, obj: dict, status: int = 200) -> None:
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):  # 静默默认访问日志
        pass

    def _authorized(self) -> bool:
        token = getattr(self.server, "auth_token", None)
        return token is None or self.headers.get("X-Auth-Token") == token

    # ---- GET: /health /api/rules / 使用说明页 ----
    def do_GET(self):
        if not self._authorized():
            return self._send_json({"ok": False, "error": "未授权：缺少或错误的 X-Auth-Token"}, 401)
        path = self.path.split("?")[0]
        if path in ("/health", "/api/health"):
            return self._send_json({"ok": True, "engine_version": ENGINE_VERSION,
                                    "rules_source": self.server.rules_source,
                                    "rules_file": str(self.server.rules_file)})
        if path == "/api/rules":
            r = self.server.rules
            return self._send_json({"ok": True,
                                    "positive_keywords": r["positive_keywords"],
                                    "side_keywords": r["side_keywords"],
                                    "correction_mapping": r["correction_mapping"],
                                    "negative_conclusion_phrases": r["_negative_phrases"],
                                    "grading_rules": r["grading_rules"],
                                    "gender_conflict": r.get("gender_conflict"),
                                    "exam_organ_map": r.get("exam_organ_map")})
        if path == "/":
            html = ("<html><body style='font-family:Microsoft YaHei'>"
                    "<h3>Ultrasound QC Engine</h3>"
                    "<p>POST /api/check &nbsp; body: 标准 JSON "
                    "{\"exam_item\":\"检查项目\",\"gender\":\"性别\",\"age\":\"年龄\","
                    "\"finding\":\"超声所见\",\"conclusion\":\"超声提示\"}</p>"
                    "<p>POST /api/check/batch &nbsp; body: {\"reports\":[标准JSON...]} "
                    "或 {\"texts\":[报告全文...]}</p>"
                    "<p>兼容旧版 {\"text\": \"报告全文\"} 或纯文本提交</p>"
                    "<p>GET /health &nbsp;|&nbsp; GET /api/rules</p></body></html>").encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(html)))
            self.end_headers()
            self.wfile.write(html)
            return
        self._send_json({"ok": False, "error": "未知路径，可用: /health /api/rules /api/check"}, 404)

    # ---- POST: /api/check /api/check/batch ----
    def do_POST(self):
        if not self._authorized():
            return self._send_json({"ok": False, "error": "未授权：缺少或错误的 X-Auth-Token"}, 401)
        path = self.path.split("?")[0]
        if path not in ("/api/check", "/api/check/batch", "/check"):
            return self._send_json({"ok": False, "error": "未知路径"}, 404)
        try:
            length = int(self.headers.get("Content-Length", 0))
        except ValueError:
            return self._send_json({"ok": False, "error": "Content-Length 非法"}, 400)
        if length > MAX_BODY_BYTES:
            return self._send_json({"ok": False, "error": "请求体超过 8MB 上限"}, 413)
        raw = self.rfile.read(length) if length else b""
        text_all = read_text_any(raw)

        payload = None
        stripped = text_all.strip()
        if stripped.startswith("{") and stripped.endswith("}"):
            try:
                payload = json.loads(stripped)
            except json.JSONDecodeError:
                payload = None

        # 批量模式：{"reports": [标准JSON对象...]} 或 {"texts": ["报告全文"...]}（兼容旧版）
        if path.endswith("/batch") or isinstance(payload, dict) and ("reports" in payload or "texts" in payload):
            payload = payload or {}
            if payload.get("reports") is not None:
                reports = payload.get("reports")
                if not isinstance(reports, list) or not reports or not all(isinstance(r, dict) for r in reports):
                    return self._send_json({"ok": False, "error": "批量模式 reports 需为非空 JSON 对象数组"}, 400)
                results = [check_json(r, self.server.rules) for r in reports]
            else:
                texts = payload.get("texts") or []
                if not isinstance(texts, list) or not texts:
                    return self._send_json({"ok": False, "error": "批量模式需提供非空 texts 数组"}, 400)
                results = [check_report(t, self.server.rules) for t in texts]
            return self._send_json({"ok": True, "count": len(results), "results": results})

        # 标准 JSON 报告对象（工作站集成规范）：
        # {"exam_item":..,"gender":..,"age":..,"finding":..,"conclusion":..}
        if isinstance(payload, dict):
            wrapped = payload.get("report") if isinstance(payload.get("report"), dict) else payload
            if any(str(k).strip().lower() in JSON_KEY_SET for k in wrapped):
                result = check_json(wrapped, self.server.rules)
                result["ok"] = True
                result["rules_source"] = self.server.rules_source
                return self._send_json(result)

        # 兼容旧版：{"text": "报告全文"} / {"raw": ...} / 纯文本
        text = (payload or {}).get("text") or (payload or {}).get("raw") or text_all
        if not text.strip():
            return self._send_json({"ok": False,
                                    "error": "报告内容为空（标准 JSON: {\"exam_item\":..,\"gender\":..,\"age\":..,"
                                             "\"finding\":..,\"conclusion\":..}，或 {\"text\":..}）"}, 400)
        result = check_report(text, self.server.rules)
        result["ok"] = True
        result["rules_source"] = self.server.rules_source
        return self._send_json(result)


def cmd_serve(args) -> int:
    rules, source = load_rules(args.rules)
    compile_rules(rules)
    server = ThreadingHTTPServer((args.host, args.port), QCRequestHandler)
    server.daemon_threads = True
    server.rules = rules
    server.rules_source = source
    server.rules_file = _exe_dir() / "rules.json"
    server.auth_token = args.token
    print("=" * 56)
    print(f"Ultrasound QC Engine v{ENGINE_VERSION} (HTTP 服务模式)")
    print(f"监听地址   : http://{args.host}:{args.port}")
    print(f"规则来源   : {source}")
    print(f"接口列表   : POST /api/check | POST /api/check/batch | GET /health | GET /api/rules")
    print("按 Ctrl+C 停止服务")
    print("=" * 56)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[服务] 已停止")
    return 0


# ============================ 内置自检 ============================

_SELFTEST_CASES = [
    {
        "name": "用例1-结论遗漏病变→描述与结论不一致",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：36岁\n"
                 "超声所见：左侧叶甲状腺内见低回声结节，边界欠清，内见强光点。\n"
                 "超声提示：甲状腺未见明显异常。\n"),
        "expect_rules": {"描述与结论不一致"},
    },
    {
        "name": "用例2-左右侧不符+用词勘误+分级不一致",
        "text": ("检查项目：双侧乳腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：左侧乳腺低回声结节，边界不清，内见微钙化，回音欠均匀，血流较丰富。\n"
                 "超声提示：右侧乳腺结节，BI-RADS 2类。\n"),
        "expect_rules": {"描述与结论不一致", "错别字或字段缺失", "分级不规范"},
    },
    {
        "name": "用例3-英文标签格式→描述与结论不一致",
        "text": ("examItem: 甲状腺超声\ngender: 女\nage: 36Y\n"
                 "finding: 左侧叶甲状腺内见低回声结节，边界欠清。\n"
                 "conclusion: 甲状腺未见明显异常。\n"
                 "知情同意：本检查不能排除所有胎儿染色体异常的可能，未见明显异常胎儿。\n"),
        "expect_rules": {"描述与结论不一致"},
    },
    {
        "name": "用例4-结论未涵盖所见→描述与结论不一致",
        "text": ("examItem: 颈动脉彩超\ngender: 男\nage: 65Y\n"
                 "finding: 右侧颈总动脉起始处后壁可见大小约1.4×0.25cm不均质回声斑块，管腔未见明显狭窄。\n"
                 "conclusion: 右侧颈总动脉血流通畅。\n"),
        "expect_rules": {"描述与结论不一致"},
    },
    {
        "name": "用例5-建议缺时间数值→语病或表述不规范",
        "text": ("gender: 女\nage: 28Y\n"
                 "finding: 子宫前位，大小正常，肌层回声均匀，双附件区未见明显异常。\n"
                 "conclusion: 子宫双附件未见明显异常，建议周后复查。\n"),
        "expect_rules": {"语病或表述不规范"},
    },
    {
        "name": "用例6-性别与检查项目冲突→性别年龄冲突(JSON输入)",
        "json": {
            "exam_item": "前列腺彩超", "gender": "女", "age": "55",
            "finding": "前列腺大小约3.9*2.5*2.9cm，实质回声均匀，内未见明显异常回声。",
            "conclusion": "前列腺未见明显异常。",
        },
        "expect_rules": {"性别年龄冲突"},
    },
    {
        "name": "用例7-所见遗漏检查部位(JSON输入)",
        "json": {
            "exam_item": "肝胆胰脾彩超", "gender": "男", "age": "50",
            "finding": "肝脏形态大小正常，包膜光滑，实质回声均匀，肝内管道走行清晰。",
            "conclusion": "肝胆胰脾未见明显异常。",
        },
        "expect_rules": {"遗漏检查部位"},
    },
    {
        "name": "用例8-标准JSON中文键名+无冲突不误报",
        "json": {
            "检查项目": "肝胆胰脾彩超", "性别": "男", "年龄": "45",
            "超声所见": "肝脏形态正常，胆囊壁不厚，胰腺回声均匀，脾脏大小正常。",
            "超声提示": "肝胆胰脾未见明显异常。",
        },
        "expect_rules": set(),
        "expect_exact": True,
    },
    {
        "name": "用例9-测量值单位/量纲异常→数据或单位错误",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺左叶较大一枚大小约1.5*12cm，右叶较大的一枚大小约2.5*1.6m。\n"
                 "超声提示：甲状腺多发结节（C-TIRADS 3类）。\n"),
        "expect_rules": {"数据或单位错误"},
    },
    {
        "name": "用例10-叠词重复→语病或表述不规范",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶见低回声结节，大小大小约2.5*1.6cm，边界清晰。\n"
                 "超声提示：甲状腺右叶结节（C-TIRADS 3类）。\n"),
        "expect_rules": {"语病或表述不规范"},
    },
    {
        "name": "用例11-模板矛盾描述→模板文字未删除",
        "text": ("检查项目：乳腺超声\n性别：女 年龄：50岁\n"
                 "超声所见：双侧腋窝探及肿大淋巴结回声，大小约2.8*1.0cm。"
                 "双侧腋窝未见明显肿大淋巴结回声。\n"
                 "超声提示：双侧腋窝淋巴结可见。\n"),
        "expect_rules": {"模板文字未删除"},
    },
    {
        "name": "用例12-检查项目与内容不符→其他",
        "text": ("检查项目：彩色多普勒超声常规检查（妇科）\n性别：女 年龄：30岁\n"
                 "超声所见：宫内见一孕囊，大小约2.0*1.5cm，可见卵黄囊。\n"
                 "超声提示：早孕。\n"),
        "expect_rules": {"其他"},
    },
    {
        "name": "用例13-年龄与检查内容冲突(G2)→性别年龄冲突",
        "text": ("检查项目：彩色多普勒超声常规检查（产科）\n性别：女 年龄：79岁\n"
                 "超声所见：胎儿头位于耻骨联合上方，双顶径约8.6cm，羊水最大深径5.2cm。\n"
                 "超声提示：单胎 晚孕 头位。\n"),
        "expect_rules": {"性别年龄冲突"},
    },
    {
        "name": "用例14-模板残留碎片(J3)→模板文字未删除",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶见低回声结节，边界清晰(不完整，)。CDFI：结节内见血流信号。未了。\n"
                 "超声提示：甲状腺右叶结节（C-TIRADS 3类）。\n"),
        "expect_rules": {"模板文字未删除"},
    },
    {
        "name": "用例15-同一结论多分级(D1)→分级不规范",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶低回声结节，边界欠清，内见微钙化。\n"
                 "超声提示：甲状腺右叶结节（C-TIRADS 3类），甲状腺右叶结节（C-TIRADS 4a类）。\n"),
        "expect_rules": {"分级不规范"},
    },
    {
        "name": "用例16-恶性征象分级偏低(D2)→分级不规范",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶低回声结节，边界不清，内见微钙化，血流丰富。\n"
                 "超声提示：甲状腺右叶结节（C-TIRADS 3类）。\n"),
        "expect_rules": {"分级不规范"},
    },
    {
        "name": "用例17-测量值缺单位(I4)→数据或单位错误",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶结节大小约2.5*1.6。\n超声提示：甲状腺右叶结节。\n"),
        "expect_rules": {"数据或单位错误"},
    },
    {
        "name": "用例18-丢失乘号(I5)→数据或单位错误",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶结节大小约1.512cm。\n超声提示：甲状腺右叶结节。\n"),
        "expect_rules": {"数据或单位错误"},
    },
    {
        "name": "用例19-整句重复(J2)→模板文字未删除",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶见低回声结节，大小约2.5*1.6cm，边界清晰。"
                 "甲状腺右叶见低回声结节，大小约2.5*1.6cm，边界清晰。\n超声提示：甲状腺右叶结节。\n"),
        "expect_rules": {"模板文字未删除"},
    },
    {
        "name": "用例20-正常带单位测量值不误报(I4 误报回归)",
        "text": ("检查项目：双侧乳腺超声\n性别：女 年龄：40岁\n"
                 "超声所见：右侧乳腺见低回声结节，大小约8.2*5.1mm。左侧乳腺见低回声结节，大小约8.3*5.0mm。\n"
                 "超声提示：双侧乳腺结节（BI-RADS 3类）。\n"),
        "expect_rules": set(),
    },
    {
        "name": "用例21-字段缺失无结论(C2)→错别字或字段缺失",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶低回声结节，大小约2.5*1.6cm。\n"),
        "expect_rules": {"错别字或字段缺失"},
    },
    {
        "name": "用例22-标点重复与空括号(F5/F6)→语病或表述不规范",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶见低回声结节（），边界清晰。。CDFI：结节内见血流信号。\n"
                 "超声提示：甲状腺右叶结节。\n"),
        "expect_rules": {"语病或表述不规范"},
    },
    {
        "name": "用例23-所见末缺句末标点(F7)→语病或表述不规范",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶见低回声结节，边界清晰\n超声提示：甲状腺右叶结节。\n"),
        "expect_rules": {"语病或表述不规范"},
    },
    {
        "name": "用例24-测量范围倒置(I6)→数据或单位错误",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶结节大小约2.5-1.6cm，边界清晰。\n超声提示：甲状腺右叶结节。\n"),
        "expect_rules": {"数据或单位错误"},
    },
    {
        "name": "用例25-错别字扩充词表→错别字或字段缺失",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：右则甲状腺见低回省结节，双侧颈部未见肿大淋巴洁。\n超声提示：甲状腺右叶结节。\n"),
        "expect_rules": {"错别字或字段缺失"},
    },
    {
        "name": "用例26-正常括注与句末不误报(误报回归)",
        "text": ("检查项目：甲状腺超声\n性别：女 年龄：45岁\n"
                 "超声所见：甲状腺右叶低回声结节（考虑良性，建议定期复查），边界清晰。\n"
                 "超声提示：甲状腺右叶结节（C-TIRADS 3类）。\n"),
        "expect_rules": set(),
    },
]


def cmd_selftest(_args) -> int:
    rules, source = load_rules(None)
    compile_rules(rules)
    print(f"[自检] 规则来源: {source}")
    all_pass = True
    for case in _SELFTEST_CASES:
        if "json" in case:
            res = check_json(case["json"], rules)
        else:
            res = check_report(case["text"], rules)
        fired = {i["rule_id"] for i in res["issues"]}
        if case.get("expect_exact"):
            passed = fired == case["expect_rules"]
        else:
            passed = case["expect_rules"] <= fired
        all_pass &= passed
        status = "PASS" if passed else "FAIL"
        print(f"[{status}] {case['name']} | 触发规则: {sorted(fired) or '无'} "
              f"| 预期: {sorted(case['expect_rules'])} | 耗时 {res['elapsed_ms']} ms")
        for iss in res["issues"]:
            print(f"        - [{iss['rule_id']}/{iss['severity']}] {iss['message']}")
        if not passed:
            print(json.dumps(res, ensure_ascii=False, indent=2))
    print("[自检] 全部通过" if all_pass else "[自检] 存在失败用例！")
    return 0 if all_pass else 1


# ============================ 主入口 ============================

def main() -> None:
    _ensure_stdio()  # 无控制台 exe 兼容
    parser = argparse.ArgumentParser(
        prog="UltrasoundQCEngine",
        description="离线超声报告自动勘误与质控引擎（零依赖 / 毫秒级 / 可打包单文件 exe）")
    sub = parser.add_subparsers(dest="mode", required=True)

    def add_common(p):
        p.add_argument("--rules", default=None, help="rules.json 路径（默认取程序同目录）")

    p_check = sub.add_parser("check", help="批量检测报告文件并导出 Excel")
    add_common(p_check)
    p_check.add_argument("inputs", nargs="+", help="报告 .txt 文件或目录（可多个）")
    p_check.add_argument("-o", "--output", default="qc_result.xlsx", help="输出 Excel 路径")
    p_check.add_argument("--json", action="store_true", help="同时在控制台输出每份 JSON 结果")

    p_serve = sub.add_parser("serve", help="启动 HTTP API 服务（供 PACS 实时调用）")
    add_common(p_serve)
    p_serve.add_argument("--host", default="127.0.0.1", help="监听地址（对接外部 PACS 用 0.0.0.0）")
    p_serve.add_argument("--port", type=int, default=os.environ.get("QC_PORT", 8976), help="监听端口")
    p_serve.add_argument("--token", default=None, help="简单鉴权 token（设置后请求需带 X-Auth-Token 头）")

    sub.add_parser("selftest", help="运行内置自检用例，验证规则引擎")
    sub.add_parser("version", help="显示版本")

    args = parser.parse_args()
    if args.mode == "check":
        sys.exit(cmd_check(args))
    if args.mode == "serve":
        sys.exit(cmd_serve(args))
    if args.mode == "selftest":
        sys.exit(cmd_selftest(args))
    print(f"Ultrasound QC Engine v{ENGINE_VERSION}")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
债券询价双表合并工具 v1

功能：
- 输入两张询价导出表（品种一 + 品种二），输出一张合并询价明细表
- 按机构名称+业务种类匹配，同一机构在两张表中均出现则合并为一行
- 备注列标注品种来源（品种一 / 品种二 / 品种一/品种二）
- 不同品种使用不同背景色，合并行黄色高亮 B/C/G 列
- 询价进度不一致时智能合并（相似保留一条，不相似换行拼接）
- 无小结区

用法：
  python bond_merge_processor.py "品种一.xls" "品种二.xls"
  python bond_merge_processor.py "品种一.xls" "品种二.xls" -o "合并结果.xlsx"

依赖：pandas, openpyxl, xlrd
"""

import argparse
import re
import sys
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, Color
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import IconSetRule


# ===================== 配置 =====================

OUTPUT_COLUMNS = [
    "分类",
    "机构名称",
    "业务种类",
    "意向投标量（亿元）",
    "询价进度",
    "认购意向（评分）",
    "备注",
]

FORMAL_CATEGORY_ORDER = [
    "有可能参与的银行",
    "有可能参与的非银",
    "无法参与的银行",
    "无法参与的非银",
]

PROBABLE_PHRASE_ORDER = [
    "认购意向高",
    "当天报量",
    "推介中",
    "观望中",
    "认购意向低",
    "在库待沟通",
    "认购意向不明",
]

UNABLE_PHRASE_ORDER = [
    "不参与",
    "价格原因",
    "资质不符",
    "额度受限",
    "不在库",
    "其他",
    "流程/时间受限",
    "利率原因",
]

BIZ_TYPE_ORDER = ["自营", "资管", "理财", "专户", "投顾", "公募"]

COLUMN_ALIASES = {
    "机构名称": ["机构名称", "机构", "机构全称", "客户名称", "投资人名称"],
    "机构类型": ["机构类型", "机构类别", "客户类型", "投资人类型", "类型"],
    "业务种类": ["业务种类", "业务类型", "产品类型", "账户类型", "业务"],
    "审批进度": ["审批进度", "项目进度", "信评进度", "审批状态", "状态"],
    "意向投标量（亿元）": [
        "意向投标量（亿元）", "意向投标量(亿元)",
        "意向投标量（亿）", "意向投标量(亿)",
        "意向投标量（万元）", "意向投标量(万元)", "意向投标量（万）", "意向投标量(万)",
        "意向投标量", "投标量", "意向量", "拟投量", "认购量", "参与量", "金额", "认购金额", "投标金额"
    ],
    "询价进度": ["询价进度", "沟通细节", "反馈内容", "进度", "反馈", "备注", "跟进情况"],
    "更新时间": ["更新时间", "更新日期", "最后更新时间", "日期", "时间"],
}

# 品种一 备注列背景色 (light blue: theme=4, tint=0.8 → approximate RGB)
# 品种二 备注列背景色 (light orange: theme=6, tint=0.8 → approximate RGB)
# 合并行黄色高亮
FILL_VARIETY_1 = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")   # 浅蓝
FILL_VARIETY_2 = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")   # 浅橙
FILL_MERGED_YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

WRAP_ALIGN = Alignment(wrap_text=True, vertical="top")
CENTER_ALIGN = Alignment(wrap_text=True, vertical="top", horizontal="center")

FONT_DEFAULT = Font(name="微软雅黑", size=10)
FONT_HEADER = Font(name="微软雅黑", size=10, bold=True)


# ===================== 工具函数 =====================

def clean_text(value: Any) -> str:
    if pd.isna(value) or value is None:
        return ""
    if isinstance(value, (int, float)):
        if value == int(value):
            return str(int(value))
        return str(value).rstrip("0").rstrip(".")
    return str(value).strip()


def read_excel_smart(path: str) -> pd.DataFrame:
    """智能读取 Excel，自动跳过合并标题行找到真正的表头"""
    ext = Path(path).suffix.lower()

    # 第一步：用 header=None 读原始数据，找表头行
    if ext == ".xls":
        raw = pd.read_excel(path, header=None, dtype=object, engine="xlrd")
    else:
        raw = pd.read_excel(path, header=None, dtype=object, engine="openpyxl")

    # 收集所有别名用于匹配表头
    alias_pool = set()
    for aliases in COLUMN_ALIASES.values():
        for item in aliases:
            alias_pool.add(item.replace(" ", "").replace("\n", ""))

    best_row = 0
    best_score = -1
    for idx in range(min(len(raw), 40)):
        row_values = [str(x).replace(" ", "").replace("\n", "") if pd.notna(x) else "" for x in raw.iloc[idx].tolist()]
        non_empty = sum(1 for x in row_values if x and x.lower() != "nan")
        alias_hits = sum(1 for x in row_values if x in alias_pool)
        score = alias_hits * 2 + non_empty
        if score > best_score:
            best_score = score
            best_row = idx

    # 第二步：用找到的表头行重新读取
    if ext == ".xls":
        df = pd.read_excel(path, header=best_row, dtype=object, engine="xlrd")
    else:
        df = pd.read_excel(path, header=best_row, dtype=object, engine="openpyxl")

    df = df.dropna(how="all")
    # 去掉 Unnamed 列
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
    df.columns = [str(c).strip() for c in df.columns]
    return df


def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for target, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in df.columns:
                mapping[alias] = target
                break
    if mapping:
        df = df.rename(columns=mapping)
    return df


def parse_amount(value: Any) -> str:
    if pd.isna(value) or value is None:
        return ""
    text = clean_text(value)
    if text == "" or text in ("0", "0.0", "0.00"):
        return ""
    try:
        num = float(text)
        if num <= 0:
            return ""
        if num >= 1:
            return str(int(num)) if num == int(num) else f"{num:.2f}"
        else:
            return f"{num:.2f}"
    except ValueError:
        return text


def has_valid_amount(value: Any) -> bool:
    if pd.isna(value):
        return False
    text = clean_text(value)
    if text == "":
        return False
    try:
        return float(text) > 0
    except Exception:
        return bool(re.search(r"\d", text))


def parse_mmdd(value: Any, fallback_text: str = "") -> str:
    if not pd.isna(value) and clean_text(value):
        if isinstance(value, (pd.Timestamp, datetime)):
            return value.strftime("%m%d")
        text = clean_text(value)
        dt = pd.to_datetime(text, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%m%d")
        m = re.search(r"(?:20\d{2}[-/.年])?(\d{1,2})[-/.月](\d{1,2})", text)
        if m:
            return f"{int(m.group(1)):02d}{int(m.group(2)):02d}"
        m = re.search(r"(?<!\d)(\d{4})(?!\d)", text)
        if m:
            return m.group(1)
    marks = re.findall(r"【(\d{4})】", fallback_text or "")
    if marks:
        return marks[-1]
    return ""


def mmdd_sort_key(mmdd: str) -> int:
    if not mmdd or not re.match(r"^\d{4}$", mmdd):
        return 0
    return int(mmdd)


def is_project_failed(approval_progress: Any) -> bool:
    return clean_text(approval_progress) == "项目未通过"


def is_bank(row: pd.Series) -> bool:
    org_type = clean_text(row.get("机构类型", ""))
    org_name = clean_text(row.get("机构名称", ""))
    bank_keywords = ["银行", "农商", "农信", "信用社", "村镇银行", "农合"]
    return any(k in org_type for k in bank_keywords) or any(k in org_name for k in bank_keywords)


def has_explicit_amount_or_ratio(text: str) -> bool:
    lower_text = text.lower()
    patterns = [
        r"\d+(\.\d+)?\s*(e|亿|亿元)",
        r"\d+(\.\d+)?\s*(kw|k|w|万)",
        r"\d+(\.\d+)?\s*%",
        r"参与\s*\d+",
        r"预计\s*\d+",
        r"暂定\s*\d+",
        r"申请.*\d+.*额度",
        r"上报.*\d+(\.\d+)?\s*(亿|万|e|kw|k|w|%)",   # 上报量需带单位，避免"5Y"误匹配
        r"按\s*\d+.*上报",
        r"报\s*\d+(\.\d+)?\s*(亿|万|e|kw|k|w|%)",
    ]
    return any(re.search(p, lower_text, flags=re.IGNORECASE) for p in patterns)


# ===================== 分类与评分 =====================

def classify_probable_phrase(text: str, amount_value: Any) -> str:
    t = clean_text(text)
    if t == "":
        return "认购意向不明"
    if (
        has_valid_amount(amount_value)
        or has_explicit_amount_or_ratio(t)
        or re.search(r"认购意向高|认购意愿高|参与意向高|会参与|可能会参与|可以参与", t)
    ):
        return "认购意向高"
    if re.search(r"当天报量|当天喊", t):
        return "当天报量"
    if re.search(r"推介|推荐|沟通|推动|同步|准入|信评中|审核中|报授信|上会|分行推动|支行|上报", t):
        return "推介中"
    if re.search(r"当天看|当天定|当天来|看看|关注|看价格|看需求|看经理|簿记当天|再看|待定|确认量|可以看|价格合适", t):
        return "观望中"
    if re.search(r"兴趣不大|看得少|意向低|意愿低|参与意向低|可能性低|期限偏短|性价比低|规模小|价格太低|买不起|跟不上", t):
        return "认购意向低"
    if "在库" in t:
        return "在库待沟通"
    return "在库待沟通"


def classify_unable_phrase(text: str) -> str:
    t = clean_text(text)
    if t == "":
        return "其他"
    if "不在库" in t:
        return "不在库"
    if re.search(r"价格|买不起|跟不上|收益|成本|利差|区间.*低", t):
        return "价格原因"
    if re.search(r"信评过不了|信评不过|没授信|无授信|只能看AAA|不看CP|没有担保|主体不符|资质|不看交易所|主体评级|无评级", t):
        return "资质不符"
    if re.search(r"额度|无额度|没有额度|额度不足|批复失效", t):
        return "额度受限"
    if re.search(r"分行没有报|分行没报|未上报|没有报|来不及|未提供|联系人|流程|时间", t):
        return "流程/时间受限"
    if "利率" in t:
        return "利率原因"
    if re.search(r"不参与|不看|投不了|不投|不能参与|不考虑|无法参与|不能投", t):
        return "不参与"
    return "其他"


def build_updated_progress_parts(mmdd: str, phrase: str, inquiry_progress: str) -> List[Tuple[str, bool]]:
    """生成"更新询价进度"的分段文本（对齐v6工具）。

    规则：
    1. 日期不加粗，分类话术加粗，后续原始进度不加粗。
    2. 避免"认购意向低：认购意向低"这类重复。
    3. 询价进度为空时不加逗号。
    4. 所有冒号替换为逗号。
    """
    t = clean_text(inquiry_progress).replace("：", "，").replace(":", "，")
    prefix = f"【{mmdd}】" if mmdd else ""

    if t == "":
        return [(prefix, False), (phrase, True)]

    if t == phrase:
        return [(prefix, False), (phrase, True)]

    duplicate_prefixes = [
        phrase + "，", phrase + ",",
        phrase + "、", phrase + " ",
    ]
    for pfx in duplicate_prefixes:
        if t.startswith(pfx):
            rest = t[len(phrase):]
            return [(prefix, False), (phrase, True), (rest, False)]

    return [(prefix, False), (phrase, True), ("，", False), (t, False)]


def parts_to_excel_rich_text(parts: List[Tuple[str, bool]], font_name: str = "楷体", font_size: int = 12):
    """将分段文本写成 Excel 富文本（对齐v6工具）。"""
    rich_text = CellRichText()
    for text, bold in parts:
        if not text:
            continue
        rich_text.append(
            TextBlock(
                InlineFont(rFont=font_name, sz=font_size, b=bold),
                text,
            )
        )
    return rich_text


def build_updated_progress(mmdd: str, phrase: str, inquiry_progress: str) -> str:
    """构建更新询价进度纯文本（对齐v6工具）。"""
    parts = build_updated_progress_parts(mmdd, phrase, inquiry_progress)
    return "".join(text for text, _ in parts if text)


def classify_row(row: pd.Series) -> Dict[str, Any]:
    org_name = clean_text(row.get("机构名称", ""))
    business_type = clean_text(row.get("业务种类", ""))
    approval_progress = clean_text(row.get("审批进度", ""))
    inquiry_progress = clean_text(row.get("询价进度", ""))
    amount = parse_amount(row.get("意向投标量（亿元）", ""))
    mmdd = parse_mmdd(row.get("更新时间", ""), inquiry_progress)

    bank = is_bank(row)
    failed = is_project_failed(approval_progress)

    if failed:
        category = "无法参与的银行" if bank else "无法参与的非银"
        phrase = classify_unable_phrase(inquiry_progress)
        phrase_order = UNABLE_PHRASE_ORDER.index(phrase) if phrase in UNABLE_PHRASE_ORDER else 999
    else:
        category = "有可能参与的银行" if bank else "有可能参与的非银"
        phrase = classify_probable_phrase(inquiry_progress, row.get("意向投标量（亿元）", ""))
        phrase_order = PROBABLE_PHRASE_ORDER.index(phrase) if phrase in PROBABLE_PHRASE_ORDER else 999

    updated_progress = build_updated_progress(mmdd, phrase, inquiry_progress)
    updated_parts = build_updated_progress_parts(mmdd, phrase, inquiry_progress)

    biz_type_key = business_type
    biz_type_order = BIZ_TYPE_ORDER.index(biz_type_key) if biz_type_key in BIZ_TYPE_ORDER else 999

    amount_num = None
    try:
        amount_num = float(clean_text(amount))
    except (ValueError, TypeError):
        pass

    # 评分
    score = None
    if category in ("有可能参与的银行", "有可能参与的非银"):
        t = clean_text(inquiry_progress)
        if phrase == "认购意向高" or phrase == "当天报量":
            score = 1
        elif phrase == "认购意向低" or re.search(r"认购意向低|参与意向低|投标意向低|意愿低", t):
            score = -1
        else:
            score = 0

    return {
        "分类": category,
        "机构名称": org_name,
        "业务种类": business_type,
        "意向投标量（亿元）": amount,
        "询价进度": inquiry_progress,
        "更新询价进度": updated_progress,
        "认购意向（评分）": score,
        "_phrase": phrase,
        "_parts": updated_parts,
        "_phrase_order": phrase_order,
        "_category_order": FORMAL_CATEGORY_ORDER.index(category) if category in FORMAL_CATEGORY_ORDER else 999,
        "_date_order": mmdd_sort_key(mmdd),
        "_biz_type_order": biz_type_order,
        "_amount_order": -(amount_num if amount_num is not None else -1),
        "_org_name": org_name,
        "_mmdd": mmdd,
    }


def process_file(input_path: str) -> pd.DataFrame:
    df = read_excel_smart(input_path)
    df = map_columns(df)
    rows = []
    for _, row in df.iterrows():
        org_name = clean_text(row.get("机构名称", ""))
        if not org_name:
            continue
        rows.append(classify_row(row))
    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame()
    out = out.sort_values(
        by=["_category_order", "_phrase_order", "_date_order", "_amount_order", "_biz_type_order", "_org_name"],
        ascending=[True, True, False, True, True, True],
        kind="mergesort",
    )
    return out


# ===================== 进度相似度判断 =====================

def _strip_date_tag(text: str) -> str:
    """去掉【MMDD】前缀"""
    return re.sub(r"^【\d{4}】\s*", "", text).strip()


def _extract_keywords(text: str) -> set:
    """提取文本中的关键词用于相似度比较"""
    t = text.lower()
    keywords = set()
    # 意向级别
    for kw in ["认购意向高", "认购意向低", "推介中", "观望中", "当天报量", "在库待沟通"]:
        if kw in t:
            keywords.add(kw)
    # 具体动作
    for kw in ["报量", "上报", "推荐", "参与", "决策", "决定", "当天", "看价格",
                "额度", "信评", "授信", "担保", "入库", "评级", "主体", "理财子",
                "支行", "总行", "分行", "经理", "发起", "申报", "沟通", "流程",
                "不参与", "无法参与", "没有", "价格原因", "期限",
                "双切", "落地", "协商", "待定", "暂定"]:
        if kw in t:
            keywords.add(kw)
    return keywords


def _char_ngram_similarity(s1: str, s2: str, n: int = 3) -> float:
    """字符 n-gram Jaccard 相似度"""
    if not s1 or not s2:
        return 0.0
    ngrams1 = {s1[i:i+n] for i in range(len(s1)-n+1)}
    ngrams2 = {s2[i:i+n] for i in range(len(s2)-n+1)}
    if not ngrams1 or not ngrams2:
        return 0.0
    return len(ngrams1 & ngrams2) / len(ngrams1 | ngrams2)


def _normalize_progress_for_compare(text: str) -> str:
    """规范化询价进度文本用于比较：统一同义表达后再判等。
    只做高置信度的同义替换，不做模糊匹配，避免误合并。"""
    t = clean_text(text)
    # 统一"没有"→"无"（如"没有主体评级"→"无主体评级"）
    t = t.replace("没有", "无")
    # 统一结尾标点
    t = t.rstrip("，。,.")
    return t


def is_progress_similar(p1: str, p2: str) -> bool:
    """判断两条询价进度是否相似（措辞不同但意思一样）"""
    if p1 == p2:
        return True
    # 去掉日期标签再比较
    s1 = _strip_date_tag(p1)
    s2 = _strip_date_tag(p2)
    if s1 == s2:
        return True
    # 一个包含另一个（如"认购意向低" vs "认购意向低，支行..."）
    if s1 in s2 or s2 in s1:
        return True
    # 关键词重叠度判断
    kw1 = _extract_keywords(p1)
    kw2 = _extract_keywords(p2)
    if kw1 and kw2:
        intersection = kw1 & kw2
        union = kw1 | kw2
        jaccard = len(intersection) / len(union) if union else 0
        # 放宽阈值：两个领域特定关键词重叠就足够
        if jaccard > 0.25 and len(intersection) >= 2:
            return True
    # 字符 n-gram 相似度兜底
    char_sim = _char_ngram_similarity(s1, s2)
    if char_sim > 0.4:
        return True
    return False


def merge_progress(p1: str, p2: str) -> str:
    """合并两条进度：相似保留较完整的一条，不相似换行拼接"""
    if is_progress_similar(p1, p2):
        # 保留较长/较完整的那条
        return p1 if len(p1) >= len(p2) else p2
    # 不相似 → 换行拼接
    return f"{p1}\n{p2}"


def merge_amount(a1: str, a2: str) -> str:
    """合并金额：取较大值，或非空值"""
    if a1 and not a2:
        return a1
    if a2 and not a1:
        return a2
    if not a1 and not a2:
        return ""
    try:
        n1 = float(a1)
        n2 = float(a2)
        return str(max(n1, n2))
    except (ValueError, TypeError):
        return a1 if a1 else a2


def merge_score(s1, s2):
    """合并评分：取较高的"""
    if s1 is None and s2 is None:
        return None
    if s1 is None:
        return s2
    if s2 is None:
        return s1
    return max(s1, s2)


def merge_updated_progress(row1, row2):
    """合并更新询价进度：返回 (text, parts) 元组。

    意向相同时合并文本；意向不同时双行展示（品种一在上，品种二在下）。
    """
    phrase1 = row1.get("_phrase", "")
    phrase2 = row2.get("_phrase", "")
    parts1 = row1.get("_parts", [])
    parts2 = row2.get("_parts", [])

    if phrase1 == phrase2:
        # 意向相同：按原逻辑合并文本
        merged_text = merge_progress(row1["询价进度"], row2["询价进度"])
        mmdd = row1.get("_mmdd", "")
        parts = build_updated_progress_parts(mmdd, phrase1, merged_text)
        text = "".join(t for t, _ in parts if t)
        return text, parts, phrase1
    else:
        # 意向不同：双行展示，品种一在上、品种二在下
        # 把换行符追加到品种一末段，不单独成段，确保 Excel 正确渲染换行
        if parts1:
            last_text, last_bold = parts1[-1]
            parts1[-1] = (last_text + "\n", last_bold)
        combined_parts = list(parts1) + list(parts2)
        text = "".join(t for t, _ in combined_parts if t)
        # 排序用 phrase：取较积极的
        p1_order = PROBABLE_PHRASE_ORDER.index(phrase1) if phrase1 in PROBABLE_PHRASE_ORDER else 999
        p2_order = PROBABLE_PHRASE_ORDER.index(phrase2) if phrase2 in PROBABLE_PHRASE_ORDER else 999
        best_phrase = phrase1 if p1_order <= p2_order else phrase2
        return text, combined_parts, best_phrase


# ===================== 双表合并 =====================

def merge_two_tables(path1: str, path2: str) -> pd.DataFrame:
    """读取并合并两张表"""
    df1 = process_file(path1)
    df2 = process_file(path2)

    if df1.empty and df2.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    if df1.empty:
        df2["备注"] = "品种二"
        return df2[OUTPUT_COLUMNS]
    if df2.empty:
        df1["备注"] = "品种一"
        return df1[OUTPUT_COLUMNS]

    # 建立品种二索引：key = (机构名称, 业务种类)
    idx2 = {}
    for i, row in df2.iterrows():
        key = (row["机构名称"], row["业务种类"])
        idx2[key] = row

    merged_rows = []
    matched_keys = set()

    # 遍历品种一
    for i, row1 in df1.iterrows():
        key = (row1["机构名称"], row1["业务种类"])
        if key in idx2:
            row2 = idx2[key]
            matched_keys.add(key)

            # 比较询价进度（清洗后文本）是否完全一致
            progress1_raw = str(row1.get("询价进度", "")).strip()
            progress2_raw = str(row2.get("询价进度", "")).strip()
            same_progress = (_normalize_progress_for_compare(progress1_raw) == _normalize_progress_for_compare(progress2_raw))

            if same_progress:
                # 进度一致 → 合并为一行
                merged_text, merged_parts, merged_phrase = merge_updated_progress(row1, row2)

                merged = {
                    "分类": row1["分类"],
                    "机构名称": row1["机构名称"],
                    "业务种类": row1["业务种类"],
                    "意向投标量（亿元）": merge_amount(row1["意向投标量（亿元）"], row2["意向投标量（亿元）"]),
                    "询价进度": merge_progress(row1["询价进度"], row2["询价进度"]),
                    "更新询价进度": merged_text,
                    "认购意向（评分）": merge_score(row1["认购意向（评分）"], row2["认购意向（评分）"]),
                    "备注": "品种一/品种二",
                    "_is_merged": True,
                    "_mmdd": row1.get("_mmdd", ""),
                    "_phrase": merged_phrase,
                    "_parts": merged_parts,
                    # 排序用
                    "_category_order": row1["_category_order"],
                    "_merge_group": f"{row1['_org_name']}_{row1['_biz_type_order']}",  # 单行无需特殊分组
                    "_phrase_order": min(row1["_phrase_order"], row2["_phrase_order"]),
                    "_date_order": max(row1["_date_order"], row2["_date_order"]),
                    "_amount_order": max(row1["_amount_order"], row2["_amount_order"]),
                    "_biz_type_order": row1["_biz_type_order"],
                    "_org_name": row1["_org_name"],
                }
                merged_rows.append(merged)
            else:
                # 进度不一致 → 分两行展示，但标记为同一组（相邻排列 + 黄色高亮）
                group_key = f"MATCH_{row1['_org_name']}_{row1['_biz_type_order']}"

                # Row a: 品种一
                merged_rows.append({
                    "分类": row1["分类"],
                    "机构名称": row1["机构名称"],
                    "业务种类": row1["业务种类"],
                    "意向投标量（亿元）": row1["意向投标量（亿元）"],
                    "询价进度": row1["询价进度"],
                    "更新询价进度": row1.get("更新询价进度", row1["询价进度"]),
                    "认购意向（评分）": row1["认购意向（评分）"],
                    "备注": "品种一",
                    "_is_merged": True,
                    "_mmdd": row1.get("_mmdd", ""),
                    "_phrase": row1.get("_phrase", ""),
                    "_parts": row1.get("_parts", None),
                    "_category_order": row1["_category_order"],
                    "_merge_group": group_key,
                    "_phrase_order": row1["_phrase_order"],
                    "_date_order": row1["_date_order"],
                    "_amount_order": row1["_amount_order"],
                    "_biz_type_order": row1["_biz_type_order"],
                    "_org_name": row1["_org_name"],
                })
                # Row b: 品种二
                merged_rows.append({
                    "分类": row2["分类"],
                    "机构名称": row2["机构名称"],
                    "业务种类": row2["业务种类"],
                    "意向投标量（亿元）": row2["意向投标量（亿元）"],
                    "询价进度": row2["询价进度"],
                    "更新询价进度": row2.get("更新询价进度", row2["询价进度"]),
                    "认购意向（评分）": row2["认购意向（评分）"],
                    "备注": "品种二",
                    "_is_merged": True,
                    "_mmdd": row2.get("_mmdd", ""),
                    "_phrase": row2.get("_phrase", ""),
                    "_parts": row2.get("_parts", None),
                    "_category_order": row2["_category_order"],
                    "_merge_group": group_key,
                    "_phrase_order": row2["_phrase_order"],
                    "_date_order": row2["_date_order"],
                    "_amount_order": row2["_amount_order"],
                    "_biz_type_order": row2["_biz_type_order"],
                    "_org_name": row2["_org_name"],
                })
        else:
            merged = {
                "分类": row1["分类"],
                "机构名称": row1["机构名称"],
                "业务种类": row1["业务种类"],
                "意向投标量（亿元）": row1["意向投标量（亿元）"],
                "询价进度": row1["询价进度"],
                "更新询价进度": row1.get("更新询价进度", row1["询价进度"]),
                "认购意向（评分）": row1["认购意向（评分）"],
                "备注": "品种一",
                "_is_merged": False,
                "_mmdd": row1.get("_mmdd", ""),
                "_phrase": row1.get("_phrase", ""),
                "_parts": row1.get("_parts", None),
                "_category_order": row1["_category_order"],
                "_merge_group": f"{row1['_org_name']}_{row1['_biz_type_order']}",
                "_phrase_order": row1["_phrase_order"],
                "_date_order": row1["_date_order"],
                "_amount_order": row1["_amount_order"],
                "_biz_type_order": row1["_biz_type_order"],
                "_org_name": row1["_org_name"],
            }
            merged_rows.append(merged)

    # 品种二中未被匹配的行
    for i, row2 in df2.iterrows():
        key = (row2["机构名称"], row2["业务种类"])
        if key not in matched_keys:
            merged_rows.append({
                "分类": row2["分类"],
                "机构名称": row2["机构名称"],
                "业务种类": row2["业务种类"],
                "意向投标量（亿元）": row2["意向投标量（亿元）"],
                "询价进度": row2["询价进度"],
                "更新询价进度": row2.get("更新询价进度", row2["询价进度"]),
                "认购意向（评分）": row2["认购意向（评分）"],
                "备注": "品种二",
                "_is_merged": False,
                "_mmdd": row2.get("_mmdd", ""),
                "_phrase": row2.get("_phrase", ""),
                "_parts": row2.get("_parts", None),
                "_category_order": row2["_category_order"],
                "_merge_group": f"{row2['_org_name']}_{row2['_biz_type_order']}",
                "_phrase_order": row2["_phrase_order"],
                "_date_order": row2["_date_order"],
                "_amount_order": row2["_amount_order"],
                "_biz_type_order": row2["_biz_type_order"],
                "_org_name": row2["_org_name"],
            })

    out = pd.DataFrame(merged_rows)

    # 按机构+业务种类分组计算组级排序键：意向高的组在前，同组内相邻
    group_stats = out.groupby("_merge_group").agg(
        grp_best_phrase=("_phrase_order", "min"),
        grp_best_date=("_date_order", "max"),
        grp_best_amount=("_amount_order", "min"),
    ).reset_index()
    out = out.merge(group_stats, on="_merge_group", how="left")

    out = out.sort_values(
        by=["_category_order", "grp_best_phrase", "grp_best_date", "grp_best_amount",
            "_merge_group", "_phrase_order", "_date_order", "_amount_order", "_biz_type_order"],
        ascending=[True, True, False, True, True, True, False, True, True],
        kind="mergesort",
    )
    return out


# ===================== Excel 输出 =====================

# 列宽（7列）
FORMAL_COLUMN_WIDTHS = {
    "A": 19.41,   # 分类
    "B": 41.69,   # 机构名称
    "C": 11.42,   # 业务种类
    "D": 13.53,   # 意向投标量
    "E": 86.35,   # 询价进度
    "F": 14.00,   # 认购意向（评分）
    "G": 16.00,   # 备注
}


def extract_title_from_path(path: str) -> str:
    """从文件路径提取债券名称用于标题"""
    stem = Path(path).stem
    # 去掉导出时间后缀
    stem = re.sub(r"\s*-\s*20\d{2}-\d{2}-\d{2}T.*$", "", stem)
    # 去掉"询价导出"等后缀
    stem = re.sub(r"询价导出.*$", "", stem).strip()
    # 去掉"品种一/品种二/品种三"等标记
    stem = re.sub(r"[（(]?品种[一二三四五六七八九十]+[）)]?", "", stem).strip()
    # 去掉末尾的分隔符
    stem = re.sub(r"[\s_\-]+$", "", stem)
    return stem


def write_output_excel(df: pd.DataFrame, output_path: str, title: str = "",
                       info_text: str = ""):
    """按模板格式输出：标题行 + 债券信息行 + 表头行 + A列合并分类 + 数据行"""
    wb = Workbook()
    ws = wb.active
    ws.title = "询价明细表"

    FONT_TITLE = Font(name="楷体", size=14)
    FONT_INFO_LABEL = Font(name="楷体", size=14)  # A2 "债券信息"
    FONT_INFO = Font(name="楷体", size=12)         # B2 债券详情
    FONT_HEADER = Font(name="楷体", size=12)
    FONT_DATA = Font(name="楷体", size=12)
    FONT_TOTAL_RED = Font(name="楷体", size=12, color="FF0000")  # 合计行红色
    CENTER = Alignment(horizontal="center", vertical="center")
    CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
    LEFT_VCENTER = Alignment(horizontal="left", vertical="center")
    LEFT_VCENTER_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
    THIN = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    # 备注颜色 - 使用主题色（匹配模板表1）
    # 品种一: light blue theme(?) → 模板用 theme=4, tint=0.8
    # 品种一/品种二: 黄色 FFFF00
    FILL_YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    # G列品种一用主题色 theme=4 tint=0.8
    FILL_G_V1 = PatternFill(fill_type="solid", fgColor=Color(theme=4, tint=0.8))
    FILL_G_V2 = PatternFill(fill_type="solid", fgColor=Color(theme=6, tint=0.8))
    # A列分类标题用主题色
    # 有可能参与: 银行 theme=9 tint=0.8, 非银 theme=5 tint=0.8
    # 无法参与: theme=0 tint=-0.15
    FILL_A_PROBABLE_BANK = PatternFill(fill_type="solid", fgColor=Color(theme=9, tint=0.8))
    FILL_A_PROBABLE_NONBANK = PatternFill(fill_type="solid", fgColor=Color(theme=5, tint=0.8))
    FILL_A_UNABLE = PatternFill(fill_type="solid", fgColor=Color(theme=0, tint=-0.15))
    FILL_WHITE = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")

    # ---- 第1行：标题 A1:G1 ----
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = title
    c.font = FONT_TITLE
    c.alignment = CENTER_WRAP
    c.border = THIN
    ws.row_dimensions[1].height = 30

    # ---- 第2行：债券信息 B2:G2（A2单独放"债券信息"）----
    # A2: "债券信息"
    a2 = ws["A2"]
    a2.value = "债券信息"
    a2.font = FONT_INFO_LABEL
    a2.alignment = CENTER_WRAP
    a2.border = THIN
    # B2:G2: 债券详情
    ws.merge_cells("B2:G2")
    c = ws["B2"]
    c.value = info_text if info_text else ""
    c.font = FONT_INFO
    c.alignment = LEFT_WRAP
    c.border = THIN
    ws.row_dimensions[2].height = 218 if info_text else 30

    # ---- 第3行：表头 ----
    headers = ["参与意向", "机构名称", "业务种类", "意向投标量（亿）", "询价进度", "认购意向（评分）", "备注"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx)
        cell.value = header
        cell.font = FONT_HEADER
        cell.alignment = CENTER_WRAP
        cell.border = THIN
        # 无填充色，保持白底
    ws.row_dimensions[3].height = 30

    # ---- 数据行：A 列按分类合并（含合计行）----
    # 先统计每个分类的起止行（含合计行）
    categories = df["分类"].tolist()
    category_ranges = {}  # cat -> [start_row, end_row] (end_row includes 合计 if applicable)
    category_needs_total = {}  # cat -> bool: "有可能参与"才需要合计

    current_cat = None
    for cat in categories:
        if cat != current_cat:
            category_needs_total[cat] = "有可能参与" in str(cat)
            current_cat = cat

    # 计算每个分类的起止行（考虑合计行）
    row_num = 4
    cat_list = list(dict.fromkeys(categories))  # unique order-preserving
    for cat in cat_list:
        count = categories.count(cat)
        needs_t = category_needs_total.get(cat, False)
        end_row = row_num + count - 1 + (1 if needs_t else 0)
        category_ranges[cat] = (row_num, end_row)
        row_num = end_row + 1

    # 写数据行（含合计行）
    row_num = 4
    prev_cat = None
    for idx, (_, data_row) in enumerate(df.iterrows()):
        category = data_row.get("分类", "")
        remark = data_row.get("备注", "")
        is_merged = data_row.get("_is_merged", False)
        score = data_row.get("认购意向（评分）", "")
        amount = data_row.get("意向投标量（亿元）", "")

        # 检测是否为本分类第一行
        is_first_in_cat = (category != prev_cat)
        prev_cat = category

        # 优先使用更新询价进度（带日期标签），回退到原始
        progress = data_row.get("更新询价进度", "") or data_row.get("询价进度", "")

        row_data = [
            category if is_first_in_cat else "",
            data_row.get("机构名称", ""),
            data_row.get("业务种类", ""),
            amount if amount is not None else "",
            progress,
            score if score is not None else "",
            remark,
        ]

        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_idx)
            if col_idx == 5:
                # E列询价进度：对齐v6工具——日期不加粗，分类话术加粗，楷体12
                parts = data_row.get("_parts", None)
                if parts and isinstance(parts, list):
                    cell.value = parts_to_excel_rich_text(parts, font_name="楷体", font_size=12)
                else:
                    mmdd = data_row.get("_mmdd", "")
                    phrase = data_row.get("_phrase", "")
                    progress = data_row.get("询价进度", "")
                    fallback_parts = build_updated_progress_parts(mmdd, phrase, progress)
                    cell.value = parts_to_excel_rich_text(fallback_parts, font_name="楷体", font_size=12)
            else:
                cell.value = val if val is not None else ""
            if col_idx == 5:
                pass  # font already set via rich text
            else:
                cell.font = FONT_DATA
            cell.border = THIN
            if col_idx in (1, 3, 4, 6, 7):
                cell.alignment = CENTER_WRAP
            elif col_idx == 5:
                cell.alignment = LEFT_WRAP
            else:
                cell.alignment = LEFT_VCENTER_WRAP

        # 备注列着色（G列 - 使用主题色）
        g_cell = ws.cell(row=row_num, column=7)
        if is_merged:
            g_cell.fill = FILL_YELLOW
            # 合并行的 B/C 列也黄色高亮
            ws.cell(row=row_num, column=2).fill = FILL_YELLOW
            ws.cell(row=row_num, column=3).fill = FILL_YELLOW
        elif remark == "品种一":
            g_cell.fill = FILL_G_V1
        elif remark == "品种二":
            g_cell.fill = FILL_G_V2
        else:
            g_cell.fill = FILL_WHITE

        # A 列分类标题着色（仅本分类第一行）
        if is_first_in_cat:
            a_cell = ws.cell(row=row_num, column=1)
            if "有可能参与" in str(category) and "银行" in str(category):
                a_cell.fill = FILL_A_PROBABLE_BANK
            elif "有可能参与" in str(category) and "非银" in str(category):
                a_cell.fill = FILL_A_PROBABLE_NONBANK
            elif "无法参与" in str(category):
                a_cell.fill = FILL_A_UNABLE
            else:
                a_cell.fill = FILL_WHITE

        row_num += 1

        # 本分类最后一行（数据行）后写合计行
        cat_start, cat_end = category_ranges.get(category, (0, 0))
        # cat_end 已包含合计行，所以 row_num == cat_end 时表示刚写完最后一数据行，下一行该写合计
        if row_num == cat_end and category_needs_total.get(category, False):
            # 汇总本分类 意向投标量
            total_amount = 0.0
            cat_indices = [j for j, c in enumerate(categories) if c == category]
            for ci in cat_indices:
                amt = df.iloc[ci].get("意向投标量（亿元）", "")
                try:
                    total_amount += float(amt) if amt and amt != "" else 0.0
                except (ValueError, TypeError):
                    pass

            # 合计行各列
            total_row = ["", "合计", "", total_amount if total_amount > 0 else "", "", "", ""]
            for col_idx, val in enumerate(total_row, 1):
                cell = ws.cell(row=row_num, column=col_idx)
                cell.value = val if val is not None else ""
                if col_idx == 2 or col_idx == 4:
                    cell.font = FONT_TOTAL_RED  # 合计和金额标红
                else:
                    cell.font = FONT_DATA
                cell.border = THIN
                if col_idx in (1, 3, 4, 6, 7):
                    cell.alignment = CENTER_WRAP
                else:
                    cell.alignment = LEFT_VCENTER_WRAP
            ws.row_dimensions[row_num].height = 20
            row_num += 1

    # ---- A 列合并（含合计行在内的区间）----
    for cat, (r_start, r_end) in category_ranges.items():
        if r_end > r_start:
            ws.merge_cells(f"A{r_start}:A{r_end}")
            ws.cell(row=r_start, column=1).alignment = CENTER_WRAP

    # ---- 列宽（与模板一致）----
    col_widths = {"A": 22.4636363636364, "B": 41.6909090909091, "C": 11.4181818181818, "D": 13.5272727272727, "E": 86.3454545454545, "F": 14.0, "G": 10.1181818181818}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # ---- 数据行高（双行合并行用更高行高）----
    for r in range(4, row_num):
        e_cell = ws.cell(row=r, column=5)
        e_val = e_cell.value
        has_newline = False
        if hasattr(e_val, '__iter__') and not isinstance(e_val, str):
            for block in e_val:
                if hasattr(block, 'text') and '\n' in (block.text or ''):
                    has_newline = True
                    break
        elif isinstance(e_val, str) and '\n' in e_val:
            has_newline = True
        ws.row_dimensions[r].height = 40 if has_newline else 20

    # ---- F 列条件格式：红绿灯（对齐v9）----
    ws.conditional_formatting.add(
        f"F4:F{row_num - 1}",
        IconSetRule(icon_style="3TrafficLights1", type="num", values=[-1, 0, 1])
    )

    # ---- 小结区：对齐 v9 标准版 8 行结构 ----
    # 统计（按机构名称去重，避免进度不同分行后重复计数）
    probable_bank_rows = df[df["分类"].str.contains("有可能参与.*银行", na=False)]
    probable_nonbank_rows = df[df["分类"].str.contains("有可能参与.*非银", na=False)]

    # 高意向银行：按（机构名称+业务种类）去重
    high_bank_rows = probable_bank_rows[probable_bank_rows["_phrase"].isin(["认购意向高", "当天报量"])]
    bank_high_pairs = set(zip(high_bank_rows["机构名称"], high_bank_rows["业务种类"]))
    bank_high_count = len(bank_high_pairs)

    # 非银关注家数：按（机构名称+业务种类）去重
    nonbank_count = len(probable_nonbank_rows[["机构名称", "业务种类"]].drop_duplicates())

    # 银行报量：按（机构名称+业务种类）去重后汇总
    seen_bank = set()
    bank_amount = 0.0
    for _, r in probable_bank_rows.iterrows():
        key = (r.get("机构名称", ""), r.get("业务种类", ""))
        if key not in seen_bank:
            seen_bank.add(key)
            try:
                bank_amount += float(r.get("意向投标量（亿元）", 0) or 0)
            except (ValueError, TypeError):
                pass

    # 合计报量 + 非银报量：按（机构名称+业务种类）去重
    total_amount = 0.0
    nonbank_amount = 0.0
    seen_total = set()
    seen_nonbank = set()
    for _, r in df.iterrows():
        cat = str(r.get("分类", ""))
        if "有可能参与" in cat:
            key = (r.get("机构名称", ""), r.get("业务种类", ""))
            if key not in seen_total:
                seen_total.add(key)
                try:
                    amt = float(r.get("意向投标量（亿元）", 0) or 0)
                    total_amount += amt
                    if "非银" in cat and key not in seen_nonbank:
                        nonbank_amount += amt
                        seen_nonbank.add(key)
                except (ValueError, TypeError):
                    pass

    FILL_SUMMARY = PatternFill(fill_type="solid", fgColor=Color(theme=8, tint=0.8))
    LEFT_VCENTER = Alignment(horizontal="left", vertical="center")
    CENTER_VCENTER = Alignment(horizontal="center", vertical="center")

    summary_start = row_num
    summary_end = summary_start + 7

    # Row 1: 小结｜参与意向较高的银行家数｜count｜合计报量｜value
    ws.cell(summary_start, 1, "小结")
    ws.cell(summary_start, 2, "参与意向较高的银行家数")
    ws.cell(summary_start, 3, bank_high_count if bank_high_count else None)
    ws.cell(summary_start, 4, "合计报量")
    ws.cell(summary_start, 5, total_amount if total_amount > 0 else None)
    ws.merge_cells(start_row=summary_start, start_column=5, end_row=summary_start, end_column=7)

    # Row 2-3: 银行报量 / 估值 (B-D merge, E-G merge)
    ws.cell(summary_start + 1, 2, "银行报量")
    ws.cell(summary_start + 1, 3, bank_amount if bank_amount > 0 else None)
    ws.cell(summary_start + 1, 4, "估值")
    for c in range(2, 5):
        ws.merge_cells(start_row=summary_start + 1, start_column=c, end_row=summary_start + 2, end_column=c)
    ws.merge_cells(start_row=summary_start + 1, start_column=5, end_row=summary_start + 2, end_column=7)

    # Row 4-5: 占比 / 发行人预期 (B-D merge, E-G merge)
    ws.cell(summary_start + 3, 2, "占比")
    ws.cell(summary_start + 3, 4, "发行人预期")
    for c in range(2, 5):
        ws.merge_cells(start_row=summary_start + 3, start_column=c, end_row=summary_start + 4, end_column=c)
    ws.merge_cells(start_row=summary_start + 3, start_column=5, end_row=summary_start + 4, end_column=7)

    # Row 6-7: 非银关注家数 / 备注 (B-D merge, E-G merge)
    ws.cell(summary_start + 5, 2, "非银关注家数")
    ws.cell(summary_start + 5, 3, nonbank_count if nonbank_count else None)
    ws.cell(summary_start + 5, 4, "备注")
    for c in range(2, 5):
        ws.merge_cells(start_row=summary_start + 5, start_column=c, end_row=summary_start + 6, end_column=c)
    ws.merge_cells(start_row=summary_start + 5, start_column=5, end_row=summary_start + 6, end_column=7)

    # Row 8: 非银报量（C列居中，对齐小结其他数值）
    ws.cell(summary_start + 7, 2, "非银报量")
    ws.cell(summary_start + 7, 3, nonbank_amount if nonbank_amount > 0 else None)

    # 小结样式
    for r in range(summary_start, summary_end + 1):
        ws.row_dimensions[r].height = 24
        for c in range(1, 8):
            cell = ws.cell(r, c)
            cell.font = FONT_DATA
            cell.fill = FILL_SUMMARY
            cell.border = THIN
            if c == 1:
                cell.alignment = CENTER_VCENTER
            elif c in (2, 5):
                cell.alignment = LEFT_VCENTER
            else:
                cell.alignment = CENTER_VCENTER

    # A 列合并（小结标签纵跨 8 行）
    ws.merge_cells(start_row=summary_start, start_column=1, end_row=summary_end, end_column=1)
    # 合计报量数值居中
    ws.cell(summary_start, 5).alignment = CENTER_VCENTER

    row_num = summary_end + 1
    # ---- 小结结束 ----

    wb.save(output_path)
    return row_num - 1


# ===================== CLI =====================

def main():
    parser = argparse.ArgumentParser(description="债券询价双表合并工具")
    parser.add_argument("file1", help="品种一询价导出文件 (.xls/.xlsx)")
    parser.add_argument("file2", help="品种二询价导出文件 (.xls/.xlsx)")
    parser.add_argument("-o", "--output", help="输出 Excel 路径", default=None)
    parser.add_argument("--title", help="自定义标题", default=None)
    args = parser.parse_args()

    path1 = args.file1
    path2 = args.file2

    if not Path(path1).exists():
        print(f"错误：文件不存在 - {path1}")
        sys.exit(1)
    if not Path(path2).exists():
        print(f"错误：文件不存在 - {path2}")
        sys.exit(1)

    print(f"品种一: {path1}")
    print(f"品种二: {path2}")

    df = merge_two_tables(path1, path2)

    if df.empty:
        print("错误：两张表均无可处理的数据")
        sys.exit(1)

    # 统计
    merged_count = df["_is_merged"].sum() if "_is_merged" in df else 0
    variety1_only = (df["备注"] == "品种一").sum()
    variety2_only = (df["备注"] == "品种二").sum()
    variety_both = (df["备注"] == "品种一/品种二").sum()

    print(f"总行数: {len(df)}")
    print(f"  品种一独有: {variety1_only}")
    print(f"  品种二独有: {variety2_only}")
    print(f"  合并行: {variety_both}")

    # 标题
    if args.title:
        title = args.title
    else:
        t1 = extract_title_from_path(path1)
        t2 = extract_title_from_path(path2)
        today = datetime.now().strftime("%Y%m%d")
        # 取第一个有意义的名称，如果相同只用一次
        if t1 == t2:
            title = f"{t1}询价表{today}"
        elif t1:
            title = f"{t1}询价表{today}"
        else:
            title = f"询价表{today}"

    # 输出路径
    if args.output:
        output_path = args.output
    else:
        today = datetime.now().strftime("%Y%m%d")
        output_path = f"询价明细表（）-{today}.xlsx"

    total_rows = write_output_excel(df, output_path, title=title, info_text="债券信息")
    print(f"输出文件: {output_path}")
    print(f"Excel 总行数: {total_rows}")


if __name__ == "__main__":
    main()

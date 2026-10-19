#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
债券/固收询价导出表自动整理工具 v6

功能：
1. 读取原始询价导出表（.xls/.xlsx）
2. 按固定规则分类：
   - 只有“审批进度=项目未通过”才进入无法参与
   - 无法参与继续拆分为银行/非银
   - 自动识别“意向投标量（万元）/意向投标量（亿元）/投标量/拟投量”等字段，保留原始数值
   - 询价进度为空：
       * 项目未通过：其他
       * 非项目未通过：认购意向不明
3. 生成正式“询价明细表”格式 Excel：
   - 不冻结窗格
   - 全表全框线
   - 单元格自动换行，避免文字视觉上溢出到旁边单元格
   - 字体、列宽、标题区尽量与模板保持一致
   - 更新询价进度列自动避免“认购意向低：认购意向低”这类重复
   - 更新询价进度列中的分类话术在 Excel 内局部加粗
   - 有可能参与银行/非银末尾合计行使用红色字体
   - 小结区按标准版 7 行结构和合并单元格输出

用法：
  python bond_inquiry_processor.py "原始询价导出.xls"
  python bond_inquiry_processor.py "原始询价导出.xls" -o "整理结果.xlsx"
  python bond_inquiry_processor.py "原始询价导出.xls" -t "询价明细表模板.xlsx" -o "整理结果.xlsx"

依赖：
  pip install -r requirements.txt
"""

import argparse
import math
import os
import re
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    RICH_TEXT_AVAILABLE = True
except Exception:
    # 兼容较旧 openpyxl：没有富文本能力时，程序仍能输出普通文本，避免直接报错。
    CellRichText = None
    TextBlock = None
    InlineFont = None
    RICH_TEXT_AVAILABLE = False


# ===================== 1. 可调整配置 =====================

OUTPUT_COLUMNS = [
    "分类",
    "机构名称",
    "业务种类",
    "意向投标量（亿元）",
    "询价进度",
    "更新询价进度",
    "认购意向（评分）",
]

FORMAL_CATEGORY_ORDER = [
    "有概率参与的银行",
    "有概率参与的非银",
    "无法参与的银行",
    "无法参与的非银",
]

# 有概率参与的话术排序；“认购意向不明”为新增兜底分类，放在最后
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

# 业务种类排序（自营→资管→理财→专户→投顾→公募）
BIZ_TYPE_ORDER = ["自营", "资管", "理财", "专户", "投顾", "公募"]

# True = 同一话术内日期新的排前；False = 日期旧的排前
SORT_NEW_DATE_FIRST = True

# 原表字段别名。如果导出表列名不一样，可在这里追加
COLUMN_ALIASES = {
    "机构名称": ["机构名称", "机构", "机构全称", "客户名称", "投资人名称"],
    "机构类型": ["机构类型", "机构类别", "客户类型", "投资人类型", "类型"],
    "业务种类": ["业务种类", "业务类型", "产品类型", "账户类型", "业务"],
    "审批进度": ["审批进度", "项目进度", "信评进度", "审批状态", "状态"],
    # 很多系统导出的列名会写成“意向投标量（万元）”，但实际录入值通常仍是 0.1、0.5 这类口径；此处只做字段识别，不改写原始数值。
    "意向投标量（亿元）": [
        "意向投标量（亿元）", "意向投标量(亿元)",
        "意向投标量（亿）", "意向投标量(亿)",
        "意向投标量（万元）", "意向投标量(万元)", "意向投标量（万）", "意向投标量(万)",
        "意向投标量", "投标量", "意向量", "拟投量", "认购量", "参与量", "金额", "认购金额", "投标金额"
    ],
    "询价进度": ["询价进度", "沟通细节", "反馈内容", "进度", "反馈", "备注", "跟进情况"],
    "更新时间": ["更新时间", "更新日期", "最后更新时间", "日期", "时间"],
}

# 输出 Excel 列宽，基本参照你提供的模板
FORMAL_COLUMN_WIDTHS = {
    "A": 19.41,
    "B": 41.69,
    "C": 11.42,
    "D": 13.53,
    "E": 86.35,
    "F": 14.00,
}

DEFAULT_TEMPLATE_NAME = "询价明细表模板.xlsx"


# ===================== 2. 文本与字段处理 =====================

def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_column_name(name: Any) -> str:
    return clean_text(name).replace(" ", "").replace("\n", "")


def find_header_row(raw: pd.DataFrame) -> int:
    alias_pool = set()
    for aliases in COLUMN_ALIASES.values():
        for item in aliases:
            alias_pool.add(normalize_column_name(item))

    best_row = 0
    best_score = -1

    for idx in range(min(len(raw), 40)):
        row_values = [normalize_column_name(x) for x in raw.iloc[idx].tolist()]
        non_empty = sum(1 for x in row_values if x and x.lower() != 'nan')
        alias_hits = sum(1 for x in row_values if x in alias_pool)
        score = alias_hits * 2 + non_empty

        if score > best_score:
            best_score = score
            best_row = idx

    return best_row


def read_excel_smart(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=object)

    # 特殊情况：第一行是大标题（合并单元格），第二行才是表头
    first_row = raw.iloc[0].tolist()
    non_empty_first = sum(1 for x in first_row if pd.notna(x) and str(x).strip())

    if non_empty_first <= 1:
        # 第一行是合并标题，再检查第二行是否表头还是债券信息
        header_row = find_header_row(raw)
        if header_row <= 1:
            header_row = 1
    else:
        header_row = find_header_row(raw)

    df = pd.read_excel(path, header=header_row, dtype=object)
    df = df.dropna(how="all").copy()

    # 去掉Unnamed列
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]

    df.columns = [clean_text(c) for c in df.columns]
    return df


def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    normalized_to_original = {normalize_column_name(c): c for c in df.columns}
    rename_map = {}
    for standard_col, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            key = normalize_column_name(alias)
            if key in normalized_to_original:
                rename_map[normalized_to_original[key]] = standard_col
                break
    df = df.rename(columns=rename_map)
    for col in COLUMN_ALIASES.keys():
        if col not in df.columns:
            df[col] = ""
    return df


def parse_amount(value: Any) -> str:
    """意向投标量保留原始数值表达。"""
    if pd.isna(value):
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).rstrip("0").rstrip(".")
    if isinstance(value, int):
        return str(value)
    return clean_text(value)


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
    """将更新时间转为 MMDD，例如 2026-06-17 -> 0617。"""
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
    """判断询价进度中是否存在明确金额/比例/上报量。"""
    lower_text = text.lower()
    patterns = [
        r"\d+(\.\d+)?\s*(e|亿|亿元)",
        r"\d+(\.\d+)?\s*(kw|k|w|万)",
        r"\d+(\.\d+)?\s*%",
        r"参与\s*\d+",
        r"预计\s*\d+",
        r"暂定\s*\d+",
        r"申请.*\d+.*额度",
        r"上报.*\d+",
        r"按\s*\d+.*上报",
        r"报\s*\d+",
    ]
    return any(re.search(p, lower_text, flags=re.IGNORECASE) for p in patterns)


# ===================== 3. 分类规则 =====================

def classify_probable_phrase(text: str, amount_value: Any) -> str:
    """
    非“项目未通过”的机构均属于有概率参与。
    注意：如果询价进度为空，按用户新口径归为“认购意向不明”，即使意向量不为空也先按空进度处理。
    """
    t = clean_text(text)

    if t == "":
        return "认购意向不明"

    # 明确高意向、明确金额/比例优先级最高，避免被“推动/当天看”等词覆盖。
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
    """仅“审批进度=项目未通过”的机构使用无法参与话术。"""
    t = clean_text(text)

    if t == "":
        return "其他"
    if "不在库" in t:
        return "不在库"
    if re.search(r"价格|买不起|跟不上|收益|成本|利差|区间.*低", t):
        return "价格原因"
    if re.search(r"信评过不了|信评不过|没授信|无授信|只能看AAA|不看CP|没有担保|主体不符|资质|不看交易所", t):
        return "资质不符"
    if re.search(r"额度|无额度|没有额度|额度不足|批复失效", t):
        return "额度受限"
    if re.search(r"分行没有报|分行没报|未上报|没有报|来不及|未提供|联系人|流程|时间", t):
        return "流程/时间受限"
    if "利率" in t:
        return "利率原因"
    if re.search(r"不参与|不看|投不了|不投|不能参与|不考虑", t):
        return "不参与"
    return "其他"



def build_updated_progress_parts(mmdd: str, phrase: str, inquiry_progress: str) -> List[Tuple[str, bool]]:
    """
    生成“更新询价进度”的分段文本。

    规则：
    1. 日期不加粗，分类话术加粗，后续原始进度不加粗。
    2. 避免“认购意向低：认购意向低”这类重复。
    3. 询价进度为空时不加冒号，例如：【0617】认购意向不明。
    """
    t = clean_text(inquiry_progress)
    prefix = f"【{mmdd}】" if mmdd else ""

    if t == "":
        return [(prefix, False), (phrase, True)]

    if t == phrase:
        return [(prefix, False), (phrase, True)]

    duplicate_prefixes = [
        phrase + "：",
        phrase + ":",
        phrase + "，",
        phrase + ",",
        phrase + "、",
        phrase + " ",
    ]

    for pfx in duplicate_prefixes:
        if t.startswith(pfx):
            rest = t[len(phrase):]
            return [(prefix, False), (phrase, True), (rest, False)]

    return [(prefix, False), (phrase, True), (",", False), (t, False)]


def parts_to_plain(parts: List[Tuple[str, bool]]) -> str:
    return "".join(text for text, _ in parts if text)


def parts_to_markdown(parts: List[Tuple[str, bool]]) -> str:
    out = []
    for text, bold in parts:
        if not text:
            continue
        safe = text.replace("|", "｜")
        out.append(f"**{safe}**" if bold else safe)
    return "".join(out)


def parts_to_excel_rich_text(parts: List[Tuple[str, bool]], font_name: str = "Noto Sans CJK SC", font_size: int = 12):
    """将分段文本写成 Excel 富文本；openpyxl 版本过旧时自动降级为普通文本。"""
    if not RICH_TEXT_AVAILABLE:
        return parts_to_plain(parts)

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
        category = "有概率参与的银行" if bank else "有概率参与的非银"
        phrase = classify_probable_phrase(inquiry_progress, row.get("意向投标量（亿元）", ""))
        phrase_order = PROBABLE_PHRASE_ORDER.index(phrase) if phrase in PROBABLE_PHRASE_ORDER else 999

    updated_progress_parts = build_updated_progress_parts(mmdd, phrase, inquiry_progress)
    updated_progress = parts_to_plain(updated_progress_parts)

    # 业务种类排序
    biz_type_key = business_type
    biz_type_order = BIZ_TYPE_ORDER.index(biz_type_key) if biz_type_key in BIZ_TYPE_ORDER else 999

    # 意向投标量数值（降序用）
    amount_num = to_number_or_none(amount)

    # 认购意向评分：仅对"有可能参与"类评分，无法参与留空
    # 评分：当天报量/认购意向高 → 1，认购意向低 → -1，其余 → 0
    # 注意：原始询价进度中可能混合信号（如"推介中，但认购意向低"），需二次检查原始文本
    score = None
    if category in ("有概率参与的银行", "有概率参与的非银"):
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
        "_updated_progress_parts": updated_progress_parts,
        "_category_order": FORMAL_CATEGORY_ORDER.index(category) if category in FORMAL_CATEGORY_ORDER else 999,
        "_phrase_order": phrase_order,
        "_date_order": mmdd_sort_key(mmdd),
        "_biz_type_order": biz_type_order,
        "_amount_order": -(amount_num if amount_num is not None else -1),
        "_org_name": org_name,
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
        return pd.DataFrame(columns=OUTPUT_COLUMNS + ["_updated_progress_parts"])

    # 排序优先级：1 意向(高→低) 2 时间(近→远) 3 金额(高→低) 4 业务种类(自营→公募)
    out = out.sort_values(
        by=["_category_order", "_phrase_order", "_date_order", "_amount_order", "_biz_type_order", "_org_name"],
        ascending=[True, True, False, True, True, True],
        kind="mergesort",
    )
    return out[OUTPUT_COLUMNS + ["_phrase", "_updated_progress_parts"]]


# ===================== 4. 模板与样式 =====================

def default_title_from_input(input_path: str) -> str:
    stem = Path(input_path).stem
    # 去掉导出时间后缀，避免标题过长
    stem = re.sub(r"\s*-\s*20\d{2}-\d{2}-\d{2}T.*$", "", stem)
    today = datetime.now().strftime("%Y%m%d")
    return f"{stem}询价表{today}"


def load_template_text(template_path: Optional[str], input_path: str) -> Tuple[str, str, List[str]]:
    title = default_title_from_input(input_path)
    bond_info = (
        "债券名称：\n"
        "发行人：\n"
        "债券类型：\n"
        "发行规模：\n"
        "发行期限：\n"
        "信用评级：\n"
        "利率区间：\n"
        "独立主承销商/簿记管理人：\n"
        "簿记时间：\n"
        "上市/挂牌场所：\n"
        "募集资金用途："
    )
    # 按用户给的标准版，第三列表头默认为“业务种类”。如模板第 3 行有表头，则优先读取模板。
    headers = ["参与意向", "机构名称", "业务种类", "意向投标量（亿）", "询价进度", "认购意向（评分）"]

    if template_path and os.path.exists(template_path):
        try:
            twb = load_workbook(template_path)
            tws = twb.active
            t_title = tws["A1"].value
            t_info = tws["B2"].value
            if t_title and "请在此填写" not in str(t_title):
                title = str(t_title)
            if t_info and "请在此填写" not in str(t_info):
                bond_info = str(t_info)
            t_headers = [clean_text(tws.cell(3, i).value) for i in range(1, 7)]
            if any(t_headers):
                headers = [t_headers[i] or headers[i] for i in range(6)]
        except Exception:
            pass

    return title, bond_info, headers


def estimate_row_height(*texts: str) -> float:
    """估算行高，确保长文本换行后大体能显示完整。"""
    max_lines = 1
    # B列约 20-22 个中文字符一行；E列约 48-52 个中文字符一行
    specs = [(texts[0] if len(texts) > 0 else "", 22), (texts[1] if len(texts) > 1 else "", 50)]
    for text, chars_per_line in specs:
        text = text or ""
        explicit_lines = text.count("\n") + 1
        calc_lines = max(1, math.ceil(len(text) / chars_per_line))
        max_lines = max(max_lines, explicit_lines, calc_lines)
    return min(max(20, max_lines * 18), 409)


def apply_all_borders(ws, min_row: int, max_row: int, min_col: int = 1, max_col: int = 6) -> None:
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = border


def style_formal_sheet(ws) -> None:
    """统一全表样式：不冻结窗格、全框线、换行、字体接近模板。"""
    ws.freeze_panes = None
    ws.sheet_view.pane = None
    ws.sheet_view.selection = []

    for col, width in FORMAL_COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = width

    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5


def to_number_or_none(value: Any) -> Optional[float]:
    """将意向投标量转成可求和数字；无法转换则返回 None。"""
    if pd.isna(value):
        return None
    text = clean_text(value)
    if text == "":
        return None
    try:
        return float(text)
    except Exception:
        m = re.search(r"\d+(?:\.\d+)?", text)
        if m:
            try:
                return float(m.group(0))
            except Exception:
                return None
    return None


def export_formal_excel(df: pd.DataFrame, output_xlsx: str, input_path: str, template_path: Optional[str] = None) -> None:
    title, bond_info, headers = load_template_text(template_path, input_path)

    wb = Workbook()
    ws = wb.active
    ws.title = "询价明细表"
    style_formal_sheet(ws)

    # 基础样式
    # 标准版字体为“楷体”。
    font_title = Font(name="楷体", size=14, bold=False)
    font_header = Font(name="楷体", size=12, bold=False)
    font_body = Font(name="楷体", size=12, bold=False)
    font_amount = Font(name="楷体", size=12, bold=False)
    font_total = Font(name="楷体", size=12, bold=False, color="FF0000")
    fill_white = PatternFill("solid", fgColor="FFFFFF")
    # 模板使用主题色（theme color），与模板严格一致：
    #   theme 9 tint 0.8 → 浅蓝（有可能参与的银行）
    #   theme 5 tint 0.8 → 浅橙（有可能参与的非银）
    #   theme 0 tint -0.15 → 浅灰（无法参与）
    #   theme 8 tint 0.8 → 浅绿/青（小结）
    fill_bank = PatternFill("solid", fgColor=Color(theme=9, tint=0.8))
    fill_nonbank = PatternFill("solid", fgColor=Color(theme=5, tint=0.8))
    fill_unable = PatternFill("solid", fgColor=Color(theme=0, tint=-0.15))
    fill_summary = PatternFill("solid", fgColor=Color(theme=8, tint=0.8))

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 标题区
    ws.merge_cells("A1:F1")
    ws["A1"] = title
    ws["A1"].font = font_title
    ws["A1"].alignment = center
    ws["A1"].fill = fill_white
    ws.row_dimensions[1].height = 30

    ws["A2"] = "债券信息"
    ws["A2"].font = font_title
    ws["A2"].alignment = center
    ws["A2"].fill = fill_white
    ws.merge_cells("B2:F2")
    ws["B2"] = bond_info
    ws["B2"].font = font_body
    ws["B2"].alignment = left
    ws["B2"].fill = fill_white
    ws.row_dimensions[2].height = min(max(120, (str(bond_info).count("\n") + 1) * 18 + 20), 409)

    for col_idx, value in enumerate(headers, 1):
        cell = ws.cell(3, col_idx, value)
        cell.font = font_header
        cell.alignment = center
        cell.fill = fill_white
    ws.row_dimensions[3].height = 30

    # 明细区：正式模板只输出五列，其中 E 列使用“更新询价进度”。
    # 标准版要求：有可能参与的银行/非银每个分组末尾增加“合计”行，合计 D 列意向投标量。
    row_idx = 4
    probable_total_rows: List[int] = []
    probable_total_value = 0.0
    probable_bank_value = 0.0
    probable_nonbank_value = 0.0
    probable_bank_count = 0
    probable_nonbank_count = 0

    for category in FORMAL_CATEGORY_ORDER:
        group = df[df["分类"] == category] if not df.empty else pd.DataFrame(columns=df.columns)
        if group.empty:
            continue

        start_row = row_idx
        group_amount_total = 0.0
        group_amount_has_value = False

        for _, r in group.iterrows():
            ws.cell(row_idx, 2, clean_text(r.get("机构名称", "")))
            ws.cell(row_idx, 3, clean_text(r.get("业务种类", "")))
            amount = r.get("意向投标量（亿元）", "")
            ws.cell(row_idx, 4, amount if amount != "" else None)
            amount_num = to_number_or_none(amount)
            if amount_num is not None:
                group_amount_total += amount_num
                group_amount_has_value = True

            updated_parts = r.get("_updated_progress_parts", None)
            if not isinstance(updated_parts, list):
                updated_parts = build_updated_progress_parts("", "", clean_text(r.get("更新询价进度", "")))
            ws.cell(row_idx, 5).value = parts_to_excel_rich_text(updated_parts, font_name="楷体", font_size=12)
            score = r.get("认购意向（评分）", None)
            if score is not None:
                ws.cell(row_idx, 6, score)
            ws.row_dimensions[row_idx].height = estimate_row_height(
                clean_text(r.get("机构名称", "")),
                clean_text(r.get("更新询价进度", "")),
            )
            row_idx += 1

        # 标准版只在“有可能参与”的银行/非银分组末尾增加合计行。
        if category in ("有概率参与的银行", "有概率参与的非银"):
            total_row = row_idx
            ws.cell(total_row, 2, "合计")
            ws.cell(total_row, 4, group_amount_total if group_amount_has_value else None)
            ws.row_dimensions[total_row].height = 24
            probable_total_rows.append(total_row)
            probable_total_value += group_amount_total
            if category == "有概率参与的银行":
                probable_bank_value += group_amount_total
            else:
                probable_nonbank_value += group_amount_total
            if category == "有概率参与的银行":
                probable_bank_count = len(group[group["_phrase"] == "认购意向高"])
            else:
                probable_nonbank_count = len(group)
            row_idx += 1

        end_row = row_idx - 1

        if start_row == end_row:
            ws.cell(start_row, 1, category.replace("有概率", "有可能"))
        else:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
            ws.cell(start_row, 1, category.replace("有概率", "有可能"))

        group_cell = ws.cell(start_row, 1)
        group_cell.font = font_body
        group_cell.alignment = center
        # 按模板使用主题色
        if category == "有概率参与的银行":
            group_cell.fill = fill_bank
        elif category == "有概率参与的非银":
            group_cell.fill = fill_nonbank
        else:
            group_cell.fill = fill_unable

    if row_idx == 4:
        ws.cell(4, 1, "无数据")
        ws.merge_cells("A4:F4")
        row_idx = 5

    last_table_row = row_idx - 1

    # 统一明细单元格样式
    for r in range(4, last_table_row + 1):
        for c in range(1, 7):
            cell = ws.cell(r, c)
            if c != 5:
                cell.font = font_amount if c == 4 else font_body
            if c in (1, 3, 4, 6):
                cell.alignment = center
            else:
                cell.alignment = left
            # 不给明细区填充底色，保持模板简洁

    # 合计行样式：严格匹配标准版，B:E 为红色字体，不做底色。
    for tr in probable_total_rows:
        ws.row_dimensions[tr].height = 24
        for c in range(2, 7):
            cell = ws.cell(tr, c)
            cell.font = font_total
            cell.alignment = center
            cell.fill = fill_white

    # 小结区：严格按标准版 8 行结构（模板精确对应）。
    # 第 1 行：参与意向较高的银行家数 / 合计报量
    # 第 2 行：银行报量(与第3行合并 B~E) / 估值
    # 第 3 行：空行（B~E 与第2行合并）
    # 第 4 行：占比(与第5行合并 B~E) / 发行人预期
    # 第 5 行：空行（B~E 与第4行合并）
    # 第 6 行：非银关注家数(与第7行合并 B~E) / 备注
    # 第 7 行：空行（B~E 与第6行合并）
    # 第 8 行：非银报量（B~E 不合并）
    summary_start = row_idx
    summary_end = summary_start + 7

    bank_count = probable_bank_count
    bank_amount = probable_bank_value
    nonbank_count = probable_nonbank_count
    nonbank_amount = probable_nonbank_value

    # 第1行：小结｜参与意向较高的银行家数｜数值｜合计报量｜数值
    ws.cell(summary_start, 1, "小结")
    ws.cell(summary_start, 2, "参与意向较高的银行家数")
    ws.cell(summary_start, 3, bank_count if bank_count else None)
    ws.cell(summary_start, 4, "合计报量")
    ws.cell(summary_start, 5, probable_total_value if probable_total_rows else None)
    ws.merge_cells(start_row=summary_start, start_column=5, end_row=summary_start, end_column=6)

    # 第2-3行：银行报量（B~D 各自合并，E~F 合并）
    ws.cell(summary_start + 1, 2, "银行报量")
    ws.cell(summary_start + 1, 3, bank_amount if bank_amount else None)
    ws.cell(summary_start + 1, 4, "估值")
    for c in range(2, 5):
        ws.merge_cells(start_row=summary_start + 1, start_column=c, end_row=summary_start + 2, end_column=c)
    ws.merge_cells(start_row=summary_start + 1, start_column=5, end_row=summary_start + 2, end_column=6)

    # 第4-5行：占比 / 发行人预期（B~D 各自合并，E~F 合并）
    ws.cell(summary_start + 3, 2, "占比")
    ws.cell(summary_start + 3, 4, "发行人预期")
    for c in range(2, 5):
        ws.merge_cells(start_row=summary_start + 3, start_column=c, end_row=summary_start + 4, end_column=c)
    ws.merge_cells(start_row=summary_start + 3, start_column=5, end_row=summary_start + 4, end_column=6)

    # 第6-7行：非银关注家数 / 备注（B~D 各自合并，E~F 合并）
    ws.cell(summary_start + 5, 2, "非银关注家数")
    ws.cell(summary_start + 5, 3, nonbank_count if nonbank_count else None)
    ws.cell(summary_start + 5, 4, "备注")
    for c in range(2, 5):
        ws.merge_cells(start_row=summary_start + 5, start_column=c, end_row=summary_start + 6, end_column=c)
    ws.merge_cells(start_row=summary_start + 5, start_column=5, end_row=summary_start + 6, end_column=6)

    # 第8行：非银报量（E~F 合并）
    ws.cell(summary_start + 7, 2, "非银报量")
    ws.cell(summary_start + 7, 3, nonbank_amount if nonbank_amount else None)
    ws.merge_cells(start_row=summary_start + 7, start_column=5, end_row=summary_start + 7, end_column=6)

    # 样式：小结区使用主题浅绿底色，模板一致
    for r in range(summary_start, summary_end + 1):
        ws.row_dimensions[r].height = 24
        for c in range(1, 7):
            cell = ws.cell(r, c)
            cell.font = font_body
            cell.fill = fill_summary
            cell.alignment = left if c in (2, 5) else center

    # A 列合并（小结标签纵跨8行）
    ws.merge_cells(start_row=summary_start, start_column=1, end_row=summary_end, end_column=1)

    # 合计报量数值居中
    ws.cell(summary_start, 5).alignment = center

    max_row = summary_end

    # 确保标题、债券信息、表头、明细区、小结区都有边框；全表全框线
    apply_all_borders(ws, 1, max_row, 1, 6)

    # 认购意向（评分）列条件格式：三色交通灯图标集（1=绿灯、0=黄灯、-1=红灯）
    from openpyxl.formatting.rule import IconSetRule
    ws.conditional_formatting.add("F4:F200",
        IconSetRule(icon_style="3TrafficLights1", type="num", values=[-1, 0, 1], showValue=True, reverse=False))

    # 再次确保不冻结窗格
    ws.freeze_panes = None
    ws.sheet_view.pane = None

    # 同步生成一个“整理数据”sheet，方便复核原始六列逻辑
    data_ws = wb.create_sheet("整理数据")
    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
        data_ws.cell(1, col_idx, col_name)
    for row_num, (_, r) in enumerate(df.iterrows(), 2):
        for col_idx, col_name in enumerate(OUTPUT_COLUMNS, 1):
            if col_name == "更新询价进度":
                updated_parts = r.get("_updated_progress_parts", None)
                if isinstance(updated_parts, list):
                    data_ws.cell(row_num, col_idx).value = parts_to_excel_rich_text(updated_parts, font_name="楷体", font_size=12)
                else:
                    data_ws.cell(row_num, col_idx, clean_text(r.get(col_name, "")))
            else:
                data_ws.cell(row_num, col_idx, clean_text(r.get(col_name, "")))
    data_ws.freeze_panes = None
    data_ws.sheet_view.pane = None
    for col_idx in range(1, len(OUTPUT_COLUMNS) + 1):
        data_ws.column_dimensions[get_column_letter(col_idx)].width = 22 if col_idx not in (5, 6) else 60
    apply_all_borders(data_ws, 1, max(1, len(df) + 1), 1, len(OUTPUT_COLUMNS))
    for row in data_ws.iter_rows():
        for cell in row:
            if cell.column != 6:
                cell.font = font_body
            cell.alignment = left if cell.column in (5, 6) else center

    wb.save(output_xlsx)

def dataframe_to_markdown(df: pd.DataFrame) -> str:
    cols = OUTPUT_COLUMNS
    lines = []
    lines.append("| " + " | ".join(cols) + " |")
    lines.append("|" + "|".join(["---"] * len(cols)) + "|")
    for _, row in df.iterrows():
        values = []
        for col in cols:
            if col == "更新询价进度" and isinstance(row.get("_updated_progress_parts", None), list):
                val = parts_to_markdown(row.get("_updated_progress_parts"))
            else:
                val = clean_text(row[col]).replace("|", "｜")
            values.append(val)
        lines.append("| " + " | ".join(values) + " |")
    return "\n".join(lines)


def save_outputs(df: pd.DataFrame, output_xlsx: str, input_path: str, template_path: Optional[str]) -> None:
    output_xlsx = str(output_xlsx)

    export_formal_excel(df, output_xlsx, input_path=input_path, template_path=template_path)

    print(f"已生成正式 Excel：{output_xlsx}")


# ===================== 5. 命令行入口 =====================

def main() -> None:
    parser = argparse.ArgumentParser(description="债券/固收询价导出表自动整理工具 v6")
    parser.add_argument("input", help="原始询价导出 Excel 文件路径，支持 .xls / .xlsx")
    parser.add_argument("-o", "--output", default="", help="输出 Excel 文件路径，默认：输出询价明细表（）-YYYYMMDD.xlsx")
    parser.add_argument("-t", "--template", default="", help="询价明细表模板路径，默认读取当前目录的询价明细表模板.xlsx")
    args = parser.parse_args()

    input_path = args.input
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"找不到输入文件：{input_path}")

    default_filename = f"输出询价明细表（）-{datetime.now().strftime('%Y%m%d')}.xlsx"
    output_xlsx = args.output or default_filename
    template_path = args.template or DEFAULT_TEMPLATE_NAME
    if not os.path.exists(template_path):
        template_path = None

    result = process_file(input_path)
    save_outputs(result, output_xlsx, input_path=input_path, template_path=template_path)


if __name__ == "__main__":
    main()

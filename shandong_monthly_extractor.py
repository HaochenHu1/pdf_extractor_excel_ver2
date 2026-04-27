from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import fitz  # PyMuPDF
import pandas as pd


@dataclass
class ShandongExtractionResult:
    info_rows: List[Dict[str, Any]]
    raw_tables: Dict[str, pd.DataFrame]
    diagnostics: List[str]
    report_month: Optional[str]


SECTION_COLUMN_MAPPING = {
    "报告月份": "报告月份",
    "section": "一级章节",
    "subsection": "二级章节",
    "field": "指标名称",
    "value": "数值",
    "unit": "单位",
    "notes": "备注",
}


def parse_report_month_from_filename(filename: str) -> Optional[str]:
    base = Path(filename).stem
    m = re.search(r"(?P<y>20\d{2})\s*年\s*(?P<m>\d{1,2})\s*月", base)
    if not m:
        return None
    year = int(m.group("y"))
    month = int(m.group("m"))
    if not (1 <= month <= 12):
        return None
    return f"{year:04d}-{month:02d}"


def remove_shandong_watermarks(text: str) -> str:
    # 晶科慧能 2025年8月12日 10:35:27 / 10：35：27, tolerate spaces/newlines.
    pattern = re.compile(
        r"晶科慧能\s*"
        r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*"
        r"(?:\n|\r|\s)*"
        r"\d{1,2}\s*[：:]\s*\d{1,2}\s*[：:]\s*\d{1,2}",
        re.IGNORECASE,
    )
    return pattern.sub(" ", text)


def normalize_shandong_text_for_regex(text: str) -> str:
    text = remove_shandong_watermarks(text)
    text = text.replace("\u3000", " ").replace("\xa0", " ")
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("：", ":").replace("％", "%")
    text = text.replace("．", ".")
    # remove footnote markers immediately before digits
    text = re.sub(r"[①②③④⑤⑥⑦⑧⑨]\s*(?=\d)", "", text)
    text = re.sub(r"(?:\(|（|\[)\s*[1-9]\s*(?:\)|）|\])\s*(?=\d)", "", text)
    text = re.sub(r"(?:注|脚注)\s*[1-9]\s*(?=\d)", "", text)

    # join OCR line breaks and repeated spaces
    text = re.sub(r"[\r\n]+", " ", text)
    text = re.sub(r"\s+", " ", text)

    # remove spaces inside numeric tokens / number+unit
    text = re.sub(r"(?<=\d)\s+(?=\d)", "", text)
    text = re.sub(r"(?<=\d)\s+(?=[.])", "", text)
    text = re.sub(r"(?<=\.)\s+(?=\d)", "", text)
    text = re.sub(r"(?<=\d)\s+(?=(?:亿千瓦时|万千瓦时|万千瓦|元/兆瓦时|%))", "", text)

    return text.strip()


def _heading_pattern(heading: str) -> re.Pattern[str]:
    normalized_heading = normalize_shandong_text_for_regex(heading)
    normalized_heading = re.sub(r"\s+", "", normalized_heading)
    pattern = r"\s*".join(re.escape(ch) for ch in normalized_heading)
    return re.compile(pattern)


def slice_section(text: str, start_heading: str, end_heading_candidates: Sequence[str]) -> str:
    start_match = _heading_pattern(start_heading).search(text)
    if not start_match:
        return ""
    start = start_match.end()
    end = len(text)
    for candidate in end_heading_candidates:
        m = _heading_pattern(candidate).search(text, pos=start)
        if m:
            end = min(end, m.start())
    return text[start:end].strip()


def _field_range(field_name: Optional[str], unit: Optional[str]) -> Optional[Tuple[float, float]]:
    key = (field_name or "", unit or "")
    ranges = {
        ("全省发电装机总容量", "万千瓦"): (1000.0, 50000.0),
        ("水电装机容量", "万千瓦"): (0.0, 20000.0),
        ("核电装机容量", "万千瓦"): (0.0, 15000.0),
        ("火电装机容量", "万千瓦"): (0.0, 40000.0),
        ("风电装机容量", "万千瓦"): (0.0, 30000.0),
        ("太阳能发电装机容量", "万千瓦"): (0.0, 40000.0),
    }
    return ranges.get(key)


def clean_shandong_numeric_value(
    raw_value: Optional[str],
    field_name: Optional[str] = None,
    unit: Optional[str] = None,
    context_text: Optional[str] = None,
) -> Tuple[Optional[str], Optional[str]]:
    if raw_value is None:
        return None, None
    cleaned = raw_value.replace(",", "").strip()
    if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?", cleaned):
        return raw_value, "数值格式异常，保留原值"

    range_hint = _field_range(field_name, unit)
    raw_float = float(cleaned)
    note: Optional[str] = None

    # Footnote-merged digit correction, conservative.
    if (
        cleaned[0] in {"1", "2", "3"}
        and len(cleaned.replace(".", "")) >= 6
        and field_name is not None
        and (field_name in (context_text or "") or (unit and unit in (context_text or "")))
    ):
        candidate = cleaned[1:]
        if candidate and re.fullmatch(r"\d+(?:\.\d+)?", candidate):
            cand_float = float(candidate)
            if range_hint:
                raw_in_range = range_hint[0] <= raw_float <= range_hint[1]
                cand_in_range = range_hint[0] <= cand_float <= range_hint[1]
                if (not raw_in_range) and cand_in_range:
                    note = f"疑似脚注数字并入数值，原始识别值为{cleaned}，已修正为{candidate}"
                    return candidate, note
                if raw_in_range and not cand_in_range:
                    return cleaned, None
            else:
                # no range hint: avoid silent aggressive changes
                if raw_float >= 10 * cand_float:
                    return cleaned, f"疑似脚注数字并入数值，候选修正值{candidate}未自动应用"

    if range_hint and not (range_hint[0] <= raw_float <= range_hint[1]):
        note = f"数值{cleaned}{unit or ''}超出经验范围{range_hint[0]}~{range_hint[1]}"
    return cleaned, note


def _extract_number_near_label(
    section_text: str,
    field_name: str,
    label_patterns: Sequence[str],
    unit_patterns: Sequence[str],
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    label_part = "(?:" + "|".join(label_patterns) + ")"
    unit_part = "(?:" + "|".join(unit_patterns) + ")"
    patterns = [
        rf"{label_part}[\s\S]{{0,120}}?(?P<value>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>{unit_part})",
        rf"(?P<value>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>{unit_part})[\s\S]{{0,80}}?{label_part}",
    ]
    for pattern in patterns:
        m = re.search(pattern, section_text)
        if not m:
            continue
        raw_value = m.group("value")
        unit = m.group("unit")
        cleaned, note = clean_shandong_numeric_value(raw_value, field_name=field_name, unit=unit, context_text=section_text)
        return cleaned, unit, note
    return None, None, None


def _extract_percent_near_label(section_text: str, field_name: str, label_patterns: Sequence[str]) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    return _extract_number_near_label(section_text, field_name, label_patterns, [r"%"])


def _add_row(
    rows: List[Dict[str, Any]],
    report_month: Optional[str],
    section: str,
    subsection: str,
    field: str,
    value: Optional[str],
    unit: Optional[str],
    notes: str = "",
) -> None:
    rows.append(
        {
            "报告月份": report_month,
            "section": section,
            "subsection": subsection,
            "field": field,
            "value": value,
            "unit": unit,
            "notes": notes,
        }
    )


def build_shandong_info_dataframe(info_rows: Sequence[Dict[str, Any]]) -> pd.DataFrame:
    base = pd.DataFrame(info_rows)
    if base.empty:
        base = pd.DataFrame(columns=list(SECTION_COLUMN_MAPPING.keys()))
    for c in SECTION_COLUMN_MAPPING.keys():
        if c not in base.columns:
            base[c] = None
    return base[list(SECTION_COLUMN_MAPPING.keys())].rename(columns=SECTION_COLUMN_MAPPING)


def _extract_fields(
    section_text: str,
    report_month: Optional[str],
    section_name: str,
    subsection_name: str,
    cfg: Sequence[Tuple[str, Sequence[str], Sequence[str], bool]],
) -> Tuple[List[Dict[str, Any]], List[str]]:
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []
    for field, labels, units, is_percent in cfg:
        if is_percent:
            value, unit, note = _extract_percent_near_label(section_text, field, labels)
            if value is not None:
                unit = "%"
        else:
            value, unit, note = _extract_number_near_label(section_text, field, labels, units)
        notes = note or ("missing" if value is None else "")
        _add_row(rows, report_month, section_name, subsection_name, field, value, unit, notes)
        if value is None:
            warnings.append(f"未提取到{field}")
        elif note:
            warnings.append(note)
    return rows, warnings


def parse_shandong_power_consumption(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    cfg = [
        ("第一产业用电量", [r"第一产业(?:用电量)?"], [r"亿千瓦时", r"万千瓦时"], False),
        ("第一产业同比增长", [r"第一产业[\s\S]{0,40}?同比(?:增长)?"], [r"%"], True),
        ("第二产业用电量", [r"第二产业(?:用电量)?"], [r"亿千瓦时", r"万千瓦时"], False),
        ("第二产业同比增长", [r"第二产业[\s\S]{0,40}?同比(?:增长)?"], [r"%"], True),
        ("第三产业用电量", [r"第三产业(?:用电量)?"], [r"亿千瓦时", r"万千瓦时"], False),
        ("第三产业同比增长", [r"第三产业[\s\S]{0,40}?同比(?:增长)?"], [r"%"], True),
        ("全社会用电量", [r"全社会用电量"], [r"亿千瓦时", r"万千瓦时"], False),
        ("城乡居民生活用电量", [r"城乡居民生活用电量"], [r"亿千瓦时", r"万千瓦时"], False),
    ]
    return _extract_fields(section_text, report_month, "一、电网概览", "（一）全省全社会用电情况", cfg)


def parse_shandong_capacity_and_generation(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    cfg = [
        ("全省发电装机总容量", [r"全省发电装机总容量|全省发电装机容量|全省装机总容量"], [r"万千瓦"], False),
        ("水电装机容量", [r"水电装机(?:容量)?|水电"], [r"万千瓦"], False),
        ("核电装机容量", [r"核电装机(?:容量)?|核电"], [r"万千瓦"], False),
        ("火电装机容量", [r"火电装机(?:容量)?|火电"], [r"万千瓦"], False),
        ("风电装机容量", [r"风电装机(?:容量)?|风电"], [r"万千瓦"], False),
        ("太阳能发电装机容量", [r"太阳能发电装机(?:容量)?|光伏装机(?:容量)?"], [r"万千瓦"], False),
        ("全省发电量", [r"全省发电量"], [r"亿千瓦时", r"万千瓦时"], False),
        ("水电发电量", [r"水电发电量|水电"], [r"亿千瓦时", r"万千瓦时"], False),
        ("火电发电量", [r"火电发电量|火电"], [r"亿千瓦时", r"万千瓦时"], False),
        ("核电发电量", [r"核电发电量|核电"], [r"亿千瓦时", r"万千瓦时"], False),
        ("风电发电量", [r"风电发电量|风电"], [r"亿千瓦时", r"万千瓦时"], False),
        ("太阳能发电量", [r"太阳能发电量|光伏发电量|太阳能发电"], [r"亿千瓦时", r"万千瓦时"], False),
    ]
    return _extract_fields(section_text, report_month, "一、电网概览", "（二）全省发电机组装机及发电总体情况", cfg)


def parse_shandong_green_power_trade(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    cfg = [
        ("省内绿电交易次数", [r"省内绿电交易|组织"], [r"次"], False),
        ("新能源场站数量", [r"新能源场站"], [r"家"], False),
        ("售电公司数量", [r"售电公司"], [r"家"], False),
        ("申报电量", [r"申报电量|交易电量"], [r"亿千瓦时", r"万千瓦时"], False),
        ("成交电量", [r"成交电量|成交"], [r"亿千瓦时", r"万千瓦时"], False),
        ("环境溢价", [r"环境溢价"], [r"元/兆瓦时"], False),
    ]
    return _extract_fields(section_text, report_month, "四、交易组织情况", "（三）绿电交易组织情况", cfg)


def parse_shandong_generation_side_settlement(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    cfg = [
        ("省内发电侧共结算上网电量", [r"省内发电侧共结算上网电量|发电侧共结算上网电量"], [r"亿千瓦时", r"万千瓦时"], False),
        ("合约电量", [r"合约电量"], [r"亿千瓦时", r"万千瓦时"], False),
        ("跨省跨区交易结算电量", [r"跨省跨区交易结算电量"], [r"亿千瓦时", r"万千瓦时"], False),
        ("富余新能源外送电量", [r"富余新能源外送电量"], [r"万千瓦时", r"亿千瓦时"], False),
    ]
    return _extract_fields(section_text, report_month, "五、市场结算情况", "（一）发电侧交易结算情况", cfg)


def parse_shandong_user_side_settlement(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []

    month_value = re.search(r"(?P<value>\d{1,2})月(?:份)?", section_text)
    _add_row(rows, report_month, "五、市场结算情况", "2. 零售侧结算情况", "月份", month_value.group("value") if month_value else None, "月" if month_value else None, "" if month_value else "missing")

    cfg = [
        ("用电批发侧总结算电量", [r"用电批发侧总?结算电量|批发侧结算总体情况"], [r"亿千瓦时", r"万千瓦时"], False),
        ("零售用户数量", [r"零售用户"], [r"家"], False),
        ("售电公司数量", [r"售电公司"], [r"家"], False),
        ("虚拟电厂数量", [r"虚拟电厂"], [r"家"], False),
        ("线上签订零售合同结算电量", [r"零售合同[\s\S]{0,50}?结算电量"], [r"亿千瓦时", r"万千瓦时"], False),
        ("结算均价", [r"结算均价(?:\([^\)]{0,100}\))?"], [r"元/兆瓦时"], False),
    ]
    parsed, warns = _extract_fields(section_text, report_month, "五、市场结算情况", "（二）用电侧交易结算情况", cfg)
    rows.extend(parsed)
    warnings.extend(warns)

    if month_value is None:
        warnings.append("未提取到月份")
    return rows, warnings


def _table_to_text(table: Any) -> str:
    title = getattr(table, "title", None) or ""
    df = getattr(table, "df", None)
    table_text = ""
    if isinstance(df, pd.DataFrame) and not df.empty:
        table_text = " ".join(df.fillna("").astype(str).values.flatten().tolist())
    return normalize_shandong_text_for_regex(f"{title} {table_text}")


def _find_table_by_keywords(tables: Sequence[Any], keywords: Sequence[str]) -> Optional[pd.DataFrame]:
    for table in tables:
        text = _table_to_text(table)
        if all(keyword in text for keyword in keywords):
            return table.df.copy()
    for table in tables:
        text = _table_to_text(table)
        if any(keyword in text for keyword in keywords):
            return table.df.copy()
    return None


def parse_shandong_table_2_medium_long_term_trade(tables: Sequence[Any], diagnostics: List[str]) -> pd.DataFrame:
    # TODO: parse detailed schema for 表2 after business rules are provided.
    df = _find_table_by_keywords(tables, ["表2", "中长期交易情况"])
    if df is None:
        diagnostics.append("[WARN] 未找到表2：中长期交易情况")
        return pd.DataFrame(columns=["raw"])
    diagnostics.append("[OK] 检测到表2：中长期交易情况")
    return df


def parse_shandong_table_3_spot_trade(tables: Sequence[Any], diagnostics: List[str]) -> pd.DataFrame:
    # TODO: parse detailed schema for 表3 after business rules are provided.
    df = _find_table_by_keywords(tables, ["表3", "现货交易情况"])
    if df is None:
        diagnostics.append("[WARN] 未找到表3：现货交易情况")
        return pd.DataFrame(columns=["raw"])
    diagnostics.append("[OK] 检测到表3：现货交易情况")
    return df


def parse_shandong_table_8_market_operation_fee_settlement(tables: Sequence[Any], diagnostics: List[str]) -> pd.DataFrame:
    # TODO: parse detailed schema for 表8 after business rules are provided.
    df = _find_table_by_keywords(tables, ["表8", "市场运行费用总体结算情况"])
    if df is None:
        diagnostics.append("[WARN] 未找到表8：市场运行费用总体结算情况")
        return pd.DataFrame(columns=["raw"])
    diagnostics.append("[OK] 检测到表8：市场运行费用总体结算情况")
    return df


def extract_shandong_market_disclosure_monthly_report(
    pdf_path: str,
    text: Optional[str],
    tables: Sequence[Any],
    output_path: Optional[str] = None,
    diagnostics: Optional[List[str]] = None,
) -> ShandongExtractionResult:
    diags = diagnostics if diagnostics is not None else []

    raw_text = text
    if raw_text is None:
        doc = fitz.open(pdf_path)
        raw_text = "\n".join((page.get_text("text") or "") for page in doc)
        doc.close()

    normalized = normalize_shandong_text_for_regex(raw_text)
    report_month = parse_report_month_from_filename(Path(pdf_path).name)
    if report_month is None:
        diags.append("[WARN] 文件名未解析出报告月份")
    else:
        diags.append(f"[OK] 报告月份: {report_month}")

    rows: List[Dict[str, Any]] = []

    power_sec = slice_section(normalized, "（一）全省全社会用电情况", ["（二）全省发电机组装机及发电总体情况", "二、"])
    if power_sec:
        parsed, warns = parse_shandong_power_consumption(power_sec, report_month)
        rows.extend(parsed)
        diags.extend([f"[WARN] {w}" for w in warns])
        diags.append("[OK] 解析章节：（一）全省全社会用电情况")
    else:
        diags.append("[WARN] 未找到章节：（一）全省全社会用电情况")

    cap_gen_sec = slice_section(normalized, "（二）全省发电机组装机及发电总体情况", ["三、", "四、交易组织情况"])
    if cap_gen_sec:
        parsed, warns = parse_shandong_capacity_and_generation(cap_gen_sec, report_month)
        rows.extend(parsed)
        diags.extend([f"[WARN] {w}" for w in warns])
        diags.append("[OK] 解析章节：（二）全省发电机组装机及发电总体情况")
    else:
        diags.append("[WARN] 未找到章节：（二）全省发电机组装机及发电总体情况")

    green_sec = slice_section(normalized, "（三）绿电交易组织情况", ["五、市场结算情况", "（四）"])
    if green_sec:
        parsed, warns = parse_shandong_green_power_trade(green_sec, report_month)
        rows.extend(parsed)
        diags.extend([f"[WARN] {w}" for w in warns])
        diags.append("[OK] 解析章节：（三）绿电交易组织情况")
    else:
        diags.append("[WARN] 未找到章节：（三）绿电交易组织情况")

    gen_settle_sec = slice_section(normalized, "（一）发电侧交易结算情况", ["（二）用电侧交易结算情况", "（三）市场运行费用总体结算情况"])
    if gen_settle_sec:
        parsed, warns = parse_shandong_generation_side_settlement(gen_settle_sec, report_month)
        rows.extend(parsed)
        diags.extend([f"[WARN] {w}" for w in warns])
        diags.append("[OK] 解析章节：（一）发电侧交易结算情况")
    else:
        diags.append("[WARN] 未找到章节：（一）发电侧交易结算情况")

    user_settle_sec = slice_section(normalized, "（二）用电侧交易结算情况", ["（三）市场运行费用总体结算情况", "六、"])
    if user_settle_sec:
        parsed, warns = parse_shandong_user_side_settlement(user_settle_sec, report_month)
        rows.extend(parsed)
        diags.extend([f"[WARN] {w}" for w in warns])
        diags.append("[OK] 解析章节：（二）用电侧交易结算情况")
    else:
        diags.append("[WARN] 未找到章节：（二）用电侧交易结算情况")

    raw_tables = {
        "山东_表2_中长期交易情况_raw": parse_shandong_table_2_medium_long_term_trade(tables, diags),
        "山东_表3_现货交易情况_raw": parse_shandong_table_3_spot_trade(tables, diags),
        "山东_表8_市场运行费用_raw": parse_shandong_table_8_market_operation_fee_settlement(tables, diags),
    }

    return ShandongExtractionResult(
        info_rows=rows,
        raw_tables=raw_tables,
        diagnostics=diags,
        report_month=report_month,
    )

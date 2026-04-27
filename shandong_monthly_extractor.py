from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pandas as pd
import fitz  # PyMuPDF


@dataclass
class ShandongExtractionResult:
    info_rows: List[Dict[str, Any]]
    raw_tables: Dict[str, pd.DataFrame]
    diagnostics: List[str]
    report_month: Optional[str]


def normalize_cn_text(text: str) -> str:
    text = text.replace("\u3000", " ").replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[（(]", "（", text)
    text = re.sub(r"[）)]", "）", text)
    text = re.sub(r"[：:]", "：", text)
    return text.strip()


def parse_report_month_from_filename(filename: str) -> Optional[str]:
    base = Path(filename).stem
    m = re.search(r"(?P<y>20\d{2})\s*年\s*(?P<m>\d{1,2})\s*月", base)
    if not m:
        return None
    year = int(m.group("y"))
    month = int(m.group("m"))
    if month < 1 or month > 12:
        return None
    return f"{year:04d}-{month:02d}"


def _heading_pattern(heading: str) -> re.Pattern[str]:
    compact = re.sub(r"\s+", "", heading)
    # allow optional OCR spaces between chars
    pattern = r"\s*".join(re.escape(ch) for ch in compact)
    return re.compile(pattern)


def slice_section(text: str, start_heading: str, end_heading_candidates: Sequence[str]) -> str:
    normalized = normalize_cn_text(text)
    start_re = _heading_pattern(start_heading)
    start_match = start_re.search(normalized)
    if not start_match:
        return ""
    start = start_match.end()
    end = len(normalized)
    for cand in end_heading_candidates:
        end_re = _heading_pattern(cand)
        m = end_re.search(normalized, pos=start)
        if m:
            end = min(end, m.start())
    return normalized[start:end].strip()


def _extract_by_regex(text: str, patterns: Sequence[str]) -> Tuple[Optional[str], Optional[str]]:
    for p in patterns:
        m = re.search(p, text)
        if m:
            return m.group("value"), (m.groupdict().get("unit") or "").strip() or None
    return None, None


def extract_number_near_label(section_text: str, label_patterns: Sequence[str], unit_patterns: Sequence[str]) -> Tuple[Optional[str], Optional[str]]:
    label_part = "(?:" + "|".join(label_patterns) + ")"
    unit_part = "(?:" + "|".join(unit_patterns) + ")"
    patterns = [
        rf"{label_part}[^。；;，,\n]{{0,40}}?(?P<value>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>{unit_part})",
        rf"(?P<value>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>{unit_part})[^。；;，,\n]{{0,20}}?{label_part}",
    ]
    return _extract_by_regex(section_text, patterns)


def extract_percent_near_label(section_text: str, label_patterns: Sequence[str]) -> Tuple[Optional[str], Optional[str]]:
    return extract_number_near_label(section_text, label_patterns, [r"%", r"％"])


def _add_row(rows: List[Dict[str, Any]], report_month: Optional[str], section: str, subsection: str, field: str,
             value: Optional[str], unit: Optional[str], source_text: str, notes: str = "") -> None:
    rows.append(
        {
            "报告月份": report_month,
            "section": section,
            "subsection": subsection,
            "field": field,
            "value": value,
            "unit": unit,
            "source_text": source_text[:500],
            "notes": notes,
        }
    )


def parse_shandong_power_consumption(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []
    # Industry-specific paragraph patterns are more stable than generic nearby-label matching.
    industry_patterns = {
        "第一产业": re.search(
            r"第一产业.{0,80}?用电量.{0,20}?(?P<value>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>亿千瓦时|万千瓦时).{0,80}?同比(?:增长)?.{0,20}?(?P<yoy>[+-]?\d+(?:\.\d+)?)\s*[%％]",
            section_text,
        ),
        "第二产业": re.search(
            r"第二产业.{0,80}?用电量.{0,20}?(?P<value>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>亿千瓦时|万千瓦时).{0,80}?同比(?:增长)?.{0,20}?(?P<yoy>[+-]?\d+(?:\.\d+)?)\s*[%％]",
            section_text,
        ),
        "第三产业": re.search(
            r"第三产业.{0,80}?用电量.{0,20}?(?P<value>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>亿千瓦时|万千瓦时).{0,80}?同比(?:增长)?.{0,20}?(?P<yoy>[+-]?\d+(?:\.\d+)?)\s*[%％]",
            section_text,
        ),
    }
    for industry, m in industry_patterns.items():
        v = m.group("value") if m else None
        u = m.group("unit") if m else None
        yoy = m.group("yoy") if m else None
        _add_row(rows, report_month, "一、电网概览", "（一）全省全社会用电情况", f"{industry}用电量", v, u, section_text, "" if v else "missing")
        _add_row(rows, report_month, "一、电网概览", "（一）全省全社会用电情况", f"{industry}同比增长", yoy, "%" if yoy else None, section_text, "" if yoy else "missing")
        if v is None:
            warnings.append(f"未提取到{industry}用电量")
        if yoy is None:
            warnings.append(f"未提取到{industry}同比增长")

    mappings = [
        ("第一产业用电量", [r"第一产业(?:用电量)?"], [r"亿千瓦时", r"万千瓦时"]),
        ("第二产业用电量", [r"第二产业(?:用电量)?"], [r"亿千瓦时", r"万千瓦时"]),
        ("第三产业用电量", [r"第三产业(?:用电量)?"], [r"亿千瓦时", r"万千瓦时"]),
        ("全社会用电量", [r"全社会用电量"], [r"亿千瓦时", r"万千瓦时"]),
        ("城乡居民生活用电量", [r"城乡居民生活用电量"], [r"亿千瓦时", r"万千瓦时"]),
    ]
    for field, labels, units in mappings:
        if any(r["field"] == field and r["value"] is not None for r in rows):
            continue
        v, u = extract_number_near_label(section_text, labels, units)
        _add_row(rows, report_month, "一、电网概览", "（一）全省全社会用电情况", field, v, u, section_text,
                 "" if v is not None else "missing")
        if v is None:
            warnings.append(f"未提取到{field}")
    return rows, warnings


def parse_shandong_capacity_and_generation(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []
    cfg = [
        ("全省发电装机总容量", [r"全省发电装机总容量|全省发电装机容量|全省装机总容量"], [r"万千瓦"]),
        ("水电装机容量", [r"水电装机(?:容量)?|水电"], [r"万千瓦"]),
        ("核电装机容量", [r"核电装机(?:容量)?|核电"], [r"万千瓦"]),
        ("火电装机容量", [r"火电装机(?:容量)?|火电"], [r"万千瓦"]),
        ("风电装机容量", [r"风电装机(?:容量)?|风电"], [r"万千瓦"]),
        ("太阳能发电装机容量", [r"太阳能发电装机(?:容量)?|光伏装机(?:容量)?"], [r"万千瓦"]),
        ("全省发电量", [r"全省发电量"], [r"亿千瓦时", r"万千瓦时"]),
        ("水电发电量", [r"水电发电量|水电"], [r"亿千瓦时", r"万千瓦时"]),
        ("火电发电量", [r"火电发电量|火电"], [r"亿千瓦时", r"万千瓦时"]),
        ("核电发电量", [r"核电发电量|核电"], [r"亿千瓦时", r"万千瓦时"]),
        ("风电发电量", [r"风电发电量|风电"], [r"亿千瓦时", r"万千瓦时"]),
        ("太阳能发电量", [r"太阳能发电量|光伏发电量|太阳能发电"], [r"亿千瓦时", r"万千瓦时"]),
    ]
    for field, labels, units in cfg:
        v, u = extract_number_near_label(section_text, labels, units)
        _add_row(rows, report_month, "一、电网概览", "（二）全省发电机组装机及发电总体情况", field, v, u, section_text,
                 "" if v is not None else "missing")
        if v is None:
            warnings.append(f"未提取到{field}")
    return rows, warnings


def parse_shandong_green_power_trade(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []
    cfg = [
        ("省内绿电交易次数", [r"省内绿电交易|组织"], [r"次"]),
        ("新能源场站数量", [r"新能源场站"], [r"家"]),
        ("售电公司数量", [r"售电公司"], [r"家"]),
        ("申报电量", [r"申报电量|交易电量"], [r"亿千瓦时", r"万千瓦时"]),
        ("成交电量", [r"成交电量|成交"], [r"亿千瓦时", r"万千瓦时"]),
        ("环境溢价", [r"环境溢价"], [r"元/兆瓦时"]),
    ]
    for field, labels, units in cfg:
        v, u = extract_number_near_label(section_text, labels, units)
        _add_row(rows, report_month, "四、交易组织情况", "（三）绿电交易组织情况", field, v, u, section_text,
                 "" if v is not None else "missing")
        if v is None:
            warnings.append(f"未提取到{field}")
    return rows, warnings


def parse_shandong_generation_side_settlement(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []
    cfg = [
        ("省内发电侧共结算上网电量", [r"省内发电侧共结算上网电量|发电侧共结算上网电量"], [r"亿千瓦时", r"万千瓦时"], "（一）发电侧交易结算情况"),
        ("合约电量", [r"合约电量"], [r"亿千瓦时", r"万千瓦时"], "（一）发电侧交易结算情况"),
        ("跨省跨区交易结算电量", [r"跨省跨区交易结算电量"], [r"亿千瓦时", r"万千瓦时"], "2. 跨省跨区交易结算情况"),
        ("富余新能源外送电量", [r"富余新能源外送电量"], [r"万千瓦时", r"亿千瓦时"], "2. 跨省跨区交易结算情况"),
    ]
    for field, labels, units, subsection in cfg:
        if any(r["field"] == field and r["value"] is not None for r in rows):
            continue
        v, u = extract_number_near_label(section_text, labels, units)
        _add_row(rows, report_month, "五、市场结算情况", subsection, field, v, u, section_text, "" if v is not None else "missing")
        if v is None:
            warnings.append(f"未提取到{field}")
    return rows, warnings


def parse_shandong_user_side_settlement(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []
    month_v, _ = _extract_by_regex(section_text, [r"(?P<value>\d{1,2})\s*月(?:份)?"])
    _add_row(rows, report_month, "五、市场结算情况", "2. 零售侧结算情况", "月份", month_v, "月" if month_v else None, section_text, "" if month_v else "missing")
    user_count = _extract_by_regex(section_text, [r"(?P<value>\d+(?:\.\d+)?)\s*家零售用户"])[0]
    retailer_count = _extract_by_regex(section_text, [r"(?P<value>\d+(?:\.\d+)?)\s*家售电公司"])[0]
    vpp_count = _extract_by_regex(section_text, [r"(?P<value>\d+(?:\.\d+)?)\s*家虚拟电厂"])[0]
    _add_row(rows, report_month, "五、市场结算情况", "2. 零售侧结算情况", "零售用户数量", user_count, "家" if user_count else None, section_text, "" if user_count else "missing")
    _add_row(rows, report_month, "五、市场结算情况", "2. 零售侧结算情况", "售电公司数量", retailer_count, "家" if retailer_count else None, section_text, "" if retailer_count else "missing")
    _add_row(rows, report_month, "五、市场结算情况", "2. 零售侧结算情况", "虚拟电厂数量", vpp_count, "家" if vpp_count else None, section_text, "" if vpp_count else "missing")
    retail_energy = re.search(
        r"零售合同.{0,30}?结算电量\s*(?P<value>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>亿千瓦时|万千瓦时)",
        section_text,
    )
    retail_price = re.search(
        r"结算均价(?:（[^）]{0,100}）)?\s*(?P<value>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>元/兆瓦时)",
        section_text,
    )
    _add_row(
        rows,
        report_month,
        "五、市场结算情况",
        "2. 零售侧结算情况",
        "线上签订零售合同结算电量",
        retail_energy.group("value") if retail_energy else None,
        retail_energy.group("unit") if retail_energy else None,
        section_text,
        "" if retail_energy else "missing",
    )
    _add_row(
        rows,
        report_month,
        "五、市场结算情况",
        "2. 零售侧结算情况",
        "结算均价",
        retail_price.group("value") if retail_price else None,
        retail_price.group("unit") if retail_price else None,
        section_text,
        "" if retail_price else "missing",
    )

    cfg = [
        ("用电批发侧总结算电量", [r"用电批发侧总?结算电量|批发侧结算总体情况"], [r"亿千瓦时", r"万千瓦时"], "1. 批发侧结算总体情况"),
    ]
    for field, labels, units, subsection in cfg:
        v, u = extract_number_near_label(section_text, labels, units)
        _add_row(rows, report_month, "五、市场结算情况", subsection, field, v, u, section_text, "" if v is not None else "missing")
        if v is None:
            warnings.append(f"未提取到{field}")
    if month_v is None:
        warnings.append("未提取到月份")
    if user_count is None:
        warnings.append("未提取到零售用户数量")
    if retailer_count is None:
        warnings.append("未提取到售电公司数量")
    if vpp_count is None:
        warnings.append("未提取到虚拟电厂数量")
    if retail_energy is None:
        warnings.append("未提取到线上签订零售合同结算电量")
    if retail_price is None:
        warnings.append("未提取到结算均价")
    return rows, warnings


def _table_to_text(table: Any) -> str:
    title = (getattr(table, "title", None) or "")
    df = getattr(table, "df", None)
    table_text = ""
    if isinstance(df, pd.DataFrame) and not df.empty:
        table_text = " ".join(df.fillna("").astype(str).values.flatten().tolist())
    return normalize_cn_text(f"{title} {table_text}")


def _find_table_by_keywords(tables: Sequence[Any], keywords: Sequence[str]) -> Optional[pd.DataFrame]:
    for table in tables:
        txt = _table_to_text(table)
        if all(k in txt for k in keywords):
            return table.df.copy()
    for table in tables:
        txt = _table_to_text(table)
        if any(k in txt for k in keywords):
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

    normalized = normalize_cn_text(raw_text)
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

    cap_gen_sec = slice_section(normalized, "（二）全省发电机组装机及发电总体情况", ["二、", "三、", "四、交易组织情况"])
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

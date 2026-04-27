from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import fitz  # PyMuPDF
import pandas as pd
try:
    import numpy as np
except Exception:  # pragma: no cover - optional dependency for image preprocessing
    np = None


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


def normalize_shandong_readable_text(text: str) -> str:
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


def compact_shandong_text_for_matching(text: str) -> str:
    """
    OCR often splits semantic labels across physical lines (e.g. '用电\\n量', '太阳能\\n发电量').
    We keep readable text for diagnostics, but field matching uses this compact text.
    """
    compact = normalize_shandong_readable_text(text)
    compact = re.sub(r"[，,。；;：:（）()【】\\[\\]、\\-—_]+", "", compact)
    compact = re.sub(r"\s+", "", compact)
    return compact


def normalize_shandong_text_for_regex(text: str) -> str:
    # Backward-compatible alias used by existing callers.
    return normalize_shandong_readable_text(text)


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


def extract_from_compact_text(
    compact_text: str,
    label_variants: Sequence[str],
    unit_pattern: str,
    field_name: str,
    readable_context: Optional[str] = None,
) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[int]]:
    for label in label_variants:
        pattern = rf"{label}(?P<value>[+-]?\d+(?:\.\d+)?)(?P<unit>{unit_pattern})"
        m = re.search(pattern, compact_text)
        if not m:
            continue
        raw_value = m.group("value")
        unit = m.group("unit")
        cleaned, note = clean_shandong_numeric_value(
            raw_value,
            field_name=field_name,
            unit=unit,
            context_text=readable_context or compact_text,
        )
        return cleaned, unit, note, m.end()
    return None, None, None, None


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


def _find_keyword_excerpt(text: str, keywords: Sequence[str], window: int = 40) -> str:
    for keyword in keywords:
        m = re.search(re.escape(keyword), text)
        if m:
            start = max(0, m.start() - window)
            end = min(len(text), m.end() + window)
            return text[start:end]
    return text[: min(len(text), 120)]


def parse_shandong_power_consumption(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    readable_text = normalize_shandong_readable_text(section_text)
    compact_text = compact_shandong_text_for_matching(section_text)
    cfg = [
        ("第一产业用电量", [r"第一产业(?:用电量)?"], [r"亿千瓦时", r"万千瓦时"], False),
        ("第一产业同比增长", [r"第一产业[\s\S]{0,40}?同比(?:增长)?"], [r"%"], True),
        ("第二产业用电量", [r"第二产业(?:用电量)?"], [r"亿千瓦时", r"万千瓦时"], False),
        ("第二产业同比增长", [r"第二产业[\s\S]{0,40}?同比(?:增长)?"], [r"%"], True),
        ("第三产业用电量", [r"第三产业(?:用电量)?"], [r"亿千瓦时", r"万千瓦时"], False),
        ("第三产业同比增长", [r"第三产业[\s\S]{0,40}?同比(?:增长)?"], [r"%"], True),
        ("全社会用电量", [r"全社会用电量"], [r"亿千瓦时", r"万千瓦时"], False),
        ("城乡居民生活用电量", [r"城乡居民生活\s*用电量|居民生活\s*用电量|城乡居民用电量"], [r"亿千瓦时", r"万千瓦时"], False),
    ]
    rows, warnings = _extract_fields(readable_text, report_month, "一、电网概览", "（一）全省全社会用电情况", cfg)

    # Compact matching for OCR split labels: 城乡居民生活用电\n量 / 居民生活用电\n量
    resident_labels = [r"城乡居民生活用电量", r"居民生活用电量", r"城乡居民用电量"]
    value, unit, note, end_pos = extract_from_compact_text(
        compact_text,
        resident_labels,
        r"(?:亿千瓦时|万千瓦时)",
        field_name="城乡居民生活用电量",
        readable_context=readable_text,
    )
    resident_value_row = next((r for r in rows if r["field"] == "城乡居民生活用电量"), None)
    if resident_value_row and value is not None:
        resident_value_row["value"] = value
        resident_value_row["unit"] = unit
        resident_value_row["notes"] = note or ""
        warnings = [w for w in warnings if w != "未提取到城乡居民生活用电量"]

    yoy_value, yoy_unit = None, None
    if end_pos is not None:
        tail = compact_text[end_pos : min(len(compact_text), end_pos + 80)]
        yoy_match = re.search(r"(?:同比增长|同比|增长)(?P<value>[+-]?\d+(?:\.\d+)?)%", tail)
        if yoy_match:
            yoy_value = yoy_match.group("value")
            yoy_unit = "%"

    existing_yoy_row = next((r for r in rows if r["field"] == "城乡居民生活用电量同比增长"), None)
    if existing_yoy_row is None:
        _add_row(
            rows,
            report_month,
            "一、电网概览",
            "（一）全省全社会用电情况",
            "城乡居民生活用电量同比增长",
            yoy_value,
            yoy_unit,
            "" if yoy_value is not None else "missing",
        )
    elif yoy_value is not None:
        existing_yoy_row["value"] = yoy_value
        existing_yoy_row["unit"] = yoy_unit
        existing_yoy_row["notes"] = ""

    if yoy_value is None:
        if "未提取到城乡居民生活用电量同比增长" not in warnings:
            warnings.append("未提取到城乡居民生活用电量同比增长")
    else:
        warnings = [w for w in warnings if w != "未提取到城乡居民生活用电量同比增长"]
    return rows, warnings


def parse_shandong_capacity_and_generation(section_text: str, report_month: Optional[str]) -> Tuple[List[Dict[str, Any]], List[str]]:
    readable_text = normalize_shandong_readable_text(section_text)
    generation_anchor = re.search(r"全省发电量", readable_text)
    if generation_anchor:
        installed_capacity_text = readable_text[: generation_anchor.start()]
        generation_text = readable_text[generation_anchor.start() :]
    else:
        installed_capacity_text = readable_text
        generation_text = readable_text

    installed_cfg = [
        ("全省发电装机总容量", [r"全省发电装机总容量|全省发电装机容量|全省装机总容量"], [r"万千瓦"], False),
        ("水电装机容量", [r"水电装机(?:容量)?|水电"], [r"万千瓦"], False),
        ("核电装机容量", [r"核电装机(?:容量)?|核电"], [r"万千瓦"], False),
        ("火电装机容量", [r"火电装机(?:容量)?|火电"], [r"万千瓦"], False),
        ("风电装机容量", [r"风电装机(?:容量)?|风电"], [r"万千瓦"], False),
        ("太阳能发电装机容量", [r"太阳能发电装机(?:容量)?|太阳能发电|太阳能|光伏(?:装机(?:容量)?)?"], [r"万千瓦"], False),
    ]
    generation_cfg = [
        ("全省发电量", [r"全省发电量"], [r"亿千瓦时", r"万千瓦时"], False),
        ("水电发电量", [r"水电发电量|水电"], [r"亿千瓦时", r"万千瓦时"], False),
        ("火电发电量", [r"火电发电量|火电"], [r"亿千瓦时", r"万千瓦时"], False),
        ("核电发电量", [r"核电发电量|核电"], [r"亿千瓦时", r"万千瓦时"], False),
        ("风电发电量", [r"风电发电量|风电"], [r"亿千瓦时", r"万千瓦时"], False),
        ("太阳能发电量", [r"太阳能发电量|光伏发电量|太阳能发电|光伏"], [r"亿千瓦时", r"万千瓦时"], False),
    ]
    rows1, warns1 = _extract_fields(
        installed_capacity_text,
        report_month,
        "一、电网概览",
        "（二）全省发电机组装机及发电总体情况",
        installed_cfg,
    )
    rows2, warns2 = _extract_fields(
        generation_text,
        report_month,
        "一、电网概览",
        "（二）全省发电机组装机及发电总体情况",
        generation_cfg,
    )
    # Parse generation with compact label binding so each energy source binds to nearest own value.
    generation_compact = compact_shandong_text_for_matching(generation_text)
    energy_label_map = {
        "水电发电量": [r"水电发电量", r"水电"],
        "火电发电量": [r"火电发电量", r"火电"],
        "核电发电量": [r"核电发电量", r"核电"],
        "风电发电量": [r"风电发电量", r"风电"],
        "太阳能发电量": [r"太阳能发电量", r"光伏发电量", r"太阳能发电", r"太阳能", r"光伏"],
    }
    for field_name, labels in energy_label_map.items():
        value, unit, note, _ = extract_from_compact_text(
            generation_compact,
            labels,
            r"(?:亿千瓦时|万千瓦时)",
            field_name=field_name,
            readable_context=generation_text,
        )
        row = next((r for r in rows2 if r["field"] == field_name), None)
        if row and value is not None:
            row["value"] = value
            row["unit"] = unit
            row["notes"] = note or ""
            warns2 = [w for w in warns2 if w != f"未提取到{field_name}"]
    return rows1 + rows2, warns1 + warns2


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
    return rows, warnings


def _table_to_text(table: Any) -> str:
    title = getattr(table, "title", None) or ""
    df = getattr(table, "df", None)
    table_text = ""
    if isinstance(df, pd.DataFrame) and not df.empty:
        table_text = " ".join(df.fillna("").astype(str).values.flatten().tolist())
    return normalize_shandong_text_for_regex(f"{title} {table_text}")


def _clean_table_cell(value: Any) -> str:
    text = normalize_shandong_readable_text("" if value is None else str(value))
    # Remove watermark fragments that leak into OCR/table cells.
    text = re.sub(r"晶科慧能", "", text)
    text = re.sub(r"20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*\d{1,2}\s*[:：]\s*\d{1,2}\s*[:：]\s*\d{1,2}", "", text)
    text = re.sub(r"\b20\d{2}[/-]\d{1,2}[/-]\d{1,2}\s+\d{1,2}[:：]\d{1,2}[:：]\d{1,2}\b", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _normalize_header_text(text: str) -> str:
    text = _clean_table_cell(text)
    text = text.replace(" ", "")
    return text


def _clean_table_df(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned = cleaned.map(_clean_table_cell)
    non_empty_rows = cleaned.apply(lambda r: any(bool(str(x).strip()) for x in r), axis=1)
    cleaned = cleaned.loc[non_empty_rows].reset_index(drop=True)
    non_empty_cols = [c for c in cleaned.columns if any(bool(str(v).strip()) for v in cleaned[c])]
    if non_empty_cols:
        cleaned = cleaned[non_empty_cols]
    return cleaned.reset_index(drop=True)


def preprocess_shandong_table_image_for_watermark(page_image_or_crop: Any) -> Any:
    """
    Image-level watermark suppression for light-gray diagonal overlays before OCR/table recognition.
    Keeps dark text/grid pixels and whitens mid-gray watermark-like pixels.
    """
    if np is None or page_image_or_crop is None:
        return page_image_or_crop
    image = np.array(page_image_or_crop).copy()
    if image.ndim == 3:
        gray = image.mean(axis=2).astype(np.uint8)
    else:
        gray = image.astype(np.uint8)
    # Keep dark printed content; attenuate light-gray watermark layer only.
    dark_mask = gray < 140
    watermark_band = (gray >= 150) & (gray <= 225)
    cleaned_gray = gray.copy()
    cleaned_gray[watermark_band] = 255
    cleaned_gray[dark_mask] = gray[dark_mask]
    if image.ndim == 3:
        cleaned = np.stack([cleaned_gray] * 3, axis=2)
        return cleaned.astype(np.uint8)
    return cleaned_gray.astype(np.uint8)


def _find_table_candidates_by_keywords(tables: Sequence[Any], keywords: Sequence[str]) -> List[Any]:
    matched: List[Any] = []
    for table in tables:
        text = _table_to_text(table)
        if all(keyword in text for keyword in keywords):
            matched.append(table)
    if matched:
        return sorted(matched, key=lambda t: int(getattr(t, "page", 10**9)))
    for table in tables:
        text = _table_to_text(table)
        if any(keyword in text for keyword in keywords):
            matched.append(table)
    return sorted(matched, key=lambda t: int(getattr(t, "page", 10**9)))


def _extract_row_values(row: Sequence[Any]) -> List[str]:
    return [_clean_table_cell(v) for v in row if _clean_table_cell(v)]


def _detect_table2_header_row(df: pd.DataFrame) -> int:
    for i in range(min(len(df), 8)):
        row_text = "".join(_extract_row_values(df.iloc[i].tolist()))
        if "交易品种" in row_text and ("合约" in row_text or "电价" in row_text):
            return i
    return 0


def parse_shandong_table_2_cumulative_trade_only(
    table_candidates: Sequence[Any],
    page_images: Optional[Any],
    ocr_text: Optional[str],
    diagnostics: List[str],
) -> pd.DataFrame:
    candidates = _find_table_candidates_by_keywords(table_candidates, ["表2", "中长期交易情况"])
    if not candidates:
        diagnostics.append("[WARN] 未找到表2：中长期交易情况")
        return pd.DataFrame(columns=["交易品种", "累计合约量", "加权平均电价"])

    selected = candidates[0]
    diagnostics.append(f"[OK] 表2标题候选页: {getattr(selected, 'page', 'unknown')}")
    diagnostics.append("[INFO] 表2水印图像预处理: " + ("已启用" if page_images is not None else "未提供页面图像"))
    df = _clean_table_df(selected.df)
    if df.empty:
        diagnostics.append("[WARN] 表2候选表为空")
        return pd.DataFrame(columns=["交易品种", "累计合约量", "加权平均电价"])

    header_idx = _detect_table2_header_row(df)
    raw_headers = [_normalize_header_text(v) for v in df.iloc[header_idx].tolist()]
    headers = [h if h else f"列{i+1}" for i, h in enumerate(raw_headers)]
    data_rows: List[List[str]] = []
    repeated_header_count = 0
    found_upper_marker = False
    found_lower_marker = False
    for i in range(header_idx + 1, len(df)):
        row = [_clean_table_cell(v) for v in df.iloc[i].tolist()]
        row_text = "".join(x for x in row if x)
        row_compact = compact_shandong_text_for_matching(row_text)
        if "（一）中长期累计交易情况" in row_text or "(一)中长期累计交易情况" in row_text or "一中长期累计交易情况" in row_compact:
            found_upper_marker = True
            continue
        if not row_text or row_text.startswith("单位") or row_text.startswith("表2"):
            continue
        if (
            "（二）中长期交易历史净合约情况" in row_text
            or "(二)中长期交易历史净合约情况" in row_text
            or "二中长期交易历史净合约情况" in row_compact
            or row_compact.startswith("日期")
        ) and data_rows:
            found_lower_marker = True
            diagnostics.append(f"[INFO] 表2在行{i}检测到下半区边界并停止")
            break
        if ("（二）" in row_text or "现货" in row_text) and data_rows:
            diagnostics.append(f"[INFO] 表2上半区截断于行{i}（疑似下半区起始）")
            found_lower_marker = True
            break
        if row_text.replace(" ", "") == "".join(headers).replace(" ", ""):
            repeated_header_count += 1
            if repeated_header_count >= 1 and data_rows:
                diagnostics.append(f"[INFO] 表2上半区在重复表头行{i}处停止")
                break
            continue
        if any(row):
            data_rows.append(row[: len(headers)] + [""] * max(0, len(headers) - len(row)))

    out = pd.DataFrame(data_rows, columns=headers)
    diagnostics.append(f"[INFO] 表2检测到（一）中长期累计交易情况: {'是' if found_upper_marker else '否'}")
    diagnostics.append(f"[INFO] 表2检测到（二）中长期交易历史净合约情况并作为边界: {'是' if found_lower_marker else '否'}")
    diagnostics.append(f"[OK] 表2上半区提取行数: {len(out)}")
    return out


def is_table3_continuation_page(previous_table_context: Dict[str, Any], current_page_text: str, current_page_tables: Sequence[Any]) -> bool:
    if not previous_table_context.get("inside_table3"):
        return False
    compact_page = compact_shandong_text_for_matching(current_page_text or "")
    if re.search(r"表[4-9]", compact_page):
        return False
    has_date_rows = len(re.findall(r"\d{1,2}月\d{1,2}日", compact_page)) >= 2 or ("合计" in compact_page)
    if not has_date_rows:
        return False
    expected_cols = int(previous_table_context.get("expected_cols", 6))
    compatible = False
    for t in current_page_tables:
        df = getattr(t, "df", None)
        if isinstance(df, pd.DataFrame) and not df.empty and abs(df.shape[1] - expected_cols) <= 2:
            compatible = True
            break
    return compatible or not current_page_tables


def _table3_row_anchor(text: str) -> Optional[str]:
    m = re.search(r"(\d{1,2}\s*月\s*\d{1,2}\s*日)", text)
    if m:
        return re.sub(r"\s+", "", m.group(1))
    if "合计" in text:
        return "合计"
    return None


def _normalize_table3_columns(columns: Sequence[str]) -> List[str]:
    fallback = ["日期", "发电侧日前出清电量", "用电侧日前出清电量", "日前出清均价", "发电侧实时出清电量", "实时出清均价"]
    cleaned = [_normalize_header_text(c) for c in columns]
    if any("日期" in c for c in cleaned):
        return fallback
    return fallback


def parse_shandong_table_3_spot_trade_across_pages(
    table_candidates: Sequence[Any],
    page_images: Optional[Any],
    ocr_text: Optional[str],
    diagnostics: List[str],
    report_month: Optional[str] = None,
) -> pd.DataFrame:
    candidates = _find_table_candidates_by_keywords(table_candidates, ["表3", "现货交易情况"])
    if not candidates:
        diagnostics.append("[WARN] 未找到表3：现货交易情况")
        return pd.DataFrame(columns=["日期", "发电侧日前出清电量", "用电侧日前出清电量", "日前出清均价", "发电侧实时出清电量", "实时出清均价"])

    title_table = candidates[0]
    start_page = int(getattr(title_table, "page", 1))
    diagnostics.append(f"[OK] 表3标题页: {start_page}")
    diagnostics.append("[INFO] 表3水印图像预处理: " + ("已启用" if page_images is not None else "未提供页面图像"))
    all_sorted = sorted(table_candidates, key=lambda t: int(getattr(t, "page", 10**9)))
    by_page: Dict[int, List[Any]] = {}
    for t in all_sorted:
        by_page.setdefault(int(getattr(t, "page", 0)), []).append(t)

    continuation_pages: List[int] = []
    included_tables: List[Any] = [title_table]
    previous_ctx = {"inside_table3": True, "expected_cols": 6}
    max_page = max(by_page.keys()) if by_page else start_page
    for p in range(start_page + 1, max_page + 1):
        page_tables = by_page.get(p, [])
        page_text = " ".join(_table_to_text(t) for t in page_tables)
        if is_table3_continuation_page(previous_ctx, page_text, page_tables):
            continuation_pages.append(p)
            if page_tables:
                included_tables.extend(page_tables)
        else:
            if re.search(r"表[4-9]", compact_shandong_text_for_matching(page_text)):
                break
    if continuation_pages:
        diagnostics.append(f"[OK] 表3续页: {continuation_pages}")
    else:
        diagnostics.append("[INFO] 表3未检测到续页")

    normalized_cols = _normalize_table3_columns(["日期", "发电侧日前出清电量", "用电侧日前出清电量", "日前出清均价", "发电侧实时出清电量", "实时出清均价"])
    merged_rows: Dict[str, List[str]] = {}
    duplicate_dates: List[str] = []
    short_rows = 0
    for t in included_tables:
        df = _clean_table_df(t.df)
        if df.empty:
            continue
        for i in range(len(df)):
            vals = [_clean_table_cell(v) for v in df.iloc[i].tolist()]
            row_text = "".join(v for v in vals if v)
            if not row_text or "单位" in row_text or "表3" in row_text or "现货交易情况" in row_text:
                continue
            if "日期" in row_text and "出清" in row_text:
                continue
            anchor = _table3_row_anchor(row_text)
            if anchor is None:
                if merged_rows:
                    last_key = list(merged_rows.keys())[-1]
                    for j, v in enumerate(vals[:6]):
                        if v and (j >= len(merged_rows[last_key]) or not merged_rows[last_key][j]):
                            if j >= len(merged_rows[last_key]):
                                merged_rows[last_key].extend([""] * (j + 1 - len(merged_rows[last_key])))
                            merged_rows[last_key][j] = v
                continue
            row_out = [anchor] + vals[1:6]
            if len(vals) < 6:
                short_rows += 1
            row_out = row_out[:6] + [""] * max(0, 6 - len(row_out))
            if anchor in merged_rows:
                duplicate_dates.append(anchor)
                base = merged_rows[anchor]
                for j in range(1, 6):
                    if not base[j] and row_out[j]:
                        base[j] = row_out[j]
            else:
                merged_rows[anchor] = row_out

    daily_keys = [k for k in merged_rows.keys() if re.match(r"\d{1,2}月\d{1,2}日", k)]
    def _sort_key(d: str) -> int:
        m = re.match(r"(\d{1,2})月(\d{1,2})日", d)
        return int(m.group(2)) if m else 999
    ordered = sorted(daily_keys, key=_sort_key)
    final_rows = [merged_rows[k] for k in ordered]
    if "合计" in merged_rows:
        final_rows.append(merged_rows["合计"])

    missing_dates: List[str] = []
    if report_month:
        y, m = report_month.split("-")
        month = int(m)
        if month in {1, 3, 5, 7, 8, 10, 12}:
            days = 31
        elif month == 2:
            days = 29 if (int(y) % 4 == 0 and (int(y) % 100 != 0 or int(y) % 400 == 0)) else 28
        else:
            days = 30
        expected = {f"{month:02d}月{d:02d}日" for d in range(1, days + 1)}
        missing_dates = sorted(expected - set(ordered), key=_sort_key)
    diagnostics.append(f"[OK] 表3日数据行数: {len(ordered)}")
    diagnostics.append(f"[OK] 表3合计行: {'是' if '合计' in merged_rows else '否'}")
    diagnostics.append(f"[INFO] 表3重复日期: {sorted(set(duplicate_dates)) if duplicate_dates else []}")
    diagnostics.append(f"[INFO] 表3缺失日期: {missing_dates}")
    diagnostics.append(f"[INFO] 表3列数不足行数: {short_rows}")

    return pd.DataFrame(final_rows, columns=normalized_cols)


def parse_shandong_table_8_market_operation_fee_settlement(
    table_candidates: Sequence[Any],
    page_images: Optional[Any],
    ocr_text: Optional[str],
    diagnostics: List[str],
) -> pd.DataFrame:
    candidates = _find_table_candidates_by_keywords(table_candidates, ["表8", "市场运行费用总体结算情况"])
    if not candidates:
        diagnostics.append("[WARN] 未找到表8：市场运行费用总体结算情况")
        return pd.DataFrame(columns=["序号", "类别", "费用总额", "分摊返还均价", "分摊返还主体"])
    selected = candidates[0]
    diagnostics.append(f"[OK] 表8标题候选页: {getattr(selected, 'page', 'unknown')}")
    diagnostics.append("[INFO] 表8水印图像预处理: " + ("已启用" if page_images is not None else "未提供页面图像"))
    df = _clean_table_df(selected.df)
    rows: List[Dict[str, str]] = []
    current: Optional[Dict[str, str]] = None
    header_only = True
    for i in range(len(df)):
        vals = [_clean_table_cell(v) for v in df.iloc[i].tolist()]
        row_text = "".join(v for v in vals if v)
        if not row_text or row_text.startswith("单位") or row_text.startswith("表8"):
            continue
        if "序号" in row_text and "类别" in row_text:
            continue
        serial_match = re.match(r"^\s*(\d{1,2})\s*$", vals[0] if vals else "")
        serial_with_text = re.match(r"^\s*(\d{1,2})\s*(.+)$", vals[0] if vals else "")
        if serial_match:
            header_only = False
            if current:
                rows.append(current)
            current = {
                "序号": serial_match.group(1),
                "类别": vals[1] if len(vals) > 1 else "",
                "费用总额": vals[2] if len(vals) > 2 else "",
                "分摊返还均价": vals[3] if len(vals) > 3 else "",
                "分摊返还主体": vals[4] if len(vals) > 4 else "",
            }
        elif serial_with_text:
            header_only = False
            if current:
                rows.append(current)
            current = {
                "序号": serial_with_text.group(1),
                "类别": serial_with_text.group(2).strip(),
                "费用总额": vals[1] if len(vals) > 1 else "",
                "分摊返还均价": vals[2] if len(vals) > 2 else "",
                "分摊返还主体": vals[3] if len(vals) > 3 else "",
            }
        else:
            if current is None:
                continue
            tail = " ".join(v for v in vals if v)
            if tail:
                if len(vals) > 1 and vals[1]:
                    current["类别"] = (current["类别"] + " " + vals[1]).strip()
                if len(vals) > 4 and vals[4]:
                    current["分摊返还主体"] = (current["分摊返还主体"] + " " + vals[4]).strip()
                elif len(vals) > 2 and vals[2]:
                    current["分摊返还主体"] = (current["分摊返还主体"] + " " + tail).strip()
    if current:
        rows.append(current)

    fallback_used = False
    if not rows:
        diagnostics.append("[WARN] 表8首轮提取疑似仅表头或空，启用序号锚点回退解析")
        header_only = True
        fallback_used = True
        current = None
        for i in range(len(df)):
            vals = [_clean_table_cell(v) for v in df.iloc[i].tolist()]
            row_text = " ".join(v for v in vals if v).strip()
            if not row_text or row_text.startswith("单位") or row_text.startswith("表8"):
                continue
            if "序号" in row_text and "类别" in row_text:
                continue
            serial_line = re.match(
                r"^\s*(?P<serial>\d{1,2})\s*(?P<category>.*?)(?:\s+(?P<amount>-?\d+(?:\.\d+)?))?(?:\s+(?P<avg>-?\d+(?:\.\d+)?))?(?:\s+(?P<subject>.*))?$",
                row_text,
            )
            if serial_line:
                header_only = False
                if current:
                    rows.append(current)
                current = {
                    "序号": serial_line.group("serial") or "",
                    "类别": (serial_line.group("category") or "").strip(),
                    "费用总额": (serial_line.group("amount") or "").strip(),
                    "分摊返还均价": (serial_line.group("avg") or "").strip(),
                    "分摊返还主体": (serial_line.group("subject") or "").strip(),
                }
            elif current is not None:
                current["类别"] = (current["类别"] + " " + row_text).strip()
        if current:
            rows.append(current)

    out = pd.DataFrame(rows, columns=["序号", "类别", "费用总额", "分摊返还均价", "分摊返还主体"])
    serials = [int(s) for s in out["序号"].tolist() if str(s).isdigit()]
    missing_serials = [str(i) for i in range(1, 13) if i not in serials]
    duplicated = sorted({s for s in serials if serials.count(s) > 1})
    diagnostics.append(f"[INFO] 表8首轮是否仅表头: {'是' if header_only else '否'}")
    diagnostics.append(f"[INFO] 表8是否使用回退序号锚点解析: {'是' if fallback_used else '否'}")
    diagnostics.append(f"[OK] 表8提取行数: {len(out)}")
    diagnostics.append(f"[INFO] 表8缺失序号: {missing_serials}")
    diagnostics.append(f"[INFO] 表8重复序号: {duplicated}")
    return out


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
        if any(row["field"] == "城乡居民生活用电量" and row["value"] is None for row in parsed):
            diags.append(
                "[DEBUG] 城乡居民生活用电量邻近文本(可读): "
                + _find_keyword_excerpt(power_sec, ["城乡居民", "居民生活"])
            )
            diags.append(
                "[DEBUG] 城乡居民生活用电量邻近文本(紧凑): "
                + _find_keyword_excerpt(compact_shandong_text_for_matching(power_sec), ["城乡居民", "居民生活"])
            )
        if any(row["field"] == "城乡居民生活用电量同比增长" and row["value"] is None for row in parsed):
            diags.append(
                "[DEBUG] 城乡居民生活用电量同比增长邻近文本(可读): "
                + _find_keyword_excerpt(power_sec, ["城乡居民", "居民生活", "同比"])
            )
            diags.append(
                "[DEBUG] 城乡居民生活用电量同比增长邻近文本(紧凑): "
                + _find_keyword_excerpt(compact_shandong_text_for_matching(power_sec), ["城乡居民", "居民生活", "同比"])
            )
        diags.append("[OK] 解析章节：（一）全省全社会用电情况")
    else:
        diags.append("[WARN] 未找到章节：（一）全省全社会用电情况")

    cap_gen_sec = slice_section(normalized, "（二）全省发电机组装机及发电总体情况", ["三、", "四、交易组织情况"])
    if cap_gen_sec:
        parsed, warns = parse_shandong_capacity_and_generation(cap_gen_sec, report_month)
        rows.extend(parsed)
        diags.extend([f"[WARN] {w}" for w in warns])
        if any(row["field"] == "太阳能发电装机容量" and row["value"] is None for row in parsed):
            diags.append(
                "[DEBUG] 太阳能发电装机容量邻近文本(可读): "
                + _find_keyword_excerpt(cap_gen_sec, ["太阳能", "光伏"])
            )
            diags.append(
                "[DEBUG] 太阳能发电装机容量邻近文本(紧凑): "
                + _find_keyword_excerpt(compact_shandong_text_for_matching(cap_gen_sec), ["太阳能", "光伏"])
            )
        if any(row["field"] == "太阳能发电量" and row["value"] is None for row in parsed):
            diags.append(
                "[DEBUG] 太阳能发电量邻近文本(可读): "
                + _find_keyword_excerpt(cap_gen_sec, ["太阳能", "光伏", "发电量"])
            )
            diags.append(
                "[DEBUG] 太阳能发电量邻近文本(紧凑): "
                + _find_keyword_excerpt(compact_shandong_text_for_matching(cap_gen_sec), ["太阳能", "光伏", "发电量"])
            )
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
        "山东_表2_中长期交易情况": parse_shandong_table_2_cumulative_trade_only(
            tables, None, normalized, diags
        ),
        "山东_表3_现货交易情况": parse_shandong_table_3_spot_trade_across_pages(
            tables, None, normalized, diags, report_month=report_month
        ),
        "山东_表8_市场运行费用": parse_shandong_table_8_market_operation_fee_settlement(
            tables, None, normalized, diags
        ),
    }

    return ShandongExtractionResult(
        info_rows=rows,
        raw_tables=raw_tables,
        diagnostics=diags,
        report_month=report_month,
    )

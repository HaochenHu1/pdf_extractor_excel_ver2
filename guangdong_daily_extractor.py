from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pandas as pd


@dataclass
class GuangdongDailyExtractionResult:
    report_type: str
    operation_date: Optional[str]
    market_rows: List[Dict[str, Any]]
    table1_volume_rows: List[Dict[str, Any]]
    table1_price_rows: List[Dict[str, Any]]
    table2_day_ahead_price_rows: List[Dict[str, Any]]
    table2_realtime_price_rows: List[Dict[str, Any]]
    diagnostics: List[str]


def normalize_chinese_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", "" if text is None else str(text)).strip()


def is_guangdong_daily_report(filename: str) -> bool:
    normalized = normalize_chinese_whitespace(filename).replace("（", "(").replace("）", ")")
    pattern = (
        r"^广东电力现货市场.*"
        r"\d{4}年\d{1,2}月"
        r".*运行日报"
        r".*\(\d{1,2}\.\d{1,2}\)"
        r"\.pdf$"
    )
    return bool(re.match(pattern, normalized, flags=re.IGNORECASE))


def _normalize_date(y: str, m: str, d: str) -> str:
    return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"


def extract_daily_report_operation_date(filename: str, text: str) -> Optional[str]:
    normalized = filename.replace("（", "(").replace("）", ")")
    m = re.search(r"\((\d{1,2})\.(\d{1,2})\)", normalized)
    ym = re.search(r"(\d{4})年(\d{1,2})月", normalized)
    if m and ym:
        return _normalize_date(ym.group(1), ym.group(2), m.group(2))
    return extract_table_operation_date_from_title(text)


def extract_table_operation_date_from_title(table_title: str) -> Optional[str]:
    t = normalize_chinese_whitespace(table_title)
    t = t.replace("\n", "")
    m = re.search(r"(\d{4})[-年]\s*(\d{1,2})[-月]\s*(\d{1,2})", t)
    if m:
        return _normalize_date(m.group(1), m.group(2), m.group(3))
    m2 = re.search(r"(\d{4})\s*[-]\s*(\d{1,2})\s*[-]\s*(\d{1,2})", t)
    if m2:
        return _normalize_date(m2.group(1), m2.group(2), m2.group(3))
    return None


def extract_market_trading_section_text(text: str) -> str:
    t = "" if text is None else str(text)
    m = re.search(r"二、市场交易情况([\s\S]*?)(?=\n\s*三、|\Z)", t)
    return normalize_chinese_whitespace(m.group(1) if m else "")


def find_table_by_title(tables: Sequence[Any], table_title_pattern: str) -> Optional[Any]:
    p = re.compile(table_title_pattern)
    for table in tables:
        title = normalize_chinese_whitespace(getattr(table, "title", "") or "")
        title = title.replace("\n", "")
        if p.search(title):
            return table
    return None


def split_price_and_time(cell_text: str) -> Tuple[str, str, str]:
    raw = normalize_chinese_whitespace(cell_text)
    m = re.search(r"(\d{1,2}:\d{2})", raw)
    if not m:
        return raw, "", raw
    time_text = m.group(1)
    price_text = re.sub(r"[（(]?\d{1,2}:\d{2}[）)]?", "", raw).strip()
    return normalize_chinese_whitespace(price_text), time_text, raw


def _rows_to_text_rows(df: pd.DataFrame) -> List[List[str]]:
    return [[normalize_chinese_whitespace(x) for x in row] for row in df.fillna("").values.tolist()]


def _extract_price_rows(source_file: str, table_name: str, section: str, table_date: Optional[str], unit: str, rows: List[List[str]]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    metrics = ["最高电价", "最低电价", "平均电价", "电价环比"]
    for row in rows:
        label = row[0] if row else ""
        if label not in {"发电侧", "燃煤", "燃气", "新能源"}:
            continue
        for idx, metric in enumerate(metrics, start=1):
            if idx >= len(row):
                continue
            price, tm, raw = split_price_and_time(row[idx]) if metric in {"最高电价", "最低电价"} else (row[idx], "", row[idx])
            out.append({"source_file": source_file, "table_name": table_name, "table_operation_date": table_date or "", "section": section, "side_or_fuel": label, "metric": metric, "price": price, "time": tm, "unit": unit, "raw_text": raw})
    return out


def extract_table1_day_ahead_volume(source_file: str, table_name: str, table_date: Optional[str], unit: str, df: pd.DataFrame) -> List[Dict[str, Any]]:
    rows = _rows_to_text_rows(df)
    out: List[Dict[str, Any]] = []
    side = ""
    for row in rows:
        if not any(row):
            continue
        first = row[0]
        if "发电侧" in first or first == "用电侧":
            side = first
            continue
        if first in {"燃煤", "燃气", "核电", "新能源", "合计", "售电公司", "大用户"}:
            out.append({"source_file": source_file, "table_name": table_name, "table_operation_date": table_date or "", "section": "（三）日前成交电量", "side": side, "category": first, "value": row[1] if len(row) > 1 else "", "unit": unit, "raw_text": " | ".join(row)})
    return out


def extract_table1_day_ahead_price(source_file: str, table_name: str, table_date: Optional[str], unit: str, df: pd.DataFrame) -> List[Dict[str, Any]]:
    return _extract_price_rows(source_file, table_name, "（四）日前成交电价", table_date, unit, _rows_to_text_rows(df))


def extract_table2_day_ahead_price(source_file: str, table_name: str, table_date: Optional[str], unit: str, df: pd.DataFrame) -> List[Dict[str, Any]]:
    return _extract_price_rows(source_file, table_name, "（二）日前成交电价", table_date, unit, _rows_to_text_rows(df))


def extract_table2_realtime_price(source_file: str, table_name: str, table_date: Optional[str], unit: str, df: pd.DataFrame) -> List[Dict[str, Any]]:
    return _extract_price_rows(source_file, table_name, "（三）实时成交电价", table_date, unit, _rows_to_text_rows(df))


def attach_table_operation_date(rows: List[Dict[str, Any]], table_title: str) -> List[Dict[str, Any]]:
    d = extract_table_operation_date_from_title(table_title)
    for r in rows:
        r["table_operation_date"] = d or ""
    return rows

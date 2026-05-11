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
    diagnostics: List[Dict[str, Any]]


def _diag(source_file: str, stage: str, status: str, message: str, rows_extracted: int = 0) -> Dict[str, Any]:
    return {
        "source_file": source_file,
        "stage": stage,
        "status": status,
        "message": message,
        "rows_extracted": rows_extracted,
    }


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


def build_market_trading_rows(section_text: str, source_file: str) -> List[Dict[str, Any]]:
    section_title = "二、市场交易情况"
    rows: List[Dict[str, Any]] = []
    normalized = normalize_chinese_whitespace(section_text)
    if not normalized:
        return rows
    for chunk in re.split(r"(?=（[一二三四五六七八九十]+）)", normalized):
        chunk = chunk.strip(" 。")
        if not chunk:
            continue
        m_sub = re.match(r"(（[一二三四五六七八九十]+）[^0-9。]*)", chunk)
        subsection = m_sub.group(1).strip() if m_sub else ""
        item_matches = list(re.finditer(r"(?:^|。)\s*(\d+)\.(.*?)(?=(?:。\s*\d+\.)|$)", chunk))
        if not item_matches:
            item_matches = [re.match(r"(?:)(.*)", chunk)]  # type: ignore[list-item]
        for mt in item_matches:
            if mt is None:
                continue
            item_no = mt.group(1) if mt.lastindex and mt.lastindex >= 1 and mt.group(1) else ""
            body = mt.group(2) if mt.lastindex and mt.lastindex >= 2 and mt.group(2) else mt.group(0)
            body = body.strip(" 。")
            # generic fallback row
            base = {
                "source_file": source_file, "report_type": "guangdong_daily", "section_title": section_title,
                "subsection_title": subsection, "item_no": item_no, "statement_type": "", "metric_name": "",
                "value": "", "unit": "", "time": "", "fuel_type": "", "side": "", "raw_text": body,
            }
            rows.append(base.copy())
            unit_pat = r"([^\d，。；（）()]*?(?:亿\s*kWh|亿kWh|厘/千瓦时|MW|个|%))"
            for side, metric in [("用电侧", "日前总成交电量"), ("发电侧", "日前总成交电量"), ("发电侧", "日前加权平均电价")]:
                mm = re.search(rf"{side}{metric}([0-9]+(?:\.[0-9]+)?)\s*{unit_pat}", body)
                if mm:
                    rec = base.copy()
                    rec.update({"statement_type": "metric", "metric_name": metric, "value": mm.group(1), "unit": mm.group(2), "side": side})
                    rows.append(rec)
            for fuel in ["燃煤", "燃气", "核电", "新能源"]:
                mm = re.search(rf"{fuel}([0-9]+(?:\.[0-9]+)?)\s*{unit_pat}", body)
                if mm:
                    rec = base.copy()
                    rec.update({"statement_type": "metric", "metric_name": "日前成交电量", "value": mm.group(1), "unit": mm.group(2), "fuel_type": fuel})
                    rows.append(rec)
            hi = re.search(rf"最高([\-]?[0-9]+(?:\.[0-9]+)?)\s*{unit_pat}", body)
            lo = re.search(rf"最低([\-]?[0-9]+(?:\.[0-9]+)?)\s*{unit_pat}", body)
            if hi:
                rec = base.copy(); rec.update({"statement_type": "metric", "metric_name": "日前机组成交价最高", "value": hi.group(1), "unit": hi.group(2)})
                rows.append(rec)
            if lo:
                rec = base.copy(); rec.update({"statement_type": "metric", "metric_name": "日前机组成交价最低", "value": lo.group(1), "unit": lo.group(2)})
                rows.append(rec)
            extra_price_patterns = [
                (r"(日前加权平均电价)\s*([\-]?\d+(?:\.\d+)?)\s*(厘/千瓦时)", ""),
                (r"(实时加权平均电价)\s*([\-]?\d+(?:\.\d+)?)\s*(厘/千瓦时)", ""),
                (r"(燃煤均价)\s*([\-]?\d+(?:\.\d+)?)\s*(厘/千瓦时)", ""),
                (r"(燃气均价)\s*([\-]?\d+(?:\.\d+)?)\s*(厘/千瓦时)", ""),
                (r"(日前机组成交价最高)\s*([\-]?\d+(?:\.\d+)?)\s*(厘/千瓦时)", ""),
                (r"(日前机组成交价最低)\s*([\-]?\d+(?:\.\d+)?)\s*(厘/千瓦时)", ""),
                (r"(实时机组成交价最高)\s*([\-]?\d+(?:\.\d+)?)\s*(厘/千瓦时)", ""),
                (r"(实时机组成交价最低)\s*([\-]?\d+(?:\.\d+)?)\s*(厘/千瓦时)", ""),
            ]
            for pat, side in extra_price_patterns:
                for pm in re.finditer(pat, body):
                    rec = base.copy()
                    rec.update({"statement_type": "metric", "metric_name": pm.group(1), "value": pm.group(2), "unit": pm.group(3), "side": side})
                    rows.append(rec)
    return rows


def _normalize_table_title_for_match(title: str) -> str:
    t = normalize_chinese_whitespace(title or "").replace("\n", "")
    # unify date separators/variants often seen in OCR/PDF extraction
    t = t.replace("－", "-").replace("—", "-").replace("–", "-").replace("—", "-")
    t = t.replace("年", "-").replace("月", "-").replace("日", "")
    return t


def find_table_by_title(tables: Sequence[Any], table_title_pattern: str) -> Optional[Any]:
    p = re.compile(table_title_pattern)
    for table in tables:
        title_raw = getattr(table, "title", "") or ""
        title = _normalize_table_title_for_match(title_raw)
        compact = re.sub(r"\s+", "", title)
        if p.search(title) or p.search(compact):
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


def normalize_unit_text(unit: str) -> str:
    t = normalize_chinese_whitespace(unit).replace(" ", "")
    t = re.sub(r"[。；，]+$", "", t)
    t = re.split(r"[（()）]", t)[0]
    t = t.strip()
    t = t.replace("亿千瓦时", "亿kWh").replace("亿kwh", "亿kWh").replace("亿KWH", "亿kWh").replace("亿 kWh", "亿kWh")
    for k in ["亿kWh", "厘/千瓦时", "MW", "个", "%"]:
        if k in t:
            return k
    return t


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

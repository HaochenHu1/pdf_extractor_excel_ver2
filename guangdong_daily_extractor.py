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


def normalize_market_metric_terms(text: str) -> str:
    """
    Repair OCR/PDF-fragmented metric phrases in market-trading narrative text.
    Keep this as a strict allowlist to avoid accidental over-normalization.
    """
    t = "" if text is None else str(text)
    join = r"[\s/／\-]*"
    # Also allow separators inside words, e.g. 燃煤均/价, 日前加权平/均电价.
    def _flex_phrase(phrase: str) -> str:
        return join.join(re.escape(ch) for ch in phrase)

    phrase_rewrites = [
        (_flex_phrase("发电侧日前总成交电量"), "发电侧日前总成交电量"),
        (_flex_phrase("燃煤日前成交电量"), "燃煤日前成交电量"),
        (_flex_phrase("燃气日前成交电量"), "燃气日前成交电量"),
        (_flex_phrase("核电日前成交电量"), "核电日前成交电量"),
        (_flex_phrase("新能源日前成交电量"), "新能源日前成交电量"),
        (_flex_phrase("日前加权平均电价"), "日前加权平均电价"),
        (_flex_phrase("日前机组成交价最低"), "日前机组成交价最低"),
        (_flex_phrase("日前机组成交价最高"), "日前机组成交价最高"),
        (_flex_phrase("燃煤均价"), "燃煤均价"),
        (_flex_phrase("燃气均价"), "燃气均价"),
    ]
    for pat, repl in phrase_rewrites:
        t = re.sub(pat, repl, t, flags=re.IGNORECASE)
    return t


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
    # General anti-linebreak normalization for all market metrics.
    # Make metric/value/unit regex matching robust to fragmented line breaks and spaces.
    normalized = normalized.replace("\r", "\n")
    normalized = re.sub(r"[ \t]*\n[ \t]*", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized)
    normalized = re.sub(r"厘\s*/\s*千瓦时", "厘/千瓦时", normalized)
    normalized = re.sub(r"亿\s*kWh", "亿kWh", normalized, flags=re.IGNORECASE)
    normalized = normalize_market_metric_terms(normalized)
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
            body = re.sub(r"\s+", " ", body)
            body = re.sub(r"厘\s*/\s*千瓦时", "厘/千瓦时", body)
            body = re.sub(r"亿\s*kWh", "亿kWh", body, flags=re.IGNORECASE)
            body = normalize_market_metric_terms(body)
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


def _lines(text: str) -> List[str]:
    return [normalize_chinese_whitespace(line) for line in (text or "").splitlines() if normalize_chinese_whitespace(line)]


def _num_text(value: str) -> str:
    m = re.search(r"[-+]?\d+(?:\.\d+)?", normalize_chinese_whitespace(value))
    return m.group(0) if m else ""


def _percent_text(value: str) -> str:
    m = re.search(r"[-+]?\d+(?:\.\d+)?\s*%", normalize_chinese_whitespace(value))
    return m.group(0).replace(" ", "") if m else _num_text(value)


def _section_between(text: str, start_pat: str, end_pat: str) -> str:
    m = re.search(start_pat + r"([\s\S]*?)" + end_pat, text or "")
    return m.group(1) if m else ""


def _section_from(text: str, start_pat: str) -> str:
    m = re.search(start_pat + r"([\s\S]*)", text or "")
    return m.group(1) if m else ""


def _parse_price_value(raw: str) -> Tuple[str, str, str]:
    price, tm, source = split_price_and_time(raw)
    return _num_text(price), tm, source


def _parse_price_block(
    source_file: str,
    table_name: str,
    table_date: Optional[str],
    section: str,
    block_text: str,
    labels: Sequence[str],
) -> List[Dict[str, Any]]:
    """Parse one bounded price section only, preventing day-ahead/realtime leakage."""
    rows: List[Dict[str, Any]] = []
    block_lines = _lines(block_text)
    metrics = ["最高电价", "最低电价", "平均电价", "电价环比"]
    for idx, line in enumerate(block_lines):
        if line not in labels:
            continue
        values = block_lines[idx + 1 : idx + 5]
        if len(values) < 4:
            continue
        for metric, raw in zip(metrics, values):
            if metric in {"最高电价", "最低电价"}:
                price, tm, source = _parse_price_value(raw)
            elif metric == "电价环比":
                price, tm, source = _percent_text(raw), "", raw
            else:
                price, tm, source = _num_text(raw), "", raw
            rows.append(
                {
                    "source_file": source_file,
                    "table_name": table_name,
                    "table_operation_date": table_date or "",
                    "section": section,
                    "side_or_fuel": line,
                    "metric": metric,
                    "price": price,
                    "time": tm,
                    "unit": "%" if metric == "电价环比" else "厘/千瓦时",
                    "raw_text": source,
                }
            )
    return rows


def _parse_table1_volume(source_file: str, table_name: str, table_date: Optional[str], block_text: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    block_lines = _lines(block_text)
    try:
        start = block_lines.index("发电侧（含基数")
    except ValueError:
        return rows
    nums: List[str] = []
    for line in block_lines[start:]:
        if line == "（四）日前成交电价":
            break
        if re.match(r"[-+]?\d+(?:\.\d*)?(?:\([^)]+\))?$", line):
            nums.append(_num_text(line))
    for category, value in zip(["燃煤", "燃气", "核电", "新能源", "合计"], nums[:5]):
        rows.append(
            {
                "source_file": source_file,
                "table_name": table_name,
                "table_operation_date": table_date or "",
                "section": "（三）日前成交电量",
                "side": "发电侧（含基数及代购电量）",
                "category": category,
                "value": value,
                "unit": "亿kWh",
                "raw_text": value,
            }
        )
    for category, value in zip(["售电公司", "大用户", "合计"], nums[5:8]):
        rows.append(
            {
                "source_file": source_file,
                "table_name": table_name,
                "table_operation_date": table_date or "",
                "section": "（三）日前成交电量",
                "side": "用电侧",
                "category": category,
                "value": value,
                "unit": "亿kWh",
                "raw_text": value,
            }
        )
    return rows


def _market_metric_rows(section_text: str, source_file: str) -> List[Dict[str, Any]]:
    """Sheet 1 boundary: 二、市场交易情况 -> （一）运行日现货日前交易情况 only."""
    text = re.sub(r"\s+", "", section_text or "")
    rows: List[Dict[str, Any]] = []

    def add(metric_name: str, value: str, unit: str, side: str = "", fuel_type: str = "") -> None:
        if value == "":
            return
        rows.append(
            {
                "source_file": source_file,
                "report_type": "guangdong_daily",
                "section_title": "二、市场交易情况",
                "subsection_title": "（一）现货日前交易情况",
                "item_no": "",
                "statement_type": "metric",
                "metric_name": metric_name,
                "value": value,
                "unit": unit,
                "time": "",
                "fuel_type": fuel_type,
                "side": side,
                "raw_text": section_text,
            }
        )

    patterns = [
        ("用电侧", "日前总成交电量", r"用电侧日前总成交电量([-+]?\d+(?:\.\d+)?)亿kWh", "亿kWh", ""),
        ("发电侧", "日前总成交电量", r"发电侧日前总成交电量([-+]?\d+(?:\.\d+)?)亿kWh", "亿kWh", ""),
        ("", "日前成交电量", r"燃煤([-+]?\d+(?:\.\d+)?)亿kWh", "亿kWh", "燃煤"),
        ("", "日前成交电量", r"燃气([-+]?\d+(?:\.\d+)?)亿kWh", "亿kWh", "燃气"),
        ("", "日前成交电量", r"核电([-+]?\d+(?:\.\d+)?)亿kWh", "亿kWh", "核电"),
        ("", "日前成交电量", r"新能源([-+]?\d+(?:\.\d+)?)亿kWh", "亿kWh", "新能源"),
        ("", "日前加权平均电价", r"日前加权平均电价([-+]?\d+(?:\.\d+)?)厘/千瓦时", "厘/千瓦时", ""),
        ("", "燃煤均价", r"燃煤均价([-+]?\d+(?:\.\d+)?)厘/千瓦时", "厘/千瓦时", ""),
        ("", "燃气均价", r"燃气均价([-+]?\d+(?:\.\d+)?)厘/千瓦时", "厘/千瓦时", ""),
        ("", "日前机组成交价最高", r"价最高([-+]?\d+(?:\.\d+)?)厘/千瓦时", "厘/千瓦时", ""),
        ("", "日前机组成交价最低", r"最低([-+]?\d+(?:\.\d+)?)厘/千瓦时", "厘/千瓦时", ""),
    ]
    for side, metric, pat, unit, fuel in patterns:
        m = re.search(pat, text)
        add(metric, m.group(1) if m else "", unit, side=side, fuel_type=fuel)
    return rows


def extract_guangdong_daily_from_text(source_file: str, text: str) -> GuangdongDailyExtractionResult:
    """
    Extract the three workbook targets from explicit text boundaries.

    Boundaries:
    - Sheet 1: 二、市场交易情况 -> （一）...现货日前交易情况, stopped before 表1.
    - Sheet 2: 表1 ... 日前交易情况 -> （四）日前成交电价, with Table 1 date.
    - Sheet 3: 表2 ... 现货交易情况 -> （二）日前成交电价 and （三）实时成交电价,
      with the running day in the Table 2 title, even when it differs from the report date.
    """
    diagnostics: List[Dict[str, Any]] = [_diag(source_file, "detect", "INFO", f"检测广东日报: {source_file}", 0)]
    report_date = extract_daily_report_operation_date(source_file, text)

    market_block = _section_between(
        text,
        r"二、市场交易情况[\s\S]*?（一）\s*\d{1,2}\s*月\s*\d{1,2}\s*日（运行日）现货日前交易情况",
        r"表1\s*运行日",
    )
    market_rows = _market_metric_rows(market_block, source_file)
    for row in market_rows:
        row["date"] = report_date or ""
    diagnostics.append(_diag(source_file, "sheet1_market_block", "INFO" if market_block else "WARN", "找到Sheet1目标边界" if market_block else "缺少Sheet1目标边界", len(market_rows)))

    table1_title = re.search(r"(表1\s*运行日\s*\d{4}[-年]\d{1,2}[-月]\d{1,2}\s*日前交易情况)", text or "")
    table1_name = table1_title.group(1) if table1_title else "表1 运行日前交易情况"
    table1_date = extract_table_operation_date_from_title(table1_name)
    table1_text = _section_between(text, r"表1\s*运行日", r"表\d+\s*运行日\s*\d{4}[-年]\d{1,2}[-月]\d{1,2}\s*现货交易情况")
    table1_volume_block = _section_between(table1_text, r"（三）日前成交电量", r"（四）日前成交电价")
    table1_price_block = _section_from(table1_text, r"（四）日前成交电价")
    table1_volume_rows = _parse_table1_volume(source_file, table1_name, table1_date, "（三）日前成交电量\n" + table1_volume_block + "\n（四）日前成交电价")
    table1_price_rows = _parse_price_block(source_file, table1_name, table1_date, "（四）日前成交电价", table1_price_block, ["发电侧", "燃煤", "燃气", "新能源"])
    diagnostics.append(_diag(source_file, "sheet2_table1_price", "INFO" if table1_price_rows else "WARN", "找到表1日前成交电价" if table1_price_rows else "表1日前成交电价为空", len(table1_price_rows)))

    table2_title = re.search(r"(表\d+\s*运行日\s*\d{4}[-年]\d{1,2}[-月]\d{1,2}\s*现货交易情况)", text or "")
    table2_name = table2_title.group(1) if table2_title else "表2 运行日现货交易情况"
    table2_date = extract_table_operation_date_from_title(table2_name)
    table2_text = _section_between(text, r"表\d+\s*运行日\s*\d{4}[-年]\d{1,2}[-月]\d{1,2}\s*现货交易情况", r"\n\s*三、市场结算情况")
    t2_da_block = _section_between(table2_text, r"（二）日前成交电价", r"（三）实时成交电价")
    t2_rt_block = _section_from(table2_text, r"（三）实时成交电价")
    table2_day_ahead_rows = _parse_price_block(source_file, table2_name, table2_date, "（二）日前成交电价", t2_da_block, ["发电侧", "燃煤", "燃气"])
    table2_realtime_rows = _parse_price_block(source_file, table2_name, table2_date, "（三）实时成交电价", t2_rt_block, ["发电侧", "燃煤", "燃气"])
    diagnostics.append(_diag(source_file, "sheet3_table2_day_ahead", "INFO" if table2_day_ahead_rows else "WARN", "找到表2日前成交电价" if table2_day_ahead_rows else "表2日前成交电价为空", len(table2_day_ahead_rows)))
    diagnostics.append(_diag(source_file, "sheet3_table2_realtime", "INFO" if table2_realtime_rows else "WARN", "找到表2实时成交电价" if table2_realtime_rows else "表2实时成交电价为空", len(table2_realtime_rows)))

    if table2_date and report_date and table2_date != report_date:
        diagnostics.append(_diag(source_file, "table2_running_day", "INFO", f"表2运行日{table2_date}不同于报告日期{report_date}", 0))
    return GuangdongDailyExtractionResult(
        "guangdong_daily",
        report_date,
        market_rows,
        table1_volume_rows,
        table1_price_rows,
        table2_day_ahead_rows,
        table2_realtime_rows,
        diagnostics,
    )

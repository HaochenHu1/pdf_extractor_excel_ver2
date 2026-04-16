from __future__ import annotations

import re
from datetime import date
from dataclasses import dataclass, field
from typing import Callable, Iterable, List, Optional, Sequence, Tuple

import fitz  # PyMuPDF


NumberPostprocessor = Callable[[str], Optional[float]]


@dataclass(frozen=True)
class MetricConfig:
    """Configuration for one metric that should be extracted from a paragraph section."""

    canonical_name: str
    aliases: Sequence[str] = field(default_factory=tuple)
    pattern: Optional[str] = None
    unit: Optional[str] = None
    postprocess: Optional[NumberPostprocessor] = None


@dataclass(frozen=True)
class SectionConfig:
    """Configuration for one target paragraph section and its output worksheet."""

    section_title: str
    target_sheet_name: str
    metrics: Sequence[MetricConfig]


@dataclass(frozen=True)
class SectionExtractionResult:
    """Extraction output that can be written directly into a dedicated Excel sheet."""

    section_title: str
    sheet_name: str
    rows: Sequence[Tuple[Optional[int], str, Optional[float], Optional[str], Optional[date]]]


SECTION_HEADING_PATTERN = re.compile(
    r"(?m)^\s*(?:[一二三四五六七八九十]+、|[（(][一二三四五六七八九十]+[）)])\s*[^\n]{1,60}$"
)
NUMBER_PATTERN = r"(?P<value>[+-]?\d+(?:\.\d+)?)"
UNIT_PATTERN = r"(?P<unit>[^\d，。；;,:：）\(\)\s]+(?:/[^\d，。；;,:：）\(\)\s]+)?)?"
VALUE_WITH_UNIT_PATTERN = rf"{NUMBER_PATTERN}\s*{UNIT_PATTERN}"


def default_number_postprocess(raw: str) -> Optional[float]:
    """Convert extracted numeric text into float while keeping None on malformed input."""

    try:
        return float(raw)
    except (TypeError, ValueError):
        return None


def build_metric_pattern(metric: MetricConfig) -> re.Pattern[str]:
    """Build a robust metric regex using aliases and OCR-tolerant separators."""

    if metric.pattern:
        return re.compile(metric.pattern)

    alias_tokens = [metric.canonical_name, *metric.aliases]
    escaped_aliases = [re.escape(token) for token in alias_tokens if token]
    alias_part = "|".join(sorted(set(escaped_aliases), key=len, reverse=True))

    # OCR often injects irregular spaces/punctuation, so allow optional noise between name and value.
    pattern = (
        rf"(?:其中)?\s*(?:{alias_part})\s*"
        rf"(?:为|是|：|:)?\s*"
        rf"{VALUE_WITH_UNIT_PATTERN}"
    )
    return re.compile(pattern)


def normalize_section_text(text: str) -> str:
    """Normalize common OCR spacing/punctuation noise for stable regex matching."""

    cleaned = text.replace("\u3000", " ").replace("\xa0", " ")
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
    cleaned = re.sub(r"\s*([，。；：:（）()、])\s*", r"\1", cleaned)
    return cleaned


def isolate_section_block(full_text: str, section_title: str) -> str:
    """Return content belonging to a section title until the next section heading."""

    title_match = re.search(re.escape(section_title), full_text)
    if not title_match:
        return ""

    section_start = title_match.end()
    remainder = full_text[section_start:]
    next_heading = SECTION_HEADING_PATTERN.search(remainder)
    section_end = section_start + next_heading.start() if next_heading else len(full_text)
    return full_text[section_start:section_end].strip()


def extract_metric_value(section_text: str, metric: MetricConfig) -> Optional[float]:
    """Extract one configured metric value from section text, returning None if not found."""

    pattern = build_metric_pattern(metric)
    match = pattern.search(section_text)
    if not match:
        return None

    raw_value = match.group("value")
    postprocess = metric.postprocess or default_number_postprocess
    return postprocess(raw_value)


def extract_metric_unit(section_text: str, metric: MetricConfig) -> Optional[str]:
    """Extract unit from matched text or fallback to fixed unit from metric config."""

    if metric.unit:
        return metric.unit

    match = build_metric_pattern(metric).search(section_text)
    if not match:
        return None
    raw_unit = match.groupdict().get("unit")
    if not raw_unit:
        return None
    return raw_unit.strip()


def convert_lifeny_to_yuan_per_kwh(
    value: Optional[float], unit: Optional[str]
) -> Tuple[Optional[float], Optional[str]]:
    """Convert 厘/千瓦时 to 元/kWh by dividing numeric value by 1000."""

    if not unit:
        return value, unit

    normalized_unit = unit.replace(" ", "")
    if normalized_unit == "厘/千瓦时":
        converted_value = round(value / 1000.0, 6) if value is not None else None
        return converted_value, "元/kWh"
    return value, unit


def parse_report_date(full_text: str) -> Optional[date]:
    """Parse report date from title format: YYYY年M月 ... （MM.DD）."""

    year_month_match = re.search(r"(?P<year>\d{4})\s*年\s*(?P<month>\d{1,2})\s*月", full_text)
    day_hint_match = re.search(r"[（(]\s*(?P<hint_month>\d{1,2})\s*[\.．。/-]\s*(?P<day>\d{1,2})\s*[）)]", full_text)
    if not year_month_match or not day_hint_match:
        return None

    year = int(year_month_match.group("year"))
    title_month = int(year_month_match.group("month"))
    hinted_month = int(day_hint_match.group("hint_month"))
    day = int(day_hint_match.group("day"))
    month = title_month if 1 <= title_month <= 12 else hinted_month
    if hinted_month == title_month:
        month = hinted_month

    try:
        return date(year, month, day)
    except ValueError:
        return None


def parse_report_month(full_text: str) -> Optional[int]:
    """Parse report month from title format: YYYY年M月."""

    year_month_match = re.search(r"(?P<year>\d{4})\s*年\s*(?P<month>\d{1,2})\s*月", full_text)
    if not year_month_match:
        return None
    month = int(year_month_match.group("month"))
    return month if 1 <= month <= 12 else None


def extract_configured_sections(full_text: str, configs: Iterable[SectionConfig]) -> List[SectionExtractionResult]:
    """Extract only configured metrics from configured sections using rule-based regex."""

    normalized = normalize_section_text(full_text)
    report_date = parse_report_date(normalized)
    report_month = parse_report_month(normalized)
    results: List[SectionExtractionResult] = []

    for config in configs:
        block = isolate_section_block(normalized, config.section_title)
        rows: List[Tuple[Optional[int], str, Optional[float], Optional[str], Optional[date]]] = []

        for metric in config.metrics:
            raw_value = extract_metric_value(block, metric)
            raw_unit = extract_metric_unit(block, metric)
            converted_value, converted_unit = convert_lifeny_to_yuan_per_kwh(raw_value, raw_unit)
            rows.append(
                (
                    report_month,
                    metric.canonical_name,
                    converted_value,
                    converted_unit,
                    report_date,
                )
            )

        results.append(
            SectionExtractionResult(
                section_title=config.section_title,
                sheet_name=config.target_sheet_name,
                rows=rows,
            )
        )

    return results


def extract_configured_sections_from_pdf(input_pdf_path: str, configs: Iterable[SectionConfig]) -> List[SectionExtractionResult]:
    """Read all page text from PDF and run configured section extraction."""

    document = fitz.open(input_pdf_path)
    full_text = "\n".join((page.get_text("text") or "") for page in document)
    document.close()
    return extract_configured_sections(full_text, configs)


def default_section_configs() -> List[SectionConfig]:
    """Default extraction config for 当前所需章节：日前市场情况 / 实时市场情况."""

    price_metrics = [
        MetricConfig(
            canonical_name="发电侧加权平均价格",
            pattern=rf"发电侧加权平均价格\s*(?:为|是|：|:)?\s*{VALUE_WITH_UNIT_PATTERN}",
            unit="元/MWh",
        ),
        MetricConfig(
            canonical_name="出清电价最大值",
            pattern=rf"(?:出清电价)?最大值\s*(?:为|是|：|:)?\s*{VALUE_WITH_UNIT_PATTERN}",
            unit="元/MWh",
        ),
        MetricConfig(
            canonical_name="出清电价最小值",
            pattern=rf"(?:出清电价)?最小值\s*(?:为|是|：|:)?\s*{VALUE_WITH_UNIT_PATTERN}",
            unit="元/MWh",
        ),
        MetricConfig(
            canonical_name="煤电均价",
            aliases=("其中煤电均价",),
            pattern=rf"(?:其中)?\s*煤电均价\s*(?:为|是|：|:)?\s*{VALUE_WITH_UNIT_PATTERN}",
            unit="元/MWh",
        ),
        MetricConfig(
            canonical_name="气电均价",
            aliases=("其中气电均价",),
            pattern=rf"(?:其中)?\s*气电均价\s*(?:为|是|：|:)?\s*{VALUE_WITH_UNIT_PATTERN}",
            unit="元/MWh",
        ),
    ]

    return [
        SectionConfig(
            section_title="（二）日前市场情况",
            target_sheet_name="日前市场情况",
            metrics=price_metrics,
        ),
        SectionConfig(
            section_title="（三）实时市场情况",
            target_sheet_name="实时市场情况",
            metrics=price_metrics,
        ),
    ]


def demo_extract_market_section_metrics() -> List[SectionExtractionResult]:
    """Demo helper for local verification with examples for 日前/实时 sections."""

    sample_text = (
        "广东电力现货市场结算运行情况月报（2026年1月)\n"
        "（二）日前市场情况\n"
        "3、发电侧平均价格情况\n"
        "1月发电侧加权平均价格为331元/MWh（包含市场化核电机组、新能源机组），"
        "出清电价最大值为414元/MWh、最小值为177元/MWh。"
        "煤电均价342元/MWh，气电均价375元/MWh。\n"
        "（三）实时市场情况\n"
        "2、发电侧平均价格情况\n"
        "1月发电侧加权平均价格为299元/MWh（包含市场化核电机组、新能源机组）；"
        "出清电价最大值为412元/MWh、最小值为113元/MWh。"
        "煤电均价315元/MWh，气电均价342元/MWh。"
    )
    return extract_configured_sections(sample_text, default_section_configs())

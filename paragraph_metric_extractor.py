from __future__ import annotations

import re
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
    rows: Sequence[Tuple[str, Optional[float]]]


SECTION_HEADING_PATTERN = re.compile(r"(?m)^\s*[一二三四五六七八九十]+、[^\n]{1,60}$")
NUMBER_PATTERN = r"(?P<value>[+-]?\d+(?:\.\d+)?)"


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
        rf"{NUMBER_PATTERN}"
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


def extract_configured_sections(full_text: str, configs: Iterable[SectionConfig]) -> List[SectionExtractionResult]:
    """Extract only configured metrics from configured sections using rule-based regex."""

    normalized = normalize_section_text(full_text)
    results: List[SectionExtractionResult] = []

    for config in configs:
        block = isolate_section_block(normalized, config.section_title)
        rows: List[Tuple[str, Optional[float]]] = []

        for metric in config.metrics:
            rows.append((metric.canonical_name, extract_metric_value(block, metric)))

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
    """Default extraction config for currently required section: 二、市场交易情况."""

    market_metrics = [
        MetricConfig(canonical_name="日前总成交电量", aliases=("发电侧日前总成交电量",)),
        MetricConfig(canonical_name="燃煤", pattern=rf"(?:其中)?\s*燃煤(?!均价)\s*{NUMBER_PATTERN}"),
        MetricConfig(canonical_name="燃气", pattern=rf"(?:其中)?\s*燃气(?!均价)\s*{NUMBER_PATTERN}"),
        MetricConfig(canonical_name="核电"),
        MetricConfig(canonical_name="新能源"),
        MetricConfig(canonical_name="日前加权平均电价"),
        MetricConfig(canonical_name="燃煤均价"),
        MetricConfig(canonical_name="燃气均价"),
        MetricConfig(canonical_name="日前机组成交价最高", pattern=rf"(?:日前机组成交价)?最高\s*{NUMBER_PATTERN}"),
        MetricConfig(canonical_name="日前机组成交价最低", pattern=rf"(?:日前机组成交价)?最低\s*{NUMBER_PATTERN}"),
    ]
    return [
        SectionConfig(
            section_title="二、市场交易情况",
            target_sheet_name="市场交易情况",
            metrics=market_metrics,
        )
    ]


def demo_extract_market_section_metrics() -> List[Tuple[str, Optional[float]]]:
    """Demo helper for local verification with the example paragraph provided by user."""

    sample_text = (
        "二、市场交易情况\n"
        "发电侧日前总成交电量16.99亿kWh（其中燃煤10.67亿kWh，燃气2.65亿kWh，"
        "核电2.42亿kWh，新能源1.26亿kWh），日前加权平均电价335.9厘/千瓦时，"
        "其中燃煤均价336.7厘/千瓦时，燃气均价352.2厘/千瓦时。"
        "日前机组成交价最高1101.1厘/千瓦时，最低-35厘/千瓦时。"
    )
    extracted = extract_configured_sections(sample_text, default_section_configs())
    return list(extracted[0].rows) if extracted else []

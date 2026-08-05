"""Input validation and normalization."""

from typing import Any, Optional

from models import Metric


def validate_input(data: dict[str, Any]) -> dict[str, Any]:
    """Confirm the input JSON has the minimum shape we can build from.

    Unwraps an `analysis` key if present (the shape Statista's export
    uses), otherwise treats the top-level JSON itself as the analysis
    object — so the wrapper key name isn't hardcoded, only the shape
    inside it. Validated by shape, not by key name: a `title` or
    `key_insights` field must be present somewhere, or there's nothing to
    build a presentation from.

    An empty `key_insights` list is not an error — it's flagged via
    `_summary_only` so the caller can build a summary-only presentation
    instead of crashing or faking data.
    """
    analysis = data.get("analysis", data)
    if not isinstance(analysis, dict) or not (
        isinstance(analysis.get("title"), str) or isinstance(analysis.get("key_insights"), list)
    ):
        raise ValueError(
            "Input JSON doesn't look like an analysis object — expected a "
            "'title' and/or 'key_insights' field (optionally nested under "
            "an 'analysis' key) — there is nothing to build a presentation from."
        )

    key_insights = analysis.get("key_insights", [])
    if not key_insights:
        analysis["_summary_only"] = True

    return analysis


# Identity fields ("who/what the value belongs to") checked in this order
# regardless of chart hint.
_LABEL_KEYS = ("brand", "action", "statement", "label", "region", "segment", "manufacturer", "category")

# Numeric fields that still belong on the label side of a metric (e.g. a
# year is numeric but it's the x-axis label, not the measured value).
_PERIOD_KEYS = ("year", "period", "date", "quarter", "month")

# Bookkeeping fields that never qualify as the label or the value.
_IGNORED_KEYS = {"description", "rank", "percentage"}


def _metric_from_raw(raw: dict[str, Any], is_trend: bool = False) -> Optional[Metric]:
    """Convert one raw metric dict into a Metric, or None if unrecognized.

    Three shapes are recognized: rank (+ optional percentage) -> "ranking";
    a known label key + percentage -> "comparison"; otherwise, the first
    label-like field and first numeric field found -> "trend" or
    "comparison" depending on `is_trend`.
    """
    has_rank = raw.get("rank") is not None
    has_percentage = raw.get("percentage") is not None

    if has_rank:
        label = raw.get("brand") or raw.get("label") or "Unknown"
        value = raw["percentage"] if has_percentage else 0
        return Metric(label=label, value=value, rank=raw["rank"], kind="ranking")

    if has_percentage:
        label = next((raw[k] for k in _LABEL_KEYS if raw.get(k) is not None), "Unknown")
        return Metric(label=label, value=raw["percentage"], kind="comparison")

    # Generic fallback: first label-like field + first numeric field, in
    # insertion order, become the label/value pair.
    label_val = None
    label_key = next((k for k in _LABEL_KEYS if raw.get(k) is not None), None)
    if label_key is None:
        label_key = next((k for k in _PERIOD_KEYS if raw.get(k) is not None), None)
    if label_key:
        label_val = raw[label_key]

    numeric_key, numeric_val = None, None
    for k, v in raw.items():
        if k in _IGNORED_KEYS or k == label_key:
            continue
        if isinstance(v, (int, float)):
            numeric_key, numeric_val = k, v
            break

    if label_val is None:
        # No recognized label field: use the first remaining non-numeric field.
        for k, v in raw.items():
            if k in _IGNORED_KEYS or k == numeric_key:
                continue
            label_val = v
            break

    if label_val is None or numeric_val is None:
        return None

    kind = "trend" if is_trend else "comparison"
    return Metric(label=str(label_val), value=numeric_val, kind=kind)


def _mark_highlight(metrics: list[Metric]) -> None:
    """Flag the single most important metric for accent-color emphasis, in
    place. Trend series highlight their latest point; ranking/comparison
    groups highlight their largest value. No-op on an empty list.
    """
    if not metrics:
        return
    if metrics[0].kind == "trend":
        target = metrics[-1]
    else:
        target = max(metrics, key=lambda m: m.value)
    target.highlight = True


def normalize_metrics(insights: list[dict[str, Any]]) -> dict[int, list[Metric]]:
    """Convert every insight's raw metrics into Metric objects.

    Returns insight index -> list[Metric]. Metrics for an insight whose
    recommended_chart is "line" are normalized as "trend" points instead
    of "comparison" bars.
    """
    normalized: dict[int, list[Metric]] = {}
    for i, insight in enumerate(insights):
        raw_metrics = insight.get("metrics", [])
        is_trend = insight.get("recommended_chart") == "line"
        metrics = [m for m in (_metric_from_raw(r, is_trend=is_trend) for r in raw_metrics) if m is not None]
        _mark_highlight(metrics)
        normalized[i] = metrics
    return normalized


def build_context(analysis: dict[str, Any], normalized_metrics: dict[int, list[Metric]]) -> dict[str, Any]:
    """Assemble a plain dict of everything the planners need.

    `analysis` is the already-unwrapped dict returned by validate_input.
    """
    key_insights = analysis.get("key_insights", [])
    data_sources = analysis.get("data_sources", {})

    insights = []
    for i, insight in enumerate(key_insights):
        source = insight.get("source", {})
        source_str = _format_source(source)
        insights.append(
            {
                "category": insight.get("category", ""),
                "insight": insight.get("insight", ""),
                "metrics": normalized_metrics.get(i, []),
                "source": source_str,
                "recommended_chart": insight.get("recommended_chart"),
            }
        )

    return {
        "title": analysis.get("title", "Untitled Analysis"),
        "summary": analysis.get("summary", ""),
        "insights": insights,
        "methodology": analysis.get("methodology", {}),
        "caveat": data_sources.get("caveat"),
        "region": data_sources.get("primary_region"),
        "summary_only": analysis.get("_summary_only", False),
    }


def _format_source(source: dict[str, Any]) -> Optional[str]:
    """Build a single human-readable citation string from a source object."""
    if not source:
        return None
    name = source.get("name")
    conductor = source.get("conductor")
    date = source.get("publication_date")
    parts = [p for p in (name, conductor, date) if p]
    return ", ".join(parts) if parts else None

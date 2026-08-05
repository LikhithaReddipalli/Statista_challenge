"""Data models for the presentation pipeline.

Metric is a plain dataclass (internal, no validation overhead). SlideSpec
is a pydantic model because it crosses a trust boundary (LLM output) and
needs real validation.
"""

from dataclasses import dataclass
from typing import Literal, Optional

from pydantic import BaseModel, Field, model_validator


@dataclass
class Metric:
    """A single normalized data point extracted from an insight.

    kind is one of "comparison" (labeled stat + percentage), "ranking"
    (ranked entity, optional percentage), or "trend" (a point on a time
    series, ordered by position rather than rank).

    highlight marks the one point the renderer should emphasize (a shade
    darker of the same brand blue) — the #1 bar or the latest trend point.
    The normalizer defaults it to "largest value" if the LLM doesn't set it.
    """

    label: str
    value: float
    rank: Optional[int] = None
    kind: Literal["comparison", "ranking", "trend"] = "comparison"
    highlight: bool = False


class KpiCard(BaseModel):
    """One headline number for a KPI-card layout slide.

    Field lengths are capped here (the trust boundary) rather than
    truncated silently at render time, since the renderer draws each field
    into a fixed-height line with no overflow handling.
    """

    label: str = Field(max_length=40)
    value: str = Field(max_length=20)
    delta: Optional[str] = Field(default=None, max_length=60)


class SlideSpec(BaseModel):
    """Validated, renderer-ready description of a single slide.

    The only shape the renderer accepts — both the LLM and rule-based
    planners must produce output that validates against this model.
    """

    title: str
    subtitle: Optional[str] = None
    layout: Literal["title", "chart", "summary", "kpi"] = "chart"
    chart_type: Optional[Literal["ranking", "comparison", "trend", "pie"]] = None
    key_message: Optional[str] = None
    bullets: list[str] = Field(default_factory=list)
    metrics: list[Metric] = Field(default_factory=list)
    kpis: list[KpiCard] = Field(default_factory=list)
    source: Optional[str] = None
    notes: Optional[str] = None

    @model_validator(mode="after")
    def _chart_type_requires_metrics(self) -> "SlideSpec":
        """Reject a chart_type with no metrics — nothing to plot, and the
        renderer would crash on an empty chart. Forces plan_llm's caller to
        fall back to plan_rule_based instead."""
        if self.chart_type is not None and not self.metrics:
            raise ValueError(f"chart_type={self.chart_type!r} requires at least one metric")
        return self

    @model_validator(mode="after")
    def _pie_slice_count_in_range(self) -> "SlideSpec":
        """Reject a pie with fewer than 2 or more than 6 metrics — mirrors
        planner._looks_like_part_to_whole. Beyond 6 slices PIE_PALETTE
        (renderer.py) runs out of distinct tints and reuses colors."""
        if self.chart_type == "pie" and not (2 <= len(self.metrics) <= 6):
            raise ValueError(
                f"chart_type='pie' requires 2-6 metrics (a readable pie can't have "
                f"fewer or more slices than that), got {len(self.metrics)}"
            )
        return self

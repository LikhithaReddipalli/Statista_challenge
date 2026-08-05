"""Turns validated context into a list of SlideSpec objects.

Two strategies: plan_rule_based (deterministic, always succeeds — the
fallback and the --no-llm path) and plan_llm (asks Gemini to design the
narrative; returns None on any failure so the caller can fall back).
"""

import json
import logging
import os
from typing import Any, Optional

from dotenv import load_dotenv
from pydantic import ValidationError

from models import Metric, SlideSpec

logger = logging.getLogger(__name__)

load_dotenv()

# Overridable via GEMINI_MODEL. Defaults to the "-latest" alias for
# convenience, but that alias can move to a different model over time —
# pin a dated version (e.g. "gemini-2.5-flash") for reproducible output.
DEFAULT_GEMINI_MODEL = "gemini-flash-latest"

# Bounds how long a hung network call can block the pipeline before
# falling back to plan_rule_based. Overridable via GEMINI_TIMEOUT_SECONDS.
DEFAULT_GEMINI_TIMEOUT_SECONDS = 30


def _looks_like_part_to_whole(metrics: list[Metric]) -> bool:
    """True if 2-6 metrics sum to ~100 — reads as pie slices of one whole
    rather than independent side-by-side stats."""
    if not (2 <= len(metrics) <= 6):
        return False
    total = sum(m.value for m in metrics)
    return 95 <= total <= 105


def _chart_type_for_metrics(metrics: list[Metric]) -> Optional[str]:
    """Pick a chart type for an insight: ranking if any metric is ranked,
    trend if the group is a time series, pie if values sum to ~100,
    otherwise a comparison bar. No metrics -> no chart."""
    if not metrics:
        return None
    if any(m.kind == "ranking" for m in metrics):
        return "ranking"
    if any(m.kind == "trend" for m in metrics):
        return "trend"
    if _looks_like_part_to_whole(metrics):
        return "pie"
    return "comparison"


def plan_rule_based(context: dict[str, Any]) -> list[SlideSpec]:
    """Build slides with fixed structure: title slide, one slide per
    insight, summary slide. No AI — the deterministic fallback that always
    succeeds on well-formed context."""
    slides: list[SlideSpec] = []

    slides.append(
        SlideSpec(
            title=context["title"],
            subtitle="Key Insights" if not context["summary_only"] else "Summary",
            layout="title",
            source=None,
            notes=context.get("caveat"),
        )
    )

    for insight in context["insights"]:
        metrics: list[Metric] = insight["metrics"]
        chart_type = _chart_type_for_metrics(metrics)
        slides.append(
            SlideSpec(
                title=insight["category"] or "Insight",
                layout="chart",
                chart_type=chart_type,
                key_message=insight["insight"],
                bullets=[] if chart_type else [insight["insight"]],
                metrics=metrics,
                source=insight["source"] or "Source unavailable.",
            )
        )

    # No key_message on the summary slide: the source "summary" field is a
    # full paragraph and reads poorly truncated onto a recap slide. The
    # bullet list is already the crisp summary.
    summary_bullets = _summary_bullets(context)
    slides.append(
        SlideSpec(
            title="Summary",
            layout="summary",
            bullets=summary_bullets,
            source=None,
            notes=context.get("caveat"),
        )
    )

    return slides


def _summary_bullets(context: dict[str, Any]) -> list[str]:
    """One bullet per insight category for the closing recap slide."""
    return [insight["category"] for insight in context["insights"] if insight["category"]]


def plan_llm(context: dict[str, Any], user_instruction: Optional[str] = None) -> Optional[list[SlideSpec]]:
    """Ask an LLM to design the presentation narrative.

    The JSON schema comes from SlideSpec.model_json_schema() so the prompt
    contract can't drift from the actual model. Returns None on any
    failure (missing key, network error, invalid JSON, schema mismatch) —
    never raises for LLM-related failures, so the caller can fall back to
    plan_rule_based.
    """
    try:
        from langchain_google_genai import ChatGoogleGenerativeAI
    except ImportError:
        logger.warning("LLM planner failed: langchain_google_genai is not installed.")
        return None

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        logger.warning("LLM planner failed: GEMINI_API_KEY is not set.")
        return None

    schema = SlideSpec.model_json_schema()
    prompt = _build_prompt(context, schema, user_instruction)

    model_name = os.environ.get("GEMINI_MODEL", DEFAULT_GEMINI_MODEL)
    try:
        timeout_seconds = float(os.environ.get("GEMINI_TIMEOUT_SECONDS", DEFAULT_GEMINI_TIMEOUT_SECONDS))
    except ValueError:
        timeout_seconds = DEFAULT_GEMINI_TIMEOUT_SECONDS

    try:
        llm = ChatGoogleGenerativeAI(model=model_name, google_api_key=api_key, timeout=timeout_seconds)
        response = llm.invoke(prompt)
        raw_text = _extract_text(response.content)
    except Exception as e:
        logger.warning("LLM planner failed: %s: %s", type(e).__name__, e)
        return None

    return _parse_and_validate(raw_text)


def _extract_text(content: Any) -> str:
    """Normalize a LangChain message's .content into plain text. Some
    Gemini models return content as a list of typed blocks instead of a
    plain string — concatenate the text blocks, ignore the rest."""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, str):
                parts.append(block)
            elif isinstance(block, dict) and block.get("type") == "text":
                parts.append(block.get("text", ""))
        return "".join(parts)
    return str(content)


def _build_prompt(context: dict[str, Any], schema: dict[str, Any], user_instruction: Optional[str]) -> str:
    instruction_line = (
        f"User instruction: {user_instruction}"
        if user_instruction
        else "No specific user instruction — use your judgment on slide count and focus."
    )

    return f"""You are a senior presentation designer at Statista, building decks in
Statista's official report style: clean, minimal, data-first, executive. It should read
like a professional analytical report, not a marketing deck — restrained, trustworthy,
premium. Generous whitespace. Flat design. No decorative elements.

Think like a presentation designer, not a data dumper: build a coherent narrative arc across
the slides, give each slide exactly one key message, and choose the layout and chart_type
that best fits what that slide is trying to say. Do not just emit one slide per raw insight
unless that genuinely serves the narrative.

{instruction_line}

Source analysis (title, summary, and normalized insights with their metrics and sources):
{json.dumps(context, default=str, indent=2)}

Return a JSON array of slide objects. Each object MUST validate against this JSON Schema:
{json.dumps(schema, indent=2)}

Rules:
- chart_type must be "ranking" for ranked/ordered data, "trend" for a time series
  (values across years/periods — use this whenever an insight's source data has a
  recommended_chart of "line"), "pie" for a true part-to-whole breakdown (2-6 slices of
  a single whole, values summing to ~100%, e.g. a market split into segments),
  "comparison" for side-by-side percentage/share stats that are each independently
  measured and do NOT sum to a whole (e.g. "57% said X influenced them, 54% said Y" —
  separate stats, not slices of one pie, so this stays a bar even though both are
  percentages), or null for a text-only slide (key_message + bullets). When in doubt
  between "pie" and "comparison", prefer "comparison" — a wrongly-chosen pie is more
  misleading than a wrongly-chosen bar.
- Color discipline (Statista brand): every chart mark is a shade of ONE brand color —
  deep blue (#003A70) and lighter tints of that same blue. There is no second hue
  anywhere, ever: no orange, green, red, purple, or rainbow palettes. You do not set
  colors directly — the renderer decides shades — but for "trend" charts you DO decide
  which single point is the headline value an executive should see first (the most
  recent or otherwise key point) by setting "highlight": true on exactly that one
  metric; leave every other metric's highlight false/unset. Bar charts ("ranking" and
  "comparison") render every bar in the same single color and ignore highlight — do not
  rely on it there. Pie charts render each slice in a distinct tint of the same blue
  family (for legibility) — this is automatic and does not need a highlight either.
- layout "kpi": use this instead of "chart" when a slide's job is to land 2-4 standalone
  headline numbers rather than show a distribution or trend (e.g. an executive-summary
  slide up front recapping the deck's most important figures). Populate "kpis" with 2-4
  objects, each {{"label": short description, "value": a compact formatted number like
  "26.4M" or "58%", "delta": optional short plain-text caption like "Leading region" or
  "Top purchase driver" — descriptive, not a colored callout}}. Do not use "kpi" layout
  for more than one slide unless the source data clearly supports a second distinct set
  of headline numbers.
- Keep every text field crisp, not lengthy: "key_message" is a single short sentence
  (aim under 15 words) stating the headline takeaway, never a restatement of the whole
  insight. "bullets" are short fragments (aim under 10 words each), not full sentences —
  cut qualifiers and filler. KPI "label" is 2-4 words. A slide with a long key_message and
  a wall of bullets has failed the brief regardless of how accurate it is.
- Every slide except the title slide must set "source" to a citation string, or null if
  genuinely unavailable.
- layout "summary": exactly one slide, always last, closing the deck. Its title and content
  must read as a forward-looking takeaway/recap (e.g. "Key Takeaways", "Strategic Outlook",
  "What This Means"), never as a methodology, data-sources, or analyst-notes framing — this
  slide is the narrative close, not a citations/caveats appendix. Put any methodology or data
  caveats in that slide's "notes" field instead, not its title or bullets.
- Respect the user instruction above (e.g. slide count, audience, tone, focus areas) if given.
- Bullet count and detail level per slide should follow the user instruction if given (e.g.
  "just headlines" or "high-level" -> 1-2 short bullets per slide; "full detail" or
  "in-depth" -> more bullets, still within the brevity rule above). Default to 2-4 bullets
  per insight slide when no instruction speaks to detail level.
- On a slide that has a chart_type set, "bullets" render as a short summary band underneath
  the chart (up to 3 lines shown; anything beyond that is dropped, not truncated with a
  "+N more" marker, so keep it to 3 or fewer) — leave "bullets" empty on chart slides unless
  the user instruction explicitly asks for a descriptive/detailed deck (e.g. "detailed",
  "descriptive", "with summary", "explain each sector"). For a plain or unspecified request,
  chart slides should stay chart + one-line key_message only, exactly as today — do not add
  bullets by default just because the data has more to say.
- Return ONLY the JSON array. No markdown code fences, no prose, no explanation."""


def _parse_and_validate(raw_text: str) -> Optional[list[SlideSpec]]:
    """Parse the LLM's raw text as JSON and validate each item against
    SlideSpec. Any failure returns None so the caller falls back — no
    unvalidated LLM output ever reaches the renderer."""
    text = raw_text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()

    try:
        raw_slides = json.loads(text)
    except json.JSONDecodeError as e:
        logger.warning("LLM planner failed: response was not valid JSON: %s", e)
        return None

    if not isinstance(raw_slides, list) or not raw_slides:
        logger.warning(
            "LLM planner failed: response JSON was not a non-empty array (got %s).",
            type(raw_slides).__name__,
        )
        return None

    try:
        return [SlideSpec(**slide) for slide in raw_slides]
    except ValidationError as e:
        logger.warning("LLM planner failed: response did not match SlideSpec schema: %s", e)
        return None

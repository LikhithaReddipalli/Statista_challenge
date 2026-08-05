"""CLI + orchestration entry point.

generate_presentation() is the single pipeline function both the CLI and
the Streamlit UI (app.py) call — the only place that wires validator,
planner, and renderer together.
"""

import argparse
import io
import json
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from models import SlideSpec
from planner import plan_llm, plan_rule_based
from renderer import build_presentation
from validator import build_context, normalize_metrics, validate_input


@dataclass
class GenerationResult:
    """The finished deck plus a few facts about how it was built."""

    pptx_bytes: bytes
    slide_count: int
    insight_count: int
    used_llm: bool
    fell_back: bool
    elapsed_seconds: float


def _render_or_none(slides: list[SlideSpec]) -> tuple[Optional[bytes], Optional[Exception]]:
    """Try to render slides to .pptx bytes. Returns (bytes, None) on
    success or (None, exception) on failure, instead of raising — the
    caller decides whether a failure is worth retrying with a different
    set of slides."""
    buf = io.BytesIO()
    try:
        build_presentation(slides, buf)
        return buf.getvalue(), None
    except Exception as e:
        return None, e


def generate_presentation(
    analysis_json: dict,
    user_prompt: Optional[str] = None,
    use_llm: bool = True,
) -> GenerationResult:
    """Run the full pipeline: validate -> normalize -> plan -> render.

    Raises ValueError only if the input JSON is unusable (see
    validate_input). Any LLM-related failure — planning or rendering —
    falls back to plan_rule_based instead of raising, tracked via
    `fell_back` so callers can surface it to the user.
    """
    start = time.monotonic()

    analysis = validate_input(analysis_json)
    insights = analysis.get("key_insights", [])
    normalized = normalize_metrics(insights)
    context = build_context(analysis, normalized)

    fell_back = False
    used_llm_slides = False
    slides: list[SlideSpec]
    if use_llm:
        slides_or_none = plan_llm(context, user_instruction=user_prompt)
        if slides_or_none is None:
            print(
                "[warn] LLM planner failed or returned invalid output — "
                "falling back to the rule-based planner.",
                file=sys.stderr,
            )
            slides = plan_rule_based(context)
            fell_back = True
        else:
            slides = slides_or_none
            used_llm_slides = True
    else:
        slides = plan_rule_based(context)

    pptx_bytes, render_error = _render_or_none(slides)
    if render_error is not None:
        if not used_llm_slides:
            raise render_error  # rule-based slides — a render failure here is a real bug, not an LLM problem
        print(
            f"[warn] Rendering the LLM-planned deck failed ({type(render_error).__name__}: {render_error}) — "
            "falling back to the rule-based planner.",
            file=sys.stderr,
        )
        slides = plan_rule_based(context)
        fell_back = True
        pptx_bytes, retry_error = _render_or_none(slides)
        if retry_error is not None:
            # A fresh error, not chained: the deterministic path failing on
            # retry is unrelated to the original LLM-render failure, and
            # chaining would obscure the real cause as "exception while
            # handling another exception".
            raise RuntimeError(
                f"rule-based fallback render also failed: {type(retry_error).__name__}: {retry_error}"
            ) from None

    elapsed = time.monotonic() - start
    return GenerationResult(
        pptx_bytes=pptx_bytes,
        slide_count=len(slides),
        insight_count=len(insights),
        used_llm=use_llm and not fell_back,
        fell_back=fell_back,
        elapsed_seconds=elapsed,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a PowerPoint deck from a Statista AI-chat analysis JSON.")
    parser.add_argument("input_path", type=str, help="Path to the input analysis JSON file.")
    parser.add_argument(
        "--no-llm",
        dest="use_llm",
        action="store_false",
        help="Skip the LLM planner and use the rule-based planner directly (no AI, no GEMINI_API_KEY needed).",
    )
    parser.add_argument(
        "--prompt",
        type=str,
        default=None,
        help="Optional user instruction for the LLM planner (e.g. slide count, audience, focus area). Ignored with --no-llm.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default="presentation.pptx",
        help="Output .pptx path (default: presentation.pptx).",
    )
    args = parser.parse_args()

    input_path = Path(args.input_path)
    if not input_path.exists():
        print(f"error: input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    try:
        analysis_json = json.loads(input_path.read_text())
    except json.JSONDecodeError as e:
        print(f"error: input file is not valid JSON: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        result = generate_presentation(analysis_json, user_prompt=args.prompt, use_llm=args.use_llm)
    except ValueError as e:
        print(f"error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        # A genuine pipeline bug, not bad input or an LLM problem (those
        # are handled inside generate_presentation) — exit cleanly instead
        # of printing a raw traceback.
        print(f"error: presentation generation failed unexpectedly: {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(1)

    Path(args.output).write_bytes(result.pptx_bytes)

    planner_used = "LLM" if result.used_llm else ("rule-based (fallback)" if result.fell_back else "rule-based")
    print(f"Saved {args.output} — {result.slide_count} slides, planner: {planner_used}, {result.elapsed_seconds:.2f}s")


if __name__ == "__main__":
    main()

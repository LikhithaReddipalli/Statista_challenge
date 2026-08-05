# AI Presentation Generator

Turns a structured Statista AI-chat analysis (JSON) into a polished, editable PowerPoint deck.

Given the JSON and an optional instruction ("5-slide exec summary", "detailed deck"), it plans a
slide-by-slide structure — title, summary, one chart or text slide per key insight — picking a
chart type (ranking / comparison / pie / trend) that matches what each metric actually means, then
renders it to a `.pptx` with real, editable PowerPoint charts.

AI is scoped to *planning* (what the slides should say); rendering to `.pptx` is 100%
deterministic code. See [DESIGN.md](DESIGN.md) for the full rationale, chart-type rules,
validation/fallback behavior, and known limitations.

```
Input JSON → validate → normalize → plan (LLM or rule-based) → validate plan → render → .pptx
```

```
project/
  app.py         Streamlit UI 
  main.py        CLI + generate_presentation() — the one function both app.py and the CLI call
  planner.py     plan_llm() / plan_rule_based()
  renderer.py    SlideSpec → .pptx (python-pptx, native charts)
  validator.py   validate_input / normalize_metrics / build_context
  models.py      Metric (dataclass), SlideSpec (pydantic)
  sample_data/   the provided Gen Z example JSON
  results/       example output — see below
  tests/         deliberately malformed inputs + a script that exercises them
```

## Setup

```bash
git clone https://github.com/LikhithaReddipalli/Statista_challenge.git
cd Statista_challenge
```

Requires **Python 3.12** (see `.python-version`; pinned because that's what
this was built and tested against — see `requirements.txt` for details).

```bash
python3.12 -m venv virtual
./virtual/bin/pip install -r requirements.txt
```

For the LLM planner, copy `.env.example` to `.env` and fill in your `GEMINI_API_KEY`
(get one at [aistudio.google.com/apikey](https://aistudio.google.com/apikey)):

```bash
cp .env.example .env
```

`GEMINI_MODEL` and `GEMINI_TIMEOUT_SECONDS` are optional — see `.env.example` for defaults.

## Running it

### CLI

```bash
# LLM planner (default), with an instruction; falls back automatically on failure
./virtual/bin/python main.py sample_data/gen_z_purchase_behavior_analysis.json \
  --prompt "5-slide deck for C-level executives, focus on brand strategy" \
  -o presentation.pptx

# Rule-based planner only (no AI, no GEMINI_API_KEY needed)
./virtual/bin/python main.py sample_data/gen_z_purchase_behavior_analysis.json --no-llm
```

### Web UI

```bash
./virtual/bin/streamlit run app.py
```

Upload a JSON file, optionally type an instruction, leave "Use LLM Planner" checked (or
uncheck it for the rule-based planner), click Generate, download the `.pptx`.

## Example output

`results/` contains two decks generated from `sample_data/gen_z_purchase_behavior_analysis.json`,
so both planning paths can be inspected directly:

- `presentation.pptx` — rule-based planner (`--no-llm`)
- `presentation_With_LLM.pptx` — LLM planner (Gemini), including a KPI-card slide

## Libraries & LLM used

- **python-pptx** — the standard, actively-maintained Python library with a real *native*
  PowerPoint chart API (as opposed to pasting in a static image), so the output stays editable
  in PowerPoint.
- **pydantic** — validates `SlideSpec`, the one object that crosses the LLM trust boundary, and
  generates the LLM prompt's JSON schema straight from the model so the two can't drift apart.
- **Gemini, via langchain-google-genai** — a thin wrapper for a single one-shot `invoke()` call
  in `plan_llm()`; no chains/agents, just the model client. Chosen for the planning step because
  slide structure genuinely benefits from judgment (how many slides, what to cut) that a fixed
  script can't make well — see [DESIGN.md § Architecture](DESIGN.md#architecture).
- **python-dotenv** — loads `GEMINI_API_KEY` from `.env` instead of requiring a shell export.
- **streamlit** — the whole UI is one form (upload, text area, checkbox, button, download);
  Streamlit covers it in well under 100 lines with no separate frontend build.

## Known limitations

- No pytest unit suite, just end-to-end fallback/degradation checks.
- LLM planner output isn't deterministic.
- No LLM retry or response caching.
- Only four chart types (ranking, comparison, pie, trend).
- Pie/comparison split is a simple sum-check heuristic.
- `_metric_from_raw`'s generic fallback can pick the wrong field on an unfamiliar shape.
- Insight-level fields (`category`, `insight`, `source`) have no alias fallback.
- `key_message`/`bullets` aren't length-validated like `KpiCard` fields are.
- Text-overflow guards are heuristic, not a real layout engine.
- Streamlit preview is minimal — no in-browser slide preview.

**Trade-offs for each:** [DESIGN.md § Known limitations](DESIGN.md#known-limitations).

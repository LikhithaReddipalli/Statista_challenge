"""Streamlit UI — a thin shell over main.generate_presentation().

No pipeline logic here: it only collects input, calls
generate_presentation(), and displays the result.
"""

import json

import streamlit as st

from main import generate_presentation

st.set_page_config(page_title="AI Presentation Generator", page_icon="📊", layout="centered")

st.title("AI Presentation Generator")
st.caption("Turn a Statista AI-chat analysis JSON into a polished, editable PowerPoint deck.")

uploaded_file = st.file_uploader("Upload Analysis JSON", type=["json"])

user_prompt = st.text_area(
    "Presentation Instructions (optional)",
    placeholder="e.g. Create a 5-slide presentation for C-level executives, focus on digital purchase behavior.",
    height=80,
)

use_llm = st.checkbox(
    "Use LLM Planner",
    value=True,
    help="Uses Gemini to design the narrative (slide count, ordering, key messages). "
    "Falls back automatically to the rule-based planner if it fails or GEMINI_API_KEY isn't set.",
)

generate_clicked = st.button("Generate Presentation", type="primary", disabled=uploaded_file is None)

if uploaded_file is not None:
    st.caption(f"File: {uploaded_file.name}")
    try:
        analysis_preview = json.loads(uploaded_file.getvalue().decode("utf-8"))
        categories = [
            i.get("category", "Untitled")
            for i in analysis_preview.get("analysis", {}).get("key_insights", [])
        ]
        if categories:
            st.caption("Detected insight categories: " + ", ".join(categories))
        else:
            st.caption("No key_insights found — will generate a summary-only presentation.")
    except (json.JSONDecodeError, UnicodeDecodeError):
        st.caption("Could not preview this file — check it's valid JSON below.")

if generate_clicked and uploaded_file is not None:
    with st.spinner("Generating presentation..."):
        try:
            raw_bytes = uploaded_file.getvalue()
            analysis_json = json.loads(raw_bytes.decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError) as e:
            st.error(f"Uploaded file is not valid JSON: {e}")
            st.stop()

        try:
            result = generate_presentation(
                analysis_json,
                user_prompt=user_prompt.strip() or None,
                use_llm=use_llm,
            )
        except ValueError as e:
            st.error(f"Could not build a presentation from this input: {e}")
            st.stop()
        except Exception as e:
            # A genuine pipeline bug, not bad input or an LLM problem
            # (those raise ValueError, handled above).
            st.error(f"Presentation generation failed unexpectedly: {type(e).__name__}: {e}")
            st.stop()

    st.success("Presentation generated")

    planner_label = "LLM planner" if result.used_llm else ("rule-based (LLM fallback)" if result.fell_back else "rule-based planner")
    st.markdown(
        f"""
**Generation Summary**

- Input validated
- {result.insight_count} insight categor{"y" if result.insight_count == 1 else "ies"} detected
- Planner used: **{planner_label}**
- {result.slide_count} slides generated
- Time: {result.elapsed_seconds:.2f}s
"""
    )
    if result.fell_back:
        st.warning("LLM planner failed or returned invalid output — used the rule-based planner instead.")

    st.download_button(
        "Download PPTX",
        data=result.pptx_bytes,
        file_name="presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

import json
import re

import streamlit as st
from anthropic import Anthropic

MODEL = "claude-sonnet-4-6"

QUARTER_LABELS = ["Q-3", "Q-2", "Q-1", "Latest_Q"]

THEME_ANALYSIS_SYSTEM_PROMPT = """You are given four quarters of earnings call Q&A bullet-point summaries for the same company, ordered chronologically and labeled Q-3 (oldest), Q-2, Q-1, and Latest_Q (most recent).

For each consecutive delta (Q-3 vs Q-2, Q-2 vs Q-1, Q-1 vs Latest_Q):
1. Find 3 common themes that are heavily discussed during the Q&A sessions of both quarters in that delta.
2. Find 2 new or emerging themes that are heavily discussed during the Q&A session of the later quarter in that delta but were lightly or never discussed during the earlier quarter.
3. Rank all 5 of those themes (the 3 common plus the 2 emerging) by how much discussion volume each received in the LATER quarter's Q&A session only (ignore the earlier quarter's volume for this ranking). Keep only the top 3 themes by that ranking, ordered from highest to lowest volume.

Rules:
- "Heavily discussed" / theme volume is judged by the amount of Q&A dialogue (number of words) spent discussing the theme.
- Each theme must be very high-level and neutral (e.g. "Proprietary Silicon", not "Proprietary Silicon Development Timing").
- Each theme in a delta quarter must be over 80% different semantically compared to the others.
- Each theme name must be 2-3 words.

Output format - return ONLY valid JSON, no commentary, no markdown code fences, no preamble. Structure:
{
  "<period of Q-2>": {"themes": ["...", "...", "..."]},
  "<period of Q-1>": {"themes": ["...", "...", "..."]},
  "<period of Latest_Q>": {"themes": ["...", "...", "..."]}
}
Each delta's key in the JSON must be the period label (e.g. "Q1 2026") of the LATER quarter in that delta. Each delta's "themes" list must contain exactly 3 entries: the top 3 themes after the step-3 ranking and cut, ordered from highest to lowest discussion volume in the later quarter, with no indication of which were originally common vs. emerging."""


def parse_md_metadata(filename: str):
    name = filename.rsplit(".", 1)[0]
    parts = name.split("-")
    ticker = " ".join(parts[:2]) if len(parts) >= 2 else None
    period_match = re.search(r"Q[1-4]\s+20\d{2}", name)
    period = period_match.group(0) if period_match else None
    return ticker, period


def quarter_sort_key(period: str):
    quarter, year = period.split()
    return int(year), int(quarter[1])


def analyze_themes(client: Anthropic, ordered_quarters):
    header = "Input Periods: " + ", ".join(
        f"{item['label']}:{item['period']}" for item in ordered_quarters
    )
    sections = "\n\n".join(
        f"=== {item['label']} ({item['period']}) ===\n{item['content']}"
        for item in ordered_quarters
    )
    user_content = f"{header}\n\n{sections}"

    with client.messages.stream(
        model=MODEL,
        max_tokens=32000,
        temperature=1,
        thinking={"type": "enabled", "budget_tokens": 20000},
        system=THEME_ANALYSIS_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}],
    ) as stream:
        response = stream.get_final_message()
    raw_text = "".join(block.text for block in response.content if block.type == "text")
    cleaned = re.sub(r"^```(?:json)?|```$", "", raw_text.strip(), flags=re.MULTILINE).strip()
    try:
        return json.loads(cleaned), None
    except json.JSONDecodeError as e:
        return None, f"Failed to parse JSON response: {e}\n\nRaw response:\n{raw_text}"


st.set_page_config(page_title="Cross-Quarter Theme Analysis", layout="wide")
st.title("Cross-Quarter Theme Analysis")

with st.sidebar:
    api_key = st.text_input("Anthropic API key", type="password")

theme_md_files = st.file_uploader(
    "Upload exactly 4 telegraphic .md files (one company, 4 consecutive quarters)",
    type=["md"],
    accept_multiple_files=True,
)

if theme_md_files:
    parsed_quarters = []
    for f in theme_md_files:
        ticker, period = parse_md_metadata(f.name)
        if period is None:
            st.error(f"{f.name}: could not parse a quarter/period (e.g. 'Q1 2026') from filename")
            continue
        parsed_quarters.append(
            {
                "filename": f.name,
                "ticker": ticker,
                "period": period,
                "content": f.read().decode("utf-8"),
            }
        )

    if len(parsed_quarters) != 4:
        st.warning(f"Need exactly 4 valid .md files; found {len(parsed_quarters)}.")
    else:
        parsed_quarters.sort(key=lambda item: quarter_sort_key(item["period"]))
        for label, item in zip(QUARTER_LABELS, parsed_quarters):
            item["label"] = label

        st.write({item["label"]: f"{item['ticker']} {item['period']}" for item in parsed_quarters})

        if st.button("Analyze Themes", disabled=not api_key):
            client = Anthropic(api_key=api_key)
            with st.spinner("Analyzing cross-quarter themes..."):
                themes, error = analyze_themes(client, parsed_quarters)
            if error:
                st.error(error)
            else:
                st.session_state["theme_analysis_result"] = themes
                st.json(themes)
elif not api_key:
    st.info("Enter your Anthropic API key in the sidebar to begin.")
else:
    st.info("Upload 4 telegraphic .md files to begin.")

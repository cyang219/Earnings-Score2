import io
import json
import re

import openpyxl
import streamlit as st
from anthropic import Anthropic
from openpyxl.utils import column_index_from_string

MODEL = "claude-sonnet-4-6"
EFFORT = "medium"

MIN_QUARTERS = 2

COMMON_THEME_SET = [
    "Top-line Growth (e.g. Revenue, Bookings, GMV, ARR, etc.)",
    "Bottom-line Expansion/Contraction (e.g. Margins, Return on Investment, Take-rate, etc.)",
    "Financial Guidance and Forward-looking Business Commentary",
]


# ============================================================
# PROMPTS
# ============================================================

def build_quarter_labels(n: int) -> list[str]:
    return [f"Q-{n - 1 - i}" for i in range(n - 1)] + ["Latest_Q"]


DESCRIPTION_LENGTH_RULE = (
    "The two fields combined must be 350-400 characters long, including spaces and punctuation"
)

READTHROUGH_PROMPT_TEMPLATE = """You are an equity research analyst specializing in sector read-throughs from earnings calls. You will be given a set of consecutive quarterly earnings call transcripts (or telegraphic summaries) from a single publicly traded company, ordered oldest to newest. The same task applies regardless of the company, sector, fiscal calendar, or number of quarters provided.
DEFINITIONS (canonical — all later references point back here):
- EXTERNAL BUSINESS FACTORS = industry/macro conditions management describes that read across to sector peers (e.g., end-market demand, customer/end-user spending, pricing environment, input or supply availability, competitive dynamics, regulatory or macro conditions). EXCLUDES the reporting company's own KPIs and internal initiatives (e.g., its own revenue, margins, segment results, product roadmap, capital returns, restructuring). Every step below uses external business factors ONLY.
- DELTA DIRECTION = in each delta the FIRST-named quarter is the LATER quarter and the SECOND-named quarter is the EARLIER quarter (e.g., in "Q-2_vs_Q-3", Q-2 is later, Q-3 is earlier). A signal describes how the later quarter changed versus the earlier quarter.
- DELTA = a QoQ comparison between one quarter and the immediately preceding quarter. With N quarters there are N-1 deltas, each comparing adjacent quarters.

{quarter_mapping}

TASK:

Step 1: QUARTER-LEVEL READ-THROUGH. For each quarter in the QUARTER MAPPING, infer the industry-wide read-through from management's message, using external business factors only (see DEFINITIONS). Use this only as internal reasoning; do not output Step 1.
Step 2: QOQ SIGNAL SCRATCHPAD. Start a <scratchpad> to assign a signal to each delta defined in the QUARTER MAPPING. Judge each signal only by how the later quarter's read-through changed versus the earlier quarter (see DELTA DIRECTION), not by whether the later quarter is positive or negative in absolute terms.

Select the single dominant external business factor for each delta using this ordered tie-breaker (apply in order; stop at the first criterion that picks one factor):
   1. A factor management discusses in BOTH quarters of the delta.
   2. Of those, the factor most relevant to industry peers / sector read-through.
   3. Of those, the factor management emphasizes most (most airtime / strongest language).
   4. Of those, a recurring factor over a one-off factor.

Classify the change in the dominant factor's read-through using this rule, then assign the signal:
   - Strengthened: the later quarter shows a clear positive change in the dominant external factor — e.g. management escalates language (such as "strong" to "insatiable"), raises a quantified external indicator, broadens the demand base, or removes a previously cited concern.
   - Weakened: the later quarter shows a clear negative change — management softens language, introduces or amplifies an external headwind, narrows the demand base, or walks back prior optimism.
   - Stable: no clear directional change — the dominant factor's framing, language intensity, and any cited external indicators are materially unchanged, OR positive and negative shifts in that factor roughly offset.

A Strengthened or Weakened signal requires a specific, citable QoQ change in the dominant factor (a language shift, a changed external indicator, or a changed demand-base statement). If no such concrete change can be named, the signal must be Stable. When evidence is genuinely balanced or ambiguous, default to Stable rather than guessing a direction.

Copy the exact template below and fill in the brackets, producing one line per delta. Do NOT write anything else inside the scratchpad.

<scratchpad>

{scratchpad_format}
</scratchpad>
   - Signal must be exactly one of: Strengthened, Weakened, Stable
   - Driver must be concrete, evidence-linked, and an external business factor

Step 3: SIGNAL FIELD. Derive the "signal" field mechanically from the scratchpad Signal for the same delta (not an independent judgment):
   - Strengthened -> "↑:"   |   Weakened -> "↓:"   |   Stable -> "→:"
If the field and the scratchpad Signal would ever disagree, the scratchpad Signal is authoritative.
Step 4: READ_THROUGH AND RATIONALE FIELDS. For each delta, produce two fields, each exactly one sentence (external business factors only — see DEFINITIONS):
   - "read_through": the inferred macro/industry read-through for the later quarter.
   - "rationale": (a) cites one specific QoQ contrast between the two quarters, and (b) explains why that contrast caused the read-through to strengthen, weaken, or stay stable, matching the delta's signal direction.
   - {description_length_rule}
   - Omit the final period "." at the end of both fields.
Before producing the JSON, verify for each delta that the rationale's described direction matches the signal arrow and the scratchpad Signal. If they disagree, the scratchpad Signal governs; revise the rationale to match.
JSON FORMAT: After </scratchpad>, output ONLY a valid JSON block using ```json tags, following the exact schema below, with one object per delta in the same order as the scratchpad. No text before or after the JSON block. Escape any double quotes inside field values.

```json
{json_format}
```"""


BULLBEAR_PROMPT_TEMPLATE = """You are an equity analyst evaluating forward 6–9 month bull/bear thesis changes from earnings calls. You will be given a set of consecutive quarterly earnings call transcripts (or telegraphic summaries) from a single publicly traded company, ordered oldest to newest. The same task applies regardless of the company, sector, fiscal calendar, or number of quarters provided.

DEFINITIONS (canonical — all later references point back here):
- FORWARD EXPECTATION = what an investor on a given side (bull or bear) would, based on that quarter's call, expect to happen to the business over the next 6–9 months. Infer expectations from the whole call but PRIORITIZE the Q&A dialogue, where management is tested and forward-looking concerns surface.
- BULL EXPECTATION = the forward case a bullish investor would hold; BEAR EXPECTATION = the forward case a bearish investor would hold. Each delta is evaluated separately for both sides. The two sides are NOT required to move in opposite directions — independently assess each side's evidence; it is valid for both to strengthen, both to weaken, or both to stay stable on the same delta.
- DELTA DIRECTION = in each delta the FIRST-named quarter is the LATER quarter and the SECOND-named quarter is the EARLIER quarter (e.g., in "Q-2_vs_Q-3", Q-2 is later, Q-3 is earlier). A signal describes how the later quarter's forward expectation changed versus the earlier quarter.
- DELTA = a QoQ comparison between one quarter and the immediately preceding quarter. With N quarters there are N-1 deltas, each comparing adjacent quarters.

{quarter_mapping}

TASK:

Step 1: FORWARD 6–9 MONTH BULL/BEAR EXPECTATIONS. For each quarter in the QUARTER MAPPING, infer both the bull and the bear forward expectation for the next 6–9 months (see DEFINITIONS), prioritizing the Q&A dialogue. Use this only as internal reasoning; do not output Step 1.

Step 2: QOQ SIGNAL SCRATCHPAD. Start a <scratchpad> to assign a signal to each side (bull and bear) of each delta defined in the QUARTER MAPPING. You MUST evaluate every delta. Judge each signal only by how the later quarter's forward expectation changed versus the earlier quarter (see DELTA DIRECTION), not by how positive or negative the expectation is in absolute terms.

Classify the change in each side's forward expectation using this rule, then assign the signal:
   - Strengthened: the later quarter shows a clear positive change in that side's forward case — e.g. management escalates supporting language, raises or affirms forward guidance/indicators that side relies on, resolves a prior concern, or analysts' Q&A questions shift from probing a risk to confirming an improvement (and management's answer corroborates it).
   - Weakened: the later quarter shows a clear negative change in that side's forward case — management softens or walks back supporting language, introduces or amplifies a forward risk, or analysts' Q&A questions surface a new risk that management cannot fully address.
   - Stable: no clear directional change — the forward case's framing, language intensity, and relevant forward indicators are materially unchanged, OR positive and negative shifts roughly offset.
A Strengthened or Weakened signal requires a specific, citable QoQ change in that side's forward case (a language shift, a changed forward indicator/guidance, or new Q&A evidence). If no such concrete change can be named, the signal must be Stable. When evidence is genuinely balanced or ambiguous, default to Stable rather than guessing a direction.

Copy the exact template below and fill in the brackets, producing the required lines per delta. Do NOT write anything else inside the scratchpad.
<scratchpad>
{scratchpad_format}
</scratchpad>
   - Signal must be exactly one of: Strengthened, Weakened, Stable
   - Driver must be concrete, evidence-linked, and tied to that side's forward case
   - Each signal must be evidence-linked to that side's forward case for those specific quarters

Step 3: TOPIC FIELD. Derive the "topic" field mechanically from the scratchpad Signal for that same side and delta (not an independent judgment):
   - Strengthened -> "↑:"   |   Weakened -> "↓:"   |   Stable -> "→:"
If the field and the scratchpad Signal would ever disagree, the scratchpad Signal is authoritative.

Step 4: EXPECTATION AND CONTEXT FIELDS. For each delta, produce three fields per side (bull and bear): "topic" (the arrow from Step 3), "expectation", and "context".
   - "expectation": one sentence stating the inferred forward-looking 6–9 month expectation for the later quarter on that side.
   - "context": one sentence (a) citing one specific QoQ change in that side's forward case and (b) explaining why that change caused the expectation to strengthen, weaken, or stay stable, matching that side's signal direction.
   - {description_length_rule} across the "expectation" and "context" fields combined.
   - STYLE: abridged, telegraphic style with abbreviations.
   - Omit the final period "." at the end of both fields.

Before producing the JSON, verify for each side of each delta that the context's described direction matches the topic arrow and the scratchpad Signal. If they disagree, the scratchpad Signal governs; revise the context to match.

JSON FORMAT: After </scratchpad>, output ONLY a valid JSON block using ```json tags, following the exact schema below, with one object per delta in the same order as the scratchpad. No text before or after the JSON block. Escape any double quotes inside field values.

```json
{json_format}
```"""


THEME_DELTA_PROMPT_TEMPLATE = """You are an earnings-call analyst comparing management messaging and analyst tone across investment themes for one quarter-over-quarter delta. You will be given consecutive telegraphic summaries of quarterly earnings calls from a single publicly traded company.

DEFINITIONS (canonical — all later references point back here):
- MANAGEMENT MESSAGE = what management conveys about a theme in the prepared remarks and their Q&A answers: the substance, confidence, and framing of their commentary on that theme.
- ANALYST TONE = the stance reflected in analysts' Q&A questions on a theme: whether they probe risk, press skeptically, or signal confidence, and how management's answers are received.
- DELTA DIRECTION = the FIRST-named quarter in QUARTERS TO ANALYZE is the LATER quarter; the SECOND-named is the EARLIER quarter. A signal describes how the later quarter changed versus the earlier quarter.
- THEME INDEPENDENCE = each theme is analyzed in complete isolation. Evidence, tone, framing, or sentiment from one theme MUST NOT be transferred to another. Management and analyst sides are also judged independently per theme; they need not move in the same direction.
- DATA BOUNDARY = use ONLY the two quarters' transcript data provided. Do NOT use outside information.

THEME LIST:
{theme_list}

QUARTERS TO ANALYZE:
{quarter_delta}

TASK:

Step 1: PER-THEME READ. For each theme in THEME LIST, using only the two quarters (see DATA BOUNDARY), internally summarize the later quarter's management message and analyst tone for that theme, and how each compares to the earlier quarter. If a theme is absent in the earlier quarter, note this now so the special-case rule in Step 2 applies correctly. Process each theme fully and in isolation before moving to the next (see THEME INDEPENDENCE). Use this only as internal reasoning; do not output Step 1.

Step 2: SIGNAL SCRATCHPAD. Start a <scratchpad> assigning a management signal and an analyst signal to each theme. Judge each signal only by the change versus the earlier quarter (see DELTA DIRECTION), not by absolute positivity or negativity.

Classify each side (management, then analyst) for each theme using this rule, then assign the signal:
- Improved: a clear positive change vs the earlier quarter — management escalates confidence/affirms progress/resolves a prior concern on that theme, OR analyst questions shift from probing a risk toward confirming an improvement.
- Worsened: a clear negative change — management softens, walks back, or introduces a new concern on that theme, OR analyst questions surface a new risk management cannot fully address.
- Stable: no clear directional change — framing, confidence, and tone on that theme are materially unchanged, OR positive and negative shifts roughly offset.
An Improved or Worsened signal requires a specific, citable QoQ change on that theme for that side (see DATA BOUNDARY). If no such concrete change can be named, the signal must be Stable. When evidence is genuinely balanced or ambiguous, default to Stable rather than guessing a direction.
Special case — theme absent in the earlier quarter: if the theme is not discussed in the earlier quarter, judge by the later quarter's absolute message/tone: positive = Improved, negative = Worsened, neutral or limited = Stable. Do not mark Improved or Worsened solely because the theme appeared.

Copy the exact template below and fill in the brackets, producing one line per theme in THEME LIST order. Do NOT write anything else inside the scratchpad.
<scratchpad>
{scratchpad_format}
</scratchpad>
- Signal must be exactly one of: Improved, Worsened, Stable
- Each signal must be evidence-linked to that theme's own data (see THEME INDEPENDENCE)

Step 3: SIGNAL FIELD. Derive the "signal" field mechanically from the scratchpad Signal for that same theme and side (not an independent judgment):
- Improved -> "↑:"   |   Worsened -> "↓:"   |   Stable -> "→:"
If the field and the scratchpad Signal would ever disagree, the scratchpad Signal is authoritative.

Step 4: MESSAGE/TONE AND RATIONALE FIELDS. For each theme, produce three fields per side (mgmt and analyst): "signal" (the arrow from Step 3), then for mgmt "message" and for analyst "tone", and "rationale" for both:
- "message" (mgmt) / "tone" (analyst): one sentence stating the later-quarter management message or analyst tone on that theme.
- "rationale": one sentence (a) citing one specific QoQ change on that theme and (b) explaining why it improved, worsened, or stayed stable, matching that side's signal direction.
- {description_length_rule} across the "message"/"tone" field and its "rationale" field combined.
- STYLE: abridged, telegraphic style with abbreviations.
- Omit the final period "." at the end of every field.

Before producing the JSON, verify that (a) every theme in THEME LIST is present, (b) each theme's mgmt and analyst signals comply with THEME INDEPENDENCE, and (c) each rationale's described direction matches its signal arrow and the scratchpad Signal. If any disagree, the scratchpad Signal governs; revise to match.

JSON FORMAT: After </scratchpad>, output ONLY a valid JSON block using ```json tags, following the exact schema below, with one object per theme in THEME LIST order. Replace the theme placeholders with the exact theme names from THEME LIST. Do not output any text between </scratchpad> and the JSON block, or after the JSON block. Escape any double quotes inside field values.

```json
{json_format}
```"""


def build_delta_pairs(labels: list[str]) -> list[tuple[str, str]]:
    return [(labels[i + 1], labels[i]) for i in range(len(labels) - 1)]


def build_quarter_mapping(ordered_quarters) -> str:
    lines = [f"{item['label']} = {item['period']}" for item in ordered_quarters]
    return "QUARTER MAPPING:\n" + "\n".join(lines)


def build_readthrough_scratchpad_format(labels: list[str]) -> str:
    lines = [
        f"{later}_vs_{earlier} ({later} vs {earlier}) -> Signal: [Signal]; "
        f"Driver: [2–5 word external business factor that drove signal]"
        for later, earlier in build_delta_pairs(labels)
    ]
    return "\n".join(lines)


def build_readthrough_json_skeleton(labels: list[str]) -> str:
    entries = [
        f"""    {{
      "period": "{later}_vs_{earlier}",
      "signal": "↑: or ↓: or →:",
      "read_through": "one-sentence macro/industry read-through",
      "rationale": "one-sentence QoQ contrast explaining the signal direction"
    }}"""
        for later, earlier in build_delta_pairs(labels)
    ]
    return "{\n  \"historical_analysis\": [\n" + ",\n".join(entries) + "\n  ]\n}"


def build_readthrough_system_prompt(ordered_quarters) -> str:
    labels = [item["label"] for item in ordered_quarters]
    return READTHROUGH_PROMPT_TEMPLATE.format(
        quarter_mapping=build_quarter_mapping(ordered_quarters),
        scratchpad_format=build_readthrough_scratchpad_format(labels),
        description_length_rule=DESCRIPTION_LENGTH_RULE,
        json_format=build_readthrough_json_skeleton(labels),
    )


def build_bullbear_scratchpad_format(labels: list[str]) -> str:
    lines = []
    for later, earlier in build_delta_pairs(labels):
        lines.append(
            f"{later}_vs_{earlier} ({later} vs {earlier}) -> Bull Signal: [Signal]; "
            f"Bull Driver: [2–5 word forward-case evidence that drove signal]"
        )
        lines.append(
            f"{later}_vs_{earlier} ({later} vs {earlier}) -> Bear Signal: [Signal]; "
            f"Bear Driver: [2–5 word forward-case evidence that drove signal]"
        )
    return "\n".join(lines)


def build_bullbear_json_skeleton(labels: list[str]) -> str:
    entries = [
        f"""    {{
      "period": "{later}_vs_{earlier}",
      "bull": {{
        "topic": "↑: or ↓: or →:",
        "expectation": "one-sentence bullish forward-looking expectation",
        "context": "one-sentence QoQ context explaining the signal direction"
      }},
      "bear": {{
        "topic": "↑: or ↓: or →:",
        "expectation": "one-sentence bearish forward-looking expectation",
        "context": "one-sentence QoQ context explaining the signal direction"
      }}
    }}"""
        for later, earlier in build_delta_pairs(labels)
    ]
    return "{\n  \"historical_analysis\": [\n" + ",\n".join(entries) + "\n  ]\n}"


def build_bullbear_system_prompt(ordered_quarters) -> str:
    labels = [item["label"] for item in ordered_quarters]
    return BULLBEAR_PROMPT_TEMPLATE.format(
        quarter_mapping=build_quarter_mapping(ordered_quarters),
        scratchpad_format=build_bullbear_scratchpad_format(labels),
        description_length_rule=DESCRIPTION_LENGTH_RULE,
        json_format=build_bullbear_json_skeleton(labels),
    )


def build_theme_list(themes: list[str]) -> str:
    return "\n".join(f"- {theme}" for theme in themes)


def build_quarter_delta_mapping(later_item, earlier_item) -> str:
    return (
        f"{later_item['label']} = {later_item['period']} (LATER)\n"
        f"{earlier_item['label']} = {earlier_item['period']} (EARLIER)"
    )


def build_theme_delta_scratchpad_format(themes: list[str]) -> str:
    return "\n".join(
        f"{theme} -> Mgmt Signal: [Signal]; Mgmt Driver: [2–5 word evidence]; "
        f"Analyst Signal: [Signal]; Analyst Driver: [2–5 word evidence]"
        for theme in themes
    )


def build_theme_delta_json_skeleton(themes: list[str]) -> str:
    entries = [
        f"""    {{
      "theme": "{theme}",
      "mgmt": {{
        "signal": "↑: or ↓: or →:",
        "message": "one-sentence management message",
        "rationale": "one-sentence QoQ contrast explaining the signal direction"
      }},
      "analyst": {{
        "signal": "↑: or ↓: or →:",
        "tone": "one-sentence analyst tone",
        "rationale": "one-sentence QoQ contrast explaining the signal direction"
      }}
    }}"""
        for theme in themes
    ]
    return "{\n  \"theme_analysis\": [\n" + ",\n".join(entries) + "\n  ]\n}"


def build_theme_delta_system_prompt(themes: list[str], later_item, earlier_item) -> str:
    return THEME_DELTA_PROMPT_TEMPLATE.format(
        theme_list=build_theme_list(themes),
        quarter_delta=build_quarter_delta_mapping(later_item, earlier_item),
        scratchpad_format=build_theme_delta_scratchpad_format(themes),
        description_length_rule=DESCRIPTION_LENGTH_RULE,
        json_format=build_theme_delta_json_skeleton(themes),
    )


# ============================================================
# ANALYSIS CALLS
# ============================================================

def parse_fenced_json_response(raw_text: str):
    match = re.search(r"```json\s*(.*?)```", raw_text, flags=re.DOTALL)
    if not match:
        return None, f"Failed to find JSON block in response\n\nRaw response:\n{raw_text}"
    try:
        return json.loads(match.group(1).strip()), None
    except json.JSONDecodeError as e:
        return None, f"Failed to parse JSON response: {e}\n\nRaw response:\n{raw_text}"


def build_readthrough_params(ordered_quarters) -> dict:
    sections = "\n\n".join(
        f"=== {item['label']} ({item['period']}) ===\n{item['content']}"
        for item in ordered_quarters
    )
    return {
        "model": MODEL,
        "max_tokens": 32000,
        "thinking": {"type": "enabled", "budget_tokens": 10000},
        "output_config": {"effort": EFFORT},
        "system": build_readthrough_system_prompt(ordered_quarters),
        "messages": [{"role": "user", "content": sections}],
    }


def build_bullbear_params(ordered_quarters) -> dict:
    sections = "\n\n".join(
        f"=== {item['label']} ({item['period']}) ===\n{item['content']}"
        for item in ordered_quarters
    )
    return {
        "model": MODEL,
        "max_tokens": 32000,
        "thinking": {"type": "enabled", "budget_tokens": 10000},
        "output_config": {"effort": EFFORT},
        "system": build_bullbear_system_prompt(ordered_quarters),
        "messages": [{"role": "user", "content": sections}],
    }


def build_theme_delta_params(later_item, earlier_item) -> dict:
    sections = "\n\n".join(
        f"=== {item['label']} ({item['period']}) ===\n{item['content']}"
        for item in (later_item, earlier_item)
    )
    return {
        "model": MODEL,
        "max_tokens": 32000,
        "thinking": {"type": "enabled", "budget_tokens": 10000},
        "output_config": {"effort": EFFORT},
        "system": build_theme_delta_system_prompt(COMMON_THEME_SET, later_item, earlier_item),
        "messages": [{"role": "user", "content": sections}],
    }


def _stream_text(client: Anthropic, params: dict) -> str:
    with client.messages.stream(**params) as stream:
        response = stream.get_final_message()
    return "".join(block.text for block in response.content if block.type == "text")


def analyze_readthrough(client: Anthropic, ordered_quarters):
    raw_text = _stream_text(client, build_readthrough_params(ordered_quarters))
    return parse_fenced_json_response(raw_text)


def analyze_bullbear(client: Anthropic, ordered_quarters):
    raw_text = _stream_text(client, build_bullbear_params(ordered_quarters))
    return parse_fenced_json_response(raw_text)


def analyze_theme_delta(client: Anthropic, later_item, earlier_item):
    raw_text = _stream_text(client, build_theme_delta_params(later_item, earlier_item))
    return parse_fenced_json_response(raw_text)


# ============================================================
# DB EXPORT
# ============================================================

def find_ticker_worksheet(workbook: openpyxl.Workbook, ticker: str):
    for name in workbook.sheetnames:
        if name.strip().upper() == ticker.strip().upper():
            return name
    return None


DB_PERIOD_COLUMN = 1
DB_FIRST_DATA_ROW = 6


def period_to_db_format(period: str) -> str:
    quarter, year = period.split()
    return f"{year[-2:]}Q{quarter[1]}"


def find_period_row(worksheet, period: str) -> int | None:
    target = period_to_db_format(period)
    for row in worksheet.iter_rows(min_row=DB_FIRST_DATA_ROW, min_col=DB_PERIOD_COLUMN, max_col=DB_PERIOD_COLUMN):
        cell = row[0]
        value = cell.value
        if value is not None and str(value).strip().upper() == target:
            return cell.row
    return None


DB_READTHROUGH_COL = column_index_from_string("DX")
DB_BULL_COL = DB_READTHROUGH_COL + 1
DB_BEAR_COL = DB_READTHROUGH_COL + 2
DB_COMMON_THEME_COLUMNS = {
    COMMON_THEME_SET[0]: (DB_READTHROUGH_COL + 3, DB_READTHROUGH_COL + 4),
    COMMON_THEME_SET[1]: (DB_READTHROUGH_COL + 5, DB_READTHROUGH_COL + 6),
    COMMON_THEME_SET[2]: (DB_READTHROUGH_COL + 7, DB_READTHROUGH_COL + 8),
}


def format_signal_cell(signal: str, text: str, rationale: str) -> str:
    return f"{signal} {text} — {rationale}"


def build_delta_lookup(result: dict) -> dict:
    return {entry["period"]: entry for entry in result.get("historical_analysis", [])}


def write_db_row(worksheet, row: int, readthrough_entry: dict, bullbear_entry: dict, theme_entry: dict):
    if readthrough_entry is not None:
        worksheet.cell(row=row, column=DB_READTHROUGH_COL).value = format_signal_cell(
            readthrough_entry["signal"], readthrough_entry["read_through"], readthrough_entry["rationale"]
        )

    if bullbear_entry is not None:
        worksheet.cell(row=row, column=DB_BULL_COL).value = format_signal_cell(
            bullbear_entry["bull"]["topic"], bullbear_entry["bull"]["expectation"], bullbear_entry["bull"]["context"]
        )
        worksheet.cell(row=row, column=DB_BEAR_COL).value = format_signal_cell(
            bullbear_entry["bear"]["topic"], bullbear_entry["bear"]["expectation"], bullbear_entry["bear"]["context"]
        )

    if theme_entry is not None:
        theme_lookup = {entry["theme"]: entry for entry in theme_entry.get("theme_analysis", [])}
        for theme_name, (mgmt_col, analyst_col) in DB_COMMON_THEME_COLUMNS.items():
            entry = theme_lookup.get(theme_name)
            if entry is None:
                continue
            worksheet.cell(row=row, column=mgmt_col).value = format_signal_cell(
                entry["mgmt"]["signal"], entry["mgmt"]["message"], entry["mgmt"]["rationale"]
            )
            worksheet.cell(row=row, column=analyst_col).value = format_signal_cell(
                entry["analyst"]["signal"], entry["analyst"]["tone"], entry["analyst"]["rationale"]
            )


# ============================================================
# UTILITIES
# ============================================================

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


# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(page_title="Cross-Quarter Theme Analysis (Lite)", layout="wide")
st.title("Cross-Quarter Theme Analysis — Lite")

with st.sidebar:
    api_key = st.text_input("Anthropic API key", type="password")

theme_md_files = st.file_uploader(
    "Upload telegraphic .md files (one company, 2 or more consecutive quarters)",
    type=["md"],
    accept_multiple_files=True,
)

db_file = st.file_uploader(
    "Upload DB format.xlsx (workbook to update with this analysis)",
    type=["xlsx"],
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

    if len(parsed_quarters) < MIN_QUARTERS:
        st.warning(f"Need at least {MIN_QUARTERS} valid .md files; found {len(parsed_quarters)}.")
    else:
        parsed_quarters.sort(key=lambda item: quarter_sort_key(item["period"]))
        labels = build_quarter_labels(len(parsed_quarters))
        for label, item in zip(labels, parsed_quarters):
            item["label"] = label

        st.write({item["label"]: f"{item['ticker']} {item['period']}" for item in parsed_quarters})

        ticker = parsed_quarters[0]["ticker"]
        sheet_name = None
        period_rows = {}
        if db_file:
            workbook = openpyxl.load_workbook(db_file)
            sheet_name = find_ticker_worksheet(workbook, ticker)
            if sheet_name is None:
                st.error(f"No tab found matching ticker '{ticker}' in {db_file.name}")
            else:
                st.success(f"Found matching tab: '{sheet_name}'")
                worksheet = workbook[sheet_name]
                missing_periods = []
                for item in parsed_quarters:
                    row = find_period_row(worksheet, item["period"])
                    if row is None:
                        missing_periods.append(item["period"])
                    else:
                        period_rows[item["period"]] = row
                if missing_periods:
                    st.error(
                        f"No row found in '{sheet_name}' for: {', '.join(missing_periods)} "
                        f"(expected column A format like '{period_to_db_format(parsed_quarters[0]['period'])}')"
                    )
                else:
                    st.success(
                        "Matched rows: "
                        + ", ".join(f"{period} -> row {row}" for period, row in period_rows.items())
                    )
        else:
            st.info(f"Upload DB format.xlsx above to update the '{ticker}' tab with this analysis.")

        rows_ready = db_file and sheet_name and len(period_rows) == len(parsed_quarters)

        if st.button("Analyze", disabled=not (api_key and rows_ready)):
            client = Anthropic(api_key=api_key)

            with st.spinner("Analyzing sector read-through..."):
                readthrough, readthrough_error = analyze_readthrough(client, parsed_quarters)
            readthrough_lookup = {}
            if readthrough_error:
                st.error(readthrough_error)
            else:
                st.session_state["readthrough_analysis_result"] = readthrough
                st.json(readthrough)
                readthrough_lookup = build_delta_lookup(readthrough)

            with st.spinner("Analyzing bull/bear forward expectations..."):
                bullbear, bullbear_error = analyze_bullbear(client, parsed_quarters)
            bullbear_lookup = {}
            if bullbear_error:
                st.error(bullbear_error)
            else:
                st.session_state["bullbear_analysis_result"] = bullbear
                st.json(bullbear)
                bullbear_lookup = build_delta_lookup(bullbear)

            theme_delta_by_key = {}
            for later_item, earlier_item in zip(parsed_quarters[1:], parsed_quarters[:-1]):
                delta_key = f"{later_item['label']}_vs_{earlier_item['label']}"
                with st.spinner(f"Analyzing common themes for {delta_key}..."):
                    theme_result, theme_error = analyze_theme_delta(client, later_item, earlier_item)
                if theme_error:
                    st.error(f"{delta_key}: {theme_error}")
                theme_delta_by_key[delta_key] = theme_result
                st.write(f"Common-theme analysis — {delta_key}")
                st.json(theme_result)

            st.session_state["theme_delta_analysis_result"] = theme_delta_by_key

            for later_item, earlier_item in zip(parsed_quarters[1:], parsed_quarters[:-1]):
                delta_key = f"{later_item['label']}_vs_{earlier_item['label']}"
                write_db_row(
                    worksheet,
                    period_rows[later_item["period"]],
                    readthrough_lookup.get(delta_key),
                    bullbear_lookup.get(delta_key),
                    theme_delta_by_key.get(delta_key),
                )

            output_buffer = io.BytesIO()
            workbook.save(output_buffer)
            st.download_button(
                label=f"Download updated {db_file.name}",
                data=output_buffer.getvalue(),
                file_name=db_file.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
elif not api_key:
    st.info("Enter your Anthropic API key in the sidebar to begin.")
else:
    st.info(f"Upload at least {MIN_QUARTERS} telegraphic .md files to begin.")

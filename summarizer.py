import concurrent.futures
import io
import zipfile

import streamlit as st
from pypdf import PdfReader
from anthropic import Anthropic

MODEL = "claude-sonnet-4-6"

SYSTEM_PROMPT = """You are given the raw text of a corporate earnings call transcript.

Step 1 - Filter out, do not include in your output:
- Operator dialogue (e.g. "Operator: Thank you, please go ahead with your question.")
- Legal disclaimers, safe-harbor / forward-looking-statement boilerplate, and any header/footer text (page numbers, copyright notices, transcript provider branding, "for internal use only", etc.)
- Simple greetings/thankings and their reciprocal responses (e.g. "Good morning everyone", "Thank you for joining the call", "Thanks, operator", "Thank you, next question please")

Step 2 - Rewrite everything that remains as a telegraphic-style rewrite:
- Convert every remaining sentence into a bullet point
- Each bullet must be written in telegraphic style: drop filler words, articles, and connective phrasing; keep it terse and clipped, like a notetaker's shorthand
- Do not drop or summarize away any substantive content, facts, numbers, or claims - every piece of real information from the transcript must appear in some bullet
- Preserve the speaker attribution and the original order of the discussion (group bullets under the speaker's name/role if speaker names are identifiable in the transcript)

Output only the resulting bulleted telegraphic rewrite. Do not add commentary, preamble, or a summary of what you did."""


def extract_pdf_text(data: bytes) -> str:
    reader = PdfReader(io.BytesIO(data))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def rewrite_transcript(client: Anthropic, transcript_text: str) -> str:
    response = client.messages.create(
        model=MODEL,
        max_tokens=16000,
        temperature=1,
        thinking={"type": "enabled", "budget_tokens": 10000},
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": transcript_text}],
    )
    return "".join(block.text for block in response.content if block.type == "text")


def process_file(data, client):
    try:
        transcript_text = extract_pdf_text(data)
        if not transcript_text.strip():
            return None, "no extractable text found"
        result = rewrite_transcript(client, transcript_text)
        return result, None
    except Exception as e:
        return None, str(e)


st.set_page_config(page_title="Earnings Call Telegraphic Rewriter", layout="wide")
st.title("Earnings Call Transcript - Telegraphic Bullet Rewriter")

with st.sidebar:
    api_key = st.text_input("Anthropic API key", type="password")

uploaded_files = st.file_uploader(
    "Drag and drop earnings call transcript PDFs",
    type=["pdf"],
    accept_multiple_files=True,
)

if st.button("Process transcripts", disabled=not (uploaded_files and api_key)):
    client = Anthropic(api_key=api_key)
    # Read all file bytes on main thread before parallelizing -
    # Streamlit's UploadedFile is not thread-safe to read from concurrently.
    file_data = [(f.name, f.read()) for f in uploaded_files]

    with st.spinner("Processing all transcripts in parallel..."):
        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = {
                executor.submit(process_file, data, client): i
                for i, (name, data) in enumerate(file_data)
            }
            results = {}
            for future in concurrent.futures.as_completed(futures):
                index = futures[future]
                results[index] = future.result()

    successes = [
        (name, results[index][0])
        for index, (name, _) in enumerate(file_data)
        if results[index][1] is None
    ]
    if successes:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for name, result in successes:
                base_name = name.rsplit(".", 1)[0]
                zip_file.writestr(f"{base_name}_telegraphic.md", result)
        st.download_button(
            label="Download all as .zip",
            data=zip_buffer.getvalue(),
            file_name="telegraphic_rewrites.zip",
            mime="application/zip",
            key="dl_all",
        )

    for index, (name, _) in enumerate(file_data):
        result, error = results[index]
        if error:
            st.error(f"{name}: {error}")
            continue

        base_name = name.rsplit(".", 1)[0]
        with st.expander(name, expanded=True):
            st.markdown(result)
            st.download_button(
                label="Download as .md",
                data=result,
                file_name=f"{base_name}_telegraphic.md",
                mime="text/markdown",
                key=f"dl_{index}",
            )
elif not api_key:
    st.info("Enter your Anthropic API key in the sidebar to begin.")
elif not uploaded_files:
    st.info("Drag and drop one or more PDF transcripts to begin.")

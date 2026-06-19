import concurrent.futures
import io
import json
import zipfile
from pathlib import Path

import streamlit as st
from pypdf import PdfReader
from anthropic import Anthropic

MODEL = "claude-sonnet-4-6"
BATCH_STATE_FILENAME = "batch_state.json"

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
- Insert a standalone `**Q&A**` heading on its own line exactly once, immediately before the first analyst/audience question, marking the boundary between prepared remarks and the Q&A session. Do not fold this heading into any speaker's name (e.g. do not write "**Jane Doe — Q&A**") - it must appear by itself as its own line, then each analyst's name/role bullet group follows below it as usual.

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


def build_batch_request(custom_id: str, transcript_text: str):
    return {
        "custom_id": custom_id,
        "params": {
            "model": MODEL,
            "max_tokens": 16000,
            "temperature": 1,
            "thinking": {"type": "enabled", "budget_tokens": 10000},
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": transcript_text}],
        },
    }


def submit_batch(client: Anthropic, file_data):
    requests = []
    extract_errors = {}
    for i, (name, data) in enumerate(file_data):
        custom_id = f"file-{i}"
        try:
            transcript_text = extract_pdf_text(data)
            if not transcript_text.strip():
                extract_errors[custom_id] = "no extractable text found"
                continue
            requests.append(build_batch_request(custom_id, transcript_text))
        except Exception as e:
            extract_errors[custom_id] = str(e)

    if not requests:
        return None, extract_errors

    batch = client.messages.batches.create(requests=requests)
    return batch.id, extract_errors


def save_batch_state(folder: str, batch_id: str, filenames, extract_errors):
    path = Path(folder) / BATCH_STATE_FILENAME
    path.write_text(json.dumps({
        "batch_id": batch_id,
        "filenames": filenames,
        "extract_errors": extract_errors,
    }))


def load_batch_state(folder: str):
    path = Path(folder) / BATCH_STATE_FILENAME
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text())
    except (json.JSONDecodeError, OSError):
        return None


def fetch_batch_results(client: Anthropic, batch_id: str):
    results = {}
    for entry in client.messages.batches.results(batch_id):
        if entry.result.type == "succeeded":
            text = "".join(
                block.text
                for block in entry.result.message.content
                if block.type == "text"
            )
            results[entry.custom_id] = (text, None)
        else:
            error_message = getattr(entry.result, "error", None)
            results[entry.custom_id] = (None, str(error_message) or entry.result.type)
    return results


def render_results(filenames, results, key_prefix):
    successes = [
        (name, results[f"file-{i}"][0])
        for i, name in enumerate(filenames)
        if f"file-{i}" in results and results[f"file-{i}"][1] is None
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
            key=f"{key_prefix}_dl_all",
        )

    for i, name in enumerate(filenames):
        custom_id = f"file-{i}"
        if custom_id not in results:
            continue
        result, error = results[custom_id]
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
                key=f"{key_prefix}_dl_{i}",
            )


st.set_page_config(page_title="Earnings Call Telegraphic Rewriter", layout="wide")
st.title("Earnings Call Transcript - Telegraphic Bullet Rewriter")

with st.sidebar:
    api_key = st.text_input("Anthropic API key", type="password")

uploaded_files = st.file_uploader(
    "Drag and drop earnings call transcript PDFs",
    type=["pdf"],
    accept_multiple_files=True,
)

use_batch = st.checkbox(
    "Use batch processing (cheaper, slower - for bulk uploads)",
    help="Submits all files as one batch job at ~50% lower API cost. "
    "Results may take a few minutes (rarely longer) - you'll need to click "
    "'Check batch status' to see progress.",
)

if use_batch:
    state_folder = st.text_input(
        "Folder to save batch tracking info",
        value=st.session_state.get("batch_state_folder", ""),
        placeholder=r"e.g. C:\Users\charles.yang\ES Updater",
        help="A small file is saved here with the batch ID so you can check on "
        "progress later, even after closing or restarting this app.",
    )
    folder_valid = bool(state_folder) and Path(state_folder).is_dir()
    if state_folder and not folder_valid:
        st.error("That folder doesn't exist.")

    # On first load (e.g. after an app restart), recover any in-progress batch
    # from disk so status can still be checked.
    if folder_valid and "batch_id" not in st.session_state:
        persisted = load_batch_state(state_folder)
        if persisted:
            st.session_state["batch_id"] = persisted["batch_id"]
            st.session_state["batch_filenames"] = persisted["filenames"]
            st.session_state["batch_extract_errors"] = persisted["extract_errors"]

    if st.button(
        "Submit batch",
        disabled=not (uploaded_files and api_key and folder_valid),
    ):
        client = Anthropic(api_key=api_key)
        file_data = [(f.name, f.read()) for f in uploaded_files]
        filenames = [name for name, _ in file_data]
        with st.spinner("Submitting batch..."):
            batch_id, extract_errors = submit_batch(client, file_data)
        st.session_state["batch_state_folder"] = state_folder
        st.session_state["batch_id"] = batch_id
        st.session_state["batch_filenames"] = filenames
        st.session_state["batch_extract_errors"] = extract_errors
        st.session_state.pop("batch_results", None)
        save_batch_state(state_folder, batch_id, filenames, extract_errors)
    elif not api_key:
        st.info("Enter your Anthropic API key in the sidebar to begin.")
    elif not uploaded_files:
        st.info("Drag and drop one or more PDF transcripts to begin.")
    elif not folder_valid:
        st.info("Enter a valid folder to save batch tracking info.")

    if st.session_state.get("batch_id"):
        batch_id = st.session_state["batch_id"]
        filenames = st.session_state["batch_filenames"]
        st.write(f"Batch ID: `{batch_id}`")

        for custom_id, message in st.session_state.get("batch_extract_errors", {}).items():
            index = int(custom_id.removeprefix("file-"))
            st.error(f"{filenames[index]}: {message}")

        if st.button("Check batch status", disabled=not api_key):
            client = Anthropic(api_key=api_key)
            batch = client.messages.batches.retrieve(batch_id)
            if batch.processing_status == "ended":
                with st.spinner("Fetching results..."):
                    st.session_state["batch_results"] = fetch_batch_results(client, batch_id)
            else:
                counts = batch.request_counts
                st.info(
                    f"Status: {batch.processing_status} - "
                    f"processing: {counts.processing}, succeeded: {counts.succeeded}, "
                    f"errored: {counts.errored}"
                )

        if st.session_state.get("batch_results"):
            render_results(filenames, st.session_state["batch_results"], key_prefix="batch")
else:
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

        render_results(
            [name for name, _ in file_data],
            {f"file-{i}": result for i, result in results.items()},
            key_prefix="instant",
        )
    elif not api_key:
        st.info("Enter your Anthropic API key in the sidebar to begin.")
    elif not uploaded_files:
        st.info("Drag and drop one or more PDF transcripts to begin.")

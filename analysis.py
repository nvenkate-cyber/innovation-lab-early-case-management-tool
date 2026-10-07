"""
Analysis logic for the case conference prep tool.

Kept separate from the Streamlit UI so the UI can be swapped later
without touching the agent code.

Requires: pip install google-genai python-docx
Set your key: export GEMINI_API_KEY="AIzaSy..."
"""

import io
from typing import Iterator

from google import genai
from google.genai import types

try:
    from docx import Document
except ImportError:
    Document = None

MODEL = "gemini-3.5-flash-lite"

SYSTEM_PROMPT = """You are an assistant helping a judge prepare for an early case \
management conference. You read case documents and produce concise, neutral, \
well-organized analysis.

Rules:
- Use only what is in the documents. Do not invent facts, dates, parties, or citations.
- Where possible, point to where information appears (page number, section, or paragraph).
- If something is unclear, missing, or inconsistent, say so plainly.
- Present issues as items the judge may wish to consider raising, not as legal conclusions.
- Stay neutral between the parties."""

DOC_SUMMARY_PROMPT = """Summarize this document for a judge preparing for an early \
case management conference. Use this structure:

**Document type and purpose** (1-2 sentences)
**Key facts and claims** (bullets, with page/section references where possible)
**Relief or action requested** (if any)
**Deadlines, dates, or procedural posture mentioned**
**Gaps or ambiguities** (anything unclear or missing)

Be concise."""

OVERVIEW_PROMPT = """Below are summaries of the documents filed in a case. \
Based on them, write a short pre-conference briefing with this structure:

**Case snapshot** (2-4 sentences)
**Potential issues to consider raising at the conference** (bullets; for each, note \
which document(s) it comes from and why it may matter)
**Inconsistencies or open questions across documents**
**Scheduling or procedural items to confirm**

Frame everything as items to consider, not conclusions.

<document_summaries>
{summaries}
</document_summaries>"""


# ---------- Building model input from uploaded files ----------

def _docx_to_text(data: bytes) -> str:
    if Document is None:
        raise RuntimeError("python-docx is required to read Word files.")
    doc = Document(io.BytesIO(data))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            parts.append(" | ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(parts)


def build_content(name: str, data: bytes) -> list:
    """Convert an uploaded file into content parts accepted by Gemini.

    PDFs and images are passed natively as raw bytes; Word and text files are
    sent as structured text blocks.
    """
    ext = name.rsplit(".", 1)[-1].lower()

    if ext == "pdf":
        return [types.Part.from_bytes(data=data, mime_type="application/pdf")]

    if ext in ("png", "jpg", "jpeg"):
        mime_type = "image/png" if ext == "png" else "image/jpeg"
        return [types.Part.from_bytes(data=data, mime_type=mime_type)]

    if ext == "docx":
        text = _docx_to_text(data)
    else:  # txt, md
        text = data.decode("utf-8", errors="replace")

    return [f'<document name="{name}">\n{text}\n</document>']


# ---------- Agent calls ----------

def get_client() -> genai.Client:
    """Reads GEMINI_API_KEY (or GOOGLE_API_KEY) from the environment."""
    return genai.Client()


def _stream(client: genai.Client, contents: list) -> Iterator[str]:
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        max_output_tokens=2000,
    )
    response = client.models.generate_content_stream(
        model=MODEL,
        contents=contents,
        config=config,
    )
    for chunk in response:
        if chunk.text:
            yield chunk.text


def stream_document_summary(client: genai.Client, name: str, data: bytes) -> Iterator[str]:
    """Stream a summary of a single document."""
    contents = build_content(name, data) + [DOC_SUMMARY_PROMPT]
    return _stream(client, contents)


def stream_case_overview(client: genai.Client, summaries: dict[str, str]) -> Iterator[str]:
    """Stream a cross-document briefing built from the per-document summaries."""
    joined = "\n\n".join(f'<summary document="{n}">\n{s}\n</summary>' for n, s in summaries.items())
    contents = [OVERVIEW_PROMPT.format(summaries=joined)]
    return _stream(client, contents)
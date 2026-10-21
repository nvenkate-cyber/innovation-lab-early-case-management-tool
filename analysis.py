"""
Analysis logic for the case conference prep tool.

Kept separate from the Streamlit UI so the UI can be swapped later
without touching the agent code.

Requires: pip install google-genai python-docx tenacity
Set your key: export GEMINI_API_KEY="AIzaSy..."
"""

import io
from typing import Iterator

from google import genai
from google.genai import types
from google.genai.errors import APIError
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

try:
    from docx import Document
except ImportError:
    Document = None

MODEL = "gemini-3.5-flash-lite"
MAX_TOKENS = 3000

SYSTEM_PROMPT = """You are an assistant helping Judge Xavier Rodriguez prepare for an early case management conference. You read case documents and produce concise, neutral, well-organized summaries.

Rules:
- Use only what is in the documents. Do not invent facts, dates, parties, or citations.
- Cite document names and specific page/paragraph numbers where possible.
- If something is unclear, missing, or inconsistent, say so plainly.
- Stay completely neutral between the parties."""

DOC_SUMMARY_PROMPT = """Summarize this document for a judge preparing for an early \
case management conference. Use this structure:

**Document type and purpose** (1-2 sentences)
**Key facts and claims** (bullets, with page/section references where possible)
**Relief or action requested** (if any)
**Deadlines, dates, or procedural posture mentioned**
**Gaps or ambiguities** (anything unclear or missing)

Be concise."""

OVERVIEW_PROMPT = """The documents for this case are provided below. Synthesize and condense all provided documents into one comprehensive pre-conference summary briefing for Judge Xavier Rodriguez.

Do NOT include recommendations, suggested questions, or judicial attention topics—provide ONLY the factual summary of the filings.

Use this structure:

**Summary**
A concise overview (3-5 sentences) summarizing the parties, nature of the suit, primary claims, defenses, and current procedural posture.

**Key Facts & Claims by Party**
Bullet points summarizing plaintiff allegations and defendant position/defenses with page/paragraph citations.

Write concisely and professionally for a federal district judge. Ensure every factual claim includes a precise document citation (e.g., Doc. 10, p. 2 ¶ 3)."""

TIMELINE_PROMPT = """Extract all significant dates, deadlines, and milestones mentioned across the provided documents and arrange them in strict chronological order.

For each entry, format the output in HTML matching this structure:

<div class="timeline-item">
    <div class="timeline-date">[Date, e.g., May 23, 2024 or By May 23, 2024]</div>
    <b>[Short Title / Description of Event or Deadline]</b><br>
    <span class="cite-tag">📄 [Exact Source, e.g., Doc. 10 - p. 2, ¶ 5]</span>
</div>

Rules:
- Order entries from earliest date to latest date.
- Include past historical events/filings AND proposed or future deadlines (e.g., discovery cutoffs, initial disclosures, motion deadlines, mediation, trial dates).
- Every item MUST include a precise citation to the source document, page, and paragraph/section where the date appears.
- Do NOT output any markdown headers, conversational prose, or standard lists—output ONLY the repeating `<div class="timeline-item">` blocks wrapped in a single parent container: `<div class="figma-card">...</div>`."""


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
        return [
            f'<document name="{name}">',
            types.Part.from_bytes(data=data, mime_type="application/pdf"),
            '</document>'
        ]

    if ext in ("png", "jpg", "jpeg"):
        mime_type = "image/png" if ext == "png" else "image/jpeg"
        return [
            f'<document name="{name}">',
            types.Part.from_bytes(data=data, mime_type=mime_type),
            '</document>'
        ]

    if ext == "docx":
        text = _docx_to_text(data)
    else:  # txt, md
        text = data.decode("utf-8", errors="replace")

    return [f'<document name="{name}">\n{text}\n</document>']


# ---------- Agent calls ----------

def get_client() -> genai.Client:
    """Reads GEMINI_API_KEY (or GOOGLE_API_KEY) from the environment."""
    return genai.Client()


@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=1, min=2, max=10),
    retry=retry_if_exception_type(APIError),
    reraise=True,
)
def _call_stream(client: genai.Client, contents: list):
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        max_output_tokens=MAX_TOKENS,
    )
    return client.models.generate_content_stream(
        model=MODEL,
        contents=contents,
        config=config,
    )


def _stream(client: genai.Client, contents: list) -> Iterator[str]:
    response = _call_stream(client, contents)
    for chunk in response:
        if chunk.text:
            yield chunk.text


def stream_document_summary(client: genai.Client, name: str, data: bytes) -> Iterator[str]:
    """Stream a summary of a single document."""
    contents = build_content(name, data) + [DOC_SUMMARY_PROMPT]
    return _stream(client, contents)


def stream_case_overview(client: genai.Client, documents: list[tuple[str, bytes]]) -> Iterator[str]:
    """Passes all uploaded documents directly into Gemini to generate one consolidated summary briefing."""
    contents = []
    
    # Pack all uploaded documents into the input content array
    for name, data in documents:
        contents.extend(build_content(name, data))
    
    # Append the overall cross-document briefing prompt
    contents.append(OVERVIEW_PROMPT)
    
    return _stream(client, contents)

def stream_case_timeline(client: genai.Client, documents: list[tuple[str, bytes]]) -> Iterator[str]:
    """Extracts chronological dates and generates timeline HTML blocks."""
    contents = []
    for name, data in documents:
        contents.extend(build_content(name, data))
    contents.append(TIMELINE_PROMPT)
    return _stream(client, contents)
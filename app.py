"""
Early Case Management Conference Prep (prototype)

Upload case documents; an agent summarizes each one and produces a
pre-conference briefing of potential issues to consider.

Run with:
    pip install streamlit python-docx google-genai
    export GEMINI_API_KEY="AIzaSy..."
    streamlit run app.py
"""

import base64
import io
import os

from google.genai.errors import APIError
import streamlit as st

import analysis

try:
    from docx import Document
except ImportError:
    Document = None

SUPPORTED_TYPES = ["pdf", "docx", "txt", "md", "png", "jpg", "jpeg"]


# ---------- Rendering helpers (UI only) ----------

def render_pdf(data: bytes) -> None:
    if hasattr(st, "pdf"):
        try:
            st.pdf(data, height=700)
            return
        except Exception:
            pass
    b64 = base64.b64encode(data).decode("utf-8")
    st.markdown(
        f'<iframe src="data:application/pdf;base64,{b64}" '
        f'width="100%" height="700" type="application/pdf"></iframe>',
        unsafe_allow_html=True,
    )


def render_docx(data: bytes, key: str) -> None:
    if Document is None:
        st.warning("Install python-docx to preview Word files.")
        return
    doc = Document(io.BytesIO(data))
    text = "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
    st.text_area("Document text", text or "(no text found)", height=400, key=f"docx_{key}")


def render_document(uploaded_file, key: str) -> None:
    data = uploaded_file.getvalue()
    ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
    if ext == "pdf":
        render_pdf(data)
    elif ext == "docx":
        render_docx(data, key)
    elif ext in ("txt", "md"):
        st.text_area("Document text", data.decode("utf-8", errors="replace"), height=400, key=f"txt_{key}")
    elif ext in ("png", "jpg", "jpeg"):
        st.image(data, use_container_width=True)


def file_key(f) -> str:
    """Stable identifier so results survive Streamlit reruns."""
    return f"{f.name}:{f.size}"


# ---------- App ----------

def main() -> None:
    st.set_page_config(page_title="Case Conference Prep", page_icon="⚖️", layout="wide")
    st.title("⚖️ Early Case Management Conference Prep")

    if not (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")):
        st.error("GEMINI_API_KEY or GOOGLE_API_KEY is not set. Set it in your environment and restart the app.")
        st.stop()

    # Persist results across reruns
    st.session_state.setdefault("summaries", {})  # file_key -> summary text
    st.session_state.setdefault("overview", {})   # frozenset of keys -> overview text

    uploaded_files = st.file_uploader(
        "Upload case documents",
        type=SUPPORTED_TYPES,
        accept_multiple_files=True,
    )
    if not uploaded_files:
        st.info("Upload one or more documents to begin.")
        return

    with st.sidebar:
        st.header("Uploaded documents")
        for f in uploaded_files:
            st.write(f"• {f.name} ({f.size / 1024:.1f} KB)")

    run = st.button("Analyze documents", type="primary")
    st.caption(
        "AI-generated analysis may contain errors. Verify against the source documents. "
        "Items listed are for consideration, not legal conclusions."
    )

    overview_box = st.container()  # filled after per-document summaries are ready
    st.divider()
    tabs = st.tabs([f.name for f in uploaded_files])

    client = analysis.get_client()
    summaries_by_name: dict[str, str] = {}

    for tab, f in zip(tabs, uploaded_files):
        key = file_key(f)
        with tab:
            summary_slot = st.container()
            with st.expander("View original document"):
                render_document(f, key)

            with summary_slot:
                if key in st.session_state["summaries"]:
                    st.markdown(st.session_state["summaries"][key])
                elif run:
                    try:
                        text = st.write_stream(
                            analysis.stream_document_summary(client, f.name, f.getvalue())
                        )
                        st.session_state["summaries"][key] = text
                    except APIError as e:
                        st.error(f"Analysis failed for {f.name}: {e}")
                    except Exception as e:
                        st.error(f"Could not process {f.name}: {e}")
                else:
                    st.info("Click **Analyze documents** to generate a summary.")

        if key in st.session_state["summaries"]:
            summaries_by_name[f.name] = st.session_state["summaries"][key]

    # Cross-document overview (built from the per-document summaries)
    with overview_box:
        if summaries_by_name:
            st.subheader("Pre-conference briefing")
            ov_key = "|".join(sorted(file_key(f) for f in uploaded_files))
            if ov_key in st.session_state["overview"]:
                st.markdown(st.session_state["overview"][ov_key])
            elif run and len(summaries_by_name) == len(uploaded_files):
                try:
                    text = st.write_stream(analysis.stream_case_overview(client, summaries_by_name))
                    st.session_state["overview"][ov_key] = text
                except APIError as e:
                    st.error(f"Overview failed: {e}")


if __name__ == "__main__":
    main()
"""
Early Case Management Conference Prep

Redesigned Streamlit interface matching Figma blueprint layout with top-level tabs:
- Summary
- Case timeline
- Discovery
- Conference questions
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


# ---------- Helper to load external CSS ----------

def load_css(file_path: str = "styles.css") -> None:
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            css_content = f.read()
        st.markdown(f"<style>{css_content}</style>", unsafe_allow_html=True)
    else:
        st.warning(f"CSS file `{file_path}` not found. Running without custom styles.")


# ---------- Document Preview Helpers ----------

def render_pdf(data: bytes) -> None:
    if hasattr(st, "pdf"):
        try:
            st.pdf(data, height=500)
            return
        except Exception:
            pass
    b64 = base64.b64encode(data).decode("utf-8")
    st.markdown(
        f'<iframe src="data:application/pdf;base64,{b64}" '
        f'width="100%" height="500" type="application/pdf"></iframe>',
        unsafe_allow_html=True,
    )


def render_docx(data: bytes, key: str) -> None:
    if Document is None:
        st.warning("Install python-docx to preview Word files.")
        return
    doc = Document(io.BytesIO(data))
    text = "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
    st.text_area("Document text", text or "(no text found)", height=350, key=f"docx_{key}")


def render_document(uploaded_file, key: str) -> None:
    data = uploaded_file.getvalue()
    ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
    if ext == "pdf":
        render_pdf(data)
    elif ext == "docx":
        render_docx(data, key)
    elif ext in ("txt", "md"):
        st.text_area("Document text", data.decode("utf-8", errors="replace"), height=350, key=f"txt_{key}")
    elif ext in ("png", "jpg", "jpeg"):
        st.image(data, use_container_width=True)


def file_key(f) -> str:
    return f"{f.name}:{f.size}"


def run_analysis(uploaded_files, client, force: bool = False) -> None:
    """Runs document summaries, briefing overview, and timeline extraction."""
    doc_tuples = [(f.name, f.getvalue()) for f in uploaded_files]
    ov_key = "|".join(sorted(file_key(f) for f in uploaded_files))

    # 1. Per-document summaries
    for f in uploaded_files:
        k = file_key(f)
        if force or k not in st.session_state["summaries"]:
            try:
                with st.spinner(f"Analyzing {f.name}..."):
                    summary_text = "".join(list(analysis.stream_document_summary(client, f.name, f.getvalue())))
                    st.session_state["summaries"][k] = summary_text
            except APIError as e:
                st.error(f"Analysis failed for {f.name}: {e}")

    # 2. Consolidated summary overview pass
    if force or ov_key not in st.session_state["overview"]:
        try:
            with st.spinner("Generating briefing overview..."):
                ov_text = "".join(list(analysis.stream_case_overview(client, doc_tuples)))
                st.session_state["overview"][ov_key] = ov_text
        except APIError as e:
            st.error(f"Overview generation failed: {e}")

    # 3. Generate Timeline
    if force or "timeline" not in st.session_state or st.session_state.get("last_ov_key") != ov_key:
        try:
            with st.spinner("Building case timeline..."):
                tl_text = "".join(list(analysis.stream_case_timeline(client, doc_tuples)))
                st.session_state["timeline"] = tl_text
        except APIError as e:
            st.error(f"Timeline generation failed: {e}")

    st.session_state["last_ov_key"] = ov_key


# ---------- App Main Layout ----------

def main() -> None:
    st.set_page_config(page_title="Case Management Prep", page_icon="⚖️", layout="wide")
    
    # Load stylesheet from styles.css
    load_css("styles.css")

    # API Key Validation
    if not (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")):
        st.error("GEMINI_API_KEY or GOOGLE_API_KEY is not set. Set it in your environment and restart the app.")
        st.stop()

    st.session_state.setdefault("summaries", {})
    st.session_state.setdefault("overview", {})

    # Header Row
    col_header, col_btn = st.columns([4, 1])
    with col_header:
        st.title("Early Case Management Conference Prep")

    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        run = st.button("✨ Analyze documents", type="primary", use_container_width=True)

    # Disclaimer Banner
    st.markdown(
        '<div class="disclaimer-banner">'
        '⚠️ <b>AI-generated analysis may contain errors.</b> Verify against the source documents. Items listed are for consideration, not legal conclusions.'
        '</div>',
        unsafe_allow_html=True,
    )

    # Main File Uploader
    uploaded_files = st.file_uploader(
        "Upload case documents",
        type=SUPPORTED_TYPES,
        accept_multiple_files=True,
        label_visibility="collapsed" if st.session_state.get("files_uploaded") else "visible",
    )

    if uploaded_files:
        st.session_state["files_uploaded"] = True

    if not uploaded_files:
        st.info("Upload case documents above to generate the briefing and preview workspace.")
        return

    client = analysis.get_client()

    # Determine batch document key to check if files changed
    ov_key = "|".join(sorted(file_key(f) for f in uploaded_files))
    
    # Auto-analyze if new files are uploaded, or force re-analysis if user clicked button
    if run or st.session_state.get("last_ov_key") != ov_key:
        run_analysis(uploaded_files, client, force=run)

    # ---------- Main Two-Column Figma Layout ----------
    left_col, right_col = st.columns([2.2, 1.1])

    # ===== LEFT COLUMN: Main Analysis Tabs =====
    with left_col:
        tab_summary, tab_timeline, tab_discovery, tab_questions = st.tabs([
            "Summary",
            "Case timeline",
            "Discovery",
            "Conference questions",
        ])

        # --- TAB 1: SUMMARY ---
        with tab_summary:
            st.markdown('<div class="card-label">SUMMARY</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-title">Case snapshot</div>', unsafe_allow_html=True)
            st.markdown(
                '<p class="card-subtitle">A condensed view of the parties, allegations, defenses, and scheduling posture based on the submitted filings.</p>',
                unsafe_allow_html=True,
            )

            # Briefing Content
            if ov_key in st.session_state["overview"]:
                st.markdown(st.session_state["overview"][ov_key])
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.info("Click **✨ Analyze documents** to generate the case snapshot briefing.")

        # --- TAB 2: CASE TIMELINE ---
        with tab_timeline:
            st.markdown('<div class="card-label">CASE TIMELINE</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-title">Case chronology</div>', unsafe_allow_html=True)
            st.markdown('<p class="card-subtitle">A focused view of key dates, filings, and upcoming deadlines.</p>', unsafe_allow_html=True)

            if "timeline" in st.session_state and st.session_state["timeline"]:
                st.markdown(st.session_state["timeline"], unsafe_allow_html=True)
            else:
                st.info("Click **✨ Analyze documents** to generate the case chronology timeline.")

        # --- TAB 3: DISCOVERY ---
        with tab_discovery:
            st.markdown('<div class="card-label">DISCOVERY</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-title">Discovery posture & preparation</div>', unsafe_allow_html=True)
            st.markdown(
                '<p class="card-subtitle">What the parties reported, what they proposed, and what to confirm before the conference.</p>',
                unsafe_allow_html=True,
            )

            st.markdown(
                """
                <div class="figma-card">
                    <b>Reported in Joint Rule 26(f) Report</b>
                    <br><br>
                    <b>ESI & Privilege:</b> No initial ESI or Rule 502 disputes reported.<br>
                    <b>Discovery Limits:</b> Parties requested no departure from standard limits.<br>
                    <span class="cite-tag">📄 Doc. 10 - pp. 2-3, ¶¶ 5-8</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

        # --- TAB 4: CONFERENCE QUESTIONS ---
        with tab_questions:
            st.markdown('<div class="card-label">CONFERENCE QUESTIONS</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-title">Possible questions to prepare for</div>', unsafe_allow_html=True)
            st.markdown(
                '<p class="card-subtitle">A condensed, source-linked agenda for counsel and judicial preparation.</p>',
                unsafe_allow_html=True,
            )

            st.markdown(
                """
                <div class="figma-card">
                    ❓ <b>Jurisdiction & Service</b><br>
                    Do jurisdiction and service remain undisputed, with no unserved parties?<br>
                    <span class="cite-tag">📄 Doc. 10 - pp. 1-2, ¶¶ 1-2</span>
                    <hr>
                    ❓ <b>Governing Dates & Mediation</b><br>
                    Was the proposed October 22, 2024 cutoff adopted or amended?<br>
                    <span class="cite-tag">📄 Doc. 10 - p. 2, ¶ 5</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ===== RIGHT COLUMN: Sidebar Document Inspector =====
    with right_col:
        st.subheader("Case documents")
        st.caption(f"{len(uploaded_files)} uploaded")

        # Select file to view
        selected_file_name = st.selectbox(
            "Select document to inspect:",
            options=[f.name for f in uploaded_files],
            label_visibility="collapsed",
        )

        selected_file = next(f for f in uploaded_files if f.name == selected_file_name)
        selected_key = file_key(selected_file)

        st.divider()

        # Document Summary View
        st.markdown(f"**Selected:** `{selected_file.name}`")

        if selected_key in st.session_state["summaries"]:
            with st.expander("📄 View AI Summary", expanded=True):
                st.markdown(st.session_state["summaries"][selected_key])
        else:
            st.info("Run analysis to view this document's summary.")

        # Document Viewer Block
        st.markdown("**Original document preview**")
        render_document(selected_file, selected_key)


if __name__ == "__main__":
    main()
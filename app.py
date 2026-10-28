"""
Early Case Management Conference Prep

Redesigned Streamlit interface matching Figma blueprint layout with top-level tabs:
- Summary
- Case timeline
- Discovery
- Conference questions
- Damages
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

SUPPORTED_TYPES = ["pdf", "docx", "txt", "md", "png", "jpg", "jpeg", "xlsx", "xls"]


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
    elif ext in ("xlsx", "xls"):
        st.caption("Spreadsheet file reference used for analytical formula/question extraction.")


def file_key(f) -> str:
    return f"{f.name}:{f.size}"


def run_analysis(uploaded_files, client, force: bool = False) -> None:
    """Runs document summaries, briefing overview, timeline extraction, discovery plan, conference questions, and damages calculations."""
    questions_excel_bytes = None
    damages_excel_bytes = None
    case_docs = []

    for f in uploaded_files:
        ext = f.name.rsplit(".", 1)[-1].lower()
        if ext in ("xlsx", "xls"):
            lower_name = f.name.lower()
            if "damages" in lower_name:
                damages_excel_bytes = f.getvalue()
            elif "common_questions" in lower_name or "questions" in lower_name:
                questions_excel_bytes = f.getvalue()
            else:
                questions_excel_bytes = f.getvalue()
        else:
            case_docs.append(f)

    doc_tuples = [(f.name, f.getvalue()) for f in case_docs]
    ov_key = "|".join(sorted(file_key(f) for f in uploaded_files))

    # 1. Per-document summaries
    for f in case_docs:
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

    # 4. Generate Discovery Analysis
    if force or "discovery" not in st.session_state or st.session_state.get("last_ov_key") != ov_key:
        try:
            with st.spinner("Analyzing proposed discovery & building tiered discovery plan..."):
                disc_text = "".join(list(analysis.stream_discovery_analysis(client, doc_tuples)))
                st.session_state["discovery"] = disc_text
        except APIError as e:
            st.error(f"Discovery analysis failed: {e}")

    # 5. Generate Conference Questions Analysis
    if force or "conference_questions" not in st.session_state or st.session_state.get("last_ov_key") != ov_key:
        try:
            with st.spinner("Evaluating early resolution questions from Common_Questions.xlsx..."):
                cq_text = "".join(list(analysis.stream_conference_questions(client, doc_tuples, excel_bytes=questions_excel_bytes)))
                st.session_state["conference_questions"] = cq_text
        except APIError as e:
            st.error(f"Conference questions analysis failed: {e}")

    # 6. Generate Damages Analysis
    if force or "damages" not in st.session_state or st.session_state.get("last_ov_key") != ov_key:
        try:
            with st.spinner("Calculating grounded potential damages using Damages_Equations.xlsx..."):
                dam_text = "".join(list(analysis.stream_damages_analysis(client, doc_tuples, damages_excel_bytes=damages_excel_bytes)))
                st.session_state["damages"] = dam_text
        except APIError as e:
            st.error(f"Damages analysis failed: {e}")

    st.session_state["last_ov_key"] = ov_key


# ---------- App Main Layout ----------

def main() -> None:
    st.set_page_config(page_title="Case Management Prep", page_icon="⚖️", layout="wide")
    load_css("styles.css")

    if not (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")):
        st.error("GEMINI_API_KEY or GOOGLE_API_KEY is not set. Set it in your environment and restart the app.")
        st.stop()

    st.session_state.setdefault("summaries", {})
    st.session_state.setdefault("overview", {})

    col_header, col_btn = st.columns([4, 1])
    with col_header:
        st.title("Early Case Management Conference Prep")

    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        run = st.button("✨ Analyze documents", type="primary", use_container_width=True)

    st.markdown(
        '<div class="disclaimer-banner">'
        '⚠️ <b>AI-generated analysis may contain errors.</b> Verify against the source documents. Items listed are for consideration, not legal conclusions.'
        '</div>',
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Upload case documents & spreadsheets (Common_Questions.xlsx, Damages_Equations.xlsx)",
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
    ov_key = "|".join(sorted(file_key(f) for f in uploaded_files))

    if run or st.session_state.get("last_ov_key") != ov_key:
        run_analysis(uploaded_files, client, force=run)

    left_col, right_col = st.columns([2.2, 1.1])

    with left_col:
        tab_summary, tab_timeline, tab_discovery, tab_questions, tab_damages = st.tabs([
            "Summary",
            "Case timeline",
            "Discovery",
            "Conference questions",
            "Damages",
        ])

        # --- TAB 1: SUMMARY ---
        with tab_summary:
            st.markdown('<div class="card-label">SUMMARY</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-title">Case snapshot</div>', unsafe_allow_html=True)
            st.markdown(
                '<p class="card-subtitle">A condensed view of the parties, allegations, defenses, and scheduling posture based on the submitted filings.</p>',
                unsafe_allow_html=True,
            )

            if ov_key in st.session_state["overview"]:
                st.markdown(st.session_state["overview"][ov_key])
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
            st.markdown('<div class="card-title">Discovery Posture & Tiered Plan</div>', unsafe_allow_html=True)
            st.markdown(
                '<p class="card-subtitle">A 3-tiered discovery plan designed to achieve early factual resolution while narrowing requested depositions and document productions.</p>',
                unsafe_allow_html=True,
            )

            if "discovery" in st.session_state and st.session_state["discovery"]:
                st.markdown(st.session_state["discovery"])
            else:
                st.info("Click **✨ Analyze documents** to generate the discovery analysis.")

        # --- TAB 4: CONFERENCE QUESTIONS ---
        with tab_questions:
            st.markdown('<div class="card-label">CONFERENCE QUESTIONS</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-title">Early Resolution Report</div>', unsafe_allow_html=True)
            st.markdown(
                '<p class="card-subtitle">Evaluated questions from Common_Questions.xlsx to identify potential early dismissal, judgment, settlement, or narrowing issues.</p>',
                unsafe_allow_html=True,
            )

            if "conference_questions" in st.session_state and st.session_state["conference_questions"]:
                st.markdown(st.session_state["conference_questions"])
            else:
                st.info("Click **✨ Analyze documents** to generate the early resolution questions analysis.")

        # --- TAB 5: DAMAGES ---
        with tab_damages:
            st.markdown('<div class="card-label">DAMAGES</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-title">Potential Damages & Remedy Breakdown</div>', unsafe_allow_html=True)
            st.markdown(
                '<p class="card-subtitle">Grounded damages calculations referencing Damages_Equations.xlsx formulas and documented case inputs.</p>',
                unsafe_allow_html=True,
            )

            if "damages" in st.session_state and st.session_state["damages"]:
                damages_text = st.session_state["damages"]

                # Clean up any escaped/inline row separators into true line breaks
                damages_text = damages_text.replace("||", "\n").replace("| |", "\n|")

                # Extract and isolate table portion to wrap inside a scrollable div
                if "|" in damages_text:
                    parts = damages_text.split("|", 1)
                    before_table = parts[0].strip()
                    table_and_after = "|" + parts[1]

                    if "\n**Notes" in table_and_after or "\n\n**" in table_and_after:
                        split_key = "\n**" if "\n**" in table_and_after else "\n\n**"
                        sub_parts = table_and_after.split(split_key, 1)
                        table_portion = sub_parts[0].strip()
                        after_table = split_key + sub_parts[1]
                    else:
                        table_portion = table_and_after.strip()
                        after_table = ""

                    if before_table:
                        st.markdown(before_table)

                    st.markdown(
                        f'<div style="overflow-x: auto; width: 100%;">{table_portion}</div>',
                        unsafe_allow_html=True,
                    )

                    if after_table:
                        st.markdown(after_table)
                else:
                    st.markdown(damages_text)
            else:
                st.info("Click **✨ Analyze documents** to calculate potential damages.")

    with right_col:
        st.subheader("Case documents")
        st.caption(f"{len(uploaded_files)} uploaded")

        selected_file_name = st.selectbox(
            "Select document to inspect:",
            options=[f.name for f in uploaded_files],
            label_visibility="collapsed",
        )

        selected_file = next(f for f in uploaded_files if f.name == selected_file_name)
        selected_key = file_key(selected_file)

        st.divider()

        st.markdown(f"**Selected:** `{selected_file.name}`")

        if selected_key in st.session_state["summaries"]:
            with st.expander("📄 View AI Summary", expanded=True):
                st.markdown(st.session_state["summaries"][selected_key])
        else:
            st.info("Run analysis to view this document's summary.")

        st.markdown("**Original document preview**")
        render_document(selected_file, selected_key)


if __name__ == "__main__":
    main()
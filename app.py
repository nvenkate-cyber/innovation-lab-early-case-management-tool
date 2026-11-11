"""
Early Case Management Conference Prep

Streamlit interface with top-level tabs:
- Overview (case snapshot, claims and defenses, scheduling posture)
- Case timeline
- Discovery
- Conference questions
- Damages
"""

import base64
import html
import io
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

from google.genai.errors import APIError
import streamlit as st

import analysis

try:
    from docx import Document
except ImportError:
    Document = None

SUPPORTED_TYPES = ["pdf", "docx", "txt", "md", "png", "jpg", "jpeg", "xlsx", "xls"]
NOT_STATED = "Not stated in provided documents"


# ---------- Helper to load external CSS ----------

def load_css(file_path: str = "styles.css") -> None:
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            css_content = f.read()
        st.markdown(f"<style>{css_content}</style>", unsafe_allow_html=True)
    else:
        st.warning(f"CSS file `{file_path}` not found. Running without custom styles.")


# ---------- Small HTML helpers (no indentation / blank lines: Markdown would treat them as code) ----------

def esc(value) -> str:
    if value is None or str(value).strip() == "":
        return ""
    return html.escape(str(value))


def cite_chip(cite) -> str:
    return f'<span class="cite-tag">📄 {esc(cite)}</span>' if esc(cite) else ""


def sub_card(label: str, title: str, meta: str) -> str:
    return (
        '<div class="sub-card">'
        f'<div class="sub-label">{esc(label)}</div>'
        f'<div class="sub-title">{esc(title) or NOT_STATED}</div>'
        f'<div class="sub-meta">{esc(meta)}</div>'
        "</div>"
    )


def point_card(title: str, text: str, cite) -> str:
    return (
        '<div class="point-card">'
        f'<div class="point-title">{esc(title)}</div>'
        f'<div class="point-text">{esc(text) or NOT_STATED}</div>'
        f"{cite_chip(cite)}"
        "</div>"
    )


def join_parts(*parts) -> str:
    return " · ".join(esc(p) for p in parts if esc(p))


# ---------- Overview helpers ----------

def get_overview(ov_key: str):
    return st.session_state.get("overview", {}).get(ov_key)


def cited_items(data) -> list:
    """Flatten the structured overview into (title, text, cite) tuples for the 'Linked to this analysis' card."""
    if not isinstance(data, dict):
        return []
    items = []
    pa = data.get("plaintiff_allegations") or {}
    if pa:
        items.append(("Plaintiff's allegations", pa.get("text", ""), pa.get("cite", "")))
    dd = data.get("defendant_defenses") or {}
    if dd:
        items.append(("Defendant's defenses", dd.get("text", ""), dd.get("cite", "")))
    for p in data.get("scheduling_posture") or []:
        if isinstance(p, dict):
            items.append((p.get("title", "Scheduling point"), p.get("text", ""), p.get("cite", "")))
    return items


def topbar_html(data) -> str:
    caption = ""
    if isinstance(data, dict):
        caption = ((data.get("case_identity") or {}).get("caption") or "").strip()
    crumbs = f"Cases / {esc(caption)} / Conference prep" if caption else "Cases / Conference prep"
    return (
        '<div class="topbar">'
        '<div class="topbar-left">'
        '<span class="logo">⚖️</span>'
        '<span class="app-name">Case management</span>'
        f'<span class="crumbs">{crumbs}</span>'
        "</div>"
        '<div class="topbar-right"><span>🔒 Private workspace</span><span class="avatar">👤</span></div>'
        "</div>"
    )


def header_subtitle(data) -> str:
    if not isinstance(data, dict):
        return ""
    ident = data.get("case_identity") or {}
    court = (ident.get("court") or "").split(",")[0].strip()
    return join_parts(ident.get("caption"), court)


def render_overview(data) -> None:
    """Overview tab: case snapshot, claims and defenses, scheduling posture."""
    if not isinstance(data, dict):
        st.markdown(data)  # legacy free-text fallback
        return

    ident = data.get("case_identity") or {}
    lead = data.get("lead_document") or {}
    lead_title = esc(lead.get("title")) or "filings"

    # Card 1: Case snapshot
    filed = lead.get("filed")
    st.markdown(
        '<div class="figma-card">'
        '<div class="card-label">SUMMARY</div>'
        '<div class="card-title">Case snapshot</div>'
        '<p class="card-subtitle">A condensed view of the parties, allegations, defenses, and scheduling posture '
        f"based on the {lead_title} and supporting case documents.</p>"
        + sub_card(
            "Case identity",
            ident.get("caption", ""),
            join_parts(ident.get("case_number"), ident.get("court")),
        )
        + sub_card(
            "Report",
            lead.get("title", ""),
            join_parts(lead.get("doc_number"), f"Filed {filed}" if esc(filed) else "", lead.get("pages")),
        )
        + sub_card("Analysis state", "AI-generated summary", "Review required · Source citations preserved")
        + "</div>",
        unsafe_allow_html=True,
    )

    if data.get("snapshot"):
        with st.expander("Narrative overview", expanded=True):
            st.markdown(str(data["snapshot"]))

    # Card 2: Allegations and defenses
    pa = data.get("plaintiff_allegations") or {}
    dd = data.get("defendant_defenses") or {}
    st.markdown(
        '<div class="figma-card">'
        '<div class="card-label">ALLEGATIONS AND DEFENSES</div>'
        '<div class="card-title">Claims and defenses</div>'
        f'<p class="card-subtitle">These points summarize the parties\' stated positions in the {lead_title} '
        "and supporting case documents. They do not represent findings by the Court.</p>"
        + point_card("Plaintiff's allegations", pa.get("text", ""), pa.get("cite", ""))
        + point_card("Defendant's defenses", dd.get("text", ""), dd.get("cite", ""))
        + "</div>",
        unsafe_allow_html=True,
    )

    # Card 3: Scheduling posture
    sched = [p for p in (data.get("scheduling_posture") or []) if isinstance(p, dict)]
    st.markdown(
        '<div class="figma-card">'
        '<div class="card-label">SCHEDULING POSTURE</div>'
        '<div class="card-title">Current posture</div>'
        f'<p class="card-subtitle">The following points summarize the procedural posture visible in the {lead_title} '
        "and supporting case documents.</p>"
        + "".join(point_card(p.get("title", ""), p.get("text", ""), p.get("cite", "")) for p in sched)
        + "</div>",
        unsafe_allow_html=True,
    )


# ---------- Timeline helpers ----------

def tl_entry_html(e: dict, open_default: bool) -> str:
    """One dated entry: date column, rail with dot, and an (optionally expandable) card."""
    kind = str(e.get("kind") or "").lower()
    status_cls = "tl-status proposed" if kind == "proposed" else "tl-status"
    details = esc(e.get("details"))
    question = esc(e.get("suggested_question"))
    expandable = bool(details or question)

    head = (
        '<div class="tl-head">'
        f'<span class="tl-title">{esc(e.get("title")) or NOT_STATED}</span>'
        f'<span class="{status_cls}">{esc(e.get("status"))}</span>'
        + ('<span class="tl-chev">⌄</span>' if expandable else "")
        + "</div>"
    )
    inner = head + f'<div class="tl-desc">{esc(e.get("description"))}</div>' + cite_chip(e.get("cite"))

    if expandable:
        expanded = (
            '<div class="tl-expanded">'
            '<div class="tl-exp-head"><span>Selected event details</span><span class="tl-exp-flag">Expanded</span></div>'
            + (f'<div class="tl-desc">{details}</div>' if details else "")
            + (
                '<div class="tl-question"><div class="tl-q-label">Suggested question · To confirm</div>'
                f'<div class="tl-q-text">{question}</div></div>'
                if question
                else ""
            )
            + "</div>"
        )
        open_attr = " open" if open_default else ""
        card = f'<details class="tl-item"{open_attr}><summary>{inner}</summary>{expanded}</details>'
    else:
        card = f'<div class="tl-item"><div class="tl-static">{inner}</div></div>'

    dot_cls = "tl-dot active" if open_default else "tl-dot"
    return (
        '<div class="tl-row">'
        '<div class="tl-date-col">'
        f'<div class="tl-date">{esc(e.get("date")) or NOT_STATED}</div>'
        f'<div class="tl-date-note">{esc(e.get("date_note"))}</div>'
        "</div>"
        f'<div class="tl-rail"><span class="{dot_cls}"></span></div>'
        f"{card}"
        "</div>"
    )


def render_timeline(data, lead_title: str = "") -> None:
    """Case timeline tab: filters, dated entries, relative events, and undated milestones."""
    if not isinstance(data, dict):
        st.markdown(data, unsafe_allow_html=True)  # legacy HTML fallback
        return

    dated = [e for e in (data.get("dated_entries") or []) if isinstance(e, dict)]
    relative = [e for e in (data.get("relative_events") or []) if isinstance(e, dict)]
    undated = data.get("undated_milestones") or {}
    undated_items = [i for i in (undated.get("items") or []) if isinstance(i, dict)]
    docs = sorted({str(e.get("doc")).strip() for e in dated + relative if str(e.get("doc") or "").strip()})
    source = esc(lead_title) or "the submitted filings"

    # Header card with filters
    with st.container(key="tl_head"):
        st.markdown(
            '<div class="card-label">CASE TIMELINE</div>'
            '<div class="card-title">Case chronology</div>'
            f'<p class="card-subtitle">A focused view of dates and sequencing visible in the {source}. '
            "Proposals are not Court-ordered deadlines.</p>",
            unsafe_allow_html=True,
        )
        f_col, d_col = st.columns([4, 1.3])
        with f_col:
            flt = st.radio(
                "Timeline filter",
                ["All events", "Fixed dates", "Proposals", "Relative events"],
                horizontal=True,
                label_visibility="collapsed",
                key="tl_filter",
            )
        with d_col:
            doc_sel = st.selectbox(
                "Document", ["All documents"] + docs, label_visibility="collapsed", key="tl_doc"
            )
        st.markdown(
            '<div class="tl-note">Historical dates do not indicate completion, lateness, or current case status.</div>',
            unsafe_allow_html=True,
        )

    def doc_ok(e: dict) -> bool:
        return doc_sel == "All documents" or str(e.get("doc") or "").strip() == doc_sel

    def kind_of(e: dict) -> str:
        return str(e.get("kind") or "").lower()

    show_dated = [e for e in dated if doc_ok(e)]
    show_rel = [e for e in relative if doc_ok(e)]
    show_undated = doc_sel == "All documents" and bool(undated_items)

    if flt == "Fixed dates":
        show_dated = [e for e in show_dated if kind_of(e) == "fixed"]
        show_rel, show_undated = [], False
    elif flt == "Proposals":
        show_dated = [e for e in show_dated if kind_of(e) == "proposed"]
        show_undated = False
    elif flt == "Relative events":
        show_dated, show_undated = [], False

    if not (show_dated or show_rel or show_undated):
        st.info("No timeline entries match this filter.")
        return

    # Dated entries
    if show_dated:
        open_idx = next(
            (
                i
                for i, e in enumerate(show_dated)
                if kind_of(e) == "proposed" and (e.get("details") or e.get("suggested_question"))
            ),
            None,
        )
        rows = "".join(tl_entry_html(e, i == open_idx) for i, e in enumerate(show_dated))
        st.markdown(
            '<div class="figma-card">'
            '<div class="tl-card-head"><span class="tl-card-title">Dated entries</span>'
            '<span class="count-pill">Earliest first</span></div>'
            f"{rows}</div>",
            unsafe_allow_html=True,
        )

    # Relative events (no calendar date)
    if show_rel:
        items = "".join(
            '<div class="tl-rel">'
            f'<div class="point-title">{esc(e.get("title"))}</div>'
            f'<div class="point-text">{esc(e.get("text")) or NOT_STATED}</div>'
            f'{cite_chip(e.get("cite"))}'
            "</div>"
            for e in show_rel
        )
        st.markdown(
            '<div class="figma-card">'
            '<div class="tl-card-head"><span class="tl-card-title">Relative event · No calendar date</span>'
            '<span class="count-pill">Proposed sequence</span></div>'
            f"{items}</div>",
            unsafe_allow_html=True,
        )

    # Milestones with no date in the filings
    if show_undated:
        tiles = "".join(
            '<div class="tl-tile"><div class="tl-tile-icon">📅</div>'
            f'<div class="point-title">{esc(i.get("title"))}</div>'
            '<div class="tl-tile-note">Date to confirm</div></div>'
            for i in undated_items
        )
        st.markdown(
            '<div class="figma-card">'
            '<div class="card-label">DATES TO CONFIRM</div>'
            f'<div class="card-title">Not dated in {source}</div>'
            '<p class="card-subtitle">Keep these milestones separate from the chronology until a governing schedule is verified.</p>'
            f'<div class="tl-tile-grid">{tiles}</div>'
            f'{cite_chip(undated.get("cite"))}'
            "</div>",
            unsafe_allow_html=True,
        )


# ---------- Discovery helpers ----------

def disc_tile(item: dict) -> str:
    proposed = str(item.get("kind") or "").lower() == "proposed"
    cls = "disc-tile proposed" if proposed else "disc-tile"
    tag_cls = "disc-tag proposed" if proposed else "disc-tag"
    tag = f'<span class="{tag_cls}">{esc(item.get("tag"))}</span>' if esc(item.get("tag")) else ""
    return (
        f'<div class="{cls}">'
        f"{tag}"
        f'<div class="disc-tile-title">{esc(item.get("title")) or NOT_STATED}</div>'
        f'<div class="disc-tile-value">{esc(item.get("value")) or NOT_STATED}</div>'
        f'<div class="disc-tile-note">{esc(item.get("note"))}</div>'
        "</div>"
    )


def render_discovery(data, tiered_plan: str = "") -> None:
    """Discovery tab: posture notice, reported facts, disclosures and cutoff, and questions to confirm."""
    if not isinstance(data, dict):
        if tiered_plan:
            st.markdown(tiered_plan)
        return

    report = esc(data.get("report_label")) or "filings"

    # Card 1: header + notice
    notice = esc(data.get("notice"))
    st.markdown(
        '<div class="figma-card">'
        '<div class="card-label">DISCOVERY</div>'
        '<div class="card-title">Discovery posture &amp; preparation</div>'
        '<p class="card-subtitle">What the parties reported, what they proposed, and what to confirm before '
        "the conference—kept distinct.</p>"
        + (f'<div class="disc-notice"><span class="disc-notice-icon">ⓘ</span><span>{notice}</span></div>' if notice else "")
        + "</div>",
        unsafe_allow_html=True,
    )

    # Card 2: reported facts
    stats = [s for s in (data.get("stats") or []) if isinstance(s, dict)]
    rows = [r for r in (data.get("reported_rows") or []) if isinstance(r, dict)]
    stat_html = "".join(
        '<div class="disc-stat">'
        f'<div class="disc-stat-label">{esc(s.get("label"))}</div>'
        f'<div class="disc-stat-value">{esc(s.get("value")) or NOT_STATED}</div>'
        f'<div class="disc-stat-note">{esc(s.get("note"))}</div>'
        "</div>"
        for s in stats
    )
    row_html = "".join(
        '<div class="disc-row">'
        f'<div class="disc-row-label">{esc(r.get("label"))}</div>'
        f'<div class="disc-row-text">{esc(r.get("text")) or NOT_STATED}</div>'
        "</div>"
        for r in rows
    )
    st.markdown(
        '<div class="figma-card">'
        f'<div class="tl-card-head"><span class="tl-card-title">Reported in the {report}</span>'
        '<span class="count-pill">Reported facts</span></div>'
        + (f'<div class="disc-stat-grid">{stat_html}</div>' if stat_html else "")
        + (f'<div class="disc-rows">{row_html}</div>' if row_html else "")
        + cite_chip(data.get("reported_cite"))
        + "</div>",
        unsafe_allow_html=True,
    )

    # Card 3: disclosures & proposed cutoff
    disclosures = [d for d in (data.get("disclosures") or []) if isinstance(d, dict)]
    if disclosures:
        tiles = "".join(disc_tile(d) for d in disclosures)
        cites = []
        for d in disclosures:
            c = str(d.get("cite") or "").strip()
            if c and c not in cites:
                cites.append(c)
        st.markdown(
            '<div class="figma-card">'
            '<div class="tl-card-head"><span class="tl-card-title">Disclosures &amp; proposed cutoff</span>'
            '<span class="count-pill">Not current status</span></div>'
            f'<div class="disc-tile-grid">{tiles}</div>'
            + "".join(cite_chip(c) for c in cites)
            + "</div>",
            unsafe_allow_html=True,
        )

    # Card 4: questions to bring
    questions = [q for q in (data.get("questions") or []) if isinstance(q, dict)]
    if questions:
        q_html = "".join(
            '<div class="disc-q">'
            '<span class="disc-q-icon">💬</span>'
            "<div>"
            f'<div class="disc-q-title">{esc(q.get("title"))}</div>'
            f'<div class="disc-q-text">{esc(q.get("text"))}</div>'
            "</div></div>"
            for q in questions
        )
        st.markdown(
            '<div class="figma-card">'
            '<div class="card-label disc-accent">SUGGESTED PREPARATION · TO CONFIRM</div>'
            '<div class="card-title">Questions to bring, not established facts</div>'
            '<p class="card-subtitle">Use these prompts to update the report snapshot. '
            "No answers or completed tasks are inferred.</p>"
            f"{q_html}"
            f'{cite_chip(data.get("questions_cite"))}'
            "</div>",
            unsafe_allow_html=True,
        )

    # Tiered plan kept available below the cards
    if tiered_plan:
        with st.expander("Tiered discovery plan (narrowest to broadest scope)"):
            st.markdown(tiered_plan)


# ---------- Conference questions helpers ----------

def cq_card(item: dict) -> str:
    questions = item.get("questions")
    if isinstance(questions, str):
        questions = [questions]
    q_html = "".join(
        f'<div class="cq-question">{esc(q)}</div>' for q in (questions or []) if esc(q)
    )
    context = esc(item.get("context"))
    return (
        '<div class="figma-card cq-card">'
        '<div class="cq-head">'
        '<span class="cq-icon">💬</span>'
        f'<span class="cq-title">{esc(item.get("title")) or NOT_STATED}</span>'
        f'<span class="cq-pill">{esc(item.get("tag")) or "To confirm"}</span>'
        "</div>"
        f"{q_html}"
        + (f'<div class="cq-context"><div>{context}</div>{cite_chip(item.get("cite"))}</div>' if context else "")
        + "</div>"
    )


def render_conference_questions(data, full_report: str = "") -> None:
    """Conference questions tab: agenda of source-linked prompts for counsel's preparation."""
    if not isinstance(data, dict):
        if full_report:
            st.markdown(full_report)
        return

    notice = esc(data.get("notice"))
    classification = esc(data.get("case_classification"))
    st.markdown(
        '<div class="figma-card">'
        '<div class="card-label">CONFERENCE QUESTIONS</div>'
        '<div class="card-title">Possible questions to prepare for</div>'
        '<p class="card-subtitle">A condensed, source-linked agenda for counsel\'s preparation—not predictions '
        "of what the judge will ask.</p>"
        + (f'<div class="disc-notice"><span class="disc-notice-icon">ⓘ</span><span>{notice}</span></div>' if notice else "")
        + (f'<p class="card-subtitle cq-class">Case classification: {classification}</p>' if classification else "")
        + "</div>",
        unsafe_allow_html=True,
    )

    items = [i for i in (data.get("items") or []) if isinstance(i, dict)]
    if items:
        st.markdown("".join(cq_card(i) for i in items), unsafe_allow_html=True)
    else:
        st.info("No questions presenting a realistic opportunity for early resolution were identified.")

    if full_report:
        with st.expander("Full early resolution analysis"):
            st.markdown(full_report)


# ---------- Damages helpers ----------

def dmg_claim_card(item: dict) -> str:
    fields = [
        ("Formula", item.get("formula")),
        ("Inputs & sources", item.get("inputs")),
        ("Result", item.get("result")),
        ("Missing info & limits", item.get("missing")),
    ]
    rows = "".join(
        '<div class="disc-row">'
        f'<div class="disc-row-label">{label}</div>'
        f'<div class="disc-row-text{" dmg-result" if label == "Result" else ""}">{esc(text)}</div>'
        "</div>"
        for label, text in fields
        if esc(text)
    )
    return (
        '<div class="figma-card cq-card">'
        '<div class="cq-head">'
        '<span class="cq-icon">🧮</span>'
        f'<span class="cq-title">{esc(item.get("title")) or NOT_STATED}</span>'
        f'<span class="cq-pill">{esc(item.get("tag"))}</span>'
        "</div>"
        f'<div class="disc-rows">{rows}</div>'
        f'{cite_chip(item.get("cite"))}'
        "</div>"
    )


def render_damages(data, full_report: str = "") -> None:
    """Damages tab: summary tiles, one card per claim/remedy, and notes."""
    if not isinstance(data, dict):
        if full_report:
            st.markdown(full_report, unsafe_allow_html=True)
        return

    notice = esc(data.get("notice"))
    case_type = esc(data.get("case_type"))
    summary = [s for s in (data.get("summary") or []) if isinstance(s, dict)]
    tiles = "".join(
        '<div class="disc-stat">'
        f'<div class="disc-stat-label">{esc(s.get("label"))}</div>'
        f'<div class="disc-stat-value">{esc(s.get("value")) or NOT_STATED}</div>'
        f'<div class="disc-stat-note">{esc(s.get("note"))}</div>'
        "</div>"
        for s in summary
    )

    # Card 1: header, notice, summary tiles
    st.markdown(
        '<div class="figma-card">'
        '<div class="card-label">DAMAGES</div>'
        '<div class="card-title">Potential damages &amp; remedies</div>'
        '<p class="card-subtitle">Grounded calculations from documented inputs. Allegations and estimates are '
        "kept distinct from established facts.</p>"
        + (f'<div class="disc-notice"><span class="disc-notice-icon">ⓘ</span><span>{notice}</span></div>' if notice else "")
        + (f'<p class="card-subtitle cq-class">Case type: {case_type}</p>' if case_type else "")
        + (f'<div class="disc-stat-grid dmg-tiles">{tiles}</div>' if tiles else "")
        + "</div>",
        unsafe_allow_html=True,
    )

    # Claim cards
    claims = [c for c in (data.get("claims") or []) if isinstance(c, dict)]
    if claims:
        st.markdown("".join(dmg_claim_card(c) for c in claims), unsafe_allow_html=True)
    else:
        st.info("The documents do not support a damages calculation.")

    # Notes card
    notes = [n for n in (data.get("notes") or []) if isinstance(n, dict)]
    if notes:
        rows = "".join(
            '<div class="disc-row">'
            f'<div class="disc-row-label">{esc(n.get("label"))}</div>'
            f'<div class="disc-row-text">{esc(n.get("text")) or NOT_STATED}</div>'
            "</div>"
            for n in notes
        )
        st.markdown(
            '<div class="figma-card">'
            '<div class="tl-card-head"><span class="tl-card-title">Notes, defenses &amp; limitations</span>'
            '<span class="count-pill">Not findings</span></div>'
            f'<div class="disc-rows">{rows}</div>'
            f'{cite_chip(data.get("notes_cite"))}'
            "</div>",
            unsafe_allow_html=True,
        )

    if full_report:
        with st.expander("Full damages table"):
            st.markdown(full_report, unsafe_allow_html=True)


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
    """Runs every analysis step in parallel worker threads.

    Worker threads only make Gemini calls (no Streamlit calls); results are stored in
    session_state from the main thread as each call finishes.
    """
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
    stale = force or st.session_state.get("last_ov_key") != ov_key

    def collect(stream) -> str:
        return "".join(stream)

    def overview_job() -> object:
        try:
            return analysis.generate_case_overview(client, doc_tuples)
        except ValueError:
            return collect(analysis.stream_case_overview(client, doc_tuples))

    def timeline_job() -> object:
        try:
            return analysis.generate_case_timeline(client, doc_tuples)
        except ValueError:
            return collect(analysis.stream_case_timeline(client, doc_tuples))

    # Each job: (label, session_state bucket, key inside bucket or None, function run in a worker thread)
    jobs = []

    for f, (name, data) in zip(case_docs, doc_tuples):
        k = file_key(f)
        if force or k not in st.session_state["summaries"]:
            jobs.append((
                f"summary of {name}", "summaries", k,
                lambda name=name, data=data: collect(analysis.stream_document_summary(client, name, data)),
            ))

    if force or ov_key not in st.session_state["overview"]:
        jobs.append(("briefing overview", "overview", ov_key, overview_job))

    if stale or "timeline" not in st.session_state:
        jobs.append((
            "case timeline", "timeline", None,
            timeline_job,
        ))

    if stale or "discovery_view" not in st.session_state:
        jobs.append((
            "discovery posture", "discovery_view", None,
            lambda: analysis.generate_case_discovery(client, doc_tuples),
        ))

    if stale or "discovery" not in st.session_state:
        jobs.append((
            "discovery plan", "discovery", None,
            lambda: collect(analysis.stream_discovery_analysis(client, doc_tuples)),
        ))

    if stale or "conference_view" not in st.session_state:
        jobs.append((
            "conference question agenda", "conference_view", None,
            lambda: analysis.generate_conference_questions(client, doc_tuples, excel_bytes=questions_excel_bytes),
        ))

    if stale or "conference_questions" not in st.session_state:
        jobs.append((
            "conference questions", "conference_questions", None,
            lambda: collect(
                analysis.stream_conference_questions(client, doc_tuples, excel_bytes=questions_excel_bytes)
            ),
        ))

    if stale or "damages_view" not in st.session_state:
        jobs.append((
            "damages overview", "damages_view", None,
            lambda: analysis.generate_damages_overview(client, doc_tuples, damages_excel_bytes=damages_excel_bytes),
        ))

    if stale or "damages" not in st.session_state:
        jobs.append((
            "damages analysis", "damages", None,
            lambda: collect(
                analysis.stream_damages_analysis(client, doc_tuples, damages_excel_bytes=damages_excel_bytes)
            ),
        ))

    if jobs:
        progress = st.progress(0.0, text=f"Analyzing documents (0 of {len(jobs)} done)...")
        with ThreadPoolExecutor(max_workers=min(4, len(jobs))) as pool:
            futures = {pool.submit(fn): (label, bucket, key) for label, bucket, key, fn in jobs}
            done = 0
            for fut in as_completed(futures):
                label, bucket, key = futures[fut]
                done += 1
                progress.progress(done / len(jobs), text=f"Analyzing documents ({done} of {len(jobs)} done)...")
                try:
                    result = fut.result()
                except APIError as e:
                    st.error(f"Failed to generate {label}: {e}")
                    continue
                except Exception as e:  # e.g. unreadable file
                    st.error(f"Failed to generate {label}: {e}")
                    continue

                if key is not None:
                    st.session_state[bucket][key] = result
                else:
                    st.session_state[bucket] = result
        progress.empty()

    st.session_state["last_ov_key"] = ov_key


# ---------- App Main Layout ----------

def apply_blue_theme() -> None:
    """Fallback for when .streamlit/config.toml isn't picked up: set the theme colors at runtime.

    Takes effect for new browser sessions, so reload the page once after the first run.
    """
    try:
        from streamlit import config

        for key, value in {
            "primaryColor": "#2563eb",
            "backgroundColor": "#f8fafc",
            "secondaryBackgroundColor": "#eff6ff",
            "textColor": "#0f172a",
        }.items():
            config.set_option(f"theme.{key}", value)
    except Exception:
        pass


def main() -> None:
    apply_blue_theme()
    st.set_page_config(page_title="Case Management Prep", page_icon="⚖️", layout="wide")
    load_css("styles.css")

    if not (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")):
        st.error("GEMINI_API_KEY or GOOGLE_API_KEY is not set. Set it in your environment and restart the app.")
        st.stop()

    st.session_state.setdefault("summaries", {})
    st.session_state.setdefault("overview", {})

    # Top app bar and page header (filled in once analysis results are known)
    topbar_ph = st.empty()
    topbar_ph.markdown(topbar_html(None), unsafe_allow_html=True)

    col_header, col_btn = st.columns([4, 1])
    with col_header:
        st.markdown('<div class="page-eyebrow">EARLY CASE MANAGEMENT</div>', unsafe_allow_html=True)
        st.markdown('<h1 class="page-title">Conference Preparation</h1>', unsafe_allow_html=True)
        subtitle_ph = st.empty()
    with col_btn:
        st.markdown("<br>", unsafe_allow_html=True)
        run = st.button("✨ Analyze documents", type="primary", use_container_width=True)

    st.markdown(
        '<div class="disclaimer-banner">'
        "<span>ⓘ AI-generated analysis may contain errors. Verify against the source documents. "
        "Items listed are for consideration, not legal conclusions.</span>"
        '<span class="disclaimer-flag">AI-generated · Review required</span>'
        "</div>",
        unsafe_allow_html=True,
    )

    left_col, right_col = st.columns([2.2, 1.1], gap="large")

    # ----- Right column, part 1: document list + uploader -----
    with right_col:
        list_box = st.container()
        uploaded_files = st.file_uploader(
            "Upload case documents & spreadsheets (Common_Questions.xlsx, Damages_Equations.xlsx)",
            type=SUPPORTED_TYPES,
            accept_multiple_files=True,
        )

    if not uploaded_files:
        with left_col:
            empty_tabs = st.tabs(["Overview", "Case timeline", "Discovery", "Conference questions", "Damages"])
            with empty_tabs[0]:
                st.markdown(
                    '<div class="figma-card"><div class="card-label">SUMMARY</div>'
                    '<div class="card-title">Case snapshot</div>'
                    '<p class="card-subtitle">Upload your case documents in the panel on the right, '
                    "then click <b>Analyze documents</b> to generate the briefing.</p></div>",
                    unsafe_allow_html=True,
                )
            for t in empty_tabs[1:]:
                with t:
                    st.info("Upload case documents and click **✨ Analyze documents** to generate this section.")
        with right_col:
            st.markdown(
                '<div class="source-footer">No documents uploaded yet. Uploads reset whenever the app reloads, '
                "so re-add your files after restarting.</div>",
                unsafe_allow_html=True,
            )
        return

    client = analysis.get_client()
    ov_key = "|".join(sorted(file_key(f) for f in uploaded_files))

    if run or st.session_state.get("last_ov_key") != ov_key:
        run_analysis(uploaded_files, client, force=run)

    overview = get_overview(ov_key)
    topbar_ph.markdown(topbar_html(overview), unsafe_allow_html=True)
    subtitle_ph.markdown(f'<div class="page-subtitle">{header_subtitle(overview)}</div>', unsafe_allow_html=True)

    # ----- Right column, part 2: document list -----
    with list_box:
        st.markdown(
            '<div class="docs-head"><span class="docs-title">Case documents</span>'
            f'<span class="count-pill">{len(uploaded_files)} uploaded</span></div>'
            '<p class="card-subtitle">Select a document to review its analysis and original source.</p>',
            unsafe_allow_html=True,
        )
        selected_file_name = st.radio(
            "Select document to inspect:",
            options=[f.name for f in uploaded_files],
            label_visibility="collapsed",
            key="selected_doc",
        )

    selected_file = next(f for f in uploaded_files if f.name == selected_file_name)
    selected_key = file_key(selected_file)
    ext = selected_file.name.rsplit(".", 1)[-1].lower()

    # ----- Left column: tabs -----
    with left_col:
        tab_summary, tab_timeline, tab_discovery, tab_questions, tab_damages = st.tabs([
            "Overview",
            "Case timeline",
            "Discovery",
            "Conference questions",
            "Damages",
        ])

        # --- TAB 1: OVERVIEW ---
        with tab_summary:
            if overview:
                render_overview(overview)
            else:
                st.info("Click **✨ Analyze documents** to generate the case snapshot briefing.")

        # --- TAB 2: CASE TIMELINE ---
        with tab_timeline:
            timeline = st.session_state.get("timeline")
            if timeline:
                lead_title = ""
                if isinstance(overview, dict):
                    lead_title = (overview.get("lead_document") or {}).get("title") or ""
                render_timeline(timeline, lead_title)
            else:
                st.info("Click **✨ Analyze documents** to generate the case chronology timeline.")

        # --- TAB 3: DISCOVERY ---
        with tab_discovery:
            disc_view = st.session_state.get("discovery_view")
            disc_plan = st.session_state.get("discovery") or ""
            if disc_view or disc_plan:
                render_discovery(disc_view, disc_plan)
            else:
                st.info("Click **✨ Analyze documents** to generate the discovery analysis.")

        # --- TAB 4: CONFERENCE QUESTIONS ---
        with tab_questions:
            cq_view = st.session_state.get("conference_view")
            cq_report = st.session_state.get("conference_questions") or ""
            if cq_view or cq_report:
                render_conference_questions(cq_view, cq_report)
            else:
                st.info("Click **✨ Analyze documents** to generate the early resolution questions analysis.")

        # --- TAB 5: DAMAGES ---
        with tab_damages:
            dmg_view = st.session_state.get("damages_view")
            dmg_report = st.session_state.get("damages") or ""
            if dmg_view or dmg_report:
                render_damages(dmg_view, dmg_report)
            else:
                st.info("Click **✨ Analyze documents** to calculate potential damages.")

    # ----- Right column, part 3: source-in-view panels -----
    with right_col:
        st.markdown(
            '<div class="source-note"><div class="source-note-title">▢ Keep the source in view</div>'
            "Citations identify the document, page, and paragraph supporting each item. "
            "Check the original before relying on the analysis.</div>",
            unsafe_allow_html=True,
        )

        st.markdown('<div class="panel-title">Original document</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="doc-name">{esc(selected_file.name)}</div>'
            f'<div class="chip-row"><span class="chip">{esc(ext.upper())}</span>'
            f'<span class="chip">{selected_file.size / 1024:.1f} KB</span></div>',
            unsafe_allow_html=True,
        )
        render_document(selected_file, selected_key)
        st.download_button(
            "Open original",
            data=selected_file.getvalue(),
            file_name=selected_file.name,
            key=f"dl_{selected_key}",
            use_container_width=True,
        )

        if selected_key in st.session_state["summaries"]:
            with st.expander("📄 View AI summary of this document"):
                st.markdown(st.session_state["summaries"][selected_key])

        items = cited_items(overview)
        if items:
            st.markdown('<div class="panel-title">🔗 Linked to this analysis</div>', unsafe_allow_html=True)
            labels = [i[0] for i in items]
            choice = st.selectbox("Linked item", labels, key="linked_item", label_visibility="collapsed")
            title, text, cite = items[labels.index(choice)]
            st.markdown(
                '<div class="linked-card">'
                f"{cite_chip(cite)}"
                f'<div class="point-title">{esc(title)}</div>'
                f'<div class="point-text">{esc(text)}</div>'
                f'<div class="linked-hint">Review {esc(cite) or "the cited passage"} in the original document to verify the cited claims.</div>'
                "</div>",
                unsafe_allow_html=True,
            )

        st.markdown(
            '<div class="source-footer">🛡 The source document remains the authority. '
            "AI summaries do not replace review of the original.</div>",
            unsafe_allow_html=True,
        )


if __name__ == "__main__":
    main()
"""
Rebuild Reflective.docx manual Table of Contents + Reflective Report body format:
  - TOC: chapters 1–6 as top-level lines only (fits ~1 page); appendices 7–13 keep full subsections
  - Body: Times New Roman 12 pt, 1.5 line spacing (course Reflective convention)
  - Right tab + dot leader to page numbers
  - Only replaces the TOC block between 'Table of Contents' and body 'Introduction'
"""
from __future__ import annotations

import re
import sys
from collections.abc import Callable
from pathlib import Path

from docx import Document
from docx.enum.text import WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

DOC = Path(r"D:\桌面文档文件勿动！\Desktop\毕业论文提交\个人\Dongdong Mao\Reflective.docx")

EM = "\u2014"  # em dash in TOC display

# Main report (ch.1–6): TOC shows chapter titles only (page = start of that chapter)
TOC_SPEC_MAIN_ONLY: list[tuple[Callable[[object], bool], str, int]] = [
    (lambda p: p.text.strip() == "Introduction" and "\t" not in p.text, "1. Introduction", 0),
    (lambda p: p.text.strip() == "Project evaluation" and "\t" not in p.text, "2. Project evaluation", 0),
    (lambda p: p.text.strip() == "Link between theory and practice" and "\t" not in p.text, "3. The link between theory and practice", 0),
    (
        lambda p: p.text.strip() == "Personal and professional development" and "\t" not in p.text,
        "4. Personal and professional development",
        0,
    ),
    (lambda p: p.text.strip() == "Summary and conclusion" and "\t" not in p.text, "5. Summary and conclusion", 0),
    (lambda p: p.text.strip() == "References" and "\t" not in p.text, "6. References", 0),
]

# Appendices: keep full subsection hierarchy in the TOC
TOC_SPEC_APPENDIX: list[tuple[Callable[[object], bool], str, int]] = [
    (
        lambda p: p.text.strip().startswith("Appendix A") and "\t" not in p.text,
        f"7. Appendix A {EM} Requirements (personal technical scope)",
        0,
    ),
    (
        lambda p: p.text.strip().startswith("A.1") and "\t" not in p.text,
        "7.1 Authentication and session scope",
        1,
    ),
    (
        lambda p: p.text.strip().startswith("A.2 Real-time") and "\t" not in p.text,
        "7.2 Real-time streaming and persistence",
        1,
    ),
    (
        lambda p: p.text.strip().startswith("A.3 Instrumentation") and "\t" not in p.text,
        "7.3 Instrumentation, exports, and admin",
        1,
    ),
    (
        lambda p: p.text.strip().startswith("A.4") and "\t" not in p.text,
        "7.4 Non-functional scope and runtime environment",
        1,
    ),
    (
        lambda p: p.text.strip().startswith("Appendix B") and "\t" not in p.text,
        f"8. Appendix B {EM} High-level design",
        0,
    ),
    (
        lambda p: p.text.strip().startswith("Appendix C") and "\t" not in p.text,
        f"9. Appendix C {EM} Data design",
        0,
    ),
    (
        lambda p: p.text.strip().startswith("Appendix D") and "\t" not in p.text,
        f"10. Appendix D {EM} Implementation details",
        0,
    ),
    (
        lambda p: p.text.strip().startswith("D.1") and "\t" not in p.text,
        "10.1 Inference and quality gate",
        1,
    ),
    (
        lambda p: p.text.strip().startswith("D.2") and "\t" not in p.text,
        "10.2 WebSocket loop and persistence batching",
        1,
    ),
    (
        lambda p: p.text.strip().startswith("D.3") and "\t" not in p.text,
        "10.3 Representative REST surface",
        1,
    ),
    (
        lambda p: p.text.strip().startswith("D.4") and "\t" not in p.text,
        "10.4 Repository code excerpts",
        1,
    ),
    (
        lambda p: p.text.strip().startswith("Appendix E") and "\t" not in p.text,
        f"11. Appendix E {EM} Testing",
        0,
    ),
    (lambda p: p.text.strip().startswith("E.1") and "\t" not in p.text, "11.1 Strategy", 1),
    (lambda p: p.text.strip().startswith("E.2") and "\t" not in p.text, "11.2 Illustrative evidence", 1),
    (
        lambda p: p.text.strip().startswith("Appendix F") and "\t" not in p.text,
        f"12. Appendix F {EM} Team meetings and process (synthesised)",
        0,
    ),
    (
        lambda p: p.text.strip().startswith("Appendix G") and "\t" not in p.text,
        f"13. Appendix G {EM} Security and operations checklist",
        0,
    ),
]

TOC_SPEC: list[tuple[Callable[[object], bool], str, int]] = [
    *TOC_SPEC_MAIN_ONLY,
    *TOC_SPEC_APPENDIX,
]

# Body: only these chapter titles stay bold; subsections + refs + figure lines unbolded.
MAIN_CHAPTER_TITLES = frozenset(
    {
        "Introduction",
        "Project evaluation",
        "Link between theory and practice",
        "Personal and professional development",
        "Summary and conclusion",
        "References",
    }
)

SUBSECTION_HEADINGS = frozenset(
    {
        "Project context and problem framing",
        "Objectives and learning-agreement alignment",
        "Ethical stance on facial analytics",
        "Structure of this report",
        "Planned versus delivered",
        "Mid-term pivot and engineering backlog",
        "Mid-term mentor pressure on UI stalls and SQLite stabilisation narrative",
        "Short-term significance for the course client",
        "Long-term significance and deployment path",
        "Limitations and honest claims",
        "Evidence chain and traceability",
        "Agile iteration and vertical slices",
        "Real-time systems and WebSocket design",
        "Machine-learning pipeline and smoothing heuristics",
        "Critical evaluation of methodology",
        "Monolith-first architecture lens",
        "Critical incidents and double-loop learning",
        "Technical skills versus learning agreement",
        "Team roles and written division of labour",
        "Communication practices under pressure",
        "Emotional labour of peer leadership",
        "Future learning needs",
    }
)

_APPENDIX_SUBSECTION = re.compile(r"^[A-G]\.\d+\s")


def _is_appendix_main_banner(text: str) -> bool:
    t = text.strip()
    if not t or "\t" in t:
        return False
    return bool(re.match(r"^Appendix [A-G]\b", t))


def _is_major_heading_text(stripped: str) -> bool:
    """True for main chapter titles and Appendix A–G banner lines."""
    t = stripped.strip()
    if not t or "\t" in t:
        return False
    if t in MAIN_CHAPTER_TITLES:
        return True
    return _is_appendix_main_banner(t)


def _force_paragraph_runs_bold(p, bold: bool) -> None:
    for run in p.runs:
        run.bold = bold


def apply_main_heading_bold_only(doc: Document) -> None:
    """After Introduction: bold only main chapter + Appendix A–G banners; unbold subsections."""
    in_body = False
    for p in doc.paragraphs:
        raw = p.text
        tstrip = raw.strip()
        if tstrip == "Table of Contents":
            continue
        if tstrip == "Introduction" and "\t" not in raw:
            in_body = True
        if not in_body:
            continue
        if not tstrip or "\t" in tstrip:
            continue
        if _is_major_heading_text(tstrip):
            _force_paragraph_runs_bold(p, True)
        elif tstrip in SUBSECTION_HEADINGS or _APPENDIX_SUBSECTION.match(tstrip):
            _force_paragraph_runs_bold(p, False)
        elif re.match(r"^\[\d+\]", tstrip):
            _force_paragraph_runs_bold(p, False)
        elif tstrip.startswith("Figure "):
            _force_paragraph_runs_bold(p, False)


# Between major headings allow up to this many blank paragraphs (~5–7 lines); inside a block keep at most 1.
MAJOR_SECTION_BLANK_PARAS_MAX = 6
INTRA_BLOCK_BLANK_PARAS_MAX = 1

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _paragraph_is_collapsible_blank(p) -> bool:
    """Only true for spacer paragraphs with no text and no drawings/shapes (image-only paras must never be removed)."""
    if p.text.strip():
        return False
    el = p._element
    if el.findall(".//%s" % qn("w:drawing")):
        return False
    if el.findall(".//%s" % qn("w:pict")):
        return False
    if el.findall(".//%s" % qn("w:object")):
        return False
    if el.findall(f".//{{{_MC_NS}}}AlternateContent"):
        return False
    return True


def collapse_excessive_vertical_gaps(doc: Document) -> None:
    """From Introduction onward: collapse long runs of *true* blank paras only (never remove figure/shape paras)."""
    intro_i = None
    for idx, p in enumerate(doc.paragraphs):
        if p.text.strip() == "Introduction" and "\t" not in p.text:
            intro_i = idx
            break
    if intro_i is None:
        return
    n = len(doc.paragraphs)
    to_remove: list = []
    i = intro_i
    while i < n:
        p = doc.paragraphs[i]
        if p.text.strip() or not _paragraph_is_collapsible_blank(p):
            i += 1
            continue
        run_start = i
        while i < n and _paragraph_is_collapsible_blank(doc.paragraphs[i]):
            i += 1
        run_end = i
        run_len = run_end - run_start
        if run_len == 0:
            continue
        next_stripped = ""
        if run_end < n:
            next_stripped = doc.paragraphs[run_end].text.strip()
        max_keep = (
            MAJOR_SECTION_BLANK_PARAS_MAX if _is_major_heading_text(next_stripped) else INTRA_BLOCK_BLANK_PARAS_MAX
        )
        if run_len > max_keep:
            for k in range(run_start + max_keep, run_end):
                to_remove.append(doc.paragraphs[k]._element)
    for el in to_remove:
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)


def find_index(doc: Document, matcher) -> int:
    for i, p in enumerate(doc.paragraphs):
        if matcher(p):
            return i
    raise RuntimeError("heading not found for TOC row")


def cumulative_chars(doc: Document, through_index: int) -> int:
    s = 0
    for j in range(min(through_index, len(doc.paragraphs) - 1) + 1):
        s += len(doc.paragraphs[j].text) + 2
    return s


def approx_page(doc: Document, para_index: int) -> int:
    intro_i = find_index(doc, lambda p: p.text.strip() == "Introduction" and "\t" not in p.text)
    appendix_a_i = find_index(
        doc,
        lambda p: p.text.strip().startswith("Appendix A") and "\t" not in p.text,
    )
    doc_end = len(doc.paragraphs) - 1
    c = cumulative_chars(doc, para_index)
    intro_c = cumulative_chars(doc, intro_i)
    aa_c = cumulative_chars(doc, appendix_a_i)
    end_c = cumulative_chars(doc, doc_end)
    if c <= intro_c:
        return max(1, min(4, 1 + int(c / 550)))
    if c <= aa_c:
        return 5 + int((c - intro_c) / max(1, aa_c - intro_c) * (18 - 5))
    return 18 + int((c - aa_c) / max(1, end_c - aa_c) * (28 - 18))


def _find_body_heading_for_anchor(doc: Document, prefix: str):
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t.startswith(prefix):
            continue
        if re.search(r"\t\d+\s*$", p.text):
            continue
        return p
    raise RuntimeError(f"body heading not found for anchor: {prefix!r}")


def ensure_appendix_toc_anchor_headings(doc: Document) -> None:
    """Some backups omit A.4 / D.4 lines; TOC rows require them. Inserts minimal titles only if missing."""
    if not any(p.text.strip().startswith("A.4") and "\t" not in p.text for p in doc.paragraphs):
        _find_body_heading_for_anchor(doc, "Appendix B").insert_paragraph_before(
            "A.4 Non-functional scope and runtime environment. "
            "Non-functional commitments include interactive responsiveness on laboratory hardware, authentication as "
            "the primary confidentiality control, and maintainable modular layout of routes and static assets. "
            "Dependencies are listed in Web1/requirements.txt (FastAPI, uvicorn, opencv-python, deepface, sqlalchemy, "
            "pandas/openpyxl for exports, passlib/python-jose for JWT). SQLite persists to emotion_analysis.db in the "
            "application working directory unless reconfigured."
        )
    if not any(p.text.strip().startswith("D.4") and "\t" not in p.text for p in doc.paragraphs):
        _find_body_heading_for_anchor(doc, "Appendix E").insert_paragraph_before(
            "D.4 Repository code excerpts (evidence). Line numbers refer to Web1/main.py at archival packaging time."
        )


def strip_toc_block(doc: Document) -> None:
    while True:
        toc_i = None
        intro_i = None
        for i, p in enumerate(doc.paragraphs):
            if p.text.strip() == "Table of Contents":
                toc_i = i
            if p.text.strip() == "Introduction" and "\t" not in p.text:
                intro_i = i
        if toc_i is None or intro_i is None or intro_i <= toc_i + 1:
            break
        p = doc.paragraphs[toc_i + 1]
        p._element.getparent().remove(p._element)


def add_toc_line_before(anchor, left_text: str, page: int, indent_level: int):
    p = anchor.insert_paragraph_before("")
    pf = p.paragraph_format
    ind = Inches(0.32) * indent_level
    pf.left_indent = ind
    pf.first_line_indent = Pt(0)
    pf.tab_stops.add_tab_stop(Inches(6.42), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    p.add_run(f"{left_text}\t{page}")
    return p


def _set_run_tnr_12(run) -> None:
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    r = run._element
    rpr = r.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for k in ("ascii", "hAnsi", "cs", "eastAsia"):
        rfonts.set(qn(f"w:{k}"), "Times New Roman")


def _format_paragraph(p) -> None:
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    for run in p.runs:
        _set_run_tnr_12(run)


def _format_cell(cell) -> None:
    for p in cell.paragraphs:
        _format_paragraph(p)
    for tbl in cell.tables:
        for row in tbl.rows:
            for c in row.cells:
                _format_cell(c)


def apply_reflective_body_style(doc: Document) -> None:
    """Whole document: TNR 12 pt, 1.5 line spacing (Reflective Report course norm)."""
    for sid in (
        "Normal",
        "Heading 1",
        "Heading 2",
        "Heading 3",
        "Heading 4",
        "List Paragraph",
        "Caption",
    ):
        if sid not in doc.styles:
            continue
        st = doc.styles[sid]
        st.font.name = "Times New Roman"
        st.font.size = Pt(12)
        st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        if sid == "Heading 1":
            st.font.bold = True
        elif sid.startswith("Heading"):
            st.font.bold = False
        else:
            st.font.bold = False

    for p in doc.paragraphs:
        _format_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _format_cell(cell)
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                _format_paragraph(p)
            for table in part.tables:
                for row in table.rows:
                    for cell in row.cells:
                        _format_cell(cell)


def rebuild_toc_only() -> None:
    """Replace only the manual TOC block; do not reformat body, bold rules, or collapse blank lines."""
    doc_fix = Document(str(DOC))
    ensure_appendix_toc_anchor_headings(doc_fix)
    doc_fix.save(str(DOC))

    doc0 = Document(str(DOC))
    resolved: list[tuple[str, int, int]] = []
    for matcher, left, lv in TOC_SPEC:
        i = find_index(doc0, matcher)
        resolved.append((left, approx_page(doc0, i), lv))

    doc = Document(str(DOC))
    strip_toc_block(doc)

    intro_p = next(
        p
        for p in doc.paragraphs
        if p.text.strip() == "Introduction" and "\t" not in p.text
    )

    anchor = intro_p
    for left, page, lv in resolved:
        add_toc_line_before(anchor, left, page, lv)

    doc.save(str(DOC))
    print("OK TOC-only:", len(resolved), "lines ->", DOC)


def main() -> None:
    doc_fix = Document(str(DOC))
    ensure_appendix_toc_anchor_headings(doc_fix)
    doc_fix.save(str(DOC))

    doc0 = Document(str(DOC))
    resolved: list[tuple[str, int, int]] = []
    for matcher, left, lv in TOC_SPEC:
        i = find_index(doc0, matcher)
        resolved.append((left, approx_page(doc0, i), lv))

    doc = Document(str(DOC))
    strip_toc_block(doc)

    intro_p = next(
        p
        for p in doc.paragraphs
        if p.text.strip() == "Introduction" and "\t" not in p.text
    )

    # Insert in document order (1 … 13): each call stacks above Introduction, so iterate forward.
    anchor = intro_p
    for left, page, lv in resolved:
        add_toc_line_before(anchor, left, page, lv)

    apply_reflective_body_style(doc)
    apply_main_heading_bold_only(doc)
    collapse_excessive_vertical_gaps(doc)
    doc.save(str(DOC))
    print(
        "OK TOC lines:",
        len(resolved),
        "+ TNR 12pt 1.5 + bold rules + vertical gap collapse ->",
        DOC,
    )


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--toc-only":
        rebuild_toc_only()
    else:
        main()

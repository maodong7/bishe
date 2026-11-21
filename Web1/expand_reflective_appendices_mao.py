# -*- coding: utf-8 -*-
"""
Append-only enrichment for Reflective.docx appendices: inserts NEW paragraphs (prose + main.py excerpts).
Does not remove or rewrite existing paragraphs (never assigns p.text on whole paragraphs except run-level typos).
Does not run full-document reformat or blank-line collapse — after save, only rebuilds TOC via
    python update_reflective_toc_mao.py --toc-only
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt

DOC = Path(r"D:\桌面文档文件勿动！\Desktop\毕业论文提交\个人\Dongdong Mao\Reflective.docx")
REPO = Path(__file__).resolve().parent
MAIN_PY = REPO / "main.py"

# New marker (backup restores will not contain this); idempotent re-runs skip body inserts.
APPEND_ONLY_MARKER = "[Appendix code evidence 2026-05-07]"
# Legacy marker from earlier run — treat as “already supplemented” so we never double-insert.
_LEGACY_MARKER = "[Appendix supplement 2026-05-06]"


def monospace_paragraph(p) -> None:
    for r in p.runs:
        r.font.name = "Consolas"
        r.font.size = Pt(9)


def find_body_heading(doc: Document, prefix: str):
    """Skip Table-of-Contents lines (tab + trailing page number)."""
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t.startswith(prefix):
            continue
        if re.search(r"\t\d+\s*$", p.text):
            continue
        return p
    raise RuntimeError(f"body heading not found: {prefix!r}")


def find_paragraph(doc: Document, prefix: str):
    for p in doc.paragraphs:
        if p.text.strip().startswith(prefix):
            return p
    raise RuntimeError(f"paragraph not found: {prefix!r}")


def _read_main_lines(lo: int, hi: int) -> str:
    """1-based inclusive line numbers from Web1/main.py."""
    text = MAIN_PY.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()
    return "\n".join(lines[lo - 1 : hi])


def _insert_stack_before(anchor, blocks: list[tuple[str, bool]]) -> None:
    """blocks: top-to-bottom reading order before anchor; each (text, is_monospace)."""
    for text, is_mono in reversed(blocks):
        p = anchor.insert_paragraph_before(text)
        if is_mono:
            monospace_paragraph(p)


def _already_supplemented(doc: Document) -> bool:
    for p in doc.paragraphs:
        t = p.text or ""
        if APPEND_ONLY_MARKER in t or _LEGACY_MARKER in t:
            return True
    return False


def ensure_d4_repository_heading(doc: Document) -> bool:
    """TOC expects D.4 before Appendix E; insert title line only if missing. Returns True if doc changed."""
    if any(p.text.strip().startswith("D.4 Repository") for p in doc.paragraphs):
        return False
    e0 = find_body_heading(doc, "Appendix E")
    e0.insert_paragraph_before(
        "D.4 Repository code excerpts (evidence). Line numbers refer to Web1/main.py at archival packaging time."
    )
    return True


def main() -> None:
    if not MAIN_PY.is_file():
        raise SystemExit(f"main.py not found: {MAIN_PY}")

    doc = Document(str(DOC))
    changed = ensure_d4_repository_heading(doc)

    if _already_supplemented(doc):
        doc.save(str(DOC))
        print("Skip: appendix code block already present ->", DOC)
        subprocess.run([sys.executable, str(REPO / "update_reflective_toc_mao.py"), "--toc-only"], cwd=str(REPO))
        return

    # --- Run-level typo only (does not replace whole paragraph) ---
    for p in doc.paragraphs:
        if p.text.strip().startswith("Pipeline:") and "DeepFace.analyse" in p.text:
            for r in p.runs:
                if "DeepFace.analyse" in (r.text or ""):
                    r.text = (r.text or "").replace("DeepFace.analyse", "DeepFace.analyze")
                    changed = True
            break

    # --- D.3: append text run only (preserves any inline objects in that paragraph) ---
    for p in doc.paragraphs:
        if p.text.strip().startswith("Non-exhaustive list implemented in main.py"):
            extra = (
                " Explicit export paths: GET /api/sessions/export/excel and GET /api/sessions/{session_id}/export/excel. "
                "Further repository routes in my integration scope include GET /api/auth/me, GET /api/sessions/{session_id}, "
                "POST /api/sessions/{session_id}/speech-recognition, GET /api/emotions, GET /api/patients/{patient_id}/data, "
                "and DELETE /api/patients/{patient_id} (see FastAPI decorators in main.py)."
            )
            if "Explicit export paths" not in p.text:
                p.add_run(extra)
                changed = True
            break

    # --- A.4 (only if absent) ---
    anchor_b = find_body_heading(doc, "Appendix B")
    if not any(p.text.strip().startswith("A.4 Non-functional") for p in doc.paragraphs):
        anchor_b.insert_paragraph_before(
            "A.4 Non-functional scope and runtime environment. "
            "Non-functional commitments include interactive responsiveness on laboratory hardware, authentication as "
            "the primary confidentiality control, and maintainable modular layout of routes and static assets. "
            "Dependencies are listed in Web1/requirements.txt (FastAPI, uvicorn, opencv-python, deepface, sqlalchemy, "
            "pandas/openpyxl for exports, passlib/python-jose for JWT). SQLite persists to emotion_analysis.db in the "
            "application working directory unless reconfigured."
        )
        changed = True

    # --- Code blocks: each stack is inserted immediately BEFORE the given heading, i.e. after all prior content
    #     (including your figures) in that appendix letter — nothing is deleted. ---

    a2 = find_body_heading(doc, "A.2 Real-time")
    intro_para = (
        f"{APPEND_ONLY_MARKER} Additive code excerpts from Web1/main.py (line numbers at packaging time). "
        "They sit before the next appendix subsection heading and do not replace earlier narrative or figures."
    )
    _insert_stack_before(
        a2,
        [
            (intro_para, False),
            ("Listing A-1. Password hashing and JWT issuance (main.py L178–189, L533–551, abridged).", False),
            (_read_main_lines(178, 189) + "\n\n" + _read_main_lines(533, 551), True),
            ("Listing A-2. WebSocket gate, batch persistence, disconnect flush (main.py L664–669, L744–756, L791–798).", False),
            (
                _read_main_lines(664, 669)
                + "\n\n"
                + _read_main_lines(744, 756)
                + "\n\n"
                + _read_main_lines(791, 798),
                True,
            ),
        ],
    )
    changed = True

    a3 = find_body_heading(doc, "A.3 Instrumentation")
    _insert_stack_before(
        a3,
        [
            (
                "Listing A-3. SQLite URL/engine and emotion row schema (main.py L25–28, L59–69). "
                "Grounds WebSocket batch inserts and export queries in Appendix C–D.",
                False,
            ),
            (_read_main_lines(25, 28) + "\n\n" + _read_main_lines(59, 69), True),
        ],
    )
    changed = True

    b0 = find_body_heading(doc, "Appendix B")
    _insert_stack_before(
        b0,
        [
            (
                "Design-to-code traceability (additive): static assets at main.py L478–479; WebSocket `/ws/{session_id}` at L664. "
                "Cross-check with your Figure B-1 in the body of Appendix B above this note.",
                False,
            ),
        ],
    )
    changed = True

    c0 = find_body_heading(doc, "Appendix C")
    _insert_stack_before(
        c0,
        [
            ("Listing C-1. User and recognition session ORM (main.py L40–57, excerpt).", False),
            (_read_main_lines(40, 57), True),
        ],
    )
    changed = True

    d0 = find_body_heading(doc, "Appendix D")
    _insert_stack_before(
        d0,
        [
            (
                "Listing D-1. `recognize_emotion` quality gate and DeepFace.analyze path (main.py L352–410, excerpt).",
                False,
            ),
            (_read_main_lines(352, 410), True),
        ],
    )
    changed = True

    e0 = find_body_heading(doc, "Appendix E")
    _insert_stack_before(
        e0,
        [
            (
                "Listing E-1. Excel export entry points (main.py L1012–1020, L1326–1336, abridged). "
                "Supports manual regression evidence in E.2 without replacing existing tables or figures.",
                False,
            ),
            (_read_main_lines(1012, 1020) + "\n\n" + _read_main_lines(1326, 1336), True),
        ],
    )
    changed = True

    e2 = find_paragraph(doc, "E.2 Illustrative evidence")
    if not any("Figures E-1 and E-2 are illustrative" in p.text for p in doc.paragraphs):
        e2.insert_paragraph_before(
            "Figures E-1 and E-2 are illustrative matrices and environment summaries. If moderators require stricter "
            "visual evidence for the archival PDF, replace them with timestamped runtime screenshots and logged "
            "command outputs; repository paths under Web1 remain the primary falsification anchor."
        )
        changed = True

    g_first = find_paragraph(doc, "Rotate SECRET_KEY")
    if not any("Phase map (opening" in p.text for p in doc.paragraphs):
        p = g_first.insert_paragraph_before(
            "Table F-2. Phase map (opening → final; indicative).\n"
            "- Opening (Oct 2025): proposal storyline; initial risk list; DeepFace feasibility read.\n"
            "- Mid-term (Nov 2025): mentor-driven backlog on UI stalls and SQLite stability.\n"
            "- Integration (Jan–Mar 2026): JWT, static pages, WebSocket transport, batch persistence tuning.\n"
            "- Pre-final (Apr 2026): exports, admin read paths, regression after CORS/JWT/WebSocket changes."
        )
        p.insert_paragraph_before(
            "This subsection synthesises coordination between opening materials, mid-term feedback, and the written "
            "division-of-labour document (组员分工部分.docx). It complements the Individual Log and repository commits "
            "as dated audit trails rather than reconstructed memory."
        )
        changed = True

    if not any("Minimum pre-demo operational checks" in p.text for p in doc.paragraphs):
        p = find_paragraph(doc, "Rotate SECRET_KEY")
        p = p.insert_paragraph_before(
            "Table G-1. Minimum pre-demo operational checks (illustrative).\n"
            "- Dependencies: pip install -r requirements.txt completes on the target OS without import errors.\n"
            "- Auth: invalid password rejected; JWT expiry behaviour matches main.py configuration.\n"
            "- WebSocket: ≥30 s continuous stream without server traceback under demo load.\n"
            "- Export: Excel opens; EmotionRecord rows populate for a seeded session."
        )
        p.insert_paragraph_before(
            "Appendix G is not a penetration-test certificate. It records actionable hardening items acknowledged "
            "during integration and aligns with coursework expectations for operations awareness alongside biometric inference."
        )
        changed = True

    doc.save(str(DOC))
    print("OK appendix append-only + save ->", DOC)

    r = subprocess.run([sys.executable, str(REPO / "update_reflective_toc_mao.py"), "--toc-only"], cwd=str(REPO))
    if r.returncode != 0:
        print("WARN: TOC-only exited", r.returncode)
    else:
        print("OK TOC-only refresh")


if __name__ == "__main__":
    main()

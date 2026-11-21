# -*- coding: utf-8 -*-
"""
In-place: keep everything BEFORE the Table of Contents heading unchanged (XML untouched
for that block). From the TOC heading onward, replace with a richer TOC + expanded body.

If your TOC title is not exactly "Table of Contents", extend find_toc_index().
"""
from __future__ import annotations

import os
import shutil
from datetime import datetime

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

# 「个人部分2」下的 BS_001 Reflective Report 终稿（与 个人部分 无关）
USER_PATH = (
    r"D:\桌面文档文件勿动！\Desktop\毕业设计\论文\个人部分2\毕业论文全套_生成"
    r"\10356_080901_24259701_BS_001_Reflective_Report.docx"
)
FIG_DIR = (
    r"D:\桌面文档文件勿动！\Desktop\毕业设计\论文\个人部分2\清单配图_generated"
)


def _style_or_normal(doc: Document, name: str) -> str:
    try:
        _ = doc.styles[name]
        return name
    except KeyError:
        return "Normal"


def find_toc_index(src: Document) -> int:
    for i, p in enumerate(src.paragraphs):
        t = (p.text or "").strip()
        st = p.style.name if p.style else ""
        if t == "Table of Contents" and "Heading" in st:
            return i
        if t in ("目录", "Table of contents") and "Heading" in st:
            return i
    for i, p in enumerate(src.paragraphs):
        tx = p.text or ""
        if "Table of Contents" in tx or tx.strip() == "目录":
            return i
    raise RuntimeError("Could not find TOC heading (Table of Contents / 目录). Edit find_toc_index().")


def remove_from_paragraph_to_end(doc: Document, start_para_index: int) -> None:
    """Remove from TOC paragraph through all following body children until w:sectPr."""
    body = doc.element.body
    start_el = doc.paragraphs[start_para_index]._element
    to_remove: list = []
    el = start_el
    while el is not None and el.tag != qn("w:sectPr"):
        to_remove.append(el)
        el = el.getnext()
    for el in to_remove:
        body.remove(el)


def add_h(doc: Document, text: str, level: int = 1) -> None:
    sn = f"Heading {level}"
    if _style_or_normal(doc, sn) != "Normal":
        doc.add_paragraph(text, style=sn)
        return
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(16 if level <= 1 else 14 if level == 2 else 12)


def add_p(doc: Document, text: str) -> None:
    doc.add_paragraph(text)


def toc_line(doc: Document, title: str, page: int) -> None:
    dots = max(2, 72 - min(60, len(title)))
    p = doc.add_paragraph(style=_style_or_normal(doc, "List Paragraph"))
    p.add_run(f"{title}{'.' * dots}{page}")


def try_fig(doc: Document, fname: str, caption: str) -> None:
    path = os.path.join(FIG_DIR, fname)
    if os.path.isfile(path):
        doc.add_paragraph()
        doc.add_picture(path, width=Inches(6.2))
        c = doc.add_paragraph(caption)
        for run in c.runs:
            run.italic = True
            run.font.size = Pt(10)


def build_toc(doc: Document) -> None:
    add_h(doc, "Table of Contents", 1)
    entries = [
        ("Introduction", 6),
        ("  Project context and problem framing", 6),
        ("  Objectives and learning-agreement alignment", 7),
        ("  Ethical stance on facial analytics", 8),
        ("  Structure of this report", 8),
        ("Project evaluation", 9),
        ("  Planned versus delivered", 9),
        ("  Mid-term pivot and engineering backlog", 10),
        ("  Short-term significance for the course client", 11),
        ("  Long-term significance and deployment path", 11),
        ("  Limitations and honest claims", 12),
        ("  Evidence chain and traceability", 12),
        ("Link between theory and practice", 13),
        ("  Agile iteration and vertical slices", 13),
        ("  Real-time systems and WebSocket design", 14),
        ("  Machine-learning pipeline and smoothing heuristics", 15),
        ("  Critical evaluation of methodology", 16),
        ("  Monolith-first architecture lens", 16),
        ("  Critical incidents and double-loop learning", 17),
        ("Personal and professional development", 18),
        ("  Technical skills versus learning agreement", 18),
        ("  Team roles and written division of labour", 19),
        ("  Communication practices under pressure", 20),
        ("  Emotional labour of peer leadership", 21),
        ("  Future learning needs", 22),
        ("Summary and conclusion", 23),
        ("References", 25),
        ("Appendix A — Requirements (personal technical scope)", 27),
        ("  A.1 Authentication and session scope", 27),
        ("  A.2 Real-time streaming and persistence", 28),
        ("  A.3 Instrumentation, exports, and admin", 28),
        ("Appendix B — High-level design", 29),
        ("Appendix C — Data design", 31),
        ("Appendix D — Implementation details", 33),
        ("  Inference and quality gate", 33),
        ("  WebSocket batching and overlays", 34),
        ("  REST surface (representative)", 35),
        ("Appendix E — Testing", 37),
        ("  Strategy and regression focus", 37),
        ("  Sample matrix and environment", 38),
        ("Appendix F — Team meetings and process", 39),
        ("Appendix G — Security and operations checklist", 41),
    ]
    for title, pg in entries:
        toc_line(doc, title, pg)
    add_p(
        doc,
        "Note: Page numbers are estimates for draft layout. In Word, switch to Print Layout, "
        "then select the table of contents block and press F9 → Update entire table if you convert "
        "this section to an automatic TOC field; otherwise renumber after final pagination.",
    )


def build_body(doc: Document) -> None:
    add_h(doc, "Introduction", 1)
    add_h(doc, "Project context and problem framing", 2)
    for t in [
        "The CN10 R&D product is a browser-assisted facial affect analytics service built around FastAPI, "
        "SQLite, and a WebSocket-first media loop. The framing throughout remained assistive: compress "
        "review time for supervised demonstrations rather than automate clinical judgement.",
        "Early desktop narratives from opening and mid-term materials were deliberately superseded once "
        "camera, HTTPS, and latency constraints became visible in integrated testing. That pivot is not "
        "scope failure; it is honest convergence to a deployable teaching product.",
    ]:
        add_p(doc, t)

    add_h(doc, "Objectives and learning-agreement alignment", 2)
    add_p(
        doc,
        "As team lead I committed to relational persistence, inference-to-UI integration, and coordination "
        "of backlog, risks, and integration windows. The Web1 repository reflects those commitments through "
        "SQLAlchemy models, JWT-protected routes, the /ws/{session_id} pathway, static operator pages, and "
        "export endpoints for audit-ready spreadsheets.",
    )
    add_p(
        doc,
        "Learning-agreement skills (database design, real-time integration, leadership) map to artefacts: "
        "emotion_analysis.db schema, batched commits inside the WebSocket loop, and the division-of-labour "
        "document that reduced overlap with deployment and paper-evidence owners.",
    )

    add_h(doc, "Ethical stance on facial analytics", 2)
    for t in [
        "Facial inference intersects consent, retention, and misuse risk. Course demonstrations stayed on "
        "authorised participants or synthetic clips; engineering mitigations included authentication, "
        "session scoping, and minimising stored imagery relative to labels and timestamps.",
        "A production deployment would require independent ethics review, DPIA-style analysis, and explicit "
        "operator training; here I document controls we actually implemented rather than aspirational claims.",
    ]:
        add_p(doc, t)

    add_h(doc, "Structure of this report", 2)
    add_p(
        doc,
        "Following programme guidance, I evaluate delivery against the proposal, connect practice to "
        "literature, analyse personal and team growth, and synthesise commitments beyond the degree. "
        "Technical evidence requested for Chinese archival review sits in appendices A–G, aligned with "
        "main.py and static assets in Web1.",
    )

    add_h(doc, "Project evaluation", 1)
    add_h(doc, "Planned versus delivered", 2)
    for t in [
        "The proposal trajectory emphasised seven basic emotions, DeepFace-assisted inference, and "
        "repeatable demonstrations. Delivered behaviour includes registration and token issuance, session "
        "lifecycle APIs, questionnaire persistence, recording metadata, screenshot capture, Excel exports, "
        "and the real-time overlay loop with disgust-aware smoothing parameters observed during trials.",
        "Mid-term mentor pressure on UI stalls and SQLite stability translated into batching (three-frame "
        "commits), JPEG quality trade-offs on uplink, and clearer operator flows on static pages.",
    ]:
        add_p(doc, t)

    add_h(doc, "Mid-term pivot and engineering backlog", 2)
    add_p(
        doc,
        "The backlog lived as a short, dated list tied to mentor comments: reduce synchronous writes in "
        "the hot path; clarify patient_id contracts between UI and API; document Haar cascade paths for "
        "Windows deployments. Closing those items was a condition for calling the local integration “done.”",
    )

    add_h(doc, "Short-term significance for the course client", 2)
    add_p(
        doc,
        "Markers can authenticate, run a session, stream frames, observe annotated previews, export an "
        "Excel workbook, and cite repository paths—evidence of engineering depth rather than a throwaway demo.",
    )

    add_h(doc, "Long-term significance and deployment path", 2)
    add_p(
        doc,
        "The monolith remains compatible with Linux+Nginx+wss handoff owned by a teammate; separation of "
        "local integration from public operations reduced single-point knowledge while preserving collaborative "
        "incident response when logs span environments.",
    )

    add_h(doc, "Limitations and honest claims", 2)
    add_p(
        doc,
        "Low light, occlusion, and crowd density remain accuracy risks; exports are pragmatic not "
        "standards-complete; CORS and secret rotation need further hardening. These limits shape how I would "
        "communicate with any stakeholder beyond the classroom.",
    )

    add_h(doc, "Evidence chain and traceability", 2)
    add_p(
        doc,
        "Claims in this report map to symbols in Web1: ORM classes and routes in main.py, static HTML/JS "
        "under static/, and generated schematic figures in appendices. That mapping is intentional so "
        "moderators can falsify statements quickly.",
    )

    add_h(doc, "Link between theory and practice", 1)
    add_h(doc, "Agile iteration and vertical slices", 2)
    add_p(
        doc,
        "Weekly integration favoured vertical slices (auth → sessions → WebSocket → persistence) over "
        "big-bang merges, consistent with agile emphasis on working software (Beck et al., 2001) and incremental "
        "risk reduction (Sommerville, 2016).",
    )

    add_h(doc, "Real-time systems and WebSocket design", 2)
    add_p(
        doc,
        "The producer–consumer relationship between browser frames and server inference illustrates "
        "stateful session resources in the sense discussed by Fielding (2000): REST alone is insufficient for "
        "latency-sensitive media; hybrid designs carry explicit trade-offs.",
    )

    add_h(doc, "Machine-learning pipeline and smoothing heuristics", 2)
    add_p(
        doc,
        "DeepFace bridges research-grade models and delivery constraints (Taigman et al., 2014). Session-scoped "
        "smoothing and disgust heuristics show how probabilistic outputs meet operational thresholds—an applied ML "
        "deployment theme (Goodfellow et al., 2016).",
    )

    add_h(doc, "Critical evaluation of methodology", 2)
    add_p(
        doc,
        "Waterfall would have failed when HTTPS camera policies surfaced; ad-hoc process would have lost "
        "traceability. Lightweight ceremonies plus explicit ownership documents matched ecosystem fit over "
        "ritual compliance (Highsmith, 2002).",
    )

    add_h(doc, "Monolith-first architecture lens", 2)
    add_p(
        doc,
        "Newman (2021) warns against premature distribution; a single FastAPI deployable with clear internal "
        "boundaries matched our risk profile while leaving decomposition as a future option.",
    )

    add_h(doc, "Critical incidents and double-loop learning", 2)
    for t in [
        "Incident A — bursty SQLite writes during streaming: surfaced theories-in-use about synchronous "
        "commits; resolved through batching and transactional discipline, consistent with Argyris and Schön "
        "(1974) on double-loop learning.",
        "Incident B — wss validation assumptions split between local and server logs: resolved through paired "
        "log correlation, reinforcing that boundaries aid workload but must not block cross-layer diagnosis.",
    ]:
        add_p(doc, t)

    add_h(doc, "Personal and professional development", 1)
    add_h(doc, "Technical skills versus learning agreement", 2)
    add_p(
        doc,
        "Outcomes strengthened SQLAlchemy modelling, JWT usage, WebSocket concurrency patterns, OpenCV+DeepFace "
        "integration, and pragmatic performance tuning. Gaps include automated async tests and formal benchmarking.",
    )

    add_h(doc, "Team roles and written division of labour", 2)
    add_p(
        doc,
        "Ye Yichen owns comparative model evidence and group thesis writing; Menghao Luo owns deployment and "
        "training artefacts; Xingli Shao contributes defect-led quality passes; I own local Web1 integration. "
        "The written roster became the single source of truth when schedules tightened.",
    )

    add_h(doc, "Communication practices under pressure", 2)
    add_p(
        doc,
        "Repro steps plus expected log signals reduced churn during heated debugging—a norm aligned with remote-"
        "friendly engineering practice.",
    )

    add_h(doc, "Emotional labour of peer leadership", 2)
    add_p(
        doc,
        "Coordinating peers without formal authority required explicit documents and occasional one-to-one "
        "check-ins to avoid defensive routines (Argyris, 1996).",
    )

    add_h(doc, "Future learning needs", 2)
    add_p(
        doc,
        "Invest in pytest with async clients, structured model comparison experiments, and deeper privacy "
        "engineering for biometric pipelines.",
    )

    add_h(doc, "Summary and conclusion", 1)
    for t in [
        "The delivered Web1 service matches assistive intent while remaining candid about environmental limits. "
        "The transferable habit is coupling decisions with explicit assumptions, evidence, and team contracts.",
        "I will carry forward artefact-first communication, ownership boundaries with collaborative incident "
        "response, and continuous ethical questioning whenever inference touches faces.",
    ]:
        add_p(doc, t)

    add_h(doc, "References", 1)
    refs = [
        "Argyris, C. (1996). Unrecognized defenses of scholars: Impact on theory and research. Organization Science, 7(1), 79–87.",
        "Argyris, C., & Schön, D. A. (1974). Theory in practice: Increasing professional effectiveness. Jossey-Bass.",
        "Beck, K., et al. (2001). Manifesto for agile software development. https://agilemanifesto.org/",
        "Fielding, R. T. (2000). Architectural styles and the design of network-based software architectures (Doctoral dissertation, UC Irvine).",
        "Fincher, S., Petre, M., & Clark, M. (2001). Computer science project work: Principles and pragmatics. Springer.",
        "Fook, J., & Cooper, L. (2003). Critical incident technique (cited in Monash University Bachelor of Social Work Fieldwork Manual).",
        "Goodfellow, I., Bengio, Y., & Courville, A. (2016). Deep learning. MIT Press.",
        "Highsmith, J. (2002). Agile software development ecosystems. Addison-Wesley.",
        "Newman, S. (2021). Building microservices (2nd ed.). O'Reilly.",
        "Schön, D. A. (1987). Educating the reflective practitioner. Jossey-Bass.",
        "Sommerville, I. (2016). Software engineering (10th ed.). Pearson.",
        "Taigman, Y., Yang, M., Ranzato, M., & Wolf, L. (2014). DeepFace. In CVPR (pp. 1701–1708).",
    ]
    for r in refs:
        doc.add_paragraph(r, style=_style_or_normal(doc, "List Paragraph"))

    doc.add_page_break()
    add_h(doc, "Appendix A — Requirements (personal technical scope)", 1)
    add_h(doc, "A.1 Authentication and session scope", 2)
    for t in [
        "Register and login with PBKDF2-hashed passwords; JWT bearer tokens; OAuth2PasswordBearer dependency on mutating routes.",
        "Session identifiers scoped to authenticated operators; patient_id treated as a demonstration pseudonym field.",
    ]:
        doc.add_paragraph(t, style=_style_or_normal(doc, "List Paragraph"))
    add_h(doc, "A.2 Real-time streaming and persistence", 2)
    for t in [
        "WebSocket /ws/{session_id} accepts JSON with base64 image payload; returns annotated JPEG with dominant emotion label.",
        "EmotionRecord rows persisted in batches (every three frames) with optional question_id / question_text linkage.",
    ]:
        doc.add_paragraph(t, style=_style_or_normal(doc, "List Paragraph"))
    add_h(doc, "A.3 Instrumentation, exports, and admin", 2)
    for t in [
        "Questionnaire CRUD under /api/questions; recording and screenshot routes; Excel export for sessions and global export.",
        "Admin-style GET routes for patients and session summaries to support review workflows.",
    ]:
        doc.add_paragraph(t, style=_style_or_normal(doc, "List Paragraph"))

    doc.add_page_break()
    add_h(doc, "Appendix B — High-level design", 1)
    add_p(doc, "Component context and layered decomposition (schematics).")
    try_fig(doc, "Fig-01_architecture.png", "Figure B-1. System context (browser, API, inference, database).")
    try_fig(doc, "Fig-02_layers.png", "Figure B-2. Layered structure.")

    doc.add_page_break()
    add_h(doc, "Appendix C — Data design", 1)
    add_p(
        doc,
        "RecognitionSession aggregates EmotionRecord, ScreenshotRecord, RecordingRecord, QuestionEmotion; "
        "Question stores reusable instrument definitions. Foreign keys enforce session integrity.",
    )
    try_fig(doc, "Fig-04_er.png", "Figure C-1. ER schematic.")
    try_fig(doc, "Fig-05_tables.png", "Figure C-2. Table summary.")

    doc.add_page_break()
    add_h(doc, "Appendix D — Implementation details", 1)
    add_h(doc, "D.1 Inference and quality gate", 2)
    for t in [
        "Pipeline: base64 decode → OpenCV Haar ROI with histogram equalisation → quality gate (Laplacian variance, "
        "brightness, contrast) → DeepFace.analyse with detector_backend skip or full-frame MTCNN fallback with "
        "face-area ratio guard.",
        "Smoothing: five-frame window per session; disgust boosting and secondary confirmation mitigate class confusion.",
    ]:
        add_p(doc, t)
    add_h(doc, "D.2 WebSocket loop and persistence batching", 2)
    add_p(
        doc,
        "Yellow bounding boxes and labels drawn on decoded frames; batched commits reduce SQLite contention; "
        "JPEG encode at quality 70 balances uplink size; disconnect handler flushes pending EmotionRecord rows.",
    )
    add_h(doc, "D.3 Representative REST surface", 2)
    add_p(
        doc,
        "Non-exhaustive list implemented in main.py: /api/auth/register, /api/auth/token, /api/sessions, "
        "/api/sessions/{id}/end, /api/sessions/{id}/screenshot, /api/sessions/{id}/emotions, export routes, "
        "/api/patients, question-emotion and recording POST routes, /api/video/process, /api/ai/analyze, "
        "and questionnaire routes as referenced in the repository.",
    )
    try_fig(doc, "Fig-03_flow.png", "Figure D-1. Real-time pipeline.")

    doc.add_page_break()
    add_h(doc, "Appendix E — Testing", 1)
    add_h(doc, "E.1 Strategy", 2)
    add_p(
        doc,
        "Manual matrices covered authentication, session lifecycle, sustained streaming, screenshot upload, and "
        "Excel integrity after WebSocket/CORS changes. Regression passes followed any change to JWT expiry or "
        "WebSocket close handling.",
    )
    add_h(doc, "E.2 Illustrative evidence", 2)
    try_fig(doc, "Fig-09_test.png", "Figure E-1. Illustrative test matrix.")
    try_fig(doc, "Fig-10_env.png", "Figure E-2. Environment summary.")

    doc.add_page_break()
    add_h(doc, "Appendix F — Team meetings and process (synthesised)", 1)
    tbl = doc.add_table(rows=6, cols=4)
    tbl.style = "Table Grid"
    hdr = ("Date (indicative)", "Participants", "Focus", "Decision")
    for j, h in enumerate(hdr):
        tbl.rows[0].cells[j].text = h
    rows = [
        ("2025-10-21", "CN10", "Scope", "Seven emotions; demo ethics"),
        ("2025-11-08", "CN10 + supervisor", "Mid-term", "Reduce stalls; stabilise DB"),
        ("2026-01-29", "CN10", "Realtime", "WebSocket + batch=3"),
        ("2026-03-12", "CN10 + ops", "HTTPS/wss", "Split app vs Nginx tasks"),
        ("2026-04-20", "CN10 + QA", "Regression", "Exports + screenshots"),
    ]
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            tbl.rows[i].cells[j].text = v

    doc.add_page_break()
    add_h(doc, "Appendix G — Security and operations checklist", 1)
    for t in [
        "Rotate SECRET_KEY; load from environment.",
        "Require HTTPS for camera deployments; validate wss upgrades behind proxies.",
        "Replace wildcard CORS with explicit origins.",
        "Rate-limit auth endpoints at proxy or middleware.",
        "Avoid printing tokens in shared logs.",
        "Define retention and cleanup for screenshots/recordings directories.",
    ]:
        doc.add_paragraph(t, style=_style_or_normal(doc, "List Paragraph"))


def main() -> None:
    backup = (
        USER_PATH.replace(
            ".docx",
            f"_backup_inplace_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
        )
    )
    shutil.copy2(USER_PATH, backup)

    doc = Document(USER_PATH)
    toc_idx = find_toc_index(doc)
    remove_from_paragraph_to_end(doc, toc_idx)

    build_toc(doc)
    doc.add_page_break()
    build_body(doc)

    doc.save(USER_PATH)
    print("OK", USER_PATH)
    print("backup", backup)


if __name__ == "__main__":
    main()

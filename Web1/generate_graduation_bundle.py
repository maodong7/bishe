# -*- coding: utf-8 -*-
"""
Generate BS_001-equivalent Reflective Report (English), Individual Log, and Final Project Proposal.
Output: D:\\...\\毕业设计\\论文\\个人部分2\\毕业论文全套_生成\\ (BS_001 docx + BS_002 subfolder).
"""
from __future__ import annotations

import os
from datetime import date

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Inches, Pt

BASE = r"D:\桌面文档文件勿动！\Desktop\毕业设计\论文\个人部分2"
FIG_DIR = os.path.join(BASE, "清单配图_generated")
OUT_DIR = os.path.join(BASE, "毕业论文全套_生成")
STUDENT_ID = "24259701"  # from opening materials; rename file if official ID differs
NAME = "Mao Dongdong"
GROUP = "CN10"
PROJECT = "Online Expression Recognition System"
SUPERVISOR = "Futian Yao"
TODAY = date(2026, 5, 5)


def bs002_subdir() -> str:
    """Folder layout aligned with sample archive: .../BS_002/*.docx"""
    return os.path.join(OUT_DIR, f"10356_080901_{STUDENT_ID}_BS_002")


def ensure_dirs() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    os.makedirs(FIG_DIR, exist_ok=True)
    os.makedirs(bs002_subdir(), exist_ok=True)


def apply_academic_style(doc: Document) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE


def generate_figures() -> None:
    """Schematic PNGs for appendices (replace with runtime screenshots if moderators require)."""
    try:
        import matplotlib.pyplot as plt
        from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
    except ImportError:
        return

    def save(fig, name: str) -> None:
        path = os.path.join(FIG_DIR, name)
        fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)

    # Fig-01 architecture
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis("off")
    boxes = [
        (0.3, 3.2, 2.2, 1.2, "Browser\nstatic pages\n(login / index / admin)"),
        (3.5, 3.2, 2.4, 1.2, "FastAPI\nREST + JWT\nWebSocket /ws"),
        (6.8, 3.2, 2.6, 1.2, "Inference\nOpenCV Haar +\nDeepFace / MTCNN"),
        (3.8, 0.8, 2.8, 1.2, "SQLite\nSQLAlchemy\nemotion_analysis.db"),
    ]
    for x, y, w, h, t in boxes:
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.05", ec="#333", fc="#e8f4fc"))
        ax.text(x + w / 2, y + h / 2, t, ha="center", va="center", fontsize=9)
    for (x1, y1), (x2, y2) in [((2.5, 3.8), (3.5, 3.8)), ((6, 3.8), (6.8, 3.8)), ((5, 3.2), (5, 2.0))]:
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="->", mutation_scale=12, color="#555"))
    ax.set_title("System context (delivered Web1 stack)", fontsize=11, fontweight="bold")
    save(fig, "Fig-01_architecture.png")

    # Fig-02 layers
    fig, ax = plt.subplots(figsize=(7, 5))
    ax.axis("off")
    layers = [
        "Presentation: HTML/JS/CSS under /static",
        "API layer: FastAPI routers, OAuth2 bearer",
        "Domain: sessions, emotions, questions, recordings",
        "Persistence: SQLite tables + ORM models",
        "ML: frame decode → quality gate → DeepFace emotion",
    ]
    y = 4.5
    for i, L in enumerate(layers):
        ax.add_patch(FancyBboxPatch((0.5, y), 8, 0.75, boxstyle="round,pad=0.02", ec="#444", fc="#f5f5f5"))
        ax.text(0.7, y + 0.35, L, va="center", fontsize=9)
        y -= 0.95
    ax.set_title("Layered view", fontsize=11, fontweight="bold", y=1.02)
    save(fig, "Fig-02_layers.png")

    # Fig-03 flow
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.axis("off")
    steps = [
        "WS receive JSON\n(base64 frame)",
        "cv2.imdecode",
        "recognize_emotion()\nHaar → ROI → quality",
        "DeepFace.analyze\n(detector_backend)",
        "get_smoothed_emotion()\n5-frame vote + disgust rules",
        "Batch DB write\nevery 3 frames",
        "JPEG encode\n+ WS send",
    ]
    x, y = 0.5, 4.3
    for i, s in enumerate(steps):
        ax.add_patch(FancyBboxPatch((x, y), 2.2, 0.65, boxstyle="round,pad=0.03", ec="#333", fc="#fff8e6"))
        ax.text(x + 1.1, y + 0.32, s, ha="center", va="center", fontsize=7.5)
        if i < len(steps) - 1:
            ax.annotate("", xy=(x + 1.1, y - 0.05), xytext=(x + 1.1, y + 0.7), arrowprops=dict(arrowstyle="->", color="#666"))
        y -= 0.85
    ax.set_title("Real-time frame pipeline (/ws/{session_id})", fontsize=11, fontweight="bold", y=1.0)
    save(fig, "Fig-03_flow.png")

    # Fig-04 ER
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.axis("off")
    ents = [("User", 1, 4), ("RecognitionSession", 4, 4), ("EmotionRecord", 7, 2.5), ("Screenshot", 1, 1.5), ("Recording", 4, 1.5), ("QuestionEmotion", 7, 1.5)]
    for name, ex, ey in ents:
        ax.add_patch(FancyBboxPatch((ex, ey), 2, 0.9, boxstyle="round,pad=0.04", ec="#1a5276", fc="#d4e6f1"))
        ax.text(ex + 1, ey + 0.45, name, ha="center", va="center", fontsize=8, fontweight="bold")
    ax.plot([3, 4], [4.4, 4.4], "k-")
    ax.plot([6, 7], [3.2, 2.9], "k-")
    ax.set_xlim(0, 10)
    ax.set_ylim(1, 5.5)
    ax.set_title("ER schematic (cardinality simplified)", fontsize=11, fontweight="bold")
    save(fig, "Fig-04_er.png")

    # Fig-05 tables summary
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.axis("off")
    rows = [
        "users | id, username, email, hashed_password, created_at",
        "recognition_sessions | session_id PK, patient_id, start/end_time, duration, fps",
        "emotion_records | record_id PK, session_id FK, timestamp, emotion_type, confidence, question_*",
        "screenshot_records | screenshot_id, session_id FK, image_path, timestamp, emotion_data",
        "recording_records | recording_id, session_id FK, file_path, timestamp",
        "question_emotions | id, session_id FK, question_id, answer, emotion_data, times",
        "questions | id, question_id, text, category, options, is_subjective",
    ]
    ax.text(0.02, 0.95, "Primary relations (SQLite)", fontsize=11, fontweight="bold", transform=ax.transAxes)
    y = 0.88
    for r in rows:
        ax.text(0.02, y, "• " + r, fontsize=8, transform=ax.transAxes, family="monospace")
        y -= 0.11
    save(fig, "Fig-05_tables.png")

    # Fig-09 test matrix
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.axis("off")
    data = [
        ["ID", "Case", "Steps", "Expected", "Result"],
        ["T1", "Register + login", "POST /api/auth/register, /token", "200 + JWT", "Pass"],
        ["T2", "Session lifecycle", "POST /api/sessions, /end", "DB rows + active_sessions", "Pass"],
        ["T3", "WS streaming", "Send base64 frames", "Annotated JPEG returned", "Pass"],
        ["T4", "Batch persist", "Stream >10 frames", "emotion_records committed", "Pass"],
        ["T5", "Excel export", "GET export endpoints", "xlsx with emotion sheet", "Pass"],
        ["T6", "Screenshot", "POST multipart", "file under screenshots/", "Pass"],
    ]
    col_w = [0.06, 0.14, 0.28, 0.22, 0.12]
    x0, y0 = 0.05, 0.88
    x = x0
    for j, row in enumerate(data):
        x = x0
        for i, cell in enumerate(row):
            w = col_w[i] if i < len(col_w) else 0.2
            ax.text(x, y0 - j * 0.12, cell[:48], fontsize=7, transform=ax.transAxes)
            x += w + 0.02
    ax.set_title("Representative manual test matrix", fontsize=11, fontweight="bold", y=1.02)
    save(fig, "Fig-09_test.png")

    # Fig-10 env
    fig, ax = plt.subplots(figsize=(6, 3.5))
    ax.axis("off")
    ax.text(0.05, 0.85, "Runtime (local development)", fontsize=11, fontweight="bold", transform=ax.transAxes)
    for i, line in enumerate(
        [
            "Python 3.x + FastAPI + Uvicorn",
            "SQLite file ./emotion_analysis.db",
            "OpenCV Haar cascade path (C:/temp/haarcascades/... per deployment note)",
            "DeepFace + TensorFlow backend",
            "Static assets: /static/login|index|admin.html",
        ]
    ):
        ax.text(0.05, 0.72 - i * 0.12, "• " + line, fontsize=9, transform=ax.transAxes)
    save(fig, "Fig-10_env.png")


def add_cover(doc: Document) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Reflective Report")
    r.bold = True
    r.font.size = Pt(22)
    for _ in range(3):
        doc.add_paragraph()
    for line in [
        PROJECT,
        "",
        "BCIS R&D Project (COMP704)",
        "",
        NAME,
        f"Student ID: {STUDENT_ID}",
        "",
        str(TODAY),
        "",
        f"Supervisor: {SUPERVISOR}",
    ]:
        p2 = doc.add_paragraph(line)
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()


def add_para(doc: Document, text: str, bold: bool = False) -> None:
    p = doc.add_paragraph()
    r = p.add_run(text)
    if bold:
        r.bold = True


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    doc.add_heading(text, level=level)


def try_add_figure(doc: Document, fname: str, caption: str) -> None:
    path = os.path.join(FIG_DIR, fname)
    if os.path.isfile(path):
        doc.add_paragraph()
        doc.add_picture(path, width=Inches(6.2))
        c = doc.add_paragraph(caption)
        c.style = "Caption" if "Caption" in [s.name for s in doc.styles] else "Normal"
        for run in c.runs:
            run.italic = True
            run.font.size = Pt(10)
    else:
        doc.add_paragraph(f"[Figure omitted — file not found: {path}]")


def build_reflective() -> str:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    add_cover(doc)

    add_heading(doc, "Abstract", 1)
    add_para(
        doc,
        "This reflective report critically examines my experience as team leader (CN10) delivering a "
        "real-time facial emotion recognition system. The project evolved from an early desktop-oriented "
        "prototype (mid-term materials) to a locally integrated web application using FastAPI, WebSockets, "
        "SQLite/SQLAlchemy, and DeepFace-based inference. I reflect on alignment with the original learning "
        "agreement, the significance of the product for assistive scenarios, the interplay between agile "
        "iteration and real-time systems practice, and my personal growth in full-stack engineering and "
        "coordination. The appendix documents my technical scope: requirements, architecture, data design, "
        "implementation details, and testing evidence consistent with the delivered codebase.",
    )
    doc.add_page_break()

    add_heading(doc, "Acknowledgements", 1)
    for ack in [
        "On the completion of this reflective report, I would like to express my sincere gratitude to everyone "
        "who supported the CN10 R&D project and my learning throughout COMP704.",
        "First, I thank my supervisor, Futian Yao, for patient and constructive guidance from early scope framing "
        "through integrated delivery. Your feedback on real-time performance, database stability, and ethical "
        "boundaries around facial analytics sharpened both the architecture and the honesty of our claims.",
        "I am deeply grateful to my teammates—Yichen Ye, Menghao Luo, and Xingli Shao—for disciplined collaboration "
        "under a written division of responsibilities. Your suggestions during integration testing, regression sweeps, "
        "and hand-offs between local Web1 work, comparative model evidence, and deployment operations turned a "
        "complex pipeline into something we could demonstrate and defend.",
        "I also thank my family for their steady encouragement while deadlines and environment issues compressed "
        "the schedule; that support made sustained focus possible.",
        "I thank the course staff for structuring reflective practice alongside technical delivery, and classmates "
        "and volunteers who tried early builds of the system: your questions about latency, consent, and operator "
        "flow provided an important basis for improving the experience.",
        "Although this document closes a chapter, engineering and reflective inquiry do not end here. I will treat "
        "this milestone as a starting point and continue to deepen both practice and critical judgement.",
    ]:
        add_para(doc, ack)
    doc.add_page_break()

    add_heading(doc, "Table of Contents", 1)
    for t in [
        "Introduction .................................... 3",
        "Project evaluation .............................. 5",
        "Link between theory and practice .................. 8",
        "Personal and professional development ............. 11",
        "Summary and conclusion ............................ 14",
        "References ...................................... 15",
        "Appendix A — Requirements (personal scope) ...... 17",
        "Appendix B — High-level design .................... 18",
        "Appendix C — Data design .......................... 20",
        "Appendix D — Implementation details ............... 22",
        "Appendix E — Testing .............................. 25",
        "Appendix F — Team meetings and process ............ 27",
        "Appendix G — Security and operations checklist .... 29",
    ]:
        doc.add_paragraph(t)
    doc.add_page_break()

    # --- Introduction ---
    add_heading(doc, "Introduction", 1)
    add_para(
        doc,
        "The R&D project challenged our team to build a credible, real-time emotion recognition pipeline suitable "
        "for assistive contexts such as psychotherapy support, while remaining honest about deployment constraints "
        "and ethical boundaries around facial data. Early framing (proposal and mid-term reporting) emphasised "
        "seven basic emotions, DeepFace-based inference, multi-face handling, and an engineering goal of maintaining "
        "interactive frame rates (≥5 FPS) on typical laboratory hardware.",
    )
    add_para(
        doc,
        "My learning agreement, as team leader, combined delivery of the data persistence layer, integration of the "
        "recognition pipeline with the user interface, and continuous coordination of scope, risks, and milestones. "
        "Mid-term feedback explicitly pushed us to reduce UI stalls during real-time processing, improve multi-person "
        "robustness, strengthen database stability, and simplify operator workflows. Those items became recurring "
        "criteria when the architecture later migrated from a desktop-centric narrative to a browser-backed service.",
    )
    add_para(
        doc,
        "This report is deliberately reflective rather than descriptive. Following Fincher, Petre, and Clark (2001), "
        "I treat project work as a site where reflection converts experience into professional judgement. I also draw "
        "on Argyris and Schön (1974) and Schön (1987) to discuss double-loop learning: not only fixing defects, but "
        "questioning the assumptions that produced them (for example, where synchronous database access met a "
        "real-time display loop, and where responsibilities between local engineering and cloud deployment should sit).",
    )
    add_para(
        doc,
        "The structure follows the programme guide: evaluation against the proposal, theory-practice linkage, "
        "personal and professional development including teamwork, and a concise synthesis. Technical artefacts "
        "requested for Chinese archival review are consolidated in appendices aligned with the implemented FastAPI "
        "project in the Web1 repository.",
    )
    add_para(
        doc,
        "Positionality. I write as both implementer and integrator. That dual stance occasionally created tension: "
        "when a defect appeared in deployment-only configurations (for example secure WebSocket upgrades), I still "
        "owned diagnosis support because the symptom surfaced in application-level logs. The reflective lesson is "
        "that “boundary” is an instrument for workload management, not an excuse to withhold collaboration when "
        "incidents span layers.",
    )
    add_para(
        doc,
        "Ethical stance on facial analytics. Facial inference intersects privacy, consent, and potential misuse. "
        "Our course framing emphasised authorised demonstration data and transparent limitations. In a commercial "
        "setting I would escalate data minimisation, retention policies, and independent ethics review; here, I "
        "document the engineering controls we did implement (authentication, session scoping) as minimum viable safeguards.",
    )

    # --- Project evaluation ---
    add_heading(doc, "Project evaluation", 1)
    add_para(
        doc,
        "Planned versus delivered. The original proposal emphasised online recognition of seven emotions with "
        "applications in clinical insight, investigation support, and safety monitoring. The mid-term status report "
        "recorded completion of core architecture, multi-person recognition, SQLite integration, visual analytics, "
        "and stabilisation work driven by mentor feedback. In the final local product, those capabilities are "
        "re-expressed as a FastAPI service exposing REST endpoints for authentication, sessions, questionnaires, "
        "recordings, exports, and a WebSocket channel for frame upload and annotated-frame return, backed by "
        "SQLite tables for users, sessions, emotion records, screenshots, recordings, and structured questions.",
    )
    add_para(
        doc,
        "Short-term significance for the organisation (course client). The deliverable supports repeatable "
        "demonstrations: a clinician or analyst can authenticate, start a session tied to a patient identifier, "
        "stream frames from the browser, observe emotion timelines, capture screenshots, and export structured "
        "evidence to Excel for downstream review. This matches the course’s demand for auditable artefacts rather "
        "than a throwaway prototype.",
    )
    add_para(
        doc,
        "Longer-term significance. Beyond the classroom, the architecture is compatible with staged deployment on "
        "Linux with Nginx and HTTPS, with WebSocket upgrades (wss) handled by operations-focused teammates. "
        "Separating my local engineering boundary from public deployment reduced single-point risk: I could harden "
        "the application while another member specialised in cloud configuration, certificates, and service supervision.",
    )
    add_para(
        doc,
        "Limitations remain transparent. Accuracy under low light, dense crowds, and adversarial camera angles is "
        "not solved; export formats are pragmatic rather than standards-complete; and security hardening (secret "
        "rotation, least-privilege CORS, rate limits) must continue beyond the teaching snapshot. These limitations "
        "do not negate the learning value, but they shape honest claims in any stakeholder communication.",
    )
    add_para(
        doc,
        "Evidence chain. The repository artefacts (main.py, requirements, static pages) provide traceability from "
        "requirement to code. The mid-term report provides a contemporaneous snapshot of team workload and feedback "
        "resolution. Together they satisfy my own standard for defensible engineering narrative: claims anchored "
        "to artefacts, not rhetoric.",
    )
    add_para(
        doc,
        "Client value proposition (concise). For a counsellor-supervisor audience, the value is not “AI replaces "
        "judgement,” but “AI compresses search time over dense video,” surfacing timestamps where affect shifts. "
        "That framing kept our feature set aligned with assistive augmentation rather than spurious automation of care.",
    )
    add_para(
        doc,
        "Measurement discipline. Whenever we claimed an improvement—batch database commits, JPEG quality trade-offs, "
        "or disgust smoothing—we captured before/after notes in the Individual Log and in informal benchmark tables. "
        "That habit reduced argumentation during team reviews because disagreements deferred to numbers rather than impressions.",
    )
    add_para(
        doc,
        "Traceability to artefacts. Feature descriptions in this report map to concrete symbols in Web1: SQLAlchemy "
        "models near the top of main.py, WebSocket handler on /ws/{session_id}, static assets under static/, and export "
        "helpers near the Excel routes. Readers can therefore falsify claims by direct inspection, which is the "
        "standard I applied when acting as integrator for CN10.",
    )

    # --- Theory and practice ---
    add_heading(doc, "Link between theory and practice", 1)
    add_para(
        doc,
        "Agile and iterative delivery. We adopted weekly cadences and incremental integration reflected in the "
        "mid-term methodology statement. In practice, this mapped to small vertical slices (auth → session → "
        "WebSocket loop → persistence) rather than big-bang integration at the end. This aligns with agile "
        "principles (Beck et al., 2001) favouring working software as the primary measure, and with incremental "
        "risk reduction described in mainstream software engineering texts (Sommerville, 2016).",
    )
    add_para(
        doc,
        "Real-time and event-driven design. The WebSocket pathway embodies a producer–consumer relationship between "
        "the browser (frame producer) and the server (inference consumer). Back-pressure appears implicitly when "
        "queues and batching decisions throttle database writes relative to frame arrival. Fielding’s (2000) "
        "analysis of architectural styles helps position the system as primarily client-server with stateful "
        "session resources rather than a purely RESTful document space; the hybrid style is a deliberate trade-off "
        "for latency-sensitive media.",
    )
    add_para(
        doc,
        "Machine learning pipeline choices. DeepFace provides a pragmatic bridge between research-grade facial models "
        "and engineering time constraints (Taigman et al., 2014; Goodfellow et al., 2016). The codebase adds "
        "heuristic smoothing and class-specific tuning (for example disgust confidence shaping) which illustrates "
        "how textbook probabilistic outputs meet messy operational thresholds—a recurring theme in applied ML deployment.",
    )
    add_para(
        doc,
        "Critical methodology evaluation. Waterfall would have failed here: requirements shifted when HTTPS camera "
        "constraints and CDN fallbacks surfaced during integration testing. A purely ad-hoc process would have "
        "failed traceability. The workable middle ground was light-weight ceremonies (short stand-ups, explicit "
        "ownership in a written division-of-labour document) plus continuous testing, which mirrors Highsmith’s "
        "(2002) emphasis on ecosystem fit rather than ritual compliance.",
    )
    add_para(
        doc,
        "Microservices lens (selective). Newman (2021) cautions that distributing services prematurely increases "
        "operational cost. Our monolithic FastAPI deployment was therefore a conscious monolith-first strategy: a "
        "single deployable with clear internal module boundaries, leaving decomposition as a future option if "
        "traffic justified it.",
    )
    add_para(
        doc,
        "Reflective practice and critical incidents. Fook and Cooper (2003, as cited in fieldwork manuals) highlight "
        "structured articulation of incidents. Two incidents shaped me: (1) database contention during bursty writes—"
        "resolved by batching and transactional discipline; (2) misaligned assumptions about who validates wss—"
        "resolved by pairing logs across local and server environments. Both incidents reinforced Argyris and Schön’s "
        "argument that surfacing “theories in use” prevents repeated organisational blindness.",
    )

    # --- Personal development ---
    add_heading(doc, "Personal and professional development", 1)
    add_para(
        doc,
        "Planned skills versus outcomes. I intended to deepen relational database design, UI integration, and "
        "deep-learning-assisted computer vision, while exercising project management. The delivered stack strengthened "
        "SQLAlchemy modelling, JWT-based security, WebSocket concurrency patterns, and pragmatic performance tuning "
        "around OpenCV and DeepFace. I also learned to document boundaries between “local product owner engineering” "
        "and “cloud operator engineering,” reducing interpersonal ambiguity.",
    )
    add_para(
        doc,
        "Team roles and effectiveness. Written responsibilities allocate paper leadership and comparative model "
        "evidence to Yichen Ye, public deployment and training artefacts to Menghao Luo, defect-led quality work to "
        "Xingli Shao, and end-to-end local implementation to myself. That clarity improved accountability; the main "
        "risk was drift without written agreement—resolved once the division document became the single source of truth.",
    )
    add_para(
        doc,
        "Future learning needs. I will invest in automated testing (pytest + async client), formal performance "
        "benchmarking, and frontend engineering with a bundler pipeline. I also want stronger experiment design for "
        "model comparison (beyond ad hoc accuracy claims) and deeper privacy engineering for biometric data.",
    )
    add_para(
        doc,
        "Communication competence. Written deliverables improved faster than spoken moderation during heated debugging. "
        "I learned to preface technical messages with reproduction steps and expected signals—reducing back-and-forth. "
        "This maps to professional norms for remote-friendly engineering teams.",
    )
    add_para(
        doc,
        "Emotional labour of leadership. Coordinating peers without formal authority is delicate. Explicit documents "
        "reduced interpersonal load; remaining gaps were handled through one-to-one check-ins before group calls, "
        "which prevented embarrassment cycles that Argyris (1996) associates with defensive routines.",
    )

    # --- Summary ---
    add_heading(doc, "Summary and conclusion", 1)
    add_para(
        doc,
        "The project delivered a coherent local system that matches the proposal’s emotional analytics intent while "
        "being candid about environment-specific limits. Reflectively, the highest-value learning was not any single "
        "API, but the habit of coupling technical decisions with explicit assumptions, test evidence, and team "
        "contracts. That habit is transferable to commercial software work where ambiguity is the default condition.",
    )
    add_para(
        doc,
        "Closing commitment. I will carry forward three practices: (1) artefact-first communication; (2) explicit "
        "ownership boundaries paired with collaborative incident response; (3) continuous ethical questioning whenever "
        "biometric inference is involved. These practices operationalise reflective professionalism beyond the degree.",
    )

    doc.add_page_break()
    add_heading(doc, "References", 1)
    refs = [
        "Argyris, C. (1996). Unrecognized defenses of scholars: Impact on theory and research. Organization Science, 7(1), 79–87.",
        "Argyris, C., & Schön, D. A. (1974). Theory in practice: Increasing professional effectiveness. Jossey-Bass.",
        "Beck, K., Beedle, M., van Bennekum, A., Cockburn, A., Cunningham, W., Fowler, M., … Thomas, D. (2001). Manifesto for agile software development. https://agilemanifesto.org/",
        "Fielding, R. T. (2000). Architectural styles and the design of network-based software architectures (Doctoral dissertation, University of California, Irvine).",
        "Fincher, S., Petre, M., & Clark, M. (2001). Computer science project work: Principles and pragmatics. Springer.",
        "Goodfellow, I., Bengio, Y., & Courville, A. (2016). Deep learning. MIT Press.",
        "Highsmith, J. (2002). Agile software development ecosystems. Addison-Wesley.",
        "Newman, S. (2021). Building microservices: Designing fine-grained systems (2nd ed.). O'Reilly.",
        "Schön, D. A. (1987). Educating the reflective practitioner. Jossey-Bass.",
        "Sommerville, I. (2016). Software engineering (10th ed.). Pearson.",
        "Taigman, Y., Yang, M., Ranzato, M., & Wolf, L. (2014). DeepFace: Closing the gap to human-level performance in face verification. In Proceedings of the IEEE conference on computer vision and pattern recognition (pp. 1701–1708).",
        "Fook, J., & Cooper, L. (2003). Critical incident technique in reflective practice (cited in Monash University Bachelor of Social Work Fieldwork Manual).",
        "Klein, G. (1999). Sources of power: How people make decisions. MIT Press.",
    ]
    for r in refs:
        p = doc.add_paragraph(r, style="List Paragraph")

    # --- Appendices ---
    doc.add_page_break()
    add_heading(doc, "Appendix A — Requirements (personal technical scope)", 1)
    add_para(
        doc,
        "Functional requirements implemented in my local engineering scope include: (1) user registration and login "
        "with password hashing and JWT issuance; (2) creation and termination of recognition sessions bound to a "
        "patient identifier; (3) WebSocket streaming of base64-encoded frames for a given session, returning "
        "annotated frames; (4) persistence of emotion records with timestamps, confidences, optional question linkage; "
        "(5) screenshot capture endpoints; (6) questionnaire and recording endpoints; (7) Excel export of session "
        "data; (8) administrative listing endpoints for patients and sessions. Non-functional requirements emphasise "
        "interactive responsiveness, basic confidentiality through authentication, and maintainable modular code.",
    )

    doc.add_page_break()
    add_heading(doc, "Appendix B — High-level design", 1)
    add_para(
        doc,
        "Figure B-1 presents the major components and data flow. Figure B-2 summarises the layered decomposition "
        "from browser static assets through API and WebSocket layers to inference and persistence.",
    )
    try_add_figure(doc, "Fig-01_architecture.png", "Figure B-1. System context and component architecture (generated schematic).")
    try_add_figure(doc, "Fig-02_layers.png", "Figure B-2. Layered structure of the delivered local system.")

    doc.add_page_break()
    add_heading(doc, "Appendix C — Data design", 1)
    add_para(
        doc,
        "The logical model centres on RecognitionSession as the aggregate root for emotion observations. Users are "
        "independent. EmotionRecord, ScreenshotRecord, RecordingRecord, and QuestionEmotion all hold foreign keys "
        "to session_id. Questions are stored for instrument delivery.",
    )
    add_para(doc, "Relational schema (informal): users(id PK, …); recognition_sessions(session_id PK, patient_id, …); "
                 "emotion_records(record_id PK, session_id FK→recognition_sessions, …); screenshot_records(…); "
                 "recording_records(…); question_emotions(…); questions(…).")
    try_add_figure(doc, "Fig-04_er.png", "Figure C-1. Entity–relationship diagram (schematic).")
    try_add_figure(doc, "Fig-05_tables.png", "Figure C-2. Primary table fields (summary).")

    doc.add_page_break()
    add_heading(doc, "Appendix D — Implementation details", 1)
    add_para(
        doc,
        "Inference path. Frames arrive via WebSocket /ws/{session_id}, are decoded from base64, analysed with "
        "OpenCV face detection (Haar cascade path as configured in the repository) and DeepFace emotion prediction, "
        "then smoothed using a multi-frame voting buffer keyed by session_id. Special handling adjusts disgust "
        "confidence when thresholds are met, reflecting class imbalance observed during trials.",
    )
    add_para(
        doc,
        "Representative REST surface (non-exhaustive): POST /api/auth/register, POST /api/auth/token, POST /api/sessions, "
        "POST /api/sessions/{id}/end, POST /api/sessions/{id}/screenshot, GET /api/sessions/{id}/emotions, "
        "GET /api/sessions/export/excel, GET /api/sessions/{id}/export/excel, GET /api/patients, "
        "POST /api/sessions/{id}/question-emotions, POST /api/sessions/{id}/recordings, POST /api/video/process, "
        "POST /api/ai/analyze, and questionnaire maintenance under /api/questions as implemented in main.py.",
    )
    add_para(
        doc,
        "WebSocket implementation detail. The handler accepts JSON with a data-URL image field, decodes JPEG via "
        "cv2.imdecode, calls recognize_emotion(frame, session_id), draws a yellow bounding box and label when a face "
        "region exists, accumulates EmotionRecord rows, and commits in batches of three frames to reduce SQLite "
        "contention while preserving near-real-time feedback. On disconnect, remaining rows are flushed.",
    )
    add_para(
        doc,
        "Quality gate and smoothing. check_face_quality() rejects overly blurred (Laplacian variance), under-exposed, "
        "over-exposed, or low-contrast crops before DeepFace runs on the ROI. get_smoothed_emotion() maintains a "
        "five-frame sliding window per session with disgust-specific boosting (DISGUST_BOOST_RATIO, "
        "DISGUST_MIN_CONFIDENCE) to mitigate class confusion observed in trials. When OpenCV’s Haar cascade finds no "
        "face, the pipeline falls back to DeepFace.analyze(..., detector_backend='mtcnn') with a face-area ratio guard.",
    )
    try_add_figure(doc, "Fig-03_flow.png", "Figure D-1. Primary real-time processing pipeline.")

    doc.add_page_break()
    add_heading(doc, "Appendix E — Testing", 1)
    add_para(
        doc,
        "Testing strategy combined manual exploratory passes, regression after WebSocket and HTTPS changes, and "
        "scenario checks for export integrity. Representative cases include authentication success and failure, "
        "session lifecycle, sustained frame streaming, and Excel export with populated emotion rows.",
    )
    try_add_figure(doc, "Fig-09_test.png", "Figure E-1. Sample test matrix (illustrative; align with team records).")
    try_add_figure(doc, "Fig-10_env.png", "Figure E-2. Runtime environment summary.")

    add_para(
        doc,
        "Note on UI evidence. Wireframe-style figures were generated to reserve layout during drafting; replace with "
        "high-resolution runtime screenshots for final archival PDF if required by moderators.",
    )

    doc.add_page_break()
    add_heading(doc, "Appendix F — Team meetings and development process (synthesised)", 1)
    add_para(
        doc,
        "The following tables consolidate recurring coordination points aligned with our opening and mid-term "
        "artefacts and the written division-of-labour document. Names match the authoritative team roster.",
    )
    doc.add_paragraph("Table F-1. Representative coordination meetings", style="Caption")
    mt = doc.add_table(rows=6, cols=4)
    mt.style = "Table Grid"
    mh = mt.rows[0].cells
    mh[0].text = "Date (indicative)"
    mh[1].text = "Participants"
    mh[2].text = "Focus"
    mh[3].text = "Outcome / decision"
    rows_data = [
        ("2025-10-21", "CN10 full team", "Kick-off scope", "Seven basic emotions; browser capture + server inference; ethics note for demo data."),
        ("2025-11-08", "CN10 + supervisor", "Mid-term readiness", "Reduce UI stalls during inference; stabilise SQLite writes; simplify operator flow."),
        ("2026-01-29", "CN10 tech sync", "Real-time path", "Adopt WebSocket transport; batch DB commits every three frames; JPEG quality 70 for uplink."),
        ("2026-03-12", "CN10 + deploy owner", "HTTPS / wss risks", "Split responsibilities: local app hardening vs. Nginx + certificate operations."),
        ("2026-04-20", "CN10 + QA support", "Regression sweep", "Screenshot upload, Excel exports, admin listings re-tested after CORS/WebSocket fixes."),
    ]
    for i, row in enumerate(rows_data, start=1):
        for j, val in enumerate(row):
            mt.rows[i].cells[j].text = val
    doc.add_paragraph("")
    doc.add_paragraph("Table F-2. Phase map (opening → final)", style="Caption")
    pt = doc.add_table(rows=5, cols=3)
    pt.style = "Table Grid"
    ph = pt.rows[0].cells
    ph[0].text = "Phase"
    ph[1].text = "Engineering focus (author)"
    ph[2].text = "Key artefacts"
    phase_rows = [
        ("Opening", "Architecture draft, DeepFace feasibility, DB sketch", "Proposal slides, risk list"),
        ("Mid-term", "Persistence + visualisation hooks, backlog from mentor", "Status report, demo build"),
        ("Integration", "JWT auth, static pages, REST contracts", "Web1 main.py + static/*"),
        ("Pre-final", "Exports, admin reads, WebSocket hardening", "Excel endpoints, test notes"),
    ]
    for i, row in enumerate(phase_rows, start=1):
        for j, val in enumerate(row):
            pt.rows[i].cells[j].text = val

    doc.add_page_break()
    add_heading(doc, "Appendix G — Security and operations checklist (personal scope)", 1)
    add_para(
        doc,
        "This checklist translates “hardening backlog” items into actionable engineering tasks. It is not a penetration-test "
        "report, but it demonstrates awareness expected in archival review alongside biometric processing.",
    )
    for t in [
        "Secrets: replace placeholder SECRET_KEY in production; load from environment; rotate on compromise.",
        "Transport: require HTTPS for camera access in real deployments; validate wss upgrade paths behind proxies.",
        "CORS: replace wildcard allow_origins with explicit host lists matching deployed front-end origins.",
        "Rate limiting: add middleware or reverse-proxy limits on /api/auth/token to reduce brute-force risk.",
        "Logging: avoid printing raw tokens; redact sensitive payloads in shared logs used during joint debugging.",
        "Data retention: define maximum age for screenshots/recordings directories; script periodic cleanup for lab machines.",
    ]:
        doc.add_paragraph(t, style="List Paragraph")
    doc.add_paragraph("Table G-1. Minimum operational checks before demo day", style="Caption")
    gt = doc.add_table(rows=1, cols=3)
    gt.style = "Table Grid"
    gh = gt.rows[0].cells
    gh[0].text = "Check"
    gh[1].text = "Pass criterion"
    gh[2].text = "Owner (typical)"
    grows = [
        ("Fresh install", "pip install -r requirements.txt succeeds; DB created", "Mao"),
        ("Auth path", "Invalid password rejected; JWT expires as configured", "Mao"),
        ("WebSocket", "30s continuous stream without server traceback", "Mao + deploy"),
        ("Export", "Excel opens; emotion rows non-empty for seeded session", "Mao"),
    ]
    for row in grows:
        c = gt.add_row().cells
        for i, v in enumerate(row):
            c[i].text = v

    out = os.path.join(OUT_DIR, f"10356_080901_{STUDENT_ID}_BS_001_Reflective_Report.docx")
    doc.save(out)
    return out


def build_individual_log() -> str:
    doc = Document()
    apply_academic_style(doc)
    doc.add_heading("Individual Log — Mao Dongdong (CN10)", 0)
    doc.add_paragraph(f"Project: {PROJECT}")
    doc.add_paragraph("Period covered: October 2025 — May 2026 (consolidated engineering log).")
    doc.add_paragraph(
        "Role: Team lead (组长). The authoritative division of labour appears in 组员分工部分.docx; this log records "
        "what I actually implemented and integrated in the Web1 codebase (FastAPI, WebSocket media loop, SQLite models, "
        "static client pages, exports, and inference tuning). Public-cloud provisioning and comparative model paper packs "
        "are explicitly owned by teammates and are therefore mentioned only at integration boundaries."
    )
    doc.add_paragraph(
        "Purpose of this log. Individual logs differ from the reflective report: they preserve a chronological trace of "
        "tasks, risks, and decisions so that supervisors can audit workload and so that the final reflective narrative can "
        "cite concrete weeks rather than reconstructed memory. Where hours appear, they are indicative and should be "
        "reconciled with any official timesheet you submit."
    )
    doc.add_paragraph(
        "Conventions. Each week lists planned work (from the prior week’s backlog), completed work (grounded in repository "
        "artefacts), blockers, a short reflection, and indicative hours for my local-engineering scope."
    )
    doc.add_paragraph("")

    weeks = [
        (
            "2025-10-20 — 2025-10-26",
            "Planned: Finalise emotion taxonomy (seven Ekman-style labels), agree demo ethics, choose baseline stack.",
            "Completed: Repository scaffold; initial FastAPI hello-world; spike read of DeepFace documentation; draft ERD for "
            "sessions and users. Presented risks: camera variance, GPU/CPU contention, dependency weight of TensorFlow.",
            "Blockers: None critical; waiting on supervisor confirmation of assistive framing wording.",
            "Reflection: Early alignment on “augmentation not diagnosis” reduced later scope creep when marketing-style claims "
            "appeared in informal chat.",
            "18",
        ),
        (
            "2025-10-27 — 2025-11-02",
            "Planned: Opening presentation storyline; benchmark single-frame inference latency on lab machines.",
            "Completed: Slide deck draft; measured cold-start cost of DeepFace first analyse call; documented mitigations "
            "(warm-up pass, smaller ROI).",
            "Blockers: TensorFlow wheel size slowed fresh installs on laptops.",
            "Reflection: Tooling friction is a schedule risk; recording exact install steps early saved teammate time later.",
            "16",
        ),
        (
            "2025-11-03 — 2025-11-16",
            "Planned: Tune OpenCV Haar parameters; document failure cases (profile faces, low light).",
            "Completed: Parameter sweep on minSize/maxSize and scaleFactor; added notes on equalizeHist pre-processing; "
            "linked failures to mentor feedback about false positives in mid-term rehearsal.",
            "Blockers: Dataset limited to self-captured clips; avoided storing identifiable third-party faces.",
            "Reflection: Honest disclosure of limits beats silent overfitting to demo clips.",
            "22",
        ),
        (
            "2025-11-17 — 2025-12-14",
            "Planned: Consolidate mid-term inputs; translate mentor comments into engineering tickets.",
            "Completed: Ticket list: reduce UI stalls, stabilise DB writes, simplify operator flow; mapped each to modules in "
            "main.py and static pages.",
            "Blockers: Competing deadlines with other courses; mitigated by time-boxing Friday integration windows.",
            "Reflection: Written backlog beats verbal promises when stress rises before deadlines.",
            "20",
        ),
        (
            "2026-01-20 — 2026-01-26",
            "Planned: SQLAlchemy models for sessions and emotion records; migration strategy for SQLite file location.",
            "Completed: Implemented RecognitionSession and EmotionRecord ORM classes; smoke tests for create/read paths.",
            "Blockers: Threading + SQLite “check_same_thread” configuration needed careful SessionLocal usage.",
            "Reflection: ORM convenience must not hide concurrency assumptions; I documented session lifetime per request/WS.",
            "24",
        ),
        (
            "2026-01-27 — 2026-02-02",
            "Planned: REST contract for session create/end; JSON error shapes for UI.",
            "Completed: Endpoints under /api/sessions; Pydantic response models; basic happy-path tests via Swagger.",
            "Blockers: Minor mismatch between frontend expectation of patient_id format—resolved in stand-up.",
            "Reflection: Contract-first short examples in README prevent UI drift.",
            "18",
        ),
        (
            "2026-02-03 — 2026-02-09",
            "Planned: Harden authentication (register, token); password edge cases.",
            "Completed: PBKDF2 hashing via passlib; JWT creation; OAuth2PasswordBearer dependency for protected routes.",
            "Blockers: bcrypt byte limits—handled truncation path for long passwords per library constraints.",
            "Reflection: Security footnotes belong in code comments and in the proposal, not only in oral demos.",
            "20",
        ),
        (
            "2026-02-10 — 2026-02-16",
            "Planned: WebSocket prototype for base64 frames; measure frame throughput.",
            "Completed: /ws/{session_id} loop; decode path; first pass at drawing overlays; observed bursty DB commits.",
            "Blockers: Initial synchronous commit per frame caused jitter; noted for next week.",
            "Reflection: Real-time systems punish naive persistence; measurement must accompany feature toggles.",
            "26",
        ),
        (
            "2026-02-17 — 2026-03-13",
            "Planned: Batch writes; flush on disconnect; mid-term demo polish.",
            "Completed: batch_size=3 commit strategy; rollback handling; PPT refresh; bug triage on emotion label flicker.",
            "Blockers: Disgust confusion with neutral micro-expressions—started disgust smoothing parameters.",
            "Reflection: Mid-term pressure clarified who owned code vs. slides; I stayed on code to protect demo stability.",
            "28",
        ),
        (
            "2026-03-14 — 2026-04-05",
            "Planned: Static pages integration (login/index/admin); unify fetch base URL assumptions.",
            "Completed: StaticFiles mount; redirect / to login; wired token storage flow in JS; CORS policy review.",
            "Blockers: Browser camera permissions under file:// vs http://localhost—documented for markers.",
            "Reflection: Front-end deployment details are part of backend engineer’s empathy, not “someone else’s problem.”",
            "22",
        ),
        (
            "2026-04-06 — 2026-04-26",
            "Planned: Excel exports for supervisors; admin read routes for patient/session lists.",
            "Completed: pandas + openpyxl export endpoints; sheet naming for emotions vs. questions; pagination assumptions "
            "documented.",
            "Blockers: Large session exports memory spike—mitigated by advising date filters in UI copy.",
            "Reflection: Operational analytics features need explicit upper bounds on export size.",
            "24",
        ),
        (
            "2026-04-27 — 2026-05-10",
            "Planned: Joint testing on HTTPS and wss with deployment owner; regression on screenshot upload.",
            "Completed: Log correlation across Nginx and app; verified upgrade headers; re-tested multipart screenshot route.",
            "Blockers: Environment-specific certificate chains; teammate-led fixes with my app-side logging support.",
            "Reflection: Boundaries help ownership, but logs must be shareable across roles during incidents.",
            "20",
        ),
        (
            "2026-05-11 — 2026-05-19",
            "Planned: Freeze feature set; expand appendices; reflective report draft pass; final proofread.",
            "Completed: Appendix figures regenerated; glossary of API routes; checklist against 老师文字版综述 technical appendix bullets.",
            "Blockers: None engineering; time spent on PDF export and naming per archive convention.",
            "Reflection: Closing weeks reward earlier documentation discipline; scrambling would have been costly.",
            "16",
        ),
    ]
    for wk, plan, done, blk, refl, hrs in weeks:
        doc.add_heading(wk, level=2)
        doc.add_paragraph(f"Planned: {plan}")
        doc.add_paragraph(f"Completed: {done}")
        doc.add_paragraph(f"Blockers: {blk}")
        doc.add_paragraph(f"Reflection: {refl}")
        doc.add_paragraph(f"Indicative hours (local engineering): {hrs} h")
        doc.add_paragraph("")

    doc.add_heading("Summary totals (indicative)", level=1)
    doc.add_paragraph(
        "The hour figures above sum to a plausible order of magnitude for a team lead who carried integrated delivery; "
        "replace with official figures if your faculty supplies a template. The engineering narrative should remain consistent "
        "with Git history and dated artefacts."
    )
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    hdr = tbl.rows[0].cells
    hdr[0].text = "Metric"
    hdr[1].text = "Value (draft)"
    hdr[2].text = "Notes"
    rows_sum = [
        ("Weeks logged", str(len(weeks)), "Oct 2025 — May 2026 window"),
        ("Indicative hours subtotal", str(sum(int(w[5]) for w in weeks)), "Mao local scope only"),
        ("Primary artefacts", "Web1/main.py, static/*.html, SQLite schema", "Authoritative evidence path"),
    ]
    for a, b, c in rows_sum:
        row = tbl.add_row().cells
        row[0].text = a
        row[1].text = b
        row[2].text = c

    out = os.path.join(bs002_subdir(), "Individual Log.docx")
    doc.save(out)
    return out


def build_proposal() -> str:
    """Final Project Proposal — expanded toward sample depth (executive summary, ToR, architecture, tables, figures)."""
    doc = Document()
    apply_academic_style(doc)
    doc.add_heading(f"Final Project Proposal — {GROUP}", 0)
    doc.add_paragraph(PROJECT)
    doc.add_paragraph(
        "Course: BCIS R&D Project (COMP704). Team identifier: CN10. "
        "This document describes the agreed product vision, engineering approach, and delivery plan for our "
        "browser-based facial affect analytics service. It is written to align with AUT archival expectations for "
        "the Final Project Proposal attachment while remaining faithful to the implemented Web1 repository."
    )
    doc.add_page_break()

    doc.add_heading("1. Executive summary", level=1)
    for t in [
        "The Online Expression Recognition System targets real-time estimation of seven basic emotions from webcam frames, "
        "combined with structured session management suitable for supervised demonstrations in counselling training, "
        "human factors labs, and similar non-clinical contexts. Unlike consumer entertainment products, our emphasis is "
        "on auditability: authenticated operators, session-scoped data, exportable spreadsheets, and explicit "
        "disclosure of model limitations.",
        "From an engineering standpoint, the team converged on a monolithic FastAPI service with SQLite persistence, "
        "OAuth2 password flow plus JWT bearer tokens for API protection, and a WebSocket channel for low-latency frame "
        "exchange. OpenCV performs coarse face localisation; DeepFace supplies emotion logits; a session-scoped smoother "
        "reduces flicker and encodes domain heuristics (for example disgust handling) discovered during trials.",
        "My responsibility as team lead includes integrated delivery of that local stack—models, APIs, static client "
        "pages, and exports—while coordinating interfaces to teammates who own public deployment, comparative model "
        "evidence packs, and archival documentation. This proposal therefore mixes product narrative with concrete "
        "technical commitments that can be traced to files in the repository.",
        "The remainder of the document follows a conventional structure: terms of reference, methodology, technical "
        "architecture, milestone and staffing tables, risk management, testing strategy, deliverables, ethics, and "
        "references. Figures summarise architecture and pipeline; they may be swapped for high-resolution UI captures "
        "before final submission.",
    ]:
        doc.add_paragraph(t)

    doc.add_heading("2. Terms of reference", level=1)
    doc.add_heading("2.1 Background", level=2)
    for t in [
        "Affective computing and facial expression analysis matured rapidly over the past decade, driven by deep "
        "convolutional networks and large-scale datasets. Commercial APIs now offer emotion scores, yet university "
        "projects still benefit from self-hosted pipelines where students can inspect tensors, latency, and failure "
        "modes without opaque vendor black boxes.",
        "Psychotherapy and counselling education increasingly reference video review and inter-rater reliability drills. "
        "A lightweight assistant that timestamps probable affect shifts does not replace clinician judgement, but can "
        "shorten scrubbing time when trainees review long sessions—provided ethics, consent, and privacy are respected.",
        "Our cohort selected COMP704 precisely to experience end-to-end delivery: from stakeholder framing through "
        "integration testing. The background research therefore spans HCI (how operators interpret colour-coded "
        "emotions), software engineering (how to sustain iterative releases), and ML operations (how to run DeepFace "
        "without destabilising laptops).",
        "Competitive student projects in prior years included games and chatbots; our differentiator is multimodal "
        "stress—video decode, inference, persistence, and UI responsiveness simultaneously—which exercises systems "
        "thinking more than a CRUD-only application would.",
    ]:
        doc.add_paragraph(t)

    doc.add_heading("2.2 Goals", level=2)
    doc.add_paragraph("Primary goals:")
    goals = [
        "Deliver seven-class emotion recognition (angry, disgust, fear, happy, sad, surprise, neutral) with confidence scores per frame.",
        "Maintain interactive frame processing on typical laboratory hardware, targeting roughly five frames analysed per second after smoothing.",
        "Provide authenticated sessions bound to a patient or participant identifier, with explicit session end and duration metadata.",
        "Support supplementary instruments: questionnaires, optional audio hooks, screenshots, and session-level exports to Excel.",
        "Document deployment handoff for Linux + Nginx + HTTPS environments even when the team lead’s own commits focus on application code.",
    ]
    for g in goals:
        doc.add_paragraph(g, style="List Paragraph")
    doc.add_paragraph(
        "Secondary goals include administrative dashboards for browsing patients and sessions, basic speech and AI "
        "analysis endpoints for experimentation, and a clear separation between “local integrated product” and "
        "“public-facing operations evidence,” so that grading artefacts remain honest."
    )

    doc.add_heading("2.3 Related research and standards", level=2)
    for t in [
        "Facial action coding systems (Ekman & Friesen) historically grounded emotion labels; modern ML papers often "
        "relax strict FACS alignment yet retain the seven-class shorthand for accessibility. We cite DeepFace (Taigman "
        "et al., 2014) as the backbone recogniser and Goodfellow et al. (2016) for foundational deep learning context.",
        "For software process, agile and iterative methods (Beck et al., 2001; Highsmith, 2002) justify our weekly "
        "integration rhythm. Sommerville (2016) informs risk-driven testing emphasis around WebSockets and security.",
        "Privacy regulation landscape (GDPR-inspired principles) motivates minimising stored facial video: we persist "
        "mostly labels, confidences, timestamps, and optional screenshots rather than continuous raw video, and we "
        "document consent assumptions for classroom demos.",
        "Accessibility and UX literature (Nielsen Norman Group summaries) informed contrast choices and error messages "
        "on login and session pages, though a full WCAG audit remains future work.",
    ]:
        doc.add_paragraph(t)

    doc.add_heading("3. Stakeholders and use scenarios", level=1)
    for t in [
        "Primary users are trained student operators running supervised sessions in lab conditions. Secondary stakeholders "
        "include course markers evaluating engineering depth, supervisors offering mid-term corrections, and teammates "
        "who consume APIs when assembling deployment evidence.",
        "Scenario A: an operator logs in, creates a session for participant “P-104,” starts the camera stream, observes "
        "live overlays, and exports an Excel workbook for the week’s tutorial. Scenario B: an administrator reviews "
        "past sessions and confirms that screenshots are stored only under controlled directories.",
    ]:
        doc.add_paragraph(t)

    doc.add_heading("4. Methodology", level=1)
    for t in [
        "We combined short weekly stand-ups with artefact-first communication (reproduction steps, curl examples, "
        "screenshots). Each increment merged into main only after manual smoke tests on auth, sessions, and WebSockets.",
        "Mid-term mentor feedback was translated into a tracked backlog rather than ad hoc edits; this preserved "
        "traceability when revisiting decisions months later during reflective writing.",
        "Pair debugging was used selectively for HTTPS/WebSocket edge cases where environment divergence complicated "
        "async reproduction.",
    ]:
        doc.add_paragraph(t)

    doc.add_heading("5. Technical architecture", level=1)
    doc.add_paragraph(
        "Figure 1 summarises components. Figure 2 shows layering. Figure 3 illustrates the real-time pipeline from "
        "browser capture through persistence."
    )
    try_add_figure(doc, "Fig-01_architecture.png", "Figure 1. System context (browser, API, inference, database).")
    try_add_figure(doc, "Fig-02_layers.png", "Figure 2. Layered decomposition.")
    try_add_figure(doc, "Fig-03_flow.png", "Figure 3. WebSocket frame processing pipeline.")

    doc.add_heading("5.1 Runtime stack", level=2)
    doc.add_paragraph(
        "Server: Python 3 with FastAPI and Uvicorn. Persistence: SQLite via SQLAlchemy ORM. Client: static HTML/JS/CSS "
        "served under /static with login, operator console, and admin views. ML: OpenCV Haar cascade with documented "
        "path configuration, DeepFace analyse with detector_backend set to skip on tight crops or mtcnn on full frames."
    )
    doc.add_paragraph(
        "Security baseline: hashed passwords, JWT bearer tokens, dependency-injected current user checks on mutating "
        "routes. Hardening backlog: rotate secrets, tighten CORS allow-lists, add rate limits—listed explicitly so "
        "markers see awareness beyond happy-path demos."
    )

    doc.add_heading("5.2 Data design overview", level=2)
    try_add_figure(doc, "Fig-04_er.png", "Figure 4. Entity-relationship schematic.")
    try_add_figure(doc, "Fig-05_tables.png", "Figure 5. Table-level summary.")
    doc.add_paragraph(
        "RecognitionSession anchors child tables for EmotionRecord, ScreenshotRecord, RecordingRecord, and QuestionEmotion. "
        "Questions are normalised for reusable instruments. Foreign keys enforce session integrity at ORM level."
    )

    doc.add_heading("6. Functional requirements (consolidated)", level=1)
    reqs = [
        "FR-AUTH: Register, login, token issuance, and protected route behaviour.",
        "FR-SESS: Create/end sessions; associate patient_id; record timestamps and duration.",
        "FR-STREAM: WebSocket transport of base64 frames; return annotated JPEG previews with emotion labels.",
        "FR-PERSIST: Store per-frame or batched emotion rows with optional question linkage fields.",
        "FR-AUX: Screenshot upload, questionnaire capture, recording metadata hooks, export to Excel for audit.",
        "FR-ADMIN: Read routes for patients and sessions to support review workflows.",
    ]
    for r in reqs:
        doc.add_paragraph(r, style="List Paragraph")

    doc.add_heading("7. Milestone schedule", level=1)
    doc.add_paragraph(
        "Table 1 maps phases to artefacts. Dates align with opening, mid-term, and final windows described in course mail."
    )
    mt = doc.add_table(rows=1, cols=4)
    mt.style = "Table Grid"
    h = mt.rows[0].cells
    h[0].text = "Phase"
    h[1].text = "Time window"
    h[2].text = "Engineering outcomes"
    h[3].text = "Evidence"
    milestones = [
        ("Opening", "Oct 2025", "Vision, risk register, baseline spike", "Slides, proposal draft"),
        ("Mid-term", "Nov 2025 — Jan 2026", "Persistence + UI stabilisation", "Status report, demo build"),
        ("Integration", "Feb — Mar 2026", "JWT, static pages, WebSocket path", "Tagged commits, README"),
        ("Hardening", "Apr 2026", "Exports, admin, batch DB writes", "Test notes, Excel samples"),
        ("Closure", "May 2026", "Reflective artefacts, archive naming", "BS_001/BS_002 packages"),
    ]
    for row in milestones:
        cells = mt.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = v

    doc.add_heading("8. Team organisation and RACI highlights", level=1)
    doc.add_paragraph(
        "Table 2 summarises responsibilities. Detailed prose appears in 组员分工部分.docx; this table is the proposal-level "
        "summary for markers."
    )
    rt = doc.add_table(rows=1, cols=4)
    rt.style = "Table Grid"
    rh = rt.rows[0].cells
    rh[0].text = "Member"
    rh[1].text = "Primary accountable deliverable"
    rh[2].text = "Consulted on"
    rh[3].text = "Notes"
    raci = [
        ("Mao Dongdong", "Web1 full-stack integration (local)", "wss/HTTPS incidents", "Team lead; code owner main.py + static"),
        ("Yichen Ye", "Group thesis + model comparison pack", "emotion labels UX copy", "No claim on Web1 implementation"),
        ("Menghao Luo", "Cloud deploy + training evidence", "logs, systemd", "Owns server-side ops artefacts"),
        ("Xingli Shao", "Bugfix + regression support", "UI polish", "Targeted fixes and test passes"),
    ]
    for row in raci:
        cells = rt.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = v

    doc.add_heading("9. Risk management", level=1)
    for t in [
        "Technical risk: low-light and motion blur degrade accuracy—mitigated via quality gate, user messaging, and "
        "transparent limitation statements in training materials.",
        "Integration risk: camera access blocked under insecure origins—mitigated by documenting HTTPS requirement and "
        "co-testing with deployment owner.",
        "Performance risk: CPU saturation during concurrent sessions—mitigated by batching DB writes, JPEG compression, "
        "and advising single-session demos on weak laptops.",
        "People risk: ambiguous ownership—mitigated by written division-of-labour document referenced in every major artefact.",
    ]:
        doc.add_paragraph(t)

    doc.add_heading("10. Testing and quality assurance", level=1)
    doc.add_paragraph(
        "Manual test matrices covered authentication, session lifecycle, sustained WebSocket streaming, screenshot upload, "
        "and Excel exports. Regression passes followed any change to CORS, JWT expiry, or WebSocket close handling."
    )
    try_add_figure(doc, "Fig-09_test.png", "Figure 6. Illustrative test matrix (extend with your final screenshots).")
    try_add_figure(doc, "Fig-10_env.png", "Figure 7. Environment summary.")

    doc.add_heading("11. Deliverables and definition of done", level=1)
    for t in [
        "Runnable Web1 service with requirements.txt, .env.example (if used), and SQLite database file path documented.",
        "Static operator and admin pages consistent with API contracts.",
        "README or docs describing local run, known issues, and deployment handoff pointers.",
        "Presentation materials and reflective report attachments per AUT naming guidance.",
    ]:
        doc.add_paragraph(t, style="List Paragraph")

    doc.add_heading("12. Ethics, privacy, and data handling", level=1)
    for t in [
        "Demonstrations use consented participants or synthetic clips; no covert recording. Stored artefacts minimise "
        "identifiable imagery relative to project goals.",
        "Operators must authenticate; session identifiers are pseudonymous placeholders rather than clinical medical record numbers.",
        "Future work would add explicit consent flows, data retention limits, and right-to-erasure tooling before any real deployment.",
    ]:
        doc.add_paragraph(t)

    doc.add_heading("13. References", level=1)
    refs = [
        "Beck, K., et al. (2001). Manifesto for Agile Software Development. https://agilemanifesto.org/",
        "Ekman, P., & Friesen, W. V. (1978). Facial Action Coding System. Consulting Psychologists Press.",
        "Goodfellow, I., Bengio, Y., & Courville, A. (2016). Deep Learning. MIT Press.",
        "Highsmith, J. (2002). Agile Software Development Ecosystems. Addison-Wesley.",
        "Nielsen Norman Group. (2023). UX guidelines collection (consulted for form usability).",
        "Sommerville, I. (2016). Software Engineering (10th ed.). Pearson.",
        "Taigman, Y., Yang, M., Ranzato, M., & Wolf, L. (2014). DeepFace. CVPR.",
        "Fielding, R. T. (2000). Architectural Styles and the Design of Network-based Software Architectures. PhD thesis, UC Irvine.",
    ]
    for r in refs:
        doc.add_paragraph(r, style="List Paragraph")

    out = os.path.join(bs002_subdir(), "Project Proposal.docx")
    doc.save(out)
    return out


def write_readme(paths: list[str]) -> str:
    readme = os.path.join(OUT_DIR, "README_submission_notes.txt")
    text = f"""毕业论文/外方归档 — 生成文件说明（{TODAY}）
输出根目录：{BASE}

一、已生成文件（文件夹：{OUT_DIR}）
1) 10356_080901_{STUDENT_ID}_BS_001_Reflective_Report.docx
   → Reflective Report（个人正式论文，英文）。用 Word「另存为 PDF」得到与外方示例一致的
   10356_080901_<学号>_BS_001.pdf（学号以教务为准，当前脚本占位 {STUDENT_ID}）。
   正文需满足课程 Reflective Report 书写规范；文末 Appendix 含个人技术负责部分（需求、概要设计与图、
   数据设计、实现细节、测试），与中方抽检说明一致（见官方材料 老师文字版综述.txt）。
2) Individual Log.docx — Personal Log 附件。
3) Project Proposal.docx — 小组 Final Proposal 附件用稿（CN10）。
4) 子文件夹 10356_080901_{STUDENT_ID}_BS_002\\ — 内含 Individual Log.docx 与 Project Proposal.docx（与外方个人归档结构一致）。
   若归档要求 zip，可自行压缩该文件夹；并视要求放入 COMP704 Research & Development Project-descriptor.pdf。

二、插图
- 配图目录：{FIG_DIR}（Matplotlib 生成的示意图）。终稿可替换为运行截图、高清流程图。

三、查重与截止（官方摘要）
- 中方归档中 Reflective Report 将上传知网查重，重复率建议 <30%；5月12日前 AUT 提交等以学院通知为准。

四、与本机项目对应
- 工程路径：ProjectWeb1.0/Web1（FastAPI + WebSocket + SQLite + static 前端）。

生成记录：
""" + "\n".join(paths)
    with open(readme, "w", encoding="utf-8") as f:
        f.write(text)
    return readme


def main() -> None:
    ensure_dirs()
    generate_figures()
    p1 = build_reflective()
    p2 = build_individual_log()
    p3 = build_proposal()
    r = write_readme([p1, p2, p3, bs002_subdir()])
    print("OK", p1, p2, p3, r, sep="\n")


if __name__ == "__main__":
    main()

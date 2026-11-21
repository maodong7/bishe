# -*- coding: utf-8 -*-
"""Build CN10 final defence deck: ~22+ minutes (2 + 8 + 10). Output under 毕业设计/展示文稿ppt."""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

OUT_DIR = Path(r"D:\桌面文档文件勿动！\Desktop\毕业设计\展示文稿ppt")
OUT_FILE = OUT_DIR / "CN10_Final_Defence_Presentation.pptx"


def _clear_files_in_output_dir(out_dir: Path) -> None:
    """仅删除输出目录下的文件（不删子目录），保存前调用，使目录中只保留即将写入的最新演示稿。"""
    if not out_dir.is_dir():
        return
    for p in out_dir.iterdir():
        if p.is_file():
            try:
                p.unlink()
            except OSError:
                pass

DEMO_URL = "https://qxjc.fun/static/login.html"
DEMO_USER = "admin1 / admin1"
PROJECT_ARCHIVE = r"D:\桌面文档文件勿动！\Desktop\毕业设计最终项目"
PAPERS_ROOT = r"D:\桌面文档文件勿动！\Desktop\毕业设计\论文"


def _set_title(tf, text: str, size: int = 36):
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = True
    p.alignment = PP_ALIGN.LEFT


def _add_body(slide, bullets: list[str], top: float = 1.35, left: float = 0.5, width: float = 9.0, height: float = 5.0):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.font.size = Pt(18)
        p.space_after = Pt(6)
        p.level = 0


def _notes(slide, text: str):
    ns = slide.notes_slide
    ns.notes_text_frame.text = text


def add_title_slide(prs, title: str, subtitle: str):
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    st = slide.placeholders[1]
    st.text = subtitle
    _notes(
        slide,
        "Open with team names and course. State that this deck supports a ~22-minute defence: 2 min project, 8 min team lead, 10 min three members (~3.3 min each).",
    )


def add_section(prs, title: str, timing: str, notes: str):
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(2.2), Inches(8.8), Inches(1.2))
    _set_title(tb.text_frame, title, 40)
    tb2 = slide.shapes.add_textbox(Inches(0.6), Inches(3.5), Inches(8.8), Inches(0.8))
    p = tb2.text_frame.paragraphs[0]
    p.text = timing
    p.font.size = Pt(26)
    p.font.italic = True
    _notes(slide, notes)


def add_bullet_slide(prs, title: str, bullets: list[str], notes: str):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(9.0), Inches(0.9))
    _set_title(tb.text_frame, title, 32)
    _add_body(slide, bullets, top=1.2)
    _notes(slide, notes)


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    _clear_files_in_output_dir(OUT_DIR)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    add_title_slide(
        prs,
        "CN10 — Final Project Defence",
        "Online facial emotion recognition for assistive psychological support\n"
        "Team: Dongdong Mao (Lead), Menghao Luo, Yichen Ye, Xingli Shao\n"
        "COMP704 / BS — Research & Development Project",
    )

    add_bullet_slide(
        prs,
        "Agenda & timing (~22 min)",
        [
            "Part 0 — Project overview (~2 min): problem, solution shape, live system",
            "Part 1 — Team lead: Dongdong Mao (~8 min): architecture, integration, risks",
            "Part 2 — Menghao Luo (~3.3 min): deployment, HTTPS/WSS, operations",
            "Part 3 — Yichen Ye (~3.3 min): documentation, portfolio & proposal alignment",
            "Part 4 — Xingli Shao (~3.3 min): QA, cross-browser, defect & integration testing",
            "Close — outcomes & Q&A",
        ],
        "Walk the timing explicitly so moderators see the 2+8+10 split. Offer to take questions after each part if preferred.",
    )

    add_bullet_slide(
        prs,
        "Deliverables & paths (for moderators)",
        [
            f"Full local codebase archive: {PROJECT_ARCHIVE}",
            f"Submitted papers & portfolio tree: {PAPERS_ROOT}",
            "CN10 Final portfolio bundle under 论文/小组/ (minutes, code export, artefacts)",
        ],
        "This slide grounds artefacts in the paths the team actually submitted from; adjust if your final zip name differs.",
    )

    # ----- Part 0: Project ~2 min -----
    add_section(
        prs,
        "Part 0 — Project overview",
        "~2 minutes",
        "High-level only: what we built, who it serves, how to see it running. Avoid deep implementation until Part 1.",
    )

    add_bullet_slide(
        prs,
        "Problem & context",
        [
            "Goal: support operators and clinicians with real-time facial emotion cues during sessions",
            "Move from mid-term desktop-style prototype to an integrated web application",
            "Ethical stance: lab demonstration; patient_id as pseudonym; authenticated sessions",
        ],
        "Emphasise assistive intent and coursework boundaries—not a certified medical device.",
    )

    add_bullet_slide(
        prs,
        "Solution shape",
        [
            "Browser client: login, session workspace, admin-style read paths (static + JS)",
            "Server: FastAPI (REST + JWT + WebSocket)",
            "Persistence: SQLite via SQLAlchemy; emotion rows batched on the hot path",
            "Inference: OpenCV face localisation + DeepFace emotion scores + session smoothing",
        ],
        "One sentence per bullet; expand verbally with one concrete user journey (login → start session → stream).",
    )

    add_bullet_slide(
        prs,
        "Mid-term → final trajectory",
        [
            "Mid-term: desktop-style prototype and feasibility risks (UI stalls, SQLite stability)",
            "Integration phase: JWT + static pages + WebSocket transport + batch persistence tuning",
            "Pre-final: exports, admin read paths, regression passes after CORS/JWT/WebSocket changes",
        ],
        "One minute narrative arc; connects portfolio minutes to the shipped web stack.",
    )

    add_bullet_slide(
        prs,
        "Risks & mitigations (group view)",
        [
            "Real-time load: batch commits + disconnect flush; monitor CPU on lab laptops",
            "Security: bearer tokens, HTTPS/WSS on public host, avoid secrets in logs/markdown",
            "Evidence: dated minutes + reproducible requirements.txt + main.py as single integration spine",
        ],
        "Use as a buffer slide if Part 0 runs short; otherwise fold into Q&A.",
    )

    # ----- Part 1: Dongdong ~8 min -----
    add_section(
        prs,
        "Part 1 — Team lead (Dongdong Mao)",
        "~8 minutes",
        "This is the deepest technical block: architecture, integration decisions, and traceability to repository files.",
    )

    add_bullet_slide(
        prs,
        "Leadership scope",
        [
            "CN10 team lead: planning cadence, backlog prioritisation, integration ownership",
            "Single runnable vertical slice: static UI ↔ API ↔ inference ↔ database",
            "Repository discipline: naming, handover notes, regression after auth/CORS/WebSocket changes",
        ],
        "Give 2–3 concrete examples of decisions (e.g., batch commit size, smoothing buffer) and why they mattered.",
    )

    add_bullet_slide(
        prs,
        "Backend architecture",
        [
            "FastAPI application entry: routes for auth, sessions, exports, patients, questionnaires",
            "OAuth2 password flow + JWT bearer dependency on protected routes",
            "WebSocket endpoint accepts JSON with base64 JPEG; returns annotated frame + emotion JSON",
        ],
        "Point to main.py as the integration hub; mention where SessionLocal is opened/closed to avoid leaks.",
    )

    add_bullet_slide(
        prs,
        "Real-time & persistence",
        [
            "active_sessions map + queue for session lifecycle (integration pattern from codebase)",
            "Batch insert EmotionRecord rows (mitigate bursty writes under load)",
            "WebSocketDisconnect path flushes pending rows to avoid silent loss",
        ],
        "Tie each bullet to a risk you actually saw (disconnect, DB lock, UI stall) and how you mitigated it.",
    )

    add_bullet_slide(
        prs,
        "Inference pipeline (high level)",
        [
            "OpenCV Haar cascade for face box; expand ROI; optional MTCNN fallback path",
            "DeepFace.analyze on cropped face with detector_backend='skip' when quality passes",
            "Session-scoped smoothing (buffer votes; disgust heuristics in code comments)",
        ],
        "Keep equations out; show one before/after anecdote (flicker vs stability) if time allows.",
    )

    add_bullet_slide(
        prs,
        "Operator-facing surface",
        [
            "Static mount /static → login.html, index workflow, admin views",
            "Exports: Excel endpoints for all sessions and per-session workbook",
            "Questionnaire + recording hooks integrated with session timeline",
        ],
        "Mention one export validation story (empty workbook vs seeded session).",
    )

    add_bullet_slide(
        prs,
        "Security & operations awareness",
        [
            "JWT + password hashing (PBKDF2 via passlib); rotate SECRET_KEY guidance for production",
            "CORS and HTTPS/WSS considerations when moving from localhost to public host",
            "Operational checklist items: rate limits, log hygiene, retention for screenshots/recordings",
        ],
        "Explicitly state this is coursework hardening, not a penetration-test certificate.",
    )

    add_bullet_slide(
        prs,
        "Evidence chain (what moderators can open)",
        [
            "Web1/main.py — route decorators, WebSocket loop, batch persistence",
            "Web1/static/* — UI flows aligned with API contracts",
            "Web1/requirements.txt — dependency closure for reproducibility",
            "Meeting minutes under portfolio — dated decisions and action owners",
        ],
        "Offer to screen-share the repo tree for 30 seconds if allowed.",
    )

    add_bullet_slide(
        prs,
        "Reflective alignment (brief)",
        [
            "Learning agreement: planned vs delivered (call out one gap honestly)",
            "Critical incidents: WebSocket/DB pressure, mentor feedback on UI stalls",
            "Personal growth: full-stack ownership + peer leadership under deadlines",
        ],
        "Bridge to written Reflective Report; do not read the report aloud—summarise in 45–60 seconds.",
    )

    # ----- Part 2: Menghao -----
    add_section(
        prs,
        "Part 2 — Menghao Luo",
        "~3.3 minutes",
        "Focus on deployment credibility: nginx, TLS, WebSocket upgrade, and repeatable restart procedures.",
    )

    add_bullet_slide(
        prs,
        "Deployment & reverse proxy",
        [
            "nginx template: upstream to uvicorn/FastAPI; TLS termination on the public host",
            "Diagnosed 502 cases: wrong upstream, missing Upgrade / Connection headers for WebSocket",
            "Successful wss path after header fixes (paired with lead integration on app side)",
        ],
        "Tell one debugging narrative with cause → fix → verification command.",
    )

    add_bullet_slide(
        prs,
        "Operations readiness",
        [
            "systemd unit patterns for service restart; journalctl / log tail for viva-style questions",
            "Rehearsed restart + smoke steps before review milestones",
            "Optional GPU acceleration noted and deferred as non-critical for coursework demo",
        ],
        "End with what you would automate next (health checks, blue/green) in one sentence.",
    )

    # ----- Part 3: Yichen -----
    add_section(
        prs,
        "Part 3 — Yichen Ye",
        "~3.3 minutes",
        "Focus on documentation quality: proposal, portfolio index, and archival packaging expectations.",
    )

    add_bullet_slide(
        prs,
        "Documentation & portfolio alignment",
        [
            "Aligned final proposal sections with portfolio index and exemplar structure",
            "Charter / scope iterations coordinated with supervisor feedback",
            "Communications plan vII: minutes continuity, stakeholder update rhythm",
        ],
        "Reference meeting minutes where you chaired or captured actions (rotate roles across term).",
    )

    add_bullet_slide(
        prs,
        "Submission artefacts",
        [
            "Final PDF exports and figure insertion for archival packs",
            "Mid-term structure aligned with high-scoring CN exemplars",
            "Client handover documentation structure mirroring portfolio expectations",
        ],
        "If asked, show one artefact filename pattern from the portfolio zip.",
    )

    # ----- Part 4: Xingli -----
    add_section(
        prs,
        "Part 4 — Xingli Shao",
        "~3.3 minutes",
        "Focus on verification: repeatable tests, cross-browser behaviour, defect reporting, poster logistics.",
    )

    add_bullet_slide(
        prs,
        "Quality assurance",
        [
            "Cross-browser smoke matrix; Safari permission messaging logged",
            "Soak tests (e.g., 10-minute run): memory stable; CPU high on laptop — documented",
            "Questionnaire integration re-runs after datetime parsing fixes",
        ],
        "Pick one defect ticket style example: steps, expected, actual, severity.",
    )

    add_bullet_slide(
        prs,
        "Exports & wrap-up logistics",
        [
            "Export edge-case tests coordinated with lead (merged questionnaire rows scenario)",
            "Defect report finalisation before submission freeze",
            "Poster print order coordination for final showcase",
        ],
        "Close with what evidence you would show if demo fails (screenshots + logs).",
    )

    # ----- Close -----
    add_bullet_slide(
        prs,
        "Group outcomes",
        [
            "Delivered integrated web system with authenticated sessions and real-time inference",
            "Traceable engineering + documentation + QA suitable for coursework moderation",
            "Clear division of labour with dated meeting minutes and action owners",
        ],
        "30-second recap; invite questions.",
    )

    add_bullet_slide(
        prs,
        "Thank you — Q&A",
        [
            "Contact: refer to cover sheet / portfolio index for emails if required",
            f"Demo: {DEMO_URL}",
            "Backup: local archive + recorded walkthrough in portfolio (if present)",
        ],
        "Thank chair; offer to share screen on repo or deployed site.",
    )

    prs.save(str(OUT_FILE))
    print("Wrote", OUT_FILE, "slides=", len(prs.slides))


if __name__ == "__main__":
    main()

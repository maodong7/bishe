# -*- coding: utf-8 -*-
"""
Generate maximally enriched CN10 master summary Word .docx from:
  - Local Web1 (ProjectWeb1.0/Web1)
  - 毕业设计/论文 (personal + group portfolio, meeting minutes, embedded Web1 copy)
Paths are configurable below.
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt

OUT_DIR = Path(r"D:\桌面文档文件勿动！\Desktop\毕业设计\毕设全部总结（ppt生成专用）")
OUT_FILE = OUT_DIR / "CN10_毕设项目全量总结_PPT生成专用.docx"


def _clear_files_in_output_dir(out_dir: Path) -> None:
    """仅删除输出目录下的文件（不删子目录），保存前调用，使目录中只保留即将写入的最新文档。"""
    if not out_dir.is_dir():
        return
    for p in out_dir.iterdir():
        if p.is_file():
            try:
                p.unlink()
            except OSError:
                pass

WEB1 = Path(__file__).resolve().parent
PAPERS = Path(r"D:\桌面文档文件勿动！\Desktop\毕业设计\论文")
FINAL_ARCHIVE = Path(r"D:\桌面文档文件勿动！\Desktop\毕业设计最终项目")
PORTFOLIO = PAPERS / "小组" / "CN10_Final portfolio"
MINUTES_DIR = PORTFOLIO / "Teamwork and Communication" / "Minutes of the Meeting"
PORTFOLIO_WEB1 = PORTFOLIO / "Project Code" / "Web1"
GLOBAL_SPEC = PORTFOLIO_WEB1 / "Global Event System Specification（全局事件系统说明）.txt"
I18N_JS = PORTFOLIO_WEB1 / "static" / "i18n.js"

THESIS_TITLE = "Online Expression Recogniton System"
PRODUCT_TITLE = "Online Expression Recognition System · Psychological Therapy"
PRODUCT_TITLE_ZH_I18N = "在线表情识别系统・心理辅助治疗"

DEMO_URL = "https://qxjc.fun/static/login.html"


def _ea(run, font: str = "宋体"):
    r = run._element
    rpr = r.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), font)


def _style_normal(doc: Document) -> None:
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(10.5)
    try:
        st._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    except Exception:
        pass


def _h(doc: Document, text: str, level: int = 1):
    p = doc.add_heading(text, level=level)
    for r in p.runs:
        r.font.name = "Times New Roman"
        r.font.size = Pt(16 if level == 1 else 13 if level == 2 else 11)
        _ea(r, "黑体" if level <= 2 else "宋体")


def _p(doc: Document, text: str, bold: bool = False):
    p = doc.add_paragraph(text)
    for r in p.runs:
        r.bold = bold
        r.font.size = Pt(10.5)
        r.font.name = "Times New Roman"
        _ea(r, "宋体")


def _table(doc: Document, headers: list[str], rows: list[list[str]], max_rows: int = 200):
    rows = rows[:max_rows]
    if not rows:
        return
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    for j, h in enumerate(headers):
        t.rows[0].cells[j].text = h
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            if j < len(headers):
                t.rows[i].cells[j].text = val
    doc.add_paragraph()


def _read_text(p: Path, limit: int | None = None) -> str:
    if not p.is_file():
        return ""
    t = p.read_text(encoding="utf-8", errors="replace")
    if limit:
        return t[:limit]
    return t


def _docx_excerpt(path: Path, max_paragraphs: int = 80, max_chars: int = 12000) -> str:
    if not path.is_file():
        return f"（文件不存在：{path}）"
    try:
        from docx import Document as D

        d = D(str(path))
    except Exception as e:
        return f"（无法读取：{e}）"
    chunks = []
    n = 0
    for para in d.paragraphs:
        t = (para.text or "").strip()
        if not t:
            continue
        if "\t" in t and re.search(r"\t\d+\s*$", t):
            continue
        chunks.append(t)
        n += 1
        if n >= max_paragraphs:
            break
    s = "\n\n".join(chunks)
    return s[:max_chars] + ("…\n（以下省略）" if len(s) > max_chars else "")


def _list_pdf_inventory() -> list[list[str]]:
    rows = []
    if not PORTFOLIO.is_dir():
        return [["（未找到 Portfolio 目录）", "", ""]]
    for p in sorted(PORTFOLIO.rglob("*.pdf")):
        try:
            rel = p.relative_to(PORTFOLIO)
        except ValueError:
            rel = p.name
        rows.append([str(rel).replace("\\", "/"), str(p.parent.name), f"{p.stat().st_size // 1024} KB"])
    return rows


def _list_portfolio_top_dirs() -> list[list[str]]:
    if not PORTFOLIO.is_dir():
        return []
    rows = []
    for child in sorted(PORTFOLIO.iterdir()):
        if child.is_dir():
            n = sum(1 for _ in child.rglob("*") if _.is_file())
            rows.append([child.name, str(n)])
    return rows


def _parse_meeting(path: Path) -> dict:
    raw = _read_text(path)
    lines = raw.splitlines()
    meta: dict[str, str] = {"file": path.name, "date": "", "chair": "", "minutes": "", "body": ""}
    for line in lines[:30]:
        if line.startswith("Date:"):
            meta["date"] = line[5:].strip()
        elif line.startswith("Chair:"):
            meta["chair"] = line[6:].strip()
        elif line.startswith("Minutes:"):
            meta["minutes"] = line[8:].strip()
    disc: list[str] = []
    actions: list[str] = []
    in_act = False
    for line in lines:
        ls = line.strip()
        if ls.startswith("Action items"):
            in_act = True
            continue
        if in_act and ls:
            if ls.startswith(("Decisions", "Agenda", "Next meeting")):
                in_act = False
            elif ls.upper().startswith("ID\t") or ls.startswith("ID\t"):
                continue
            elif re.match(r"^A\d+\s", ls) or re.match(r"^A\d+\t", ls):
                if len(actions) < 12:
                    actions.append(ls[:200])
    capture = False
    for line in lines:
        ls = line.strip()
        if ls.startswith("Discussion"):
            capture = True
            continue
        if capture and ls.startswith(("Decisions", "Action items")):
            break
        if capture and ls:
            if ls.startswith(("Date:", "Chair:", "Minutes:", "Attendees:")):
                continue
            disc.append(ls[:240])
    meta["body"] = " | ".join(disc[:6]) if disc else ""
    meta["actions"] = "；".join(actions[:10])
    return meta


def _collect_meetings() -> list[dict]:
    out = []
    if not MINUTES_DIR.is_dir():
        return out
    for p in sorted(MINUTES_DIR.glob("Meeting Minutes *.md")):
        out.append(_parse_meeting(p))
    return out


def _list_personal_docx() -> list[list[str]]:
    rows = []
    if not PAPERS.is_dir():
        return rows
    for p in PAPERS.rglob("*.docx"):
        if ".venv" in str(p).lower():
            continue
        try:
            rel = p.relative_to(PAPERS)
        except ValueError:
            rel = Path(p.name)
        rows.append([str(rel).replace("\\", "/"), f"{p.stat().st_size // 1024} KB"])
    rows.sort(key=lambda x: x[0])
    return rows


def _orm_detail_rows() -> list[list[str]]:
    return [
        ["User", "users", "id PK, username unique, email unique, hashed_password, created_at"],
        [
            "RecognitionSession",
            "recognition_sessions",
            "session_id PK, patient_id, start_time, end_time, duration, fps; relations to EmotionRecord, ScreenshotRecord",
        ],
        [
            "EmotionRecord",
            "emotion_records",
            "record_id PK, session_id FK, timestamp, emotion_type, confidence, frame_path, question_id, question_text",
        ],
        [
            "ScreenshotRecord",
            "screenshot_records",
            "screenshot_id PK, session_id FK, image_path, timestamp, emotion_data JSON-ish string",
        ],
        [
            "QuestionEmotion",
            "question_emotions",
            "id PK, session_id FK, question_id, question_text, answer, start_time, end_time, emotion_data",
        ],
        ["RecordingRecord", "recording_records", "recording_id PK, session_id FK, file_path, timestamp"],
        ["Question", "questions", "id PK, question_id unique, text, category, options JSON, is_subjective"],
    ]


def _frontend_modules() -> list[str]:
    return [
        "登录/注册：localStorage 保存 access_token；跳转主界面。",
        "主界面 index：患者 ID 输入、开始/停止会话、截图、录音、音频导出、视频上传分析入口。",
        "双 Canvas：video-feed 展示、video-canvas 离屏捕获；WebSocket 推帧与收帧。",
        "问卷：进度条、上一题/下一题、选项容器动态渲染；编辑问卷 Modal + POST /api/questions。",
        "统计：饼图/折线图（Chart.js）、dominant emotion、FPS 显示。",
        "导出：Excel 按钮；全量/单会话后端 StreamingResponse。",
        "语音识别：状态区 + 文本区 + 保存到会话。",
        "AI 分析 Modal：患者选择、可编辑 prompt、调用 POST /api/ai/analyze 展示报告。",
        "后台 admin：患者列表、删除单患者/全部、会话详情拉取。",
        f"i18n：{I18N_JS.name} — 中英 PACK，login/index.admin 中文产品名「{PRODUCT_TITLE_ZH_I18N}」；window 派发 app:langchange。",
        "HTML <title>：index 为 Onile…（拼写笔误）；与 i18n 中文品牌「在线…」并存，答辩时可说明将统一为 Online + Recognition。",
    ]


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    _clear_files_in_output_dir(OUT_DIR)
    doc = Document()
    _style_normal(doc)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("CN10 毕业设计 — 项目与论文全量总览")
    r.bold = True
    r.font.size = Pt(20)
    _ea(r, "黑体")
    t2 = doc.add_paragraph()
    t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = t2.add_run("（PPT / 终期答辩 / 归档交叉索引 · 自动汇总稿）")
    r2.font.size = Pt(12)
    _ea(r2, "宋体")
    doc.add_paragraph()

    _p(
        doc,
        "本文档由脚本根据本地 Web1 工程、毕业设计论文目录及 CN10 Final portfolio 扫描生成，"
        "用于一次性掌握「代码事实 + 过程证据 + 论文交付物位置」。PDF 正文未做 OCR 内嵌，"
        "请在 Word 中配合超链接或资源管理器打开对应路径。",
        bold=True,
    )

    _h(doc, "0. 标准题目、产品名与演示入口", 1)
    _p(doc, f"毕业论文标准题目：{THESIS_TITLE}", bold=True)
    _p(doc, f"完整网站系统名称（英文）：{PRODUCT_TITLE}", bold=True)
    _p(doc, f"i18n 中文品牌（登录页等）：{PRODUCT_TITLE_ZH_I18N}", bold=True)
    _p(doc, f"线上演示入口：{DEMO_URL}")
    _p(doc, f"本地完整归档（用户给定）：{FINAL_ARCHIVE}")
    _p(doc, f"论文根目录：{PAPERS}")
    _p(doc, f"主工程（开发源）：{WEB1}")
    _p(doc, f"Portfolio 内嵌代码副本：{PORTFOLIO_WEB1}（与主工程应保持一致；以 Git/组长定稿为准）")

    _h(doc, "1. Portfolio 一级目录与文件规模", 1)
    _table(doc, ["子目录", "文件数（递归近似）"], _list_portfolio_top_dirs())

    _h(doc, "2. Portfolio 内 PDF 交付物清单", 1)
    pdf_rows = _list_pdf_inventory()
    _p(doc, f"下列为 CN10_Final portfolio 下全部 PDF（共 {len(pdf_rows)} 个），相对路径便于答辩逐一点名。")
    _table(doc, ["相对路径", "上级文件夹", "大小"], pdf_rows[:80])
    if len(pdf_rows) > 80:
        _p(doc, f"（另有 {len(pdf_rows) - 80} 个 PDF 未展开，请直接在 {PORTFOLIO} 下检索 *.pdf。）")

    _h(doc, "3. 个人与小组 Word 文档索引", 1)
    _table(doc, ["相对 论文/ 的路径", "大小"], _list_personal_docx())

    _h(doc, "4. 组长 Reflective Report（BS_001）正文摘录", 1)
    reflective = PAPERS / "个人" / "Dongdong Mao" / "10356_080901_22H034160115_BS_001.docx"
    _p(doc, f"源文件：{reflective}")
    _p(doc, _docx_excerpt(reflective, 55, 10000))

    _h(doc, "5. 其他成员 Word 文档（仅路径说明）", 1)
    _p(doc, "除 BS_001 外的 论文/个人 下 docx：")
    skip = "10356_080901_22H034160115_BS_001.docx"
    personal_root = PAPERS / "个人"
    if personal_root.is_dir():
        for p in sorted(personal_root.rglob("*.docx")):
            if skip in p.name:
                continue
            try:
                rel = p.relative_to(PAPERS)
            except ValueError:
                rel = p
            _p(doc, f"· {rel}")
    else:
        _p(doc, "（未找到 论文/个人 目录）")

    _h(doc, "6. 会议纪要全表（CN10，Markdown 源）", 1)
    meets = _collect_meetings()
    _table(
        doc,
        ["文件名", "Date", "Chair", "Minutes", "讨论摘录", "行动项摘录"],
        [
            [
                m["file"],
                m.get("date", ""),
                m.get("chair", ""),
                m.get("minutes", ""),
                m.get("body", "")[:300],
                m.get("actions", "")[:400],
            ]
            for m in meets
        ],
    )

    _h(doc, "7. 全局事件系统说明（Portfolio Web1 文本节选）", 1)
    gs_limit = 15000
    gs = _read_text(GLOBAL_SPEC, gs_limit) if GLOBAL_SPEC.is_file() else ""
    _p(doc, gs if gs else "（未找到 Global Event Specification 文本）")
    if GLOBAL_SPEC.is_file() and GLOBAL_SPEC.stat().st_size > gs_limit:
        _p(doc, f"（原文超过 {gs_limit} 字符，已截断；完整版见 {GLOBAL_SPEC}）")

    _h(doc, "7b. 毕业设计最终项目目录（若已归档）", 1)
    if FINAL_ARCHIVE.is_dir():
        sub = []
        for c in sorted(FINAL_ARCHIVE.iterdir()):
            sub.append(c.name + (" /" if c.is_dir() else ""))
        _p(doc, "顶层条目：" + "；".join(sub[:40]))
        if len(sub) > 40:
            _p(doc, f"… 共 {len(sub)} 项。")
    else:
        _p(doc, f"（目录不存在或未挂载：{FINAL_ARCHIVE}）")

    _h(doc, "8. 技术栈（requirements.txt 双份对照）", 1)
    for label, base in [("主工程 Web1", WEB1), ("Portfolio 副本", PORTFOLIO_WEB1)]:
        req = base / "requirements.txt"
        _p(doc, f"{label}: {req}", bold=True)
        _p(doc, _read_text(req, 2000) if req.is_file() else "（缺失）")

    _h(doc, "9. 后端模块与路由（main.py 事实）", 1)
    api = [
        ["GET", "/", "RedirectResponse → /static/login.html"],
        ["POST", "/api/auth/register", "UserCreate 校验、PBKDF2 哈希、写 users"],
        ["POST", "/api/auth/token", "OAuth2PasswordRequestForm → JWT"],
        ["GET", "/api/auth/me", "Depends(get_current_user)"],
        ["POST/GET", "/api/questions", "问卷批量写/读"],
        ["POST", "/api/sessions", "RecognitionSession + active_sessions + queue"],
        ["POST", "/api/sessions/{id}/end", "结束、删 active_sessions"],
        ["WS", "/ws/{session_id}", "JSON 帧、批量 EmotionRecord、disconnect flush"],
        ["GET", "/api/sessions/{id}/emotions", "列表序列化"],
        ["POST", "/api/sessions/{id}/screenshot", "multipart 存盘 + ScreenshotRecord"],
        ["POST", "/api/sessions/{id}/question-emotions", "问卷情绪 JSON"],
        ["POST", "/api/sessions/{id}/recordings", "录音元数据"],
        ["POST", "/api/sessions/{id}/speech-recognition", "语音识别结果"],
        ["GET", "/api/sessions/export/excel", "pandas 多 sheet"],
        ["GET", "/api/sessions/{id}/export/excel", "单会话导出"],
        ["GET/DELETE", "/api/patients*", "列表/详情/删除（存在重复 GET 定义，应合并）"],
        ["POST", "/api/video/process", "上传 mp4 抽帧批量写 emotion_records"],
        ["POST", "/api/ai/analyze", "DeepSeek Chat Completions — 密钥须环境变量"],
    ]
    _table(doc, ["类型", "路径", "说明"], api)

    _h(doc, "10. 数据库 ORM 字段级摘要", 1)
    _table(doc, ["模型", "表", "字段要点"], _orm_detail_rows())

    _h(doc, "11. 推理链与平滑", 1)
    _p(
        doc,
        "OpenCV Haar + 均衡灰度 → 最大脸外扩 ROI → check_face_quality（Laplacian/亮度/对比度）"
        " → DeepFace.analyze（skip 或全帧 MTCNN 回退）→ get_smoothed_emotion（FIFO + disgust 启发式）。",
    )
    _p(doc, "视频离线处理：recognize_emotion(..., session_id=None) 不使用会话平滑；按帧间隔抽样。")

    _h(doc, "12. 前端功能模块（index / admin / login）", 1)
    for line in _frontend_modules():
        _p(doc, "· " + line)

    _h(doc, "13. 运行、部署与运维关键词", 1)
    _p(doc, "本地：uvicorn main:app --reload（具体参数以小组 README 为准）。")
    _p(doc, "公网：nginx TLS 反代；WebSocket Upgrade/Connection；Menghao 纪要中 502/upstream 排查；systemd + journalctl 演练。")
    _p(doc, "运行时目录：screenshots/、videos/、emotion_analysis.db 于工作目录。")

    _h(doc, "14. 安全、伦理与技术债", 1)
    _table(
        doc,
        ["项", "说明"],
        [
            ["JWT SECRET", "默认值仅课程用；生产环境变量"],
            ["CORS *", "演示便利；生产收紧"],
            ["外部 LLM", "密钥禁止入库与明文投屏；轮换"],
            ["患者数据", "演示伪名；删除接口高危需确认 UX"],
            ["医疗边界", "情绪标签辅助；非诊断"],
            ["代码债", "GET /api/patients 双定义；print 日志；index title Onile 拼写"],
        ],
    )

    _h(doc, "15. 答辩时间结构（与 PPT 一致）", 1)
    _table(
        doc,
        ["段", "时长", "讲者"],
        [
            ["项目总览", "~2 min", "开场"],
            ["组长技术与集成", "~8 min", "Dongdong Mao"],
            ["组员三人", "~10 min 合计", "Menghao Luo → Yichen Ye → Xingli Shao"],
        ],
    )

    _h(doc, "16. Client Handover 与其他 Portfolio 子项（仅路径）", 1)
    for name in [
        "Client Handover/Client Handover Document.pdf",
        "Teamwork and Communication/Team Contract.pdf",
        "Development Activities & Quality Assurance/Designs -UI, Architecture, and Persistent Data.pdf",
        "Research & Upskilling/Research/Project Development Method Research.pdf",
    ]:
        p = PORTFOLIO / name.replace("/", "\\")
        if not p.exists():
            p = PORTFOLIO / Path(name)
        _p(doc, f"· {p}" if p.exists() else f"· （未找到）{name}")

    _h(doc, "17. 修订记录", 1)
    _p(doc, "2026-05-07：Word 初版。2026-05-07：极致丰富版 — 自动并入 Portfolio PDF 清单、会议纪要表、Reflective 摘录、全局事件说明、双份 requirements、前后端模块详单、路径索引。")

    doc.save(str(OUT_FILE))
    print("Wrote", OUT_FILE, "paragraphs~", len(doc.paragraphs))


if __name__ == "__main__":
    main()

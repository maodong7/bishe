# -*- coding: utf-8 -*-
"""
答辩备战：按每页 PPT 准备「导师问什么是 X」能答深一层。
非照搬幻灯片条目，非演讲稿。
"""
from __future__ import annotations

import os
from datetime import date

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

OUT_DIR = r"D:\桌面文档文件勿动！\Desktop\毕业设计\展示文稿ppt"
OUT_FILE = os.path.join(OUT_DIR, "Dongdong_Mao_答辩个人贡献_备战稿.docx")
OUT_FILE_ALT = os.path.join(OUT_DIR, "Dongdong_Mao_PPT备战稿.docx")

_BM_ID = 0


def _next_bm_id() -> int:
    global _BM_ID
    _BM_ID += 1
    return _BM_ID


def bookmark_section(doc: Document, name: str, title: str) -> None:
    p = doc.add_paragraph()
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(13)
    bid = _next_bm_id()
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    p._p.insert(0, start)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    p._p.append(end)
    p.paragraph_format.space_before = Pt(14)


def link_para(doc: Document, text: str, anchor: str) -> None:
    p = doc.add_paragraph(style="List Bullet")
    h = OxmlElement("w:hyperlink")
    h.set(qn("w:anchor"), anchor)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    h.append(r)
    p._p.append(h)


def concept(doc: Document, name: str, define: str, deeper: str, project: str, follow: str) -> None:
    h = doc.add_paragraph()
    h.add_run(name).bold = True
    for label, text in [
        ("是什么", define),
        ("再深一层", deeper),
        ("在我们项目里", project),
        ("老师可能追问", follow),
    ]:
        p = doc.add_paragraph()
        p.add_run(f"{label}：").bold = True
        p.add_run(text)


def main() -> None:
    global _BM_ID
    _BM_ID = 0
    os.makedirs(OUT_DIR, exist_ok=True)
    doc = Document()
    s = doc.styles["Normal"]
    s.font.name = "宋体"
    s.font.size = Pt(11)
    s.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    doc.add_paragraph("跳转（Ctrl+单击）")
    for anchor, label in [
        ("P1", "对应 PPT① My Role"),
        ("P2", "对应 PPT② Backend & API"),
        ("P3", "对应 PPT③ WebSockets"),
        ("P4", "对应 PPT④ SQLAlchemy"),
        ("P5", "对应 PPT⑤ AI Pipeline"),
        ("P6", "对应 PPT⑥ Key Features"),
        ("P7", "对应 PPT⑦ Challenges"),
        ("PX", "临场串联"),
    ]:
        link_para(doc, label, anchor)
    doc.add_paragraph(f"Web1/main.py　{date.today()}")

    # ===================== PPT ① =====================
    bookmark_section(doc, "P1", "对应 PPT① My Role & Key Responsibilities — 概念盘问")

    concept(
        doc,
        "Team Lead（组长）",
        "不只会写代码，还要定优先级、拆任务、对齐导师中期意见、保证小组能交出一个能演示的系统。",
        "技术上你仍做集成；管理上体现为 backlog、分工文档、联调会议、风险（延期、环境、接口不一致）由你推动关闭。",
        "CN10 里你同时是 Web1 代码主责；叶/罗/邵有各自交付物，答辩别把他们工作说成你写的。",
        "问：组长和项目经理区别？答：课程规模小，你亲自扛集成；PM 可不动代码，你必须能指 main.py。",
    )
    concept(
        doc,
        "Full-Stack Developer",
        "同一个人能覆盖「用户能点的界面」+「服务器逻辑」+「数据存哪」+「和 AI 模块怎么接」。",
        "不是精通所有前端框架，而是端到端打通：静态页调你的 API，你的 API 调数据库和 DeepFace。",
        "你：static/*.html + main.py 全链路；没有单独 Node 后端。",
        "问：前端用什么框架？答：原生 HTML/JS + Bootstrap，重点是和 WebSocket/REST 协议对齐。",
    )
    concept(
        doc,
        "System Architecture（系统架构）",
        "系统由哪些大块组成、块之间怎么传数据、谁依赖谁；是设计图，不是某一行代码。",
        "本系统是四层：浏览器 → FastAPI（REST+WS）→ 推理（OpenCV+DeepFace）→ SQLite。选单体是因为课程要可演示、可排错。",
        "mount /static；main.py 一个进程同时处理 HTTP 和 WebSocket。",
        "问：为什么不分微服务？答：用户量与工期不需要；单体集成快，见 PPT⑦。",
    )
    concept(
        doc,
        "FastAPI（本页只点名时）",
        "见 PPT② 专节；此处若被问：Python 写的 Web 框架，用来注册 URL、收请求、返回 JSON、支持 WebSocket。",
        "选它因为要做实时推流 + 自动 API 文档；详见②。",
        "main.py 里 app = FastAPI() 那一套。",
        "不要在这页展开过长，可说「后端技术栈在下一页细讲」。",
    )
    concept(
        doc,
        "WebSocket（本页只点名时）",
        "浏览器和服务器之间保持一条长连接，双方可随时发消息；适合连续传摄像头画面。",
        "和 HTTP 轮询对比：轮询是反复问「有新图吗」，WebSocket 是管道一直开着。详见 PPT③。",
        "/ws/{session_id}",
        "若问和 HTTP 区别：HTTP 一问一答；WS 持久双向。",
    )
    concept(
        doc,
        "Facial recognition / Emotion analysis（本页统称）",
        "人脸识别：在画面里找到脸的位置；情绪分析：根据面部表情判断属于哪类情绪（如 happy、sad）。",
        "两步串起来：先找脸再判情绪；判的是表情类别，不是读心，不能当医学诊断。",
        "recognize_emotion() 一条龙；见 PPT⑤ 四步拆开讲。",
        "问：准确率多少？答：依赖光照与 DeepFace 预训练；我们加了质量门控和平滑提升可用性，未做严格 benchmark。",
    )
    concept(
        doc,
        "Project backlog",
        "待办列表：要做什么、先做什么、谁负责；把导师意见变成可打勾的工程项。",
        "例：中期「UI 卡」→ 批量写库；「DB 不稳」→ 控制 commit；「流程复杂」→ 简化 static 操作步骤。",
        "体现在会议记录和你的 Individual Log，不是空话。",
        "问：怎么排优先级？答：先打通演示主路径（登录→会话→WS→有情绪数据），再加导出等。",
    )

    # ===================== PPT ② =====================
    bookmark_section(doc, "P2", "对应 PPT② Backend Architecture & API Design — 概念盘问")

    concept(
        doc,
        "FastAPI",
        "用 Python 写网站/接口的框架：你定义「访问某个网址时执行哪段函数」，框架负责解析请求、返回响应。",
        "底层基于 Starlette，支持 async（异步）和类型提示；自带 OpenAPI，访问 /docs 可看到所有接口并试调。",
        "main.py：@app.get/post/websocket；uvicorn main:app 启动。答辩可开 /docs 证明接口存在。",
        "问：和 Flask/Django？答：要原生 WebSocket+自动文档+异步，Flask 需加插件，Django 偏重全站；课程选 FastAPI 省集成时间。",
    )
    concept(
        doc,
        "REST / RESTful API",
        "一种设计风格：用 URL 表示资源（如 /api/sessions），用 HTTP 方法表示动作（GET 查、POST 建、DELETE 删），"
        "每次请求自带完整信息，服务器默认不记住上一次（无状态）。",
        "适合登录、建会话、查列表、导出文件等「做一次算一次」的操作；不适合每秒十几次传大图。",
        "register、token、sessions、patients、questions、export/excel 都是 REST。",
        "问：REST 有状态吗？答：无状态；会话状态在数据库和 active_sessions，不靠 HTTP cookie session 粘住。",
    )
    concept(
        doc,
        "API / Endpoint（接口、端点）",
        "前后端约定好的「调用地址 + 入参 + 返回 JSON」；一个 URL+方法就是一个 endpoint。",
        "例如 POST /api/auth/token 入参用户名密码，出参 access_token。",
        "PPT 上 Auth/Session/Data 三组就是按业务分的 endpoint 列表。",
        "问：怎么知道有哪些接口？答：看 /docs 或 main.py 里 @app 装饰器。",
    )
    concept(
        doc,
        "OAuth2",
        "行业标准的授权框架：规定「怎么安全地拿访问令牌」，常见模式之一是用户名密码换 token（Password Flow）。",
        "不是另一种数据库；是一套流程规范，FastAPI 用 OAuth2PasswordRequestForm 接登录表单。",
        "POST /api/auth/token，表单字段 username、password。",
        "问：OAuth2 和 JWT 关系？答：OAuth2 管「怎么登录拿令牌」；JWT 是令牌的一种实现格式。",
    )
    concept(
        doc,
        "JWT（JSON Web Token）",
        "一串编码后的字符串，里面含用户信息（如用户名）和过期时间，用服务端密钥签名，防篡改。",
        "登录后前端保存 token，以后请求 Header 带 Authorization: Bearer <token>；服务端解码验签知是谁，"
        "不必在服务器内存维护登录表（无状态，易扩展）。",
        "create_access_token；get_current_user 解码；SECRET_KEY + HS256。",
        "问：JWT 存在哪？答：前端 localStorage（演示）；生产要考虑 XSS 窃取风险。",
    )
    concept(
        doc,
        "Authentication（认证） vs Authorization（授权）",
        "认证：证明你是谁（登录）；授权：证明你能做这件事（例如只有登录才能建会话）。",
        "本项目认证=账号密码+JWT；授权=Depends(get_current_user) 挂在路由上，未登录返回 401。",
        "register 建用户；token 登录；受保护路由带 current_user。",
        "问：密码怎么存？答：只存哈希不存明文；见 PPT⑥ passlib。",
    )
    concept(
        doc,
        "OpenAPI / Swagger（自动文档）",
        "描述所有 API 路径、参数、返回格式的规范；FastAPI 根据代码自动生成可交互网页 /docs。",
        "答辩时用来证明「接口不是画的」；老师点 Try it out 就能调 register、token。",
        "访问 http://localhost:8000/docs",
        "问：文档和代码会不一致吗？答：自动生成，改路由后重启即更新；以代码为准。",
    )
    concept(
        doc,
        "WebSocket Endpoint /ws/{session_id}",
        "见 PPT③；此处记住：它是和 REST 并列的第二种入口，专管实时传帧，路径里的数字是 session_id。",
        "必须先 POST /api/sessions 创建会话，否则连不上。",
        "websocket_endpoint ~664 行。",
        "若在本页被问：答「实时通道，细节在 WebSocket 那页」。",
    )
    concept(
        doc,
        "Asynchronous（异步）",
        "程序在等待网络/磁盘时不傻等，可以去处理别的连接；FastAPI 的 async def 路由和 await 收 WS 消息为此服务。",
        "对多用户同时连 WS 更友好；单演示机仍可能 CPU 被 DeepFace 占满，异步解决不了算力瓶颈。",
        "async def websocket_endpoint；await websocket.receive_text()。",
        "问：异步等于多线程？答：不等同；是协程模型，I/O 等待时切换。",
    )

    # ===================== PPT ③ =====================
    bookmark_section(doc, "P3", "对应 PPT③ Real-time Communication with WebSockets — 概念盘问")

    concept(
        doc,
        "WebSocket",
        "在浏览器和服务器之间建立一条长期保持的连接，双方都可以主动发数据；协议在 HTTP 之上「升级」得到。",
        "像打电话一直通着，而不是发短信一问一答；适合聊天、实时行情、本项目实时视频帧。",
        "index.html：new WebSocket('ws://host/ws/'+sessionId)；服务端 accept 后循环 receive/send。",
        "问：和 HTTP 关系？答：先 HTTP 握手再 Upgrade 成 WebSocket；生产用 wss=加密版。",
    )
    concept(
        doc,
        "为什么不用 REST 传每一帧",
        "REST 每传一帧都要：建立请求→带完整 HTTP 头→等响应→结束；每秒多帧会浪费大量时间在重复开销上。",
        "WebSocket 连接建一次，后面只传消息体（JSON+base64 图），延迟和 CPU 都更低。",
        "帧走 WS；登录建会话走 REST。",
        "问：能用长轮询吗？答：可以但仍笨重；WS 是更合适方案。",
    )
    concept(
        doc,
        "Bidirectional（双向）",
        "客户端可以连续 upload 帧，服务器也可以连续 push 标注后的图，不用等对方问才回。",
        "本项目：上行 image；下行 type:frame + 标注 JPEG + emotions 数组。",
        "send_text / receive_text 交替。",
        "问：单向行不行？答：只要上传不行，必须回标注图给操作者看。",
    )
    concept(
        doc,
        "base64 与 data URL",
        "二进制图片不能直接塞进 JSON 文本，所以把字节编成 base64 字符串；data:image/jpeg;base64, 是浏览器常见前缀。",
        "服务端 split 逗号取后半段 → b64decode → 得到 JPEG 字节 → cv2.imdecode 成图像矩阵。",
        "frame_data['image'].split(',')[1]",
        "问：为何不用二进制 WS 帧？答：实现简单，JSON 一种格式统一带 question_id；代价是体积增大约 33%。",
    )
    concept(
        doc,
        "session_id 与 active_sessions",
        "session_id：一次评估场次的编号；active_sessions：内存字典，记录哪些场次允许开 WebSocket。",
        "防止乱连：未 create_session 的 id 连 WS 会被 1008 关闭；这是幻灯片 Security 的含义。",
        "create_session 写入；end_session 删除；WS 入口检查 in active_sessions。",
        "问：重启服务器 session 丢吗？答：内存字典会丢，需重新建会话；DB 里历史仍在。",
    )
    concept(
        doc,
        "Batch processing（数据库批量写）",
        "把多条写操作攒一批再一次性 commit，减少磁盘 I/O 次数。",
        "SQLite 同步写慢，每帧 commit 会卡住事件循环导致 UI 卡；batch_size=3 是实测折中。",
        "batch_records；满 3 帧 add_all+commit；断开时 flush。",
        "问：丢最后两帧数据？答：最多延迟 2 帧入库，断开时会 flush；画面仍实时回传。",
    )
    concept(
        doc,
        "Low-latency / High-volume（幻灯片用语）",
        "低延迟：从采帧到看到标注图的时间短；高吞吐：单位时间处理很多帧。",
        "瓶颈常在 DeepFace 推理和机器 CPU，不在 WS 本身；优化靠 ROI、JPEG70、批量写库。",
        "见 PPT⑦ Latency。",
        "问：多少 FPS？答：未正式测；演示够用，未做性能测试报告。",
    )

    # ===================== PPT ④ =====================
    bookmark_section(doc, "P4", "对应 PPT④ Database Design with SQLAlchemy — 概念盘问")

    concept(
        doc,
        "Database / 数据库",
        "持久化存储：关掉浏览器、重启服务器，数据还在；本项目用文件型 SQLite。",
        "存用户、每场评估、每条情绪记录，支撑导出 Excel 和事后复盘。",
        "emotion_analysis.db 在项目目录。",
        "问：数据在内存还是硬盘？答：入库在硬盘；active_sessions 在内存是临时准入名单。",
    )
    concept(
        doc,
        "SQL",
        "结构化查询语言，用来建表、插入、查询；ORM 会在底层自动生成 SQL。",
        "你答辩不必背 SQL 语句，但要知 ORM 最终仍是往表里 INSERT/SELECT。",
        "db.query(EmotionRecord).filter(session_id==…)",
        "问：你们手写 SQL 吗？答：主要用 SQLAlchemy ORM，极少裸 SQL。",
    )
    concept(
        doc,
        "SQLAlchemy / ORM",
        "ORM=用 Python 类代表数据库表，用对象操作代替手写 SQL；类属性对应列，实例对应行。",
        "好处：类型清晰、改字段集中、换数据库引擎时改动小。SessionLocal() 是一次「对话」用来 commit/rollback。",
        "class User(Base)、RecognitionSession、EmotionRecord…；engine + SessionLocal。",
        "问：ORM 慢吗？答：对本项目规模可忽略；瓶颈在推理不在 ORM。",
    )
    concept(
        doc,
        "SQLite",
        "嵌入式数据库，整个库一个文件，无需单独装 MySQL 服务；适合单机演示。",
        "并发写能力有限，所以 WS 里用批量 commit；课程足够，生产可换 PostgreSQL。",
        "DATABASE_URL = sqlite:///./emotion_analysis.db；check_same_thread=False。",
        "问：为何不用 MySQL？答：零配置、助教易复现；架构上 ORM 可迁移。",
    )
    concept(
        doc,
        "Model / Table（User、RecognitionSession…）",
        "一张表存一类实体；行是记录，列是字段。",
        "User 存账号；RecognitionSession 是一次评估；EmotionRecord 是时间轴上的情绪点；"
        "Screenshot/QuestionEmotion 等是附件与问卷。",
        "见 main.py ~40–110 行类定义。",
        "问：表之间关系？答：多表通过 session_id 外键挂到 RecognitionSession。",
    )
    concept(
        doc,
        "Foreign Key / Relationship（外键与关系）",
        "外键：子表存父表主键，保证「这条情绪属于哪场会话」；防止孤儿数据。",
        "RecognitionSession 是中心；EmotionRecord.session_id 指向它。",
        "导出时 JOIN 或按 session 查询聚合。",
        "问：删会话情绪怎么办？答：设计上级联或先删子记录；实现以代码为准。",
    )
    concept(
        doc,
        "Data integrity / Traceability",
        "完整性：约束使数据一致；可追溯：能从导出表追到某 session 的每条记录。",
        "答辩含义：不是空话，是「导师能要求你打开 DB 或 Excel 对得上说法」。",
        "Excel 导出按 session。",
        "问：谁保证不篡改？答：课程演示级；生产要审计日志与权限，未全做。",
    )

    # ===================== PPT ⑤ =====================
    bookmark_section(doc, "P5", "对应 PPT⑤ AI/ML Pipeline — 概念盘问")

    concept(
        doc,
        "Computer Vision（计算机视觉）",
        "让计算机从图像/视频里提取信息（脸在哪、表情如何）；本项目用 OpenCV + DeepFace。",
        "流水线：图像矩阵 → 找脸 → 质量判断 → 分类情绪 → 平滑 → 显示。",
        "recognize_emotion 函数。",
        "问：和 NLP 区别？答：输入是图像不是文本。",
    )
    concept(
        doc,
        "OpenCV",
        "开源计算机视觉库；本项目用其读图、画框、Haar 检测、Laplacian 算模糊。",
        "Haar Cascade 是经典人脸检测器，快但对角度光照敏感。",
        "cv2.CascadeClassifier；detectMultiScale；rectangle 画黄框。",
        "问：为何不用 YOLO？答：课程集成周期紧，Haar+DeepFace 够演示；YOLO 可作改进。",
    )
    concept(
        doc,
        "Haar Cascade",
        "用矩形特征+级联分类器在图上滑动找脸；速度快，适合实时第一关。",
        "灰度+equalizeHist 后 detectMultiScale；取最大脸；扩边 15% 当 ROI。",
        "haarcascade_frontalface_default.xml",
        "问：误检怎么办？答：质量门控+DeepFace 只在 ROI 上跑；失败走 MTCNN 兜底。",
    )
    concept(
        doc,
        "MTCNN",
        "深度学习多阶段人脸检测，比 Haar 准但慢；本项目作 Haar 失败时的 fallback。",
        "DeepFace.analyze(..., detector_backend='mtcnn') 全图检测；脸占帧 10%–80% 才采纳。",
        "在 recognize_emotion 的 else 分支。",
        "问：为何两个检测器？答：快+稳：平时 Haar，难例 MTCNN。",
    )
    concept(
        doc,
        "ROI（Region of Interest）",
        "只裁人脸区域送模型，减少计算量并减少背景干扰。",
        "face_region = frame[y:y+h, x:x+w]；DeepFace 对 ROI analyze。",
        "detector_backend='skip' 表示脸已裁好不再检测。",
        "问：不裁剪行吗？答：整图更慢且易误检背景。",
    )
    concept(
        doc,
        "DeepFace",
        "封装多种人脸分析模型（情绪、年龄等）的 Python 库；本项目只用 emotion。",
        "调用预训练网络，输出各情绪概率，取最大为 dominant_emotion；不是你们从零训练。",
        "DeepFace.analyze(..., actions=['emotion'])",
        "问：模型谁训练的？答：库自带预训练；我们做集成与后处理。",
    )
    concept(
        doc,
        "Dominant emotion / Confidence",
        "主导情绪：概率最高的那一类；置信度：该类的分数，越高模型越肯定。",
        "界面显示 happy: 85 即主导 happy、置信 85；过低时平滑可能压掉。",
        "result[0]['dominant_emotion']；emotion[emotion] 分数。",
        "问：七类情绪固定吗？答：跟 DeepFace 输出标签一致，与 Ekman 基本情绪叙事对齐。",
    )
    concept(
        doc,
        "Laplacian variance（质量检测）",
        "用拉普拉斯算子看图像边缘强度；方差小=边缘少=往往模糊。",
        "check_face_quality 里 laplacian_var < 50 判糊；另查亮度、对比度。",
        "过滤糊/暗/过曝帧，避免乱识别。",
        "问：阈值怎么定？答：实验调参，非理论公式；可答根据试演示调整。",
    )
    concept(
        doc,
        "Smoothing heuristic（平滑启发式）",
        "不用新模型，用规则对最近 N 帧结果投票/加权，让显示稳定。",
        "五帧缓冲 EMOTION_HISTORY；disgust 连续出现则加权；解决 flickering。",
        "get_smoothed_emotion()",
        "问：算 AI 还是工程？答：后处理工程技巧，为可用性服务。",
    )
    concept(
        doc,
        "Flickering（情绪闪烁）",
        "标签帧间来回跳，用户不信结果；PPT⑦ 专讲。",
        "根因是单帧分类噪声；平滑是应对手段。",
        "见 get_smoothed_emotion。",
        "问：能 100% 消除吗？答：不能，只能减轻；极端光照仍会变。",
    )

    # ===================== PPT ⑥ =====================
    bookmark_section(doc, "P6", "对应 PPT⑥ Key Features — 概念盘问")

    concept(
        doc,
        "pandas",
        "Python 数据分析库，把表格式数据放在 DataFrame 里，方便筛选、合并、统计。",
        "导出前把 SQL 查出来的多表/多行整理成表格结构，再交给 openpyxl 写文件。",
        "export/excel 路由里使用。",
        "问：和 Excel 关系？答：pandas 处理数据，openpyxl 写 xlsx 文件。",
    )
    concept(
        doc,
        "openpyxl",
        "读写 .xlsx 的库；生成导师能直接打开的 Excel。",
        "把 DataFrame 写入 workbook，可能多 sheet（情绪、问卷等，以代码为准）。",
        "GET /api/sessions/export/excel",
        "问：为何导出 Excel？答：用户习惯表格复盘，比只看实时画面更有审计感。",
    )
    concept(
        doc,
        "python-jose",
        "处理 JWT 的库：encode 生成令牌，decode 验证签名与过期。",
        "与 passlib 分工：jose 管 token，passlib 管密码哈希。",
        "jwt.encode / jwt.decode in create_access_token, get_current_user。",
        "问：JWT 加密吗？答：payload 可 base64 解码见内容，防篡改靠签名不是加密。",
    )
    concept(
        doc,
        "passlib / 密码哈希",
        "密码不以明文存库；哈希=单向变换，登录时比对哈希而非还原密码。",
        "代码用 pbkdf2_sha256（幻灯片写 bcrypt 时改口：都是 passlib 支持的哈希算法）。",
        "get_password_hash / verify_password。",
        "问：能解密密码吗？答：不能，只能验用户输入是否匹配哈希。",
    )
    concept(
        doc,
        "Evidence-based development（可查证开发）",
        "论文/答辩里每个说法能在代码或数据里找到对应，避免「只吹不做」。",
        "指 main.py 函数、/docs 路由、emotion_analysis.db、导出样例 xlsx。",
        "老师问「你做过吗」就打开文件或演示。",
        "问：怎么证明 WebSocket batch？答：指 main.py batch_size=3 行。",
    )

    # ===================== PPT ⑦ =====================
    bookmark_section(doc, "P7", "对应 PPT⑦ Challenges & Solutions — 概念盘问")

    concept(
        doc,
        "Latency（延迟）",
        "从采帧到看到结果的时间；高延迟=画面跟不上真人表情。",
        "来源：推理慢、每帧写库、大图编解码；已用 ROI、batch commit、JPEG70 缓解。",
        "PPT 提 tracking 跳帧—代码未实现，勿说已实现；可说 future work。",
        "问：优化了多少？答：主观演示流畅度改善，无定量 FPS 报告。",
    )
    concept(
        doc,
        "Decouple I/O from processing（解耦）",
        "把慢操作（写磁盘）从快路径（推理+回传）里拆开，别让写盘拖住回传。",
        "批量写库=推理线程先回图，写库攒够再 commit。",
        "batch_size=3",
        "问：异步写库？答：当前是同循环内批量同步 commit，不是独立写线程；够课程用。",
    )
    concept(
        doc,
        "Emotion flickering",
        "见 PPT⑤ smoothing；挑战页强调这是用户体验问题。",
        "解法=get_smoothed_emotion 投票+disgust 规则。",
        "对比演示：关平滑则标签跳，开则稳。",
        "问：调模型行吗？答：换模型成本高；后处理更快。",
    )
    concept(
        doc,
        "Agile / Vertical slice（敏捷、纵向切片）",
        "敏捷：小步交付、经常集成；纵向切片=一次做完一个小功能的全栈（界面→API→库），"
        "而不是先把所有 API 写完再写界面。",
        "顺序：登录→建会话→WS 通→有情绪入库→再导出/问卷。",
        "减少末期联调崩盘；组长价值在此。",
        "问：和瀑布区别？答：瀑布先全设计再全实现；我们中期就要能演示。",
    )
    concept(
        doc,
        "Early integration testing",
        "开发早期就把前后端、WS、DB 接起来测，而不是最后一周才联调。",
        "中期 mentor 反馈 UI 卡后才改 batch，说明集成测试有效。",
        "有测试记录/Individual Log 可提。",
        "问：有自动化测试吗？答：以手工+Swagger 为主；pytest 可作改进。",
    )

    # ===================== 临场 =====================
    bookmark_section(doc, "PX", "临场串联 — 七页合成一个故事")

    doc.add_paragraph(
        "答辩时按 PPT 翻页讲，老师打断问「什么是 XX」时，用上面该页概念块回答；"
        "若问跨页（如「整个系统怎么跑」），用下面 30 秒版。"
    )
    doc.add_paragraph(
        "30 秒版：用户登录拿 JWT → REST 建评估会话 → WebSocket 持续传摄像头帧 → "
        "服务端 OpenCV 找脸、DeepFace 判情绪、平滑后画框回传 → 情绪每几帧写入 SQLite → "
        "结束可 Excel 导出；公网部署与 wss 由队友完成，我负责 Web1 集成与 main.py。"
    )
    doc.add_paragraph("演示：uvicorn → /static/login.html → 建 session → 摄像头 → 黄框与情绪标签。")
    doc.add_paragraph("分工：叶论文/模型；罗部署；邵测试；你 Web1。")
    doc.add_paragraph("代码指认：认证~175｜平滑~250｜质量~304｜识别~352｜会话~619｜WS~664")

    for path in (OUT_FILE, OUT_FILE_ALT):
        try:
            doc.save(path)
            print(path)
            return
        except PermissionError:
            continue
    doc.save(OUT_FILE + ".new.docx")
    print(OUT_FILE + ".new.docx")


if __name__ == "__main__":
    main()

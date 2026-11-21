from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# 创建演示文稿
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

def add_title_slide(prs, title, subtitle=''):
    slide_layout = prs.slide_layouts[6]  # 空白布局
    slide = prs.slides.add_slide(slide_layout)
    
    # 添加标题
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.333), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 51, 102)
    p.alignment = PP_ALIGN.CENTER
    
    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(12.333), Inches(1))
        tf = sub_box.text_frame
        p = tf.paragraphs[0]
        p.text = subtitle
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(102, 102, 102)
        p.alignment = PP_ALIGN.CENTER
    
    return slide

def add_content_slide(prs, title, content_lines, speaker=''):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # 标题
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.333), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 51, 102)
    
    # 演讲者信息
    if speaker:
        speaker_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(12.333), Inches(0.5))
        tf = speaker_box.text_frame
        p = tf.paragraphs[0]
        p.text = f'演讲者: {speaker}'
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(128, 128, 128)
        p.font.italic = True
    
    # 内容
    content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.333), Inches(5.5))
    tf = content_box.text_frame
    tf.word_wrap = True
    
    for i, line in enumerate(content_lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(51, 51, 51)
        p.space_after = Pt(12)
    
    return slide

# ===== 开始页 =====
add_title_slide(prs, 
    'CN10 毕业设计答辩', 
    'Online Expression Recognition System · Psychological Therapy\n在线表情识别系统・心理辅助治疗')

# ===== 第一部分：组长项目总览（2分钟） =====
add_content_slide(prs, 
    '项目总览 - 产品定位',
    [
        '• 项目名称：Online Expression Recognition System · Psychological Therapy',
        '• 中文品牌：在线表情识别系统・心理辅助治疗',
        '• 项目类型：基于实时面部情绪识别的Web应用系统',
        '• 核心功能：实时情绪识别、问卷调查、数据记录与分析、AI辅助报告',
        '• 应用场景：心理辅助治疗、教学演示、情绪监测辅助',
        '• 演示入口：https://qxjc.fun/static/login.html'
    ],
    'Dongdong Mao (组长) - 2分钟')

add_content_slide(prs,
    '项目总览 - 技术架构',
    [
        '• 后端框架：FastAPI + WebSocket',
        '• 数据库：SQLite + SQLAlchemy ORM',
        '• 推理引擎：DeepFace (TensorFlow/Keras)',
        '• 前端技术：原生HTML/JS + Chart.js',
        '• 实时通信：WebSocket 帧传输',
        '• 部署方式：Uvicorn + Nginx 反向代理',
        '• 安全机制：JWT认证 + OAuth2PasswordBearer'
    ],
    'Dongdong Mao (组长)')

add_content_slide(prs,
    '项目总览 - 核心功能模块',
    [
        '1. 用户认证系统：注册/登录/JWT令牌管理',
        '2. 实时情绪识别：摄像头采集 → Haar/MTCNN检测 → DeepFace分析',
        '3. 会话管理：创建/暂停/结束会话，支持断线重连',
        '4. 问卷调查：动态渲染题目，进度跟踪',
        '5. 数据导出：Excel格式导出，支持单会话/全量导出',
        '6. AI分析报告：基于情绪数据的智能分析报告',
        '7. 多语言支持：中英双语切换 (i18n)',
        '8. 后台管理：患者列表、会话详情查看'
    ],
    'Dongdong Mao (组长)')

# ===== 第二部分：组长自己的工作与贡献（8分钟） =====
add_content_slide(prs,
    '我的工作与贡献 - 后端架构设计',
    [
        '• 负责整体后端架构设计与实现',
        '• 设计并实现FastAPI主应用框架 (main.py)',
        '• 构建数据库ORM模型：User, Session, EmotionRecord, Question, Answer等',
        '• 实现RESTful API路由体系：',
        '  - /api/auth/* : 认证相关 (注册/登录/令牌刷新)',
        '  - /api/sessions/* : 会话生命周期管理',
        '  - /api/questions/* : 问卷CRUD',
        '  - /api/ai/analyze : AI分析接口'
    ],
    'Dongdong Mao (组长) - 8分钟')

add_content_slide(prs,
    '我的工作与贡献 - WebSocket实时通信',
    [
        '• 设计并实现WebSocket实时情绪传输通道',
        '• 实现 /ws/{session_id} 端点：',
        '  - 帧接收：浏览器 → 服务器 (JPEG编码, quality=70)',
        '  - 推理处理：Haar定位 → 质量检查 → DeepFace分析',
        '  - 平滑算法：5帧FIFO缓冲区 + DISGUST启发式去抖动',
        '  - 帧发送：标注结果 → 浏览器实时渲染',
        '• 优化策略：三帧批量提交减少同步写压力',
        '• 安全设计：WebSocket连接需验证session_id有效性'
    ],
    'Dongdong Mao (组长)')

add_content_slide(prs,
    '我的工作与贡献 - 数据库与持久化',
    [
        '• 设计SQLite数据库结构 (emotion_analysis.db)',
        '• SQLAlchemy ORM模型实现：',
        '  - User: 用户认证信息 (bcrypt哈希)',
        '  - Session: 会话元数据 (开始/结束时间)',
        '  - EmotionRecord: 情绪记录 (时间戳/类型/置信度)',
        '  - Question/Answer: 问卷系统',
        '• 实现数据导出功能：',
        '  - 单会话导出：GET /api/sessions/{id}/export/excel',
        '  - 全量导出：GET /api/sessions/export/excel',
        '  - 使用pandas + openpyxl生成Excel文件'
    ],
    'Dongdong Mao (组长)')

add_content_slide(prs,
    '我的工作与贡献 - 推理管道优化',
    [
        '• 设计多级推理管道：',
        '  1. OpenCV Haar级联检测器 (主要路径)',
        '  2. 均衡灰度预处理',
        '  3. 最大人脸ROI外扩裁剪',
        '  4. 质量检查：Laplacian/亮度/对比度阈值',
        '  5. DeepFace.analyze (detector_backend="skip")',
        '  6. MTCNN全帧回退 (Haar未命中时)',
        '• 人脸面积比约束：0.10 ~ 0.80',
        '• 置信度阈值：> 40% 才进入平滑缓冲区',
        '• 去抖动策略：DISGUST_*规则减少厌恶-中性闪烁'
    ],
    'Dongdong Mao (组长)')

add_content_slide(prs,
    '我的工作与贡献 - 团队协作与项目管理',
    [
        '• 担任CN10团队组长，负责整体技术方向把控',
        '• 制定并维护工程待办列表 (Engineering Backlog)：',
        '  - 减少热路径同步写入',
        '  - 明确patient_id跨层契约',
        '  - 文档化Haar级联Windows部署路径',
        '• 协调前后端接口对接：',
        '  - 统一路由命名规范',
        '  - WebSocket消息格式约定',
        '  - 静态页面部署方案',
        '• 代码审查与Git版本管理',
        '• 本地集成测试与部署演练'
    ],
    'Dongdong Mao (组长)')

# ===== 第三部分：其他成员工作（10分钟） =====
add_content_slide(prs,
    '成员工作 - 前端开发 (lmh)',
    [
        '• 负责前端页面开发与UI实现',
        '• 实现登录/注册页面 (login.html)',
        '• 开发主功能页面 (index.html)：',
        '  - 双Canvas设计：video-feed展示 + video-canvas离屏捕获',
        '  - WebSocket帧推送与接收处理',
        '  - 实时情绪图表展示 (Chart.js饼图/折线图)',
        '  - 问卷交互界面：进度条、上一题/下一题',
        '  - 截图、录音、音频导出功能',
        '  - AI分析Modal：患者选择、Prompt编辑、报告展示'
    ],
    'lmh - ~3分钟')

add_content_slide(prs,
    '成员工作 - 前端开发 (lmh) 续',
    [
        '• 实现后台管理页面 (admin.html)：',
        '  - 患者列表展示与动态渲染',
        '  - 删除单患者/全部患者功能',
        '  - 会话详情拉取与展示',
        '• 国际化实现 (i18n)：',
        '  - i18n.js 中英语言包开发',
        '  - app:langchange 全局事件系统',
        '  - 跨页面语言一致性维护',
        '• 前端性能优化与用户体验改进'
    ],
    'lmh - ~3分钟')

add_content_slide(prs,
    '成员工作 - 部署与运维 (sxl)',
    [
        '• 负责系统部署与运维工作',
        '• 本地开发环境配置：',
        '  - Uvicorn服务启动与参数调优',
        '  - 依赖管理与requirements.txt维护',
        '• 公网部署方案：',
        '  - Nginx TLS反向代理配置',
        '  - WebSocket Upgrade/Connection支持',
        '  - 502错误排查与upstream优化',
        '• 系统监控与日志管理：',
        '  - systemd服务配置',
        '  - journalctl日志分析'
    ],
    'sxl - ~3.5分钟')

add_content_slide(prs,
    '成员工作 - AI分析与测试 (yyc)',
    [
        '• 负责AI分析功能与系统测试',
        '• AI分析报告模块：',
        '  - 设计POST /api/ai/analyze接口',
        '  - 实现基于情绪数据的智能分析算法',
        '  - 生成结构化分析报告',
        '• 系统测试与质量保证：',
        '  - 功能测试：覆盖登录、会话、识别、导出全流程',
        '  - 性能测试：WebSocket帧率、数据库写入压力',
        '  - 边界测试：低光照、遮挡、多人场景',
        '• 文档编写：',
        '  - Client Handover文档',
        '  - 测试报告与问题跟踪'
    ],
    'yyc - ~3.5分钟')

# ===== 结束页 =====
add_title_slide(prs,
    '谢谢聆听',
    'Questions & Answers\nCN10 Team')

print('PPT生成完成，正在保存...')
prs.save('CN10_毕业设计答辩.pptx')
print('保存成功!')

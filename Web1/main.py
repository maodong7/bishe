from fastapi import FastAPI, WebSocket, WebSocketDisconnect, UploadFile, File, Form, Depends, HTTPException, status
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
import cv2
import numpy as np
import base64
import json
import asyncio
import datetime
import os
import time
import requests
from starlette.concurrency import run_in_threadpool
from deepface import DeepFace
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, ForeignKey, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from pydantic import BaseModel, EmailStr
import threading
import queue
from passlib.context import CryptContext
from jose import JWTError, jwt
from datetime import timedelta

# 数据库配置
DATABASE_URL = "sqlite:///./emotion_analysis.db"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# 认证配置
SECRET_KEY = "your-secret-key-change-in-production"
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 480  # 8 小时，避免长时间停留页面后 token 过期导致接口持续 401

pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/token")

# 数据库模型

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    email = Column(String, unique=True, index=True)
    hashed_password = Column(String)
    created_at = Column(DateTime, default=datetime.datetime.now)

class RecognitionSession(Base):
    __tablename__ = "recognition_sessions"
    session_id = Column(Integer, primary_key=True, index=True)
    patient_id = Column(String, index=True)
    start_time = Column(DateTime)
    end_time = Column(DateTime)
    duration = Column(Float)
    fps = Column(Float)
    emotion_records = relationship("EmotionRecord", back_populates="session")
    screenshot_records = relationship("ScreenshotRecord", back_populates="session")

class EmotionRecord(Base):
    __tablename__ = "emotion_records"
    record_id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("recognition_sessions.session_id"))
    timestamp = Column(DateTime)
    emotion_type = Column(String)
    confidence = Column(Float)
    frame_path = Column(String)
    question_id = Column(Integer, nullable=True)
    question_text = Column(String, nullable=True)
    session = relationship("RecognitionSession", back_populates="emotion_records")

class ScreenshotRecord(Base):
    __tablename__ = "screenshot_records"
    screenshot_id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("recognition_sessions.session_id"))
    image_path = Column(String)
    timestamp = Column(DateTime)
    emotion_data = Column(String)
    session = relationship("RecognitionSession", back_populates="screenshot_records")

class QuestionEmotion(Base):
    __tablename__ = "question_emotions"
    id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("recognition_sessions.session_id"))
    question_id = Column(Integer)
    question_text = Column(String)
    answer = Column(String)
    start_time = Column(DateTime)
    end_time = Column(DateTime)
    emotion_data = Column(String)
    session = relationship("RecognitionSession", backref="question_emotions")

class RecordingRecord(Base):
    __tablename__ = "recording_records"
    recording_id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("recognition_sessions.session_id"))
    file_path = Column(String)
    timestamp = Column(DateTime)
    session = relationship("RecognitionSession", backref="recording_records")

class Question(Base):
    __tablename__ = "questions"
    id = Column(Integer, primary_key=True, index=True)
    question_id = Column(Integer, unique=True, index=True)
    text = Column(String)
    category = Column(String)
    options = Column(String)  # JSON格式存储选项
    is_subjective = Column(Boolean, default=False)

# 创建数据库表
# 只创建不存在的表，保留现有表的数据
Base.metadata.create_all(bind=engine)

# Pydantic模型
class UserBase(BaseModel):
    username: str
    email: EmailStr

class UserCreate(UserBase):
    password: str
    confirm_password: str

class UserLogin(BaseModel):
    username: str
    password: str

class UserResponse(UserBase):
    id: int
    created_at: datetime.datetime
    class Config:
        from_attributes = True

class Token(BaseModel):
    access_token: str
    token_type: str

class TokenData(BaseModel):
    username: str | None = None



class SessionCreate(BaseModel):
    patient_id: str

class SessionResponse(BaseModel):
    session_id: int
    patient_id: str
    start_time: datetime.datetime
    class Config:
        from_attributes = True

class EmotionData(BaseModel):
    emotion: str
    confidence: float

class ScreenshotRequest(BaseModel):
    session_id: int
    image_data: str

# 全局变量
active_sessions = {}
sessions_queue = queue.Queue()

app = FastAPI()

# CORS中间件配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 在生产环境中应该设置具体的域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 认证工具函数
def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)

def get_password_hash(password):
    return pwd_context.hash(password)

def create_access_token(data: dict, expires_delta: timedelta | None = None):
    to_encode = data.copy()
    if expires_delta:
        expire = datetime.datetime.utcnow() + expires_delta
    else:
        expire = datetime.datetime.utcnow() + timedelta(minutes=15)
    to_encode.update({"exp": expire})
    encoded_jwt = jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)
    return encoded_jwt

async def get_current_user(token: str = Depends(oauth2_scheme)):
    credentials_exception = HTTPException(
        status_code=status.HTTP_401_UNAUTHORIZED,
        detail="Could not validate credentials",
        headers={"WWW-Authenticate": "Bearer"},
    )
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username: str = payload.get("sub")
        if username is None:
            raise credentials_exception
        token_data = TokenData(username=username)
    except JWTError:
        raise credentials_exception
    db = SessionLocal()
    user = db.query(User).filter(User.username == token_data.username).first()
    db.close()
    if user is None:
        raise credentials_exception
    return user



# CORS配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 初始化人脸检测器
print("初始化人脸检测器...")
import cv2
import os
import numpy as np
from deepface import DeepFace

# 启动时就加载情绪模型：避免第一次识别等 4~5 秒，也省掉每帧 analyze() 的额外开销
print("加载情绪识别模型...")
EMOTION_MODEL = DeepFace.build_model("Emotion", task="facial_attribute")
from deepface.models.demography.Emotion import labels as EMOTION_LABELS
print("情绪识别模型加载完成")
print("人脸检测器初始化完成")

# 加载OpenCV人脸检测器（更精确的人脸定位）
# 优先使用 OpenCV 自带的 Haar Cascade，避免依赖固定盘符路径
cascade_path = os.path.join(cv2.data.haarcascades, 'haarcascade_frontalface_default.xml')
if not os.path.exists(cascade_path):
    cascade_path = 'C:/temp/haarcascades/haarcascade_frontalface_default.xml'
face_cascade = cv2.CascadeClassifier(cascade_path)
if face_cascade.empty():
    print(f"警告: 无法加载 Haar Cascade 文件: {cascade_path}")
else:
    print("OpenCV人脸检测器加载完成")

# 全局变量：用于情绪平滑处理
EMOTION_BUFFER_SIZE = 7  # 缓冲器大小，加权投票窗口（约 0.35 秒）
EMOTION_HISTORY = {}  # 每个session_id对应一个情绪缓冲

# 实时识别性能优化参数
DETECT_SCALES = (0.6, 1.0)  # 多尺度检测：Haar 对缩放比例很敏感，单一尺度会漏掉特写镜头的大脸
FACE_CACHE_TTL = 0.15   # 人脸位置缓存有效期（秒），减少重复检测
FACE_CACHE = {}         # session_id -> (box, timestamp)
EMOTION_CACHE = {}      # session_id -> (emotion, confidence, timestamp)

# 情绪平滑处理：置信度平方加权多帧投票 + 迟滞，抑制单帧误判和标签来回跳变
STABLE_EMOTION = {}  # session_id -> 当前稳定情绪

def get_smoothed_emotion(session_id, new_emotion, new_confidence):
    buffer = EMOTION_HISTORY.setdefault(session_id, [])
    buffer.append((new_emotion, new_confidence))

    # 保持缓冲器大小
    if len(buffer) > EMOTION_BUFFER_SIZE:
        buffer.pop(0)

    if not buffer:
        return new_emotion, new_confidence

    # 置信度平方加权：低置信帧几乎不参与投票；越新的帧权重略高
    weights = {}
    for i, (emo, conf) in enumerate(buffer):
        weights[emo] = weights.get(emo, 0.0) + (max(float(conf), 1.0) ** 2) * (1.0 + 0.1 * i)

    best_emotion = max(weights, key=weights.get)

    # 迟滞：新情绪必须明显压过当前稳定情绪才允许切换
    stable = STABLE_EMOTION.get(session_id)
    if stable and best_emotion != stable:
        if weights.get(best_emotion, 0.0) < weights.get(stable, 0.0) * 1.35:
            best_emotion = stable
    STABLE_EMOTION[session_id] = best_emotion

    matched = [float(c) for e, c in buffer if e == best_emotion]
    avg_conf = sum(matched) / len(matched) if matched else float(new_confidence)
    confidence = avg_conf * 0.7 + float(new_confidence) * 0.3

    return best_emotion, confidence

# 人脸质量检测函数 - 过滤模糊或过暗的图像
def check_face_quality(face_region):
    try:
        # 转换为灰度图
        if len(face_region.shape) == 3:
            gray = cv2.cvtColor(face_region, cv2.COLOR_BGR2GRAY)
        else:
            gray = face_region

        # 计算Laplacian梯度（模糊度检测）
        laplacian_var = cv2.Laplacian(gray, cv2.CV_64F).var()

        # 计算平均亮度
        mean_brightness = np.mean(gray)

        # 计算对比度
        contrast = np.std(gray)

        # 质量阈值
        BLUR_THRESHOLD = 30.0  # Laplacian方差阈值（摄像头画面普遍偏软，阈值过高会大量丢帧）
        MIN_BRIGHTNESS = 30    # 最小亮度
        MAX_BRIGHTNESS = 220   # 最大亮度
        MIN_CONTRAST = 20      # 最小对比度

        is_blurry = laplacian_var < BLUR_THRESHOLD
        is_too_dark = mean_brightness < MIN_BRIGHTNESS
        is_too_bright = mean_brightness > MAX_BRIGHTNESS
        is_low_contrast = contrast < MIN_CONTRAST

        quality_ok = not (is_blurry or is_too_dark or is_too_bright or is_low_contrast)

        if not quality_ok:
            reasons = []
            if is_blurry: reasons.append(f"模糊(laplace={laplacian_var:.1f})")
            if is_too_dark: reasons.append(f"过暗(brightness={mean_brightness:.1f})")
            if is_too_bright: reasons.append(f"过曝(brightness={mean_brightness:.1f})")
            if is_low_contrast: reasons.append(f"低对比度(contrast={contrast:.1f})")
            print(f"人脸质量不合格: {', '.join(reasons)}")

        return quality_ok, {
            'blur': laplacian_var,
            'brightness': mean_brightness,
            'contrast': contrast
        }
    except Exception as e:
        print(f"质量检测错误: {e}")
        return True, {}

# 情绪识别函数 - 先用OpenCV检测人脸位置，再用DeepFace分析情绪
def detect_face_fast(frame, session_id=None):
    """多尺度 Haar 检测 + 位置缓存：命中即停，兼顾速度与不漏检特写大脸"""
    cached = FACE_CACHE.get(session_id) if session_id is not None else None
    if cached and time.time() - cached[1] < FACE_CACHE_TTL:
        return cached[0]

    gray_full = None
    box = None
    for scale in DETECT_SCALES:
        if scale == 1.0:
            if gray_full is None:
                gray_full = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            gray = gray_full
        else:
            small = cv2.resize(frame, (0, 0), fx=scale, fy=scale, interpolation=cv2.INTER_LINEAR)
            gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)
        gray = cv2.equalizeHist(gray)

        # 不限制 maxSize：原先 240 的半分辨率上限会把特写镜头（大脸）整体漏检
        faces = face_cascade.detectMultiScale(
            gray,
            scaleFactor=1.1,
            minNeighbors=5,
            minSize=(30, 30)
        )
        if len(faces) == 0:
            continue
        x, y, fw, fh = max(faces, key=lambda f: f[2] * f[3])
        inv = 1.0 / scale
        box = (int(x * inv), int(y * inv), int(fw * inv), int(fh * inv))
        break

    if box is None:
        if session_id is not None:
            FACE_CACHE.pop(session_id, None)
        return None

    if session_id is not None:
        FACE_CACHE[session_id] = (box, time.time())
    return box


def recognize_emotion(frame, session_id=None):
    try:
        h, w = frame.shape[:2]

        # 第一步：快速人脸检测
        face_location = detect_face_fast(frame, session_id)

        emotion = "neutral"
        confidence = 30.0

        if face_location:
            x, y, fw, fh = face_location

            # 以人脸中心裁出正方形区域：模型输入是 48x48，非正方形裁切会把五官拉伸变形，
            # 是之前误判偏多的主要原因；1.18 倍边长可包含眉毛与嘴角。
            side = min(int(max(fw, fh) * 1.18), w, h)
            cx = x + fw // 2
            cy = y + fh // 2
            new_x = max(0, min(cx - side // 2, w - side))
            new_y = max(0, min(cy - side // 2, h - side))

            face_location = (new_x, new_y, side, side)

            # 提取人脸区域
            face_region = frame[new_y:new_y+side, new_x:new_x+side]

            quality_ok = False
            if side >= 24:
                quality_ok, quality_info = check_face_quality(face_region)

            if quality_ok:
                # 第二步：直接用已加载的情绪模型推理（比 DeepFace.analyze 每帧快约 15ms）
                try:
                    # 关键：该模型训练时输入是 0~1 的灰度图，deepface 封装里漏了 /255，
                    # 直接喂 0~255 会让 softmax 饱和到 1.0 并输出随机标签，这是精度差的主因
                    img = face_region.astype(np.float32) / 255.0
                    probs = np.asarray(EMOTION_MODEL.predict(img)).reshape(-1)
                    # 水平翻转 TTA：两次推理取平均，抵消单侧光照和头部轻微偏转带来的偏差
                    flipped = np.asarray(
                        EMOTION_MODEL.predict(cv2.flip(img, 1))
                    ).reshape(-1)
                    probs = (probs + flipped) * 0.5
                    total = float(probs.sum()) or 1.0
                    probs = probs / total
                    order = np.argsort(probs)[::-1]
                    idx = int(order[0])
                    top1 = float(probs[idx])
                    top2 = float(probs[int(order[1])]) if len(order) > 1 else 0.0
                    emotion = EMOTION_LABELS[idx]
                    confidence = top1 * 100.0
                    # 前两名差距过小说明模型本身不确定，压低置信度避免误导用户
                    if top1 - top2 < 0.12:
                        confidence *= 0.6

                    # 多帧加权投票平滑，抑制单帧误判
                    if session_id is not None:
                        emotion, confidence = get_smoothed_emotion(session_id, emotion, confidence)
                        EMOTION_CACHE[session_id] = (emotion, confidence, time.time())
                except Exception as model_error:
                    print(f"情绪识别错误: {model_error}")
                    emotion = "neutral"
                    confidence = 50.0
            else:
                # 质量不佳/人脸过小时复用最近一次稳定结果（不写回缓冲，避免自我强化）
                cached = EMOTION_CACHE.get(session_id) if session_id is not None else None
                if cached and time.time() - cached[2] < 1.5:
                    emotion, confidence = cached[0], max(30.0, cached[1] * 0.9)
                else:
                    emotion = "neutral"
                    confidence = 30.0  # 低质量时降低置信度
        else:
            # 没检测到人脸时不再调用 DeepFace 做全图检测（耗时 150~200ms，是卡顿主因）
            # 直接复用最近一次结果，超过 1.5 秒才回落到 neutral
            cached = EMOTION_CACHE.get(session_id) if session_id is not None else None
            if cached and time.time() - cached[2] < 1.5:
                emotion, confidence = cached[0], max(30.0, cached[1] * 0.9)
            else:
                emotion, confidence = "neutral", 30.0

        return emotion, confidence, face_location

    except Exception as e:
        print(f"情绪识别错误: {e}")
        import traceback
        traceback.print_exc()
        return "neutral", 50.0, None

# 会话处理线程
def session_processor():
    while True:
        try:
            session_id = sessions_queue.get()
            if session_id in active_sessions:
                session_info = active_sessions[session_id]
                # 处理会话数据
                # 这里可以添加更多处理逻辑
                pass
            sessions_queue.task_done()
        except Exception as e:
            print(f"Session processor error: {e}")

# 启动会话处理线程
threading.Thread(target=session_processor, daemon=True).start()

# 静态文件服务
app.mount("/static", StaticFiles(directory="static"), name="static")

# 根路径重定向到登录页面
@app.get("/")
async def root():
    from fastapi.responses import RedirectResponse
    return RedirectResponse(url="/static/login.html")

# 认证相关API
@app.post("/api/auth/register", response_model=UserResponse)
async def register(user: UserCreate):
    db = SessionLocal()
    try:
        # 检查用户名是否已存在
        if db.query(User).filter(User.username == user.username).first():
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Username already registered"
            )
        
        # 检查邮箱是否已存在
        if db.query(User).filter(User.email == user.email).first():
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Email already registered"
            )
        
        # 检查密码是否一致
        if user.password != user.confirm_password:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Passwords do not match"
            )
        
        # 检查密码长度（bcrypt限制为72字节）
        if len(user.password.encode('utf-8')) > 72:
            # 手动截断密码
            truncated_password = user.password[:72]
            hashed_password = get_password_hash(truncated_password)
        else:
            # 创建新用户
            hashed_password = get_password_hash(user.password)
        db_user = User(
            username=user.username,
            email=user.email,
            hashed_password=hashed_password
        )
        db.add(db_user)
        db.commit()
        db.refresh(db_user)
        return db_user
    finally:
        db.close()

@app.post("/api/auth/token", response_model=Token)
async def login(form_data: OAuth2PasswordRequestForm = Depends()):
    db = SessionLocal()
    try:
        # 查找用户
        user = db.query(User).filter(User.username == form_data.username).first()
        if not user or not verify_password(form_data.password, user.hashed_password):
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Incorrect username or password",
                headers={"WWW-Authenticate": "Bearer"},
            )
        
        # 创建访问令牌
        access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
        access_token = create_access_token(
            data={"sub": user.username}, expires_delta=access_token_expires
        )
        return {"access_token": access_token, "token_type": "bearer"}
    finally:
        db.close()



@app.get("/api/auth/me", response_model=UserResponse)
async def get_current_user_info(current_user: User = Depends(get_current_user)):
    return current_user

# 问题相关的Pydantic模型
class QuestionModel(BaseModel):
    id: int
    text: str
    category: str
    options: list = []
    type: str = "choice"

class QuestionsRequest(BaseModel):
    questions: list[QuestionModel]

@app.post("/api/questions")
async def save_questions(request: QuestionsRequest, current_user: User = Depends(get_current_user)):
    """保存问卷问题到数据库"""
    db = SessionLocal()
    try:
        # 删除旧的问题记录
        db.query(Question).delete()
        
        # 保存新的问题
        for q in request.questions:
            options = q.options
            if options:
                options = ["一般" if opt == "一般kkk" else opt for opt in options]
            question = Question(
                question_id=q.id,
                text=q.text,
                category=q.category,
                options=json.dumps(options) if options else None,
                is_subjective=(q.type == "subjective")
            )
            db.add(question)
        
        db.commit()
        return {"message": "Questions saved successfully"}
    except Exception as e:
        db.rollback()
        return {"error": str(e)}
    finally:
        db.close()

@app.get("/api/questions")
async def get_questions(current_user: User = Depends(get_current_user)):
    """获取问卷问题"""
    db = SessionLocal()
    try:
        questions = db.query(Question).order_by(Question.question_id).all()
        result = []
        for q in questions:
            options = json.loads(q.options) if q.options else []
            options = ["一般" if opt == "一般kkk" else opt for opt in options]
            result.append({
                "id": q.question_id,
                "text": q.text,
                "category": q.category,
                "options": options,
                "type": "subjective" if q.is_subjective else "choice"
            })
        return {"questions": result}
    finally:
        db.close()

@app.post("/api/sessions", response_model=SessionResponse)
async def create_session(session: SessionCreate, current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        db_session = RecognitionSession(
            patient_id=session.patient_id,
            start_time=datetime.datetime.now(),
            end_time=None,
            duration=0.0,
            fps=0.0
        )
        db.add(db_session)
        db.commit()
        db.refresh(db_session)
        active_sessions[db_session.session_id] = {
            "patient_id": session.patient_id,
            "start_time": db_session.start_time,
            "emotions": [],
            "frames": []
        }
        sessions_queue.put(db_session.session_id)
        return SessionResponse(
            session_id=db_session.session_id,
            patient_id=db_session.patient_id,
            start_time=db_session.start_time
        )
    finally:
        db.close()

@app.post("/api/sessions/{session_id}/end")
async def end_session(session_id: int, current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        db_session = db.query(RecognitionSession).filter(RecognitionSession.session_id == session_id).first()
        if db_session:
            db_session.end_time = datetime.datetime.now()
            db_session.duration = (db_session.end_time - db_session.start_time).total_seconds()
            db.commit()
            if session_id in active_sessions:
                del active_sessions[session_id]
            return {"message": "Session ended successfully"}
        return {"message": "Session not found"}
    finally:
        db.close()

@app.websocket("/ws/{session_id}")
async def websocket_endpoint(websocket: WebSocket, session_id: int):
    await websocket.accept()
    if session_id not in active_sessions:
        await websocket.close(code=1008, reason="Session not found")
        return
    
    # 预创建数据库会话，减少每次操作的开销
    db = SessionLocal()
    try:
        # 批量处理计数器
        batch_count = 0
        batch_records = []
        batch_size = 3  # 每3帧批量保存一次
        
        while True:
            data = await websocket.receive_text()
            # 处理接收到的视频帧
            frame_data = json.loads(data)
            if "image" in frame_data:
                # 解码base64图像
                image_data = frame_data["image"].split(",")[1]
                image_bytes = base64.b64decode(image_data)
                np_arr = np.frombuffer(image_bytes, np.uint8)
                frame = cv2.imdecode(np_arr, cv2.IMREAD_COLOR)
                
                # 获取问题信息
                question_id = frame_data.get("question_id", None)
                question_text = frame_data.get("question_text", None)
                
                # 确保问题ID是整数类型
                if question_id is not None:
                    try:
                        question_id = int(question_id)
                    except (ValueError, TypeError):
                        question_id = None
                
                # 处理检测到的人脸
                emotion_records = []

                # 人脸检测与情绪识别放到线程池执行，避免阻塞 WebSocket 事件循环
                emotion, confidence, face_location = await run_in_threadpool(recognize_emotion, frame, session_id)

                # 收集情绪记录
                emotion_record = EmotionRecord(
                    session_id=session_id,
                    timestamp=datetime.datetime.now(),
                    emotion_type=emotion,
                    confidence=confidence,
                    frame_path="",
                    question_id=question_id,
                    question_text=question_text
                )
                emotion_records.append(emotion_record)
                
                # 批量保存情绪记录
                if emotion_records:
                    batch_records.extend(emotion_records)
                    batch_count += 1
                    
                    # 每batch_size帧批量保存一次
                    if batch_count >= batch_size:
                        try:
                            db.add_all(batch_records)
                            db.commit()
                            batch_records = []
                            batch_count = 0
                        except Exception as e:
                            print(f"Database error: {e}")
                            db.rollback()
                
                # 只回传识别结果（人脸框 + 情绪），画面由前端本地实时渲染，
                # 不再把 JPEG 编码后回传，省掉每帧 ~10ms 编码和传输开销
                emotions_data = []
                for record in emotion_records:
                    emotions_data.append({
                        "emotion": record.emotion_type,
                        "confidence": record.confidence
                    })

                # 如果没有情绪记录，添加当前情绪
                if not emotions_data:
                    emotions_data.append({
                        "emotion": emotion,
                        "confidence": confidence
                    })

                await websocket.send_text(json.dumps({
                    "type": "frame",
                    "box": list(face_location) if face_location else None,
                    "emotion": emotion,
                    "confidence": round(confidence, 2),
                    "emotions": emotions_data
                }))
    except WebSocketDisconnect:
        print(f"WebSocket disconnected for session {session_id}")
        # 保存剩余的记录
        if batch_records:
            try:
                db.add_all(batch_records)
                db.commit()
            except Exception as e:
                print(f"Database error on disconnect: {e}")
    except Exception as e:
        print(f"WebSocket error: {e}")
    finally:
        db.close()

@app.post("/api/sessions/{session_id}/screenshot")
async def take_screenshot(session_id: int, file: UploadFile = File(...), current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        # 保存截图
        os.makedirs("screenshots", exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = f"screenshots/{session_id}_{timestamp}.jpg"
        
        with open(file_path, "wb") as f:
            f.write(await file.read())
        
        # 保存截图记录
        screenshot_record = ScreenshotRecord(
            session_id=session_id,
            image_path=file_path,
            timestamp=datetime.datetime.now(),
            emotion_data=""
        )
        db.add(screenshot_record)
        db.commit()
        
        return {"message": "Screenshot saved successfully", "file_path": file_path}
    finally:
        db.close()

@app.post("/api/sessions/{session_id}/question-emotions")
async def save_question_emotions(session_id: int, data: dict, current_user: User = Depends(get_current_user)):
    """保存问题情绪数据"""
    db = SessionLocal()
    try:
        # 检查会话是否存在
        session = db.query(RecognitionSession).filter(RecognitionSession.session_id == session_id).first()
        if not session:
            return {"error": "Session not found"}
        
        # 保存问题情绪数据
        question_emotion = QuestionEmotion(
            session_id=session_id,
            question_id=data.get("question_id"),
            question_text=data.get("question_text"),
            answer=data.get("answer"),
            start_time=datetime.datetime.fromisoformat(data.get("start_time")),
            end_time=datetime.datetime.fromisoformat(data.get("end_time")),
            emotion_data=json.dumps(data.get("emotions", []))
        )
        db.add(question_emotion)
        db.commit()
        
        return {"message": "Question emotions saved successfully"}
    except Exception as e:
        db.rollback()
        return {"error": str(e)}
    finally:
        db.close()

@app.post("/api/sessions/{session_id}/recordings")
async def upload_recording(session_id: int, audio: UploadFile = File(...), current_user: User = Depends(get_current_user)):
    """上传录音并进行语音识别"""
    db = SessionLocal()
    try:
        # 检查会话是否存在
        session = db.query(RecognitionSession).filter(RecognitionSession.session_id == session_id).first()
        if not session:
            return {"error": "Session not found"}
        
        # 保存录音
        import os
        os.makedirs("recordings", exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = f"recordings/{session_id}_{timestamp}.wav"
        with open(file_path, "wb") as f:
            f.write(await audio.read())
        
        # 进行语音识别
        transcript = ""
        try:
            import speech_recognition as sr
            import subprocess
            
            print(f"Processing audio file: {file_path}")
            if os.path.exists(file_path):
                print(f"File exists: True")
                print(f"File size: {os.path.getsize(file_path)} bytes")
            else:
                print(f"File exists: False")
                transcript = "音频文件不存在，请重试"
                return {"message": "Recording uploaded successfully", "transcript": transcript}
            
            # 检查文件格式，如果不是WAV，尝试转换
            if not file_path.endswith('.wav'):
                # 尝试使用ffmpeg转换（如果可用）
                wav_path = file_path.replace('.wav', '_converted.wav')
                try:
                    print(f"Attempting to convert to WAV: {wav_path}")
                    subprocess.run(['ffmpeg', '-i', file_path, wav_path], check=True, capture_output=True)
                    file_path = wav_path
                    print(f"Conversion successful: {file_path}")
                except Exception as convert_error:
                    print(f"Error converting audio: {convert_error}")
            
            r = sr.Recognizer()
            with sr.AudioFile(file_path) as source:
                print("Reading audio file...")
                audio_data = r.record(source)
                print("Audio data recorded, attempting recognition...")
                # 使用Google Web Speech API进行识别
                transcript = r.recognize_google(audio_data, language='zh-CN')
                print(f"Recognition successful: {transcript}")
        except Exception as e:
            print(f"Speech recognition error: {e}")
            # 提供更详细的错误信息
            if "request error" in str(e).lower():
                transcript = "语音识别服务无法访问，请检查网络连接"
            elif "audio file could not be read" in str(e).lower():
                transcript = "音频文件格式不支持，请重试"
            elif "no speech detected" in str(e).lower():
                transcript = "未检测到语音，请重试"
            else:
                transcript = f"语音识别失败: {str(e)}"
            print(f"Error message: {transcript}")
        
        # 保存录音记录
        recording_record = RecordingRecord(
            session_id=session_id,
            file_path=file_path,
            timestamp=datetime.datetime.now()
        )
        db.add(recording_record)
        db.commit()
        
        # 录音仅存档，不写入第10题答案（第10题答案由 Speech Recognition 文本框保存）
        return {"message": "Recording uploaded successfully", "transcript": transcript}
    except Exception as e:
        db.rollback()
        return {"error": str(e)}
    finally:
        db.close()

@app.post("/api/sessions/{session_id}/speech-recognition")
async def save_speech_recognition_result(session_id: int, request: dict, current_user: User = Depends(get_current_user)):
    """保存Windows语音识别结果"""
    db = SessionLocal()
    try:
        # 检查会话是否存在
        session = db.query(RecognitionSession).filter(RecognitionSession.session_id == session_id).first()
        if not session:
            return {"success": False, "error": "Session not found"}
        
        # 获取识别结果
        transcript = request.get('transcript', '').strip()
        if not transcript:
            return {"success": False, "error": "Transcript is empty"}
        
        # 查找该会话的主观题记录（只有第10题是主观题）
        subjective_questions = db.query(QuestionEmotion).filter(
            QuestionEmotion.session_id == session_id,
            QuestionEmotion.question_id == 10
        ).all()
        
        # 将文本框内容写入第10题（更新已有记录；若无记录则新建）
        if subjective_questions:
            for question in subjective_questions:
                question.answer = transcript
        else:
            q10 = db.query(Question).filter(Question.question_id == 10).first()
            question_emotion = QuestionEmotion(
                session_id=session_id,
                question_id=10,
                question_text=q10.text if q10 else "请用你自己的话描述最近的状态和表现。",
                answer=transcript,
                start_time=datetime.datetime.now(),
                end_time=datetime.datetime.now(),
                emotion_data=json.dumps([]),
            )
            db.add(question_emotion)
        db.commit()
        return {"success": True, "message": "Speech recognition result saved successfully"}
    except Exception as e:
        db.rollback()
        return {"success": False, "error": str(e)}
    finally:
        db.close()

@app.get("/api/sessions/{session_id}/emotions")
async def get_session_emotions(session_id: int, current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        records = db.query(EmotionRecord).filter(EmotionRecord.session_id == session_id).order_by(EmotionRecord.timestamp).all()
        emotions = [
            {
                "timestamp": record.timestamp,
                "emotion": record.emotion_type,
                "confidence": record.confidence
            }
            for record in records
        ]
        return {"emotions": emotions}
    finally:
        db.close()

@app.get("/api/sessions/export/excel")
async def export_all_sessions_excel(current_user: User = Depends(get_current_user)):
    from fastapi.responses import StreamingResponse
    import pandas as pd
    import io
    
    db = SessionLocal()
    try:
        # 获取所有会话，按患者ID分组
        sessions = db.query(RecognitionSession).all()
        
        if not sessions:
            return {"message": "No sessions found"}
        
        # 获取最新的问题列表
        questions = db.query(Question).order_by(Question.question_id).all()
        question_options = {}
        for q in questions:
            options = json.loads(q.options) if q.options else []
            question_options[q.question_id] = options
        
        # 按患者ID分组
        patient_sessions = {}
        for session in sessions:
            if session.patient_id not in patient_sessions:
                patient_sessions[session.patient_id] = []
            patient_sessions[session.patient_id].append(session)
        
        # 创建Excel文件
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 为每个患者创建一个工作表
            for patient_id, patient_session_list in patient_sessions.items():
                # 收集该患者的所有情绪记录
                all_data = []
                
                for session in patient_session_list:
                    # 获取情绪记录
                    records = db.query(EmotionRecord).filter(EmotionRecord.session_id == session.session_id).order_by(EmotionRecord.timestamp).all()
                    
                    # 获取问题情绪记录
                    question_emotions = db.query(QuestionEmotion).filter(QuestionEmotion.session_id == session.session_id).order_by(QuestionEmotion.start_time).all()
                    
                    # 为每个问题创建记录映射，用于跟踪已处理的记录
                    question_records = {}
                    
                    # 准备数据
                    for record in records:
                        question_id = ""
                        question_text = ""
                        answer = ""
                        stored_answer = ""
                        
                        # 尝试通过时间戳匹配QuestionEmotion记录
                        for qe in question_emotions:
                            if qe.start_time:
                                # 宽松的时间戳匹配，允许±30秒的误差
                                time_diff = abs((record.timestamp - qe.start_time).total_seconds())
                                if time_diff <= 30:
                                    question_id = qe.question_id
                                    question_text = qe.question_text
                                    stored_answer = qe.answer
                                    break
                        
                        # 如果没有找到，检查EmotionRecord中的问题信息
                        if not question_id:
                            # 检查EmotionRecord是否有问题信息
                            if hasattr(record, 'question_id') and record.question_id:
                                question_id = record.question_id
                            if hasattr(record, 'question_text') and record.question_text:
                                question_text = record.question_text
                            
                            # 如果有问题ID，尝试找到对应的回答
                            if question_id:
                                for qe in question_emotions:
                                    if hasattr(qe, 'question_id') and qe.question_id == question_id:
                                        stored_answer = qe.answer
                                        break
                        
                        # 转换答案：如果答案是数字索引，转换为实际选项文本
                        if stored_answer:
                            try:
                                # 尝试将答案转换为整数索引
                                answer_index = int(stored_answer)
                                # 检查是否有对应的问题选项
                                if question_id in question_options:
                                    options = question_options[question_id]
                                    if 0 <= answer_index < len(options):
                                        answer = options[answer_index]
                                    else:
                                        answer = stored_answer
                                else:
                                    answer = stored_answer
                            except ValueError:
                                # 如果不是数字，直接使用原始答案
                                answer = stored_answer
                        else:
                            answer = ""
                        
                        # 确保值不为None
                        question_id = question_id if question_id is not None else ""
                        question_text = question_text if question_text is not None else ""
                        
                        # 为每个问题记录添加到映射中
                        if question_id:
                            if question_id not in question_records:
                                question_records[question_id] = []
                            question_records[question_id].append({
                                "会话ID": session.session_id,
                                "时间戳": record.timestamp,
                                "情绪类型": record.emotion_type,
                                "置信度": record.confidence,
                                "问题ID": question_id,
                                "问题内容": question_text,
                                "回答": answer
                            })
                        else:
                            # 没有问题ID的记录直接添加
                            all_data.append({
                                "会话ID": session.session_id,
                                "时间戳": record.timestamp,
                                "情绪类型": record.emotion_type,
                                "置信度": record.confidence,
                                "问题ID": question_id,
                                "问题内容": question_text,
                                "回答": answer
                            })
                    
                    # 处理每个问题的记录，对于主观题只保留四成的行数
                    for question_id, records in question_records.items():
                        if question_id == 10:  # 只有第10题是主观题
                            # 只保留现有一半的行数（15%）
                            keep_count = max(1, int(len(records) * 0.15))
                            # 均匀采样
                            step = max(1, len(records) // keep_count)
                            selected_records = records[::step][:keep_count]
                            all_data.extend(selected_records)
                        else:
                            # 其他题目的记录全部保留
                            all_data.extend(records)
                
                # 创建DataFrame
                df = pd.DataFrame(all_data)
                
                # 写入数据
                sheet_name = patient_id[:31]  # 限制工作表名称长度
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 获取工作表对象，准备合并单元格
                worksheet = writer.sheets[sheet_name]
                
                # 处理主观题的合并单元格
                # 检查'问题ID'列是否存在
                if '问题ID' in df.columns:
                    # 按问题ID分组
                    grouped = df.groupby('问题ID')
                    
                    for question_id, group in grouped:
                        if question_id == 10:  # 只有第10题是主观题
                            if len(group) > 1:
                                # 获取起始和结束行
                                start_row = group.index.min() + 2  # +2 因为Excel从1开始，且有表头
                                end_row = group.index.max() + 2
                                
                                # 合并问题ID列
                                worksheet.merge_cells(start_row=start_row, start_column=5, end_row=end_row, end_column=5)
                                
                                # 合并问题内容列
                                worksheet.merge_cells(start_row=start_row, start_column=6, end_row=end_row, end_column=6)
                                
                                # 合并回答列
                                worksheet.merge_cells(start_row=start_row, start_column=7, end_row=end_row, end_column=7)
        
        output.seek(0)
        
        # 返回Excel文件
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=emotion_analysis_all.xlsx"}
        )
    finally:
        db.close()

@app.get("/api/patients")
async def get_patients(current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        # 获取所有会话
        sessions = db.query(RecognitionSession).all()
        
        # 按患者ID分组
        patient_sessions = {}
        for session in sessions:
            if session.patient_id not in patient_sessions:
                patient_sessions[session.patient_id] = []
            patient_sessions[session.patient_id].append({
                "session_id": session.session_id,
                "start_time": session.start_time,
                "end_time": session.end_time,
                "duration": session.duration
            })
        
        # 构建患者列表
        patients = []
        for patient_id, sessions in patient_sessions.items():
            patients.append({
                "patient_id": patient_id,
                "sessions": sessions
            })
        
        return patients
    finally:
        db.close()

@app.get("/api/sessions/{session_id}")
async def get_session(session_id: int, current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        # 获取会话信息
        session = db.query(RecognitionSession).filter(RecognitionSession.session_id == session_id).first()
        if not session:
            return {"message": "Session not found"}
        
        # 返回会话信息
        return {
            "session_id": session.session_id,
            "patient_id": session.patient_id,
            "start_time": session.start_time,
            "end_time": session.end_time,
            "duration": session.duration
        }
    finally:
        db.close()

@app.get("/api/emotions")
async def get_emotions(session_id: int, current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        # 获取指定会话的情绪记录
        records = db.query(EmotionRecord).filter(EmotionRecord.session_id == session_id).order_by(EmotionRecord.timestamp).all()
        
        # 构建返回数据
        emotions = []
        for record in records:
            emotions.append({
                "emotion": record.emotion_type,
                "confidence": record.confidence,
                "timestamp": record.timestamp
            })
        
        return emotions
    finally:
        db.close()

@app.delete("/api/patients/{patient_id}")
async def delete_patient(patient_id: str, current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        # 获取该患者的所有会话
        sessions = db.query(RecognitionSession).filter(RecognitionSession.patient_id == patient_id).all()
        
        # 为每个会话删除相关数据
        for session in sessions:
            # 删除问题情绪记录
            db.query(QuestionEmotion).filter(QuestionEmotion.session_id == session.session_id).delete()
            
            # 删除录音记录
            db.query(RecordingRecord).filter(RecordingRecord.session_id == session.session_id).delete()
            
            # 删除截图记录
            db.query(ScreenshotRecord).filter(ScreenshotRecord.session_id == session.session_id).delete()
            
            # 删除情绪记录
            db.query(EmotionRecord).filter(EmotionRecord.session_id == session.session_id).delete()
            
            # 删除会话记录
            db.query(RecognitionSession).filter(RecognitionSession.session_id == session.session_id).delete()
        
        db.commit()
        return {"message": "Patient data deleted successfully"}
    except Exception as e:
        db.rollback()
        return {"message": f"Error deleting patient data: {str(e)}"}
    finally:
        db.close()

@app.delete("/api/patients")
async def delete_all_patients(current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        # 删除所有问题情绪记录
        db.query(QuestionEmotion).delete()
        
        # 删除所有录音记录
        db.query(RecordingRecord).delete()
        
        # 删除所有截图记录
        db.query(ScreenshotRecord).delete()
        
        # 删除所有情绪记录
        db.query(EmotionRecord).delete()
        
        # 删除所有会话记录
        db.query(RecognitionSession).delete()
        
        db.commit()
        return {"message": "All patient data deleted successfully"}
    except Exception as e:
        db.rollback()
        return {"message": f"Error deleting all patient data: {str(e)}"}
    finally:
        db.close()

@app.get("/api/sessions/{session_id}/export/excel")
async def export_session_excel(session_id: int, current_user: User = Depends(get_current_user)):
    from fastapi.responses import StreamingResponse
    import pandas as pd
    import io
    
    db = SessionLocal()
    try:
        # 获取会话信息
        session = db.query(RecognitionSession).filter(RecognitionSession.session_id == session_id).first()
        if not session:
            return {"message": "Session not found"}
        
        # 获取情绪记录
        records = db.query(EmotionRecord).filter(EmotionRecord.session_id == session_id).order_by(EmotionRecord.timestamp).all()
        
        # 获取问题情绪记录
        question_emotions = db.query(QuestionEmotion).filter(QuestionEmotion.session_id == session_id).order_by(QuestionEmotion.start_time).all()
        
        # 准备情绪数据
        emotion_data = []
        for record in records:
            # 查找该记录对应的问题
            corresponding_question = None
            for qe in question_emotions:
                if qe.start_time:
                    # 宽松的时间戳匹配，允许±30秒的误差
                    time_diff = abs((record.timestamp - qe.start_time).total_seconds())
                    if time_diff <= 30:
                        corresponding_question = qe
                        break
            
            # 获取问题信息
            question_id = ""
            question_text = ""
            answer = ""
            if corresponding_question:
                question_id = corresponding_question.question_id
                question_text = corresponding_question.question_text
                answer = corresponding_question.answer
            
            emotion_data.append({
                "Timestamp": record.timestamp,
                "Emotion": record.emotion_type,
                "Confidence": record.confidence,
                "Question_ID": question_id,
                "Question_Text": question_text,
                "Answer": answer
            })
        
        # 创建情绪DataFrame
        df_emotions = pd.DataFrame(emotion_data)
        
        # 准备问题数据
        question_data = []
        for qe in question_emotions:
            question_data.append({
                "Question_ID": qe.question_id,
                "Question_Text": qe.question_text,
                "Answer": qe.answer,
                "Start_Time": qe.start_time,
                "End_Time": qe.end_time,
                "Duration_Seconds": (qe.end_time - qe.start_time).total_seconds() if qe.end_time and qe.start_time else 0
            })
        
        # 创建问题DataFrame
        df_questions = pd.DataFrame(question_data)
        
        # 创建Excel文件
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 写入情绪数据
            emotion_sheet_name = f'{session.patient_id}_Emotions'[:31]
            df_emotions.to_excel(writer, sheet_name=emotion_sheet_name, index=False)
            
            # 处理情绪数据工作表的合并单元格
            worksheet = writer.sheets[emotion_sheet_name]
            # 检查'Question_ID'列是否存在
            if 'Question_ID' in df_emotions.columns:
                grouped = df_emotions.groupby('Question_ID')
                
                for question_id, group in grouped:
                    if question_id == 10:  # 只有第10题是主观题
                        if len(group) > 1:
                            # 获取起始和结束行
                            start_row = group.index.min() + 2  # +2 因为Excel从1开始，且有表头
                            end_row = group.index.max() + 2
                            
                            # 合并问题ID列
                            worksheet.merge_cells(start_row=start_row, start_column=4, end_row=end_row, end_column=4)
                            
                            # 合并问题内容列
                            worksheet.merge_cells(start_row=start_row, start_column=5, end_row=end_row, end_column=5)
                            
                            # 合并回答列
                            worksheet.merge_cells(start_row=start_row, start_column=6, end_row=end_row, end_column=6)
            
            # 写入问题数据
            if not df_questions.empty:
                question_sheet_name = f'{session.patient_id}_Questions'[:31]
                df_questions.to_excel(writer, sheet_name=question_sheet_name, index=False)
        
        output.seek(0)
        
        # 返回Excel文件
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename=emotion_analysis_{session.patient_id}.xlsx"}
        )
    finally:
        db.close()


# 视频处理接口
@app.post("/api/video/process")
async def process_video(video: UploadFile = File(...), patient_id: str = Form(...), current_user: User = Depends(get_current_user)):
    """处理上传的视频文件"""
    try:
        # 创建会话
        db = SessionLocal()
        db_session = RecognitionSession(
            patient_id=patient_id,
            start_time=datetime.datetime.now(),
            end_time=None,
            duration=0.0,
            fps=0.0
        )
        db.add(db_session)
        db.commit()
        db.refresh(db_session)
        session_id = db_session.session_id
        
        # 保存视频文件
        os.makedirs("videos", exist_ok=True)
        video_path = f"videos/{session_id}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.mp4"
        with open(video_path, "wb") as f:
            f.write(video.file.read())
        
        # 处理视频
        cap = cv2.VideoCapture(video_path)
        fps = cap.get(cv2.CAP_PROP_FPS)
        frame_interval = max(1, int(fps / 5))  # 每5fps处理一帧
        frame_count = 0
        processed_frames = 0
        
        emotion_records = []
        
        while cap.isOpened():
            ret, frame = cap.read()
            if not ret:
                break
            
            frame_count += 1
            if frame_count % frame_interval != 0:
                continue
            
            # 人脸检测
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray, 1.3, 5)
            
            # 处理检测到的人脸
            for (x, y, w, h) in faces:
                # 提取人脸区域
                face_roi = frame[y:y+h, x:x+w]

                # 情绪识别（视频处理不需要平滑）
                emotion, confidence, _ = recognize_emotion(face_roi, session_id=None)
                
                # 收集情绪记录
                emotion_records.append(EmotionRecord(
                    session_id=session_id,
                    timestamp=datetime.datetime.now(),
                    emotion_type=emotion,
                    confidence=confidence,
                    frame_path="",
                    question_id=None,
                    question_text="Video Analysis"
                ))
                
                processed_frames += 1
            
            # 每处理10帧保存一次数据
            if len(emotion_records) >= 10:
                db.add_all(emotion_records)
                db.commit()
                emotion_records = []
        
        # 保存剩余的记录
        if emotion_records:
            db.add_all(emotion_records)
            db.commit()
        
        # 更新会话信息
        db_session.end_time = datetime.datetime.now()
        db_session.duration = (db_session.end_time - db_session.start_time).total_seconds()
        db_session.fps = processed_frames / db_session.duration if db_session.duration > 0 else 0
        db.commit()
        
        cap.release()
        
        return {"session_id": session_id, "message": "Video processed successfully"}
    except Exception as e:
        print(f"Error processing video: {e}")
        return {"message": "Error processing video", "error": str(e)}
    finally:
        db.close()

# AI综合分析报告相关API

@app.get("/api/patients")
async def get_patients(current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        # 获取所有唯一的患者ID
        patients = db.query(RecognitionSession.patient_id).distinct().all()
        patient_list = [{'patient_id': p.patient_id} for p in patients]
        return patient_list
    except Exception as e:
        return {"message": f"Error getting patients: {str(e)}"}
    finally:
        db.close()

@app.get("/api/patients/{patient_id}/data")
async def get_patient_data(patient_id: str, current_user: User = Depends(get_current_user)):
    db = SessionLocal()
    try:
        # 获取患者的所有会话
        sessions = db.query(RecognitionSession).filter(RecognitionSession.patient_id == patient_id).order_by(RecognitionSession.start_time).all()
        
        patient_data = {
            'patient_id': patient_id,
            'sessions': []
        }
        
        for session in sessions:
            # 获取会话的情绪记录，采样处理，每5条取1条，减少数据量
            all_emotions = db.query(EmotionRecord).filter(EmotionRecord.session_id == session.session_id).order_by(EmotionRecord.timestamp).all()
            # 采样处理，保留关键情绪变化
            emotions = []
            for i, e in enumerate(all_emotions):
                # 保留第一条、最后一条，以及每5条取1条
                if i == 0 or i == len(all_emotions) - 1 or i % 5 == 0:
                    emotions.append({
                        'timestamp': e.timestamp,
                        'emotion_type': e.emotion_type,
                        'confidence': e.confidence,
                        'question_id': e.question_id,
                        'question_text': e.question_text
                    })
            
            # 获取会话的问题记录
            questions = db.query(QuestionEmotion).filter(QuestionEmotion.session_id == session.session_id).order_by(QuestionEmotion.start_time).all()
            
            session_data = {
                'session_id': session.session_id,
                'start_time': session.start_time,
                'emotions': emotions,
                'questions': [
                    {
                        'question_id': q.question_id,
                        'question_text': q.question_text,
                        'answer': q.answer,
                        'start_time': q.start_time,
                        'end_time': q.end_time
                    }
                    for q in questions
                ]
            }
            
            patient_data['sessions'].append(session_data)
        
        return patient_data
    except Exception as e:
        return {"message": f"Error getting patient data: {str(e)}"}
    finally:
        db.close()

class AiAnalysisRequest(BaseModel):
    patient_data: dict
    prompt: str

class AiAnalysisResponse(BaseModel):
    result: str

@app.post("/api/ai/analyze", response_model=AiAnalysisResponse)
async def analyze_patient_data(request: AiAnalysisRequest, current_user: User = Depends(get_current_user)):
    try:
        # 构建请求数据
        patient_data = request.patient_data
        prompt = request.prompt
        
        print(f"Patient data received: {json.dumps(patient_data, ensure_ascii=False, default=str)[:500]}...")
        print(f"Prompt length: {len(prompt)}")
        
        # 替换患者数据占位符
        formatted_prompt = prompt.replace('{patient_data}', json.dumps(patient_data, ensure_ascii=False, default=str))
        
        print(f"Formatted prompt length: {len(formatted_prompt)}")
        
        # 使用DeepSeek API
        api_key = "sk-54276de9c6ee4af39c5cc94f1748abda"
        url = "https://api.deepseek.com/v1/chat/completions"
        
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        
        data = {
            "model": "deepseek-chat",
            "messages": [
                {
                    "role": "system",
                    "content": "你是一名专业的心理治疗师，擅长分析患者的情绪数据并生成详细的心理评估报告。请控制报告长度，保持在原内容的70%左右，同时确保信息完整。"
                },
                {
                    "role": "user",
                    "content": formatted_prompt
                }
            ],
            "temperature": 0.6,  # 略微降低温度，加快生成速度
            "max_tokens": 1200  # 减少最大token数，加快生成速度
        }
        
        print(f"Request data: {json.dumps(data, ensure_ascii=False, default=str)[:500]}...")
        
        # 发送请求，添加重试机制
        max_retries = 3
        retry_delay = 2  # 秒
        
        for attempt in range(max_retries):
            try:
                print(f"Attempt {attempt + 1}/{max_retries} to call DeepSeek API...")
                response = requests.post(url, headers=headers, json=data, timeout=60)  # 增加超时时间到60秒
                print(f"Response status: {response.status_code}")
                response_text = response.text
                print(f"Response text: {response_text[:500]}...")
                
                if response.status_code == 200:
                    result = response.json()
                    return AiAnalysisResponse(result=result['choices'][0]['message']['content'])
                else:
                    try:
                        error_data = response.json()
                        error_message = error_data.get('error', {}).get('message', 'AI分析失败')
                    except:
                        error_message = f"HTTP {response.status_code}: {response_text}"
                    raise HTTPException(status_code=response.status_code, detail=error_message)
            except requests.exceptions.Timeout:
                print(f"Request timed out, retrying in {retry_delay} seconds...")
                if attempt < max_retries - 1:
                    import time
                    time.sleep(retry_delay)
                else:
                    raise HTTPException(status_code=500, detail="AI分析失败: 请求超时，请检查网络连接")
            except Exception as e:
                print(f"Error sending request: {str(e)}")
                if attempt < max_retries - 1:
                    import time
                    time.sleep(retry_delay)
                else:
                    raise HTTPException(status_code=500, detail=f"AI分析失败: {str(e)}")
        
        # 返回模拟的分析结果
        # mock_result = "# 心理评估报告\n\n## 1. 患者基本信息\n- 患者ID: {patient_id}\n- 评估时间范围: 2026-04-14\n\n## 2. 情绪分析\n- 主要情绪类型: 中性、悲伤\n- 情绪变化趋势: 整体情绪较为稳定，偶尔出现悲伤情绪\n- 情绪稳定性: 情绪波动较小，整体状态平稳\n\n## 3. 问题分析\n- 患者在回答关于情绪状态的问题时，表现出轻微的悲伤情绪\n- 在回答睡眠质量问题时，情绪较为中性\n- 在回答社交活动问题时，情绪保持稳定\n\n## 4. 心理评估\n- 患者整体心理状态良好，情绪基本稳定\n- 存在轻微的情绪波动，主要表现为偶尔的悲伤情绪\n- 建议关注患者的情绪变化，特别是在压力较大的时期\n\n## 5. 建议和治疗方案\n- 建议保持规律的作息时间，保证充足的睡眠\n- 鼓励患者参与更多的社交活动，增强社交支持\n- 建议定期进行情绪监测，及时发现情绪变化\n- 如情绪持续低落，建议寻求专业心理辅导\n\n## 6. 结论\n患者当前心理状态良好，情绪基本稳定。建议继续保持健康的生活方式，定期进行情绪监测，如有需要，及时寻求专业帮助。".format(patient_id=patient_data.get('patient_id', '未知'))
        # return AiAnalysisResponse(result=mock_result)
    except Exception as e:
        print(f"Error in AI analysis: {str(e)}")
        raise HTTPException(status_code=500, detail=f"AI分析失败: {str(e)}")

@app.get("/")
async def root():
    return {"message": "Emotion Analysis API"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
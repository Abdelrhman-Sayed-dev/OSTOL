"""
Fleet Management System — Production-Ready Backend
Author: Refactored for enterprise use
"""

import os, re, uuid, logging, shutil, struct
from datetime import datetime, timedelta
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, List, Optional

import bcrypt
import jwt
import sqlite3
from fastapi import (
    FastAPI, Depends, HTTPException, Request,
    UploadFile, File, status
)
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.trustedhost import TrustedHostMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, validator
from slowapi import Limiter, _rate_limit_exceeded_handler
from slowapi.util import get_remote_address
from slowapi.errors import RateLimitExceeded
from starlette.middleware.base import BaseHTTPMiddleware

# ══════════════════════════════════════════════════════
# 1. ENVIRONMENT — REQUIRED (no fallback secrets)
# ══════════════════════════════════════════════════════

def _require_env(key: str) -> str:
    val = os.environ.get(key)
    if not val:
        raise RuntimeError(
            f"❌ Missing required environment variable: {key}\n"
            f"   Set it in your .env file or hosting dashboard."
        )
    return val

try:
    SECRET_KEY     = _require_env("SECRET_KEY")
    DATABASE_PATH  = os.environ.get("DATABASE_PATH", "trip_tracker.db")
    ALGORITHM      = "HS256"
    ACCESS_EXP_MIN = int(os.environ.get("ACCESS_TOKEN_EXP_MIN", "30"))    # 30 min
    REFRESH_EXP_DAYS = int(os.environ.get("REFRESH_TOKEN_EXP_DAYS", "7")) # 7 days
    UPLOAD_DIR     = Path(os.environ.get("UPLOAD_DIR", "uploads"))
    MAX_AUDIO_MB   = int(os.environ.get("MAX_AUDIO_MB", "10"))
    ALLOWED_ORIGINS = os.environ.get("ALLOWED_ORIGINS", "http://localhost:8000").split(",")
    ENVIRONMENT    = os.environ.get("ENVIRONMENT", "development")          # "production"
    ALLOWED_AUDIO  = {"audio/webm", "audio/ogg", "audio/mp4", "audio/mpeg", "audio/wav"}
except RuntimeError as e:
    print(e); raise SystemExit(1)

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# ══════════════════════════════════════════════════════
# 2. STRUCTURED LOGGING
# ══════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format='{"time":"%(asctime)s","level":"%(levelname)s","msg":"%(message)s"}',
    datefmt="%Y-%m-%dT%H:%M:%S"
)
log = logging.getLogger("fleet")

def log_event(event: str, **ctx):
    """Structured log — never log passwords or tokens."""
    safe = {k: v for k, v in ctx.items() if k not in ("password","token","audio_data")}
    log.info(f"{event} | {safe}")

# ══════════════════════════════════════════════════════
# 3. DATABASE — SQLite with WAL mode + proper indexing
# ══════════════════════════════════════════════════════

@contextmanager
def get_db():
    conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")   # write-ahead logging for concurrency
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def migrate_db():
    with get_db() as conn:
        c = conn.cursor()

        # USERS
        c.execute("""CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN('admin','driver')),
            created_at TEXT DEFAULT(datetime('now')),
            last_login TEXT,
            refresh_token TEXT,
            refresh_exp TEXT
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")

        # DRIVERS
        c.execute("""CREATE TABLE IF NOT EXISTS drivers(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            phone TEXT DEFAULT '',
            status TEXT DEFAULT 'active',
            user_id INTEGER UNIQUE,
            national_id TEXT DEFAULT '',
            birth_date TEXT DEFAULT '',
            driver_license_expiry TEXT DEFAULT '',
            vehicle_license_expiry TEXT DEFAULT '',
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_drivers_user ON drivers(user_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_drivers_status ON drivers(status)")

        # CARS
        c.execute("""CREATE TABLE IF NOT EXISTS cars(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plate TEXT UNIQUE NOT NULL,
            model TEXT NOT NULL,
            status TEXT DEFAULT 'available',
            chassis TEXT DEFAULT '',
            car_license_expiry TEXT DEFAULT '',
            sector TEXT DEFAULT ''
        )""")

        # PERMISSIONS
        c.execute("""CREATE TABLE IF NOT EXISTS driver_car_permissions(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER NOT NULL,
            car_id INTEGER NOT NULL,
            UNIQUE(driver_id, car_id),
            FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE CASCADE,
            FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE CASCADE
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_perm_driver ON driver_car_permissions(driver_id)")

        # TRIPS
        c.execute("""CREATE TABLE IF NOT EXISTS trips(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER NOT NULL,
            car_id INTEGER NOT NULL,
            start_time TEXT,
            end_time TEXT,
            start_odometer REAL,
            end_odometer REAL,
            start_location TEXT DEFAULT '',
            end_location TEXT DEFAULT '',
            garage_location TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE CASCADE,
            FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE CASCADE
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_trips_driver ON trips(driver_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_trips_active ON trips(driver_id, end_time)")

        # WORKSHOPS
        c.execute("""CREATE TABLE IF NOT EXISTS workshop_records(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER,
            type TEXT NOT NULL,
            quantity REAL,
            price REAL NOT NULL DEFAULT 0,
            notes TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            operation_type TEXT DEFAULT '',
            vehicle_id INTEGER,
            odometer_reading REAL,
            description TEXT DEFAULT '',
            tire_action TEXT DEFAULT '',
            location TEXT DEFAULT '',
            FOREIGN KEY(driver_id) REFERENCES drivers(id),
            FOREIGN KEY(vehicle_id) REFERENCES cars(id)
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_driver ON workshop_records(driver_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_type ON workshop_records(type)")

        # EMERGENCY REPORTS — audio stored as file path, NOT base64
        c.execute("""CREATE TABLE IF NOT EXISTS emergency_reports(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER,
            car_id INTEGER,
            type TEXT NOT NULL,
            audio_url TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            is_read INTEGER DEFAULT 0,
            location TEXT DEFAULT '',
            is_handled INTEGER DEFAULT 0,
            action_taken TEXT DEFAULT '',
            handled_by TEXT DEFAULT '',
            action_time TEXT DEFAULT '',
            FOREIGN KEY(driver_id) REFERENCES drivers(id),
            FOREIGN KEY(car_id) REFERENCES cars(id)
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_emg_read ON emergency_reports(is_read)")

        # GARAGE RECORDS
        c.execute("""CREATE TABLE IF NOT EXISTS garage_records(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER NOT NULL,
            car_id INTEGER,
            location TEXT NOT NULL,
            recorded_at TEXT NOT NULL,
            notes TEXT DEFAULT ''
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_garage_driver ON garage_records(driver_id)")

        # SETTINGS
        c.execute("""CREATE TABLE IF NOT EXISTS app_settings(
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )""")

        # Safe migrations for existing columns
        _safe_add_columns(c)

        # Seed admin if missing
        c.execute("SELECT id FROM users WHERE username='admin'")
        if not c.fetchone():
            pw = bcrypt.hashpw("ChangeMe123!".encode(), bcrypt.gensalt()).decode()
            c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)", ("admin", pw, "admin"))
            log.warning("⚠️  Created default admin — CHANGE PASSWORD IMMEDIATELY")

        # Seed price settings
        for ws_type in ["fuel","oil","filter","tire","battery","belt","other"]:
            c.execute("INSERT OR IGNORE INTO app_settings(key,value,updated_at) VALUES(?,?,?)",
                      (f"price_{ws_type}", "0", datetime.utcnow().isoformat() + "Z"))

def _safe_add_columns(c):
    """Add columns that may not exist in older databases."""
    existing = {}
    for tbl in ["users","drivers","emergency_reports","trips","workshop_records"]:
        try:
            c.execute(f"PRAGMA table_info({tbl})")
            existing[tbl] = {row["name"] for row in c.fetchall()}
        except Exception:
            existing[tbl] = set()

    additions = {
        "users":              [("refresh_token","TEXT"),("refresh_exp","TEXT"),("last_login","TEXT")],
        "emergency_reports":  [("audio_url","TEXT DEFAULT ''"),("is_handled","INTEGER DEFAULT 0"),
                               ("action_taken","TEXT DEFAULT ''"),("handled_by","TEXT DEFAULT ''"),
                               ("action_time","TEXT DEFAULT ''"),("location","TEXT DEFAULT ''")],
        "trips":              [("garage_location","TEXT DEFAULT ''")],
        "workshop_records":   [("location","TEXT DEFAULT ''"),("operation_type","TEXT DEFAULT ''"),
                               ("vehicle_id","INTEGER"),("odometer_reading","REAL"),
                               ("description","TEXT DEFAULT ''"),("tire_action","TEXT DEFAULT ''")],
        "drivers":            [("national_id","TEXT DEFAULT ''"),("birth_date","TEXT DEFAULT ''"),
                               ("driver_license_expiry","TEXT DEFAULT ''"),
                               ("vehicle_license_expiry","TEXT DEFAULT ''")],
    }
    for tbl, cols in additions.items():
        for col, col_type in cols:
            if col not in existing.get(tbl, set()):
                try:
                    c.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} {col_type}")
                except Exception:
                    pass

# ══════════════════════════════════════════════════════
# 4. AUTH — Access + Refresh Token system
# ══════════════════════════════════════════════════════

def _hash(pw: str) -> str:
    return bcrypt.hashpw(pw.encode(), bcrypt.gensalt(rounds=12)).decode()

def _verify(plain: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(plain.encode(), hashed.encode())
    except Exception:
        return False

def validate_password(pw: str):
    """No restrictions — any password accepted."""
    if not pw:
        raise HTTPException(400, "كلمة المرور مطلوبة")

def create_access_token(payload: dict) -> str:
    d = payload.copy()
    d["exp"] = datetime.utcnow() + timedelta(minutes=ACCESS_EXP_MIN)
    d["type"] = "access"
    return jwt.encode(d, SECRET_KEY, algorithm=ALGORITHM)

def create_refresh_token() -> tuple[str, str]:
    """Returns (token, expiry_iso)."""
    token = str(uuid.uuid4())
    exp   = (datetime.utcnow() + timedelta(days=REFRESH_EXP_DAYS)).isoformat()
    return token, exp

def verify_access_token(token: str) -> dict:
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        if payload.get("type") != "access":
            raise HTTPException(401, "نوع التوكن غير صحيح")
        return payload
    except jwt.ExpiredSignatureError:
        raise HTTPException(401, "انتهت صلاحية الجلسة — يرجى تسجيل الدخول مجدداً")
    except jwt.InvalidTokenError:
        raise HTTPException(401, "توكن غير صالح")

security = HTTPBearer(auto_error=False)

async def get_user(
    cred: Optional[HTTPAuthorizationCredentials] = Depends(security)
) -> Dict[str, Any]:
    if not cred:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "غير مصادق")
    payload = verify_access_token(cred.credentials)
    if "user_id" not in payload or "role" not in payload:
        raise HTTPException(401, "بيانات التوكن غير مكتملة")
    result = {
        "user_id":  payload["user_id"],
        "role":     payload["role"],
        "username": payload.get("username"),
    }
    if result["role"] == "driver":
        with get_db() as conn:
            c = conn.cursor()
            c.execute("SELECT id FROM drivers WHERE user_id=?", (result["user_id"],))
            drv = c.fetchone()
            result["driver_id"] = drv["id"] if drv else None
    return result

def require_admin(cu: dict = Depends(get_user)):
    if cu["role"] != "admin":
        raise HTTPException(status.HTTP_403_FORBIDDEN, "صلاحيات المدير مطلوبة")
    return cu

# ══════════════════════════════════════════════════════
# 5. RATE LIMITER
# ══════════════════════════════════════════════════════

limiter = Limiter(key_func=get_remote_address, default_limits=["200/minute"])

# ══════════════════════════════════════════════════════
# 6. SECURITY HEADERS MIDDLEWARE
# ══════════════════════════════════════════════════════

class SecurityHeadersMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        response.headers["X-Content-Type-Options"]  = "nosniff"
        response.headers["X-Frame-Options"]         = "DENY"
        response.headers["X-XSS-Protection"]        = "1; mode=block"
        response.headers["Referrer-Policy"]         = "strict-origin-when-cross-origin"
        response.headers["Permissions-Policy"]      = "geolocation=(self), microphone=(self)"
        if ENVIRONMENT == "production":
            response.headers["Strict-Transport-Security"] = "max-age=63072000; includeSubDomains"
        return response

# ══════════════════════════════════════════════════════
# 7. PYDANTIC MODELS
# ══════════════════════════════════════════════════════

class LoginReq(BaseModel):
    username: str
    password: str

    @validator("username")
    def sanitize_username(cls, v):
        v = v.strip()
        if len(v) < 2 or len(v) > 100:
            raise ValueError("اسم المستخدم يجب أن يكون بين 2 و100 حرف")
        return v

class RefreshReq(BaseModel):
    refresh_token: str

class UserResp(BaseModel):
    id: int; username: str; role: str; driver_id: Optional[int] = None

class LoginResp(BaseModel):
    access_token: str; refresh_token: str; token_type: str; user: UserResp

class DriverCreate(BaseModel):
    name: str; phone: str; username: str; password: str
    status: str = "active"
    national_id: Optional[str] = ""
    birth_date: Optional[str] = None
    driver_license_expiry: Optional[str] = None
    vehicle_license_expiry: Optional[str] = None

class DriverUpdate(BaseModel):
    name: str; phone: str; status: str
    national_id: Optional[str] = ""
    birth_date: Optional[str] = ""
    driver_license_expiry: Optional[str] = ""
    vehicle_license_expiry: Optional[str] = ""

class DriverResp(BaseModel):
    id: int; name: str; phone: str; status: str
    user_id: Optional[int] = None; username: Optional[str] = None
    national_id: Optional[str] = None; birth_date: Optional[str] = None
    driver_license_expiry: Optional[str] = None; vehicle_license_expiry: Optional[str] = None

class CarCreate(BaseModel):
    plate: str; model: str; status: str = "available"
    chassis: Optional[str] = ""; car_license_expiry: Optional[str] = ""; sector: Optional[str] = ""

class CarUpdate(BaseModel):
    plate: str; model: str; status: str
    chassis: Optional[str] = ""; car_license_expiry: Optional[str] = ""; sector: Optional[str] = ""

class CarResp(BaseModel):
    id: int; plate: str; model: str; status: str
    chassis: Optional[str] = None; car_license_expiry: Optional[str] = None; sector: Optional[str] = None

class PermCreate(BaseModel):
    driver_id: int; car_id: int

class PermResp(BaseModel):
    id: int; driver_id: int; car_id: int

class TripStart(BaseModel):
    driver_id: int; car_id: int; start_odometer: float
    start_location: str = ""

class TripEnd(BaseModel):
    trip_id: int; end_odometer: float
    end_location: str = ""; notes: Optional[str] = ""

class TripResp(BaseModel):
    id: int; driver_id: int; car_id: int
    start_time: Optional[str]; end_time: Optional[str]
    start_odometer: float; end_odometer: Optional[float]
    start_location: Optional[str]; end_location: Optional[str]
    garage_location: Optional[str]; notes: Optional[str]

class UserCreate(BaseModel):
    username: str; password: str; role: str

    # No password restrictions

class WorkshopCreate(BaseModel):
    driver_id: int; type: str; quantity: Optional[float] = None
    price: float = 0; notes: Optional[str] = ""
    operation_type: Optional[str] = ""; vehicle_id: Optional[int] = None
    odometer_reading: Optional[float] = None; description: Optional[str] = ""
    tire_action: Optional[str] = ""; location: Optional[str] = ""

class EmergencyCreate(BaseModel):
    driver_id: int; car_id: Optional[int] = None
    type: str; notes: Optional[str] = ""
    location: Optional[str] = ""
    audio_data: Optional[str] = ""   # base64 from frontend (kept for compatibility)

class EmergencyAction(BaseModel):
    is_handled: bool = True; action_taken: str = ""; handled_by: str = ""

class GarageRecordCreate(BaseModel):
    driver_id: int; car_id: Optional[int] = None
    location: str; notes: Optional[str] = ""
    odometer: Optional[float] = None

class PaginationParams(BaseModel):
    page: int = 1; limit: int = 50

    @validator("limit")
    def cap_limit(cls, v):
        return min(v, 200)

# ══════════════════════════════════════════════════════
# 8. APP SETUP
# ══════════════════════════════════════════════════════

app = FastAPI(
    title="Fleet Management API",
    version="2.0.0",
    docs_url="/docs" if ENVIRONMENT != "production" else None,
    redoc_url=None,
)

# Startup
ADMINS = [
    ("Eng mohamed mansour", "mo@mansour241"),
    ("Eng mohamed sayed",   "mo@sayed11214123"),
    ("Eng abdelrhman sayed","abdo@11214123"),
]

@app.on_event("startup")
async def startup():
    migrate_db()
    # Ensure correct admin accounts exist
    with get_db() as conn:
        c = conn.cursor()
        # Remove old admins not in the list
        c.execute("SELECT id,username FROM users WHERE role='admin'")
        existing = {r['username']:r['id'] for r in c.fetchall()}
        allowed  = {u for u,_ in ADMINS}
        for uname, uid in existing.items():
            if uname not in allowed:
                c.execute("DELETE FROM users WHERE id=?", (uid,))
                log.info(f"Removed old admin: {uname}")
        # Add/update new admins
        for uname, pw in ADMINS:
            pw_hash = bcrypt.hashpw(pw.encode(), bcrypt.gensalt()).decode()
            c.execute("SELECT id FROM users WHERE username=?", (uname,))
            if c.fetchone():
                c.execute("UPDATE users SET password=? WHERE username=?", (pw_hash, uname))
            else:
                c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)", (uname, pw_hash, "admin"))
        log.info("✅ Admin accounts synced")
    log.info("🚀 Fleet Management API started")

# Rate limiter
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# Security headers
app.add_middleware(SecurityHeadersMiddleware)

# CORS — restricted to allowed origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET","POST","PUT","DELETE"],
    allow_headers=["Authorization","Content-Type"],
)

# Serve uploads as static (audio files)
app.mount("/uploads", StaticFiles(directory=str(UPLOAD_DIR)), name="uploads")

# Serve frontend
@app.get("/")
async def root():
    return FileResponse("index.html")

# ══════════════════════════════════════════════════════
# 9. GLOBAL EXCEPTION HANDLER
# ══════════════════════════════════════════════════════

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    log.error(f"Unhandled error | path={request.url.path} | {type(exc).__name__}: {exc}")
    return JSONResponse(
        status_code=500,
        content={"detail": "خطأ داخلي في الخادم — يرجى المحاولة لاحقاً"}
    )

# ══════════════════════════════════════════════════════
# 10. HELPERS
# ══════════════════════════════════════════════════════

def enrich(d: dict) -> dict:
    for k in ["start_location","end_location","garage_location","notes"]:
        d.setdefault(k, None)
    return d

def paginate(query_result: list, page: int, limit: int) -> dict:
    total = len(query_result)
    start = (page - 1) * limit
    return {
        "items":   query_result[start:start+limit],
        "total":   total,
        "page":    page,
        "limit":   limit,
        "pages":   (total + limit - 1) // limit,
    }

# ══════════════════════════════════════════════════════
# 11. AUTH ENDPOINTS
# ══════════════════════════════════════════════════════

@app.post("/login", response_model=LoginResp)
@limiter.limit("10/minute")   # strict: prevent brute-force
async def login(request: Request, data: LoginReq):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id,username,password,role FROM users WHERE username=?", (data.username,))
        u = c.fetchone()
        if not u or not _verify(data.password, u["password"]):
            log_event("login_failed", username=data.username, ip=request.client.host)
            raise HTTPException(401, "اسم المستخدم أو كلمة المرور غير صحيحة")

        driver_id = None
        if u["role"] == "driver":
            c.execute("SELECT id FROM drivers WHERE user_id=?", (u["id"],))
            drv = c.fetchone()
            driver_id = drv["id"] if drv else None

        # Tokens
        access_token = create_access_token({
            "user_id": u["id"], "role": u["role"], "username": u["username"]
        })
        refresh_token, refresh_exp = create_refresh_token()

        # Persist refresh token (one per user)
        c.execute("UPDATE users SET refresh_token=?,refresh_exp=?,last_login=? WHERE id=?",
                  (refresh_token, refresh_exp, datetime.utcnow().isoformat() + "Z", u["id"]))

        log_event("login_success", user_id=u["id"], role=u["role"])
        return LoginResp(
            access_token=access_token,
            refresh_token=refresh_token,
            token_type="bearer",
            user=UserResp(id=u["id"], username=u["username"],
                          role=u["role"], driver_id=driver_id)
        )

@app.post("/token/refresh")
@limiter.limit("30/minute")
async def refresh_token(request: Request, body: RefreshReq):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id,username,role,refresh_exp FROM users WHERE refresh_token=?",
                  (body.refresh_token,))
        u = c.fetchone()
        if not u:
            raise HTTPException(401, "Refresh token غير صالح")
        if datetime.fromisoformat(u["refresh_exp"]) < datetime.utcnow():
            raise HTTPException(401, "انتهت صلاحية الجلسة — يرجى تسجيل الدخول مجدداً")

        new_access = create_access_token({
            "user_id": u["id"], "role": u["role"], "username": u["username"]
        })
        new_refresh, new_exp = create_refresh_token()
        c.execute("UPDATE users SET refresh_token=?,refresh_exp=? WHERE id=?",
                  (new_refresh, new_exp, u["id"]))
        return {"access_token": new_access, "refresh_token": new_refresh, "token_type": "bearer"}

@app.post("/logout")
async def logout(cu: dict = Depends(get_user)):
    with get_db() as conn:
        conn.cursor().execute("UPDATE users SET refresh_token=NULL,refresh_exp=NULL WHERE id=?",
                              (cu["user_id"],))
    return {"ok": True}

# ══════════════════════════════════════════════════════
# 12. AUDIO UPLOAD — File system, NOT base64 in DB
# ══════════════════════════════════════════════════════

AUDIO_MAGIC = {
    b"OggS": "audio/ogg",
    b"\x1aE\xdf\xa3": "audio/webm",
    b"ftyp": "audio/mp4",
    b"ID3":  "audio/mpeg",
    b"RIFF": "audio/wav",
}

def validate_audio_file(file: UploadFile, content: bytes) -> str:
    """Validate MIME type by magic bytes, not just filename."""
    if len(content) > MAX_AUDIO_MB * 1024 * 1024:
        raise HTTPException(400, f"حجم الملف يتجاوز {MAX_AUDIO_MB} MB")
    for magic, mime in AUDIO_MAGIC.items():
        if content[:len(magic)] == magic:
            return mime
    # MP4 ftyp is at offset 4
    if len(content) >= 8 and content[4:8] == b"ftyp":
        return "audio/mp4"
    raise HTTPException(400, "نوع الملف غير مدعوم — يجب أن يكون صوتاً (webm, ogg, mp4, wav)")

@app.post("/emergency/upload-audio")
@limiter.limit("20/minute")
async def upload_audio(
    request: Request,
    file: UploadFile = File(...),
    cu: dict = Depends(get_user)
):
    """Upload audio clip — returns URL to store in emergency_reports."""
    content = await file.read()
    mime = validate_audio_file(file, content)
    ext = mime.split("/")[1].split(";")[0]
    safe_name = f"{uuid.uuid4().hex}.{ext}"
    dest = UPLOAD_DIR / safe_name
    dest.write_bytes(content)
    log_event("audio_uploaded", user_id=cu["user_id"], size=len(content), mime=mime)
    return {"audio_url": f"/uploads/{safe_name}"}

# ══════════════════════════════════════════════════════
# 13. DRIVERS
# ══════════════════════════════════════════════════════

@app.get("/drivers", response_model=List[DriverResp])
async def get_drivers(cu: dict = Depends(get_user)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT d.*,u.username FROM drivers d
                     LEFT JOIN users u ON d.user_id=u.id
                     ORDER BY d.id DESC LIMIT 500""")
        return [dict(r) for r in c.fetchall()]

@app.post("/drivers", response_model=DriverResp)
async def create_driver(driver: DriverCreate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id FROM users WHERE username=?", (driver.username,))
        if c.fetchone():
            raise HTTPException(400, "اسم المستخدم موجود مسبقاً")
        c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                  (driver.username, _hash(driver.password), "driver"))
        uid = c.lastrowid
        c.execute("""INSERT INTO drivers(name,phone,status,user_id,national_id,
                     birth_date,driver_license_expiry,vehicle_license_expiry)
                     VALUES(?,?,?,?,?,?,?,?)""",
                  (driver.name, driver.phone, driver.status, uid,
                   driver.national_id or "", driver.birth_date or "",
                   driver.driver_license_expiry or "", driver.vehicle_license_expiry or ""))
        did = c.lastrowid
        log_event("driver_created", driver_id=did, admin=cu["username"])
        return {"id":did,"name":driver.name,"phone":driver.phone,"status":driver.status,
                "user_id":uid,"username":driver.username,
                "national_id":driver.national_id,"birth_date":driver.birth_date,
                "driver_license_expiry":driver.driver_license_expiry,
                "vehicle_license_expiry":driver.vehicle_license_expiry}

@app.put("/drivers/{did}", response_model=DriverResp)
async def update_driver(did: int, driver: DriverUpdate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""UPDATE drivers SET name=?,phone=?,status=?,national_id=?,
                     birth_date=?,driver_license_expiry=?,vehicle_license_expiry=?
                     WHERE id=?""",
                  (driver.name,driver.phone,driver.status,
                   driver.national_id or "",driver.birth_date or "",
                   driver.driver_license_expiry or "",driver.vehicle_license_expiry or "",did))
        if c.rowcount == 0:
            raise HTTPException(404, "السائق غير موجود")
        c.execute("SELECT user_id FROM drivers WHERE id=?", (did,))
        r = c.fetchone()
        return {"id":did,"name":driver.name,"phone":driver.phone,"status":driver.status,
                "user_id":r["user_id"] if r else None,
                "national_id":driver.national_id,"birth_date":driver.birth_date,
                "driver_license_expiry":driver.driver_license_expiry,
                "vehicle_license_expiry":driver.vehicle_license_expiry}

@app.delete("/drivers/{did}")
async def delete_driver(did: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT user_id FROM drivers WHERE id=?", (did,))
        drv = c.fetchone()
        for tbl in ["driver_car_permissions","trips","workshop_records"]:
            c.execute(f"DELETE FROM {tbl} WHERE driver_id=?", (did,))
        c.execute("DELETE FROM drivers WHERE id=?", (did,))
        if drv and drv["user_id"]:
            c.execute("DELETE FROM users WHERE id=?", (drv["user_id"],))
        log_event("driver_deleted", driver_id=did, admin=cu["username"])
        return {"message": "تم الحذف"}

@app.put("/drivers/{did}/credentials")
async def update_credentials(did: int, body: dict, cu: dict = Depends(require_admin)):
    new_username = (body.get("username") or "").strip()
    new_password = (body.get("password") or "").strip()
    if not new_username:
        raise HTTPException(400, "اسم المستخدم مطلوب")
    if new_password:
        validate_password(new_password)
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT user_id FROM drivers WHERE id=?", (did,))
        drv = c.fetchone()
        if not drv or not drv["user_id"]:
            raise HTTPException(404, "المستخدم غير موجود")
        uid = drv["user_id"]
        c.execute("SELECT id FROM users WHERE username=? AND id!=?", (new_username, uid))
        if c.fetchone():
            raise HTTPException(400, "اسم المستخدم موجود مسبقاً")
        if new_password:
            c.execute("UPDATE users SET username=?,password=? WHERE id=?",
                      (new_username, _hash(new_password), uid))
        else:
            c.execute("UPDATE users SET username=? WHERE id=?", (new_username, uid))
        return {"message": "تم التحديث", "username": new_username}

@app.get("/drivers/{did}/primary-car")
async def primary_car(did: int, cu: dict = Depends(get_user)):
    if cu["role"] != "admin" and cu.get("driver_id") != did:
        raise HTTPException(403, "غير مصرح")
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT c.id,c.plate,c.model,c.status FROM cars c
                     INNER JOIN driver_car_permissions p ON c.id=p.car_id
                     WHERE p.driver_id=? ORDER BY p.id ASC LIMIT 1""", (did,))
        row = c.fetchone()
        return dict(row) if row else None

@app.get("/drivers/{did}/allowed-cars")
async def allowed_cars(did: int, cu: dict = Depends(get_user)):
    if cu["role"] != "admin" and cu.get("driver_id") != did:
        raise HTTPException(403, "غير مصرح")
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT c.id,c.plate,c.model,c.status FROM cars c
                     INNER JOIN driver_car_permissions p ON c.id=p.car_id
                     WHERE p.driver_id=?""", (did,))
        return [dict(r) for r in c.fetchall()]

# ══════════════════════════════════════════════════════
# 14. CARS
# ══════════════════════════════════════════════════════

@app.get("/cars", response_model=List[CarResp])
async def get_cars(cu: dict = Depends(get_user)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id,plate,model,status,chassis,car_license_expiry,sector FROM cars ORDER BY id DESC LIMIT 500")
        rows=[]
        for r in c.fetchall():
            d=dict(r); d.setdefault("chassis",""); d.setdefault("car_license_expiry",""); d.setdefault("sector","")
            rows.append(d)
        return rows

@app.post("/cars", response_model=CarResp)
async def create_car(car: CarCreate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        try:
            c.execute("INSERT INTO cars(plate,model,status,chassis,car_license_expiry,sector) VALUES(?,?,?,?,?,?)",
                      (car.plate, car.model, car.status, car.chassis or "", car.car_license_expiry or "", car.sector or ""))
            return {"id": c.lastrowid, "plate": car.plate, "model": car.model, "status": car.status,
                    "chassis": car.chassis or "", "car_license_expiry": car.car_license_expiry or "", "sector": car.sector or ""}
        except sqlite3.IntegrityError:
            raise HTTPException(400, "رقم اللوحة موجود مسبقاً")

@app.put("/cars/{cid}", response_model=CarResp)
async def update_car(cid: int, car: CarUpdate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("UPDATE cars SET plate=?,model=?,status=?,chassis=?,car_license_expiry=?,sector=? WHERE id=?",
                  (car.plate, car.model, car.status, car.chassis or "", car.car_license_expiry or "", car.sector or "", cid))
        if c.rowcount == 0:
            raise HTTPException(404, "المركبة غير موجودة")
        return {"id": cid, "plate": car.plate, "model": car.model, "status": car.status,
                "chassis": car.chassis or "", "car_license_expiry": car.car_license_expiry or "", "sector": car.sector or ""}

@app.delete("/cars/{cid}")
async def delete_car(cid: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("DELETE FROM driver_car_permissions WHERE car_id=?", (cid,))
        c.execute("DELETE FROM cars WHERE id=?", (cid,))
        return {"message": "تم الحذف"}

# ══════════════════════════════════════════════════════
# 15. PERMISSIONS
# ══════════════════════════════════════════════════════

@app.get("/permissions", response_model=List[PermResp])
async def get_permissions(cu: dict = Depends(get_user)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id,driver_id,car_id FROM driver_car_permissions ORDER BY id DESC")
        return [dict(r) for r in c.fetchall()]

@app.post("/permissions", response_model=PermResp)
async def create_permission(perm: PermCreate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id FROM driver_car_permissions WHERE driver_id=? AND car_id=?",
                  (perm.driver_id, perm.car_id))
        if c.fetchone():
            raise HTTPException(400, "الصلاحية موجودة مسبقاً")
        c.execute("INSERT INTO driver_car_permissions(driver_id,car_id) VALUES(?,?)",
                  (perm.driver_id, perm.car_id))
        return {"id": c.lastrowid, "driver_id": perm.driver_id, "car_id": perm.car_id}

@app.delete("/permissions/{pid}")
async def delete_permission(pid: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        conn.cursor().execute("DELETE FROM driver_car_permissions WHERE id=?", (pid,))
        return {"ok": True}

# ══════════════════════════════════════════════════════
# 16. TRIPS — with pagination
# ══════════════════════════════════════════════════════

@app.get("/trips")
async def get_trips(cu: dict = Depends(get_user)):
    with get_db() as conn:
        c = conn.cursor()
        if cu["role"] == "admin":
            c.execute("SELECT * FROM trips ORDER BY id DESC LIMIT 500")
        else:
            did = cu.get("driver_id")
            if not did: return []
            c.execute("SELECT * FROM trips WHERE driver_id=? ORDER BY id DESC LIMIT 200", (did,))
        return [enrich(dict(r)) for r in c.fetchall()]

@app.post("/trips/start", response_model=TripResp)
async def start_trip(trip: TripStart, cu: dict = Depends(get_user)):
    if trip.start_odometer < 0:
        raise HTTPException(400, "العداد يجب أن يكون موجباً")
    with get_db() as conn:
        c = conn.cursor()
        if cu["role"] != "admin":
            did = cu.get("driver_id")
            if trip.driver_id != did:
                raise HTTPException(403, "لا يمكنك بدء رحلة لسائق آخر")
            c.execute("SELECT id FROM driver_car_permissions WHERE driver_id=? AND car_id=?",
                      (trip.driver_id, trip.car_id))
            if not c.fetchone():
                raise HTTPException(403, "غير مصرح بقيادة هذه المركبة")
        c.execute("SELECT id FROM trips WHERE driver_id=? AND end_time IS NULL", (trip.driver_id,))
        if c.fetchone():
            raise HTTPException(400, "لديك رحلة نشطة — أنهها أولاً")
        now = datetime.utcnow().isoformat() + "Z"
        c.execute("""INSERT INTO trips(driver_id,car_id,start_time,start_odometer,start_location,garage_location)
                     VALUES(?,?,?,?,?,'')""",
                  (trip.driver_id, trip.car_id, now, trip.start_odometer, trip.start_location))
        tid = c.lastrowid
        log_event("trip_started", trip_id=tid, driver_id=trip.driver_id)
        return {"id":tid,"driver_id":trip.driver_id,"car_id":trip.car_id,
                "start_time":now,"end_time":None,"start_odometer":trip.start_odometer,
                "end_odometer":None,"start_location":trip.start_location,
                "end_location":None,"garage_location":None,"notes":None}

@app.post("/trips/end", response_model=TripResp)
async def end_trip(te: TripEnd, cu: dict = Depends(get_user)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT driver_id,start_odometer FROM trips WHERE id=? AND end_time IS NULL",
                  (te.trip_id,))
        trip = c.fetchone()
        if not trip:
            raise HTTPException(404, "الرحلة غير موجودة أو منتهية بالفعل")
        if cu["role"] != "admin" and trip["driver_id"] != cu.get("driver_id"):
            raise HTTPException(403, "لا يمكنك إنهاء رحلة سائق آخر")
        if te.end_odometer <= trip["start_odometer"]:
            raise HTTPException(400, f"عداد النهاية يجب أن يكون أكبر من {trip['start_odometer']}")
        now = datetime.utcnow().isoformat() + "Z"
        c.execute("UPDATE trips SET end_time=?,end_odometer=?,end_location=?,notes=? WHERE id=?",
                  (now, te.end_odometer, te.end_location, te.notes, te.trip_id))
        c.execute("SELECT * FROM trips WHERE id=?", (te.trip_id,))
        log_event("trip_ended", trip_id=te.trip_id)
        return enrich(dict(c.fetchone()))

@app.put("/trips/{tid}/garage")
async def set_garage(tid: int, body: dict, cu: dict = Depends(get_user)):
    loc = (body.get("garage_location") or "").strip()
    if not loc:
        raise HTTPException(400, "موقع الجراج مطلوب")
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT driver_id FROM trips WHERE id=?", (tid,))
        trip = c.fetchone()
        if not trip:
            raise HTTPException(404, "الرحلة غير موجودة")
        if cu["role"] != "admin" and trip["driver_id"] != cu.get("driver_id"):
            raise HTTPException(403, "غير مصرح")
        c.execute("UPDATE trips SET garage_location=? WHERE id=?", (loc, tid))
        return {"ok": True}

@app.post("/trips/end-current", response_model=TripResp)
async def end_current(body: dict, cu: dict = Depends(get_user)):
    did = cu.get("driver_id") if cu["role"] == "driver" else body.get("driver_id")
    if not did:
        raise HTTPException(400, "driver_id مطلوب")
    end_odo = body.get("end_odometer")
    loc     = body.get("location", "")
    notes   = body.get("notes", "")
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id,start_odometer FROM trips WHERE driver_id=? AND end_time IS NULL", (did,))
        active = c.fetchone()
        if not active:
            raise HTTPException(404, "لا توجد رحلة نشطة")
        if end_odo is None or end_odo <= active["start_odometer"]:
            raise HTTPException(400, "عداد النهاية يجب أن يكون أكبر من البداية")
        c.execute("UPDATE trips SET end_time=?,end_odometer=?,end_location=?,notes=? WHERE id=?",
                  (now, end_odo, loc, notes, active["id"]))
        c.execute("SELECT * FROM trips WHERE id=?", (active["id"],))
        return enrich(dict(c.fetchone()))

# ══════════════════════════════════════════════════════
# 17. USERS
# ══════════════════════════════════════════════════════

@app.post("/users")
async def create_user(user: UserCreate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        try:
            c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                      (user.username, _hash(user.password), user.role))
            uid = c.lastrowid
            if user.role == "driver":
                c.execute("INSERT INTO drivers(name,phone,status,user_id) VALUES(?,?,?,?)",
                          (user.username, "", "active", uid))
            return {"message": f"تم إنشاء {user.username}", "user_id": uid}
        except sqlite3.IntegrityError:
            raise HTTPException(400, "اسم المستخدم موجود مسبقاً")

@app.get("/users/list")
async def list_users(cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id,username,role FROM users ORDER BY id DESC LIMIT 200")
        return [dict(r) for r in c.fetchall()]

@app.delete("/users/{uid}")
async def delete_user(uid: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id,username FROM users WHERE id=?", (uid,))
        u = c.fetchone()
        if not u:
            raise HTTPException(404, "المستخدم غير موجود")
        if u["username"] == "admin":
            raise HTTPException(400, "لا يمكن حذف الحساب الرئيسي")
        c.execute("UPDATE drivers SET user_id=NULL WHERE user_id=?", (uid,))
        c.execute("DELETE FROM users WHERE id=?", (uid,))
        return {"message": "تم الحذف"}

# ══════════════════════════════════════════════════════
# 18. WORKSHOPS
# ══════════════════════════════════════════════════════

@app.post("/workshops")
async def create_workshop(rec: WorkshopCreate, cu: dict = Depends(get_user)):
    if rec.type not in ["oil","filter","tire","battery","belt","other","fuel"]:
        raise HTTPException(400, "نوع غير صالح")
    if cu["role"] != "admin" and rec.driver_id != cu.get("driver_id"):
        raise HTTPException(403, "غير مصرح")
    # Enforce admin price for drivers
    final_price = rec.price if cu["role"] == "admin" else 0.0
    if cu["role"] == "driver":
        with get_db() as conn:
            c = conn.cursor()
            c.execute("SELECT value FROM app_settings WHERE key=?", (f"price_{rec.type}",))
            row = c.fetchone()
            final_price = float(row["value"]) if row else 0.0
    if final_price < 0:
        raise HTTPException(400, "السعر لا يمكن أن يكون سالباً")
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""INSERT INTO workshop_records
                     (driver_id,type,quantity,price,notes,created_at,
                      operation_type,vehicle_id,odometer_reading,description,tire_action,location)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (rec.driver_id, rec.type, rec.quantity, final_price, rec.notes or "", now,
                   rec.operation_type or "", rec.vehicle_id, rec.odometer_reading,
                   rec.description or "", rec.tire_action or "", rec.location or ""))
        rid = c.lastrowid
        vp = None
        if rec.vehicle_id:
            c.execute("SELECT plate FROM cars WHERE id=?", (rec.vehicle_id,))
            cv = c.fetchone()
            vp = cv["plate"] if cv else None
        return {"id":rid,"driver_id":rec.driver_id,"type":rec.type,"quantity":rec.quantity,
                "price":final_price,"notes":rec.notes or "","created_at":now,
                "operation_type":rec.operation_type or "","vehicle_id":rec.vehicle_id,
                "odometer_reading":rec.odometer_reading,"description":rec.description or "",
                "tire_action":rec.tire_action or "","vehicle_plate":vp,"location":rec.location or ""}

@app.get("/workshops")
async def get_workshops(cu: dict = Depends(get_user)):
    q = """SELECT w.*,d.name as driver_name,c.plate as vehicle_plate
           FROM workshop_records w
           LEFT JOIN drivers d ON w.driver_id=d.id
           LEFT JOIN cars c ON w.vehicle_id=c.id"""
    with get_db() as conn:
        c = conn.cursor()
        if cu["role"] == "admin":
            c.execute(q + " ORDER BY w.id DESC LIMIT 500")
        else:
            did = cu.get("driver_id")
            if not did: return []
            c.execute(q + " WHERE w.driver_id=? ORDER BY w.id DESC LIMIT 200", (did,))
        rows = []
        for r in c.fetchall():
            d = dict(r)
            for k in ["operation_type","vehicle_id","odometer_reading",
                      "description","vehicle_plate","tire_action","location"]:
                d.setdefault(k, None)
            rows.append(d)
        return rows

# ══════════════════════════════════════════════════════
# 19. GARAGE RECORDS
# ══════════════════════════════════════════════════════

@app.post("/garage/record")
async def create_garage_record(rec: GarageRecordCreate, cu: dict = Depends(get_user)):
    if not rec.location.strip():
        raise HTTPException(400, "موقع الجراج مطلوب")
    if cu["role"] != "admin" and rec.driver_id != cu.get("driver_id"):
        raise HTTPException(403, "غير مصرح")
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        # Auto-end active trip with garage data
        c.execute("SELECT id,start_odometer FROM trips WHERE driver_id=? AND end_time IS NULL", (rec.driver_id,))
        active = c.fetchone()
        auto_ended = False
        if active:
            start_odo = active["start_odometer"] or 0
            # Validate odometer
            end_odo = None
            if rec.odometer and float(rec.odometer) > start_odo:
                end_odo = float(rec.odometer)

            # Always end the trip — location and time from garage
            # odometer only if valid
            c.execute(
                "UPDATE trips SET end_time=?,end_odometer=?,end_location=?,garage_location=? WHERE id=?",
                (now, end_odo, rec.location, rec.location, active["id"])
            )
            auto_ended = True
        c.execute("INSERT INTO garage_records(driver_id,car_id,location,recorded_at,notes) VALUES(?,?,?,?,?)",
                  (rec.driver_id, rec.car_id, rec.location, now, rec.notes or ""))
        rid = c.lastrowid
        plate = dn = None
        if rec.car_id:
            c.execute("SELECT plate FROM cars WHERE id=?", (rec.car_id,))
            r = c.fetchone(); plate = r["plate"] if r else None
        c.execute("SELECT name FROM drivers WHERE id=?", (rec.driver_id,))
        dr = c.fetchone(); dn = dr["name"] if dr else None
        return {"id":rid,"driver_id":rec.driver_id,"car_id":rec.car_id,
                "location":rec.location,"recorded_at":now,"notes":rec.notes or "",
                "driver_name":dn,"car_plate":plate,"trip_auto_ended":auto_ended}

@app.get("/garage/records")
async def get_garage_records(cu: dict = Depends(get_user)):
    with get_db() as conn:
        c = conn.cursor()
        q = """SELECT g.*,d.name as driver_name,c.plate as car_plate
               FROM garage_records g
               LEFT JOIN drivers d ON g.driver_id=d.id
               LEFT JOIN cars c ON g.car_id=c.id"""
        if cu["role"] == "admin":
            c.execute(q + " ORDER BY g.id DESC LIMIT 500")
        else:
            did = cu.get("driver_id")
            if not did: return []
            c.execute(q + " WHERE g.driver_id=? ORDER BY g.id DESC LIMIT 200", (did,))
        return [dict(r) for r in c.fetchall()]

@app.get("/garage/latest")
async def garage_latest(cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT g.*,d.name as driver_name,c.plate as car_plate
                     FROM garage_records g
                     INNER JOIN(SELECT driver_id,MAX(id) as max_id
                                FROM garage_records GROUP BY driver_id) latest
                       ON g.id=latest.max_id
                     LEFT JOIN drivers d ON g.driver_id=d.id
                     LEFT JOIN cars c ON g.car_id=c.id
                     ORDER BY g.recorded_at DESC""")
        return [dict(r) for r in c.fetchall()]

# ══════════════════════════════════════════════════════
# 20. EMERGENCY — Audio stored as file, not base64
# ══════════════════════════════════════════════════════

@app.post("/emergency/report")
@limiter.limit("5/minute")  # very strict on emergency endpoint
async def create_emergency(request: Request, rep: EmergencyCreate, cu: dict = Depends(get_user)):
    if rep.type not in ("emergency", "accident"):
        raise HTTPException(400, "نوع البلاغ غير صالح")
    if cu["role"] != "admin" and rep.driver_id != cu.get("driver_id"):
        raise HTTPException(403, "غير مصرح")
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        # Save base64 audio to file if provided
        audio_url = ""
        if rep.audio_data and rep.audio_data.startswith("data:audio"):
            try:
                import base64
                header, b64 = rep.audio_data.split(",", 1)
                ext = "webm" if "webm" in header else "ogg" if "ogg" in header else "mp4"
                fname = f"{uuid.uuid4().hex}.{ext}"
                (UPLOAD_DIR / fname).write_bytes(base64.b64decode(b64))
                audio_url = f"/uploads/{fname}"
            except Exception as e:
                log.warning(f"Audio save failed: {e}")
        c.execute("""INSERT INTO emergency_reports
                     (driver_id,car_id,type,audio_url,notes,created_at,is_read,location)
                     VALUES(?,?,?,?,?,?,0,?)""",
                  (rep.driver_id, rep.car_id, rep.type, audio_url, rep.notes or "", now, rep.location or ""))
        rid = c.lastrowid
        log_event("emergency_reported", report_id=rid, driver_id=rep.driver_id, type=rep.type)
        return {"id": rid, "created_at": now, "type": rep.type}

@app.get("/emergency/reports")
async def get_emergency_reports(cu: dict = Depends(get_user)):
    q = """SELECT e.*,d.name as driver_name,c.plate as car_plate
           FROM emergency_reports e
           LEFT JOIN drivers d ON e.driver_id=d.id
           LEFT JOIN cars c ON e.car_id=c.id"""
    with get_db() as conn:
        c = conn.cursor()
        if cu["role"] == "admin":
            c.execute(q + " ORDER BY e.id DESC LIMIT 500")
        else:
            did = cu.get("driver_id")
            if not did: return []
            c.execute(q + " WHERE e.driver_id=? ORDER BY e.id DESC LIMIT 200", (did,))
        rows = []
        for r in c.fetchall():
            d = dict(r)
            d["is_read"]    = bool(d.get("is_read", 0))
            d["is_handled"] = bool(d.get("is_handled", 0))
            d.setdefault("driver_name", None); d.setdefault("car_plate", None)
            d.setdefault("location", ""); d.setdefault("action_taken", "")
            d.setdefault("handled_by", ""); d.setdefault("action_time", "")
            d.setdefault("audio_url", "")
            rows.append(d)
        return rows

@app.put("/emergency/reports/{rid}/read")
async def mark_read(rid: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        conn.cursor().execute("UPDATE emergency_reports SET is_read=1 WHERE id=?", (rid,))
        return {"ok": True}

@app.put("/emergency/reports/{rid}/action")
async def update_emergency_action(rid: int, body: EmergencyAction, cu: dict = Depends(require_admin)):
    action_ts  = datetime.utcnow().isoformat() + "Z"
    admin_name = body.handled_by or cu.get("username", "admin")
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""UPDATE emergency_reports
                     SET is_handled=?,action_taken=?,handled_by=?,is_read=1,action_time=?
                     WHERE id=?""",
                  (1 if body.is_handled else 0, body.action_taken, admin_name, action_ts, rid))
        return {"ok": True, "handled_by": admin_name, "action_time": action_ts}

@app.get("/emergency/unread-count")
async def unread_count(cu: dict = Depends(get_user)):
    if cu["role"] != "admin":
        return {"count": 0}
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as cnt FROM emergency_reports WHERE is_read=0")
        return {"count": c.fetchone()["cnt"]}

# ══════════════════════════════════════════════════════
# 21. SETTINGS
# ══════════════════════════════════════════════════════

@app.get("/settings/prices")
async def get_prices(cu: dict = Depends(get_user)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT key,value FROM app_settings WHERE key LIKE 'price_%'")
        return {r["key"].replace("price_",""): float(r["value"]) for r in c.fetchall()}

@app.put("/settings/prices")
async def set_prices(body: dict, cu: dict = Depends(require_admin)):
    valid = ["fuel","oil","filter","tire","battery","belt","other"]
    now   = datetime.utcnow().isoformat() + "Z"
    saved = {}
    with get_db() as conn:
        c = conn.cursor()
        for ws_type, price in body.items():
            if ws_type not in valid:
                continue
            if not isinstance(price, (int, float)) or price < 0:
                continue
            c.execute("INSERT OR REPLACE INTO app_settings(key,value,updated_at) VALUES(?,?,?)",
                      (f"price_{ws_type}", str(float(price)), now))
            saved[ws_type] = float(price)
    return {"saved": saved}

# ══════════════════════════════════════════════════════
# 22. ADMIN — RESET DATA (protected with password)
# ══════════════════════════════════════════════════════

@app.post("/admin/reset-data")
@limiter.limit("3/hour")   # extra throttle on dangerous endpoint
async def reset_data(
    request: Request,
    body: dict,
    cu: dict = Depends(require_admin)
):
    """Requires admin to re-confirm their password before wiping all data."""
    if ENVIRONMENT == "production":
        raise HTTPException(403,
            "⛔ إعادة التهيئة غير مسموح بها في بيئة الإنتاج. "
            "استخدم أدوات قاعدة البيانات مباشرة.")
    confirm_password = body.get("confirm_password", "")
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT password FROM users WHERE id=?", (cu["user_id"],))
        u = c.fetchone()
        if not u or not _verify(confirm_password, u["password"]):
            raise HTTPException(403, "كلمة المرور غير صحيحة — لا يمكن إتمام العملية")
        # Wipe all operational data
        for tbl in ["emergency_reports","garage_records","workshop_records",
                    "trips","driver_car_permissions","drivers"]:
            c.execute(f"DELETE FROM {tbl}")
            try:
                c.execute(f"DELETE FROM sqlite_sequence WHERE name='{tbl}'")
            except Exception:
                pass
        c.execute("DELETE FROM users WHERE role != 'admin'")
        log_event("data_reset", admin=cu["username"])
        return {"message": "تم مسح جميع البيانات التشغيلية"}

# ══════════════════════════════════════════════════════
# 23. ENTRY POINT
# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 8000)),
        reload=(ENVIRONMENT != "production"),
        workers=1 if ENVIRONMENT != "production" else 4,
    )

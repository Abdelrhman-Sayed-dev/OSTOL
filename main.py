"""
Fleet Management System — Production-Ready Backend
Author: Refactored for enterprise use
"""

import base64
import csv
import io
import os
import uuid
import logging
from datetime import datetime, timedelta
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

import bcrypt
import jwt
import sqlite3
from fastapi import (
    FastAPI, Depends, HTTPException, Request,
    UploadFile, File, status
)
from fastapi.middleware.cors import CORSMiddleware
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
            role TEXT NOT NULL CHECK(role IN('superuser','admin','driver','reporter')),
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
            branch TEXT DEFAULT '',
            fixed_number TEXT DEFAULT '',
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
            car_name TEXT DEFAULT '',
            car_code TEXT DEFAULT '',
            chassis TEXT DEFAULT '',
            engine_number TEXT DEFAULT '',
            year TEXT DEFAULT '',
            project TEXT DEFAULT '',
            branch TEXT DEFAULT '',
            car_license_expiry TEXT DEFAULT '',
            equipment_type TEXT DEFAULT '',
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
        c.execute("CREATE INDEX IF NOT EXISTS idx_garage_driver_loc ON garage_records(driver_id, location, recorded_at)")

        # SETTINGS
        c.execute("""CREATE TABLE IF NOT EXISTS app_settings(
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )""")

        # AUDIT LOGS
        c.execute("""CREATE TABLE IF NOT EXISTS audit_logs(
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id    INTEGER,
            username   TEXT NOT NULL,
            role       TEXT NOT NULL,
            action     TEXT NOT NULL,
            details    TEXT DEFAULT '',
            ip_address TEXT DEFAULT '',
            created_at TEXT NOT NULL
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_audit_user    ON audit_logs(user_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_logs(created_at)")

        # Safe migrations for existing columns
        _safe_add_columns(c)

        # Seed admin if missing
        c.execute("SELECT id FROM users WHERE username='admin'")
        if not c.fetchone():
            pw = bcrypt.hashpw("ChangeMe123!".encode(), bcrypt.gensalt()).decode()
            c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)", ("admin", pw, "admin"))
            log.warning("⚠️  Created default admin — CHANGE PASSWORD IMMEDIATELY")

        # Seed price settings
        for ws_type in WORKSHOP_TYPES:
            c.execute("INSERT OR IGNORE INTO app_settings(key,value,updated_at) VALUES(?,?,?)",
                      (f"price_{ws_type}", "0", datetime.utcnow().isoformat() + "Z"))

def _safe_add_columns(c):
    """Add columns that may not exist in older databases."""
    # ── Migrate users role CHECK to include 'reporter' and 'superuser' ──
    try:
        c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='users'")
        row = c.fetchone()
        tbl_sql = row['sql'] or '' if row else ''
        needs_role_fix   = 'superuser' not in tbl_sql
        needs_unique_fix = 'username TEXT UNIQUE' in tbl_sql  # نشيل UNIQUE من username
        if needs_role_fix or needs_unique_fix:
            c.execute("PRAGMA foreign_keys=OFF")
            c.execute("""CREATE TABLE IF NOT EXISTS users_migrated(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN('superuser','admin','driver','reporter')),
                branch TEXT DEFAULT '',
                created_at TEXT DEFAULT(datetime('now')),
                last_login TEXT,
                refresh_token TEXT,
                refresh_exp TEXT
            )""")
            c.execute("INSERT OR IGNORE INTO users_migrated SELECT id,username,password,role,COALESCE(branch,''),created_at,last_login,refresh_token,refresh_exp FROM users")
            c.execute("DROP TABLE users")
            c.execute("ALTER TABLE users_migrated RENAME TO users")
            c.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")
            c.execute("PRAGMA foreign_keys=ON")
            log.info("✅ Migrated users: removed UNIQUE on username, added superuser role")
    except Exception as e:
        log.warning(f"Users migration skipped: {e}")

    # ── UNIQUE index على fixed_number في drivers (بيتجاهل الفارغ) ──
    try:
        c.execute("""CREATE UNIQUE INDEX IF NOT EXISTS idx_drivers_fixed_number
                     ON drivers(fixed_number)
                     WHERE fixed_number IS NOT NULL AND fixed_number != ''""")
        log.info("✅ UNIQUE index on drivers.fixed_number created")
    except Exception as e:
        log.warning(f"fixed_number index: {e}")

    existing = {}
    for tbl in ["users","drivers","cars","emergency_reports","trips","workshop_records"]:
        try:
            c.execute(f"PRAGMA table_info({tbl})")
            existing[tbl] = {row["name"] for row in c.fetchall()}
        except Exception:
            existing[tbl] = set()

    # Create driver_requests table if not exists
    c.execute("""CREATE TABLE IF NOT EXISTS driver_requests (
        id                  INTEGER PRIMARY KEY AUTOINCREMENT,
        driver_id           INTEGER NOT NULL,
        type                TEXT NOT NULL,
        notes               TEXT DEFAULT '',
        new_odometer        REAL,
        trip_id             INTEGER,
        status              TEXT DEFAULT 'pending',
        priority            TEXT DEFAULT 'medium',
        admin_notes         TEXT DEFAULT '',
        admin_message       TEXT DEFAULT '',
        handled_by          TEXT DEFAULT '',
        handled_at          TEXT DEFAULT '',
        forwarded_to_super  INTEGER DEFAULT 0,
        super_decision      TEXT DEFAULT '',
        super_notes         TEXT DEFAULT '',
        super_handled_by    TEXT DEFAULT '',
        super_handled_at    TEXT DEFAULT '',
        created_at          TEXT NOT NULL,
        FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE CASCADE
    )""")

    # Migrate driver_requests new columns
    try:
        c.execute("PRAGMA table_info(driver_requests)")
        existing_req_cols = {row["name"] for row in c.fetchall()}
        req_new_cols = [
            ("priority",           "TEXT DEFAULT 'medium'"),
            ("forwarded_to_super", "INTEGER DEFAULT 0"),
            ("super_decision",     "TEXT DEFAULT ''"),
            ("super_notes",        "TEXT DEFAULT ''"),
            ("super_handled_by",   "TEXT DEFAULT ''"),
            ("super_handled_at",   "TEXT DEFAULT ''"),
        ]
        for col, col_type in req_new_cols:
            if col not in existing_req_cols:
                c.execute(f"ALTER TABLE driver_requests ADD COLUMN {col} {col_type}")
    except Exception as e:
        log.warning(f"driver_requests migration: {e}")
    c.execute("CREATE INDEX IF NOT EXISTS idx_req_driver ON driver_requests(driver_id)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_req_status ON driver_requests(status)")

    additions = {
        "users":              [("refresh_token","TEXT"),("refresh_exp","TEXT"),("last_login","TEXT"),
                               ("branch","TEXT DEFAULT ''")],
        "emergency_reports":  [("audio_url","TEXT DEFAULT ''"),("is_handled","INTEGER DEFAULT 0"),
                               ("action_taken","TEXT DEFAULT ''"),("handled_by","TEXT DEFAULT ''"),
                               ("action_time","TEXT DEFAULT ''"),("location","TEXT DEFAULT ''"),
                               ("driver_message","TEXT DEFAULT ''")] ,
        "trips":              [("garage_location","TEXT DEFAULT ''")],
        "workshop_records":   [("location","TEXT DEFAULT ''"),("operation_type","TEXT DEFAULT ''"),
                               ("vehicle_id","INTEGER"),("odometer_reading","REAL"),
                               ("description","TEXT DEFAULT ''"),("tire_action","TEXT DEFAULT ''")],
        "drivers":            [("national_id","TEXT DEFAULT ''"),("birth_date","TEXT DEFAULT ''"),
                               ("driver_license_expiry","TEXT DEFAULT ''"),
                               ("vehicle_license_expiry","TEXT DEFAULT ''"),
                               ("branch","TEXT DEFAULT ''"),
                               ("fixed_number","TEXT DEFAULT ''")],
        "cars":               [("chassis","TEXT DEFAULT ''"),("car_license_expiry","TEXT DEFAULT ''"),
                               ("sector","TEXT DEFAULT ''"),("car_name","TEXT DEFAULT ''"),
                               ("car_code","TEXT DEFAULT ''"),("engine_number","TEXT DEFAULT ''"),
                               ("year","TEXT DEFAULT ''"),("project","TEXT DEFAULT ''"),
                               ("branch","TEXT DEFAULT ''"),("equipment_type","TEXT DEFAULT ''")],
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
        "branch":   payload.get("branch", ""),
    }
    if result["role"] == "driver":
        with get_db() as conn:
            c = conn.cursor()
            c.execute("SELECT id FROM drivers WHERE user_id=?", (result["user_id"],))
            drv = c.fetchone()
            result["driver_id"] = drv["id"] if drv else None
    return result

def require_admin(cu: dict = Depends(get_user)):
    if cu["role"] not in ("admin", "superuser"):
        raise HTTPException(status.HTTP_403_FORBIDDEN, "صلاحيات المدير مطلوبة")
    return cu

def require_superuser(cu: dict = Depends(get_user)):
    if cu["role"] != "superuser":
        raise HTTPException(status.HTTP_403_FORBIDDEN, "صلاحيات السوبر يوزر مطلوبة")
    return cu

def require_admin_or_reporter(cu: dict = Depends(get_user)):
    """Allows superuser, admin and reporter (read-only) roles."""
    if cu["role"] not in ("superuser", "admin", "reporter"):
        raise HTTPException(status.HTTP_403_FORBIDDEN, "صلاحيات غير كافية")
    return cu

def _branch_filter(cu: dict) -> Optional[str]:
    """
    يرجع اسم الفرع اللي المستخدم مرتبط بيه، أو None لو superuser أو مفيش فرع.
    - superuser: يشوف كل الفروع → None
    - admin/reporter بدون branch: يشوف كل حاجة (backward compatible) → None
    - admin/reporter عنده branch: يشوف فرعه فقط → اسم الفرع
    """
    if cu["role"] == "superuser":
        return None
    return cu.get("branch") or None

def _effective_branch(cu: dict, super_branch: Optional[str] = None) -> Optional[str]:
    """
    للسوبر يوزر: يقدر يمرر super_branch (query param) عشان يفلتر بفرع معين.
    للأدمن/المراقب: دايماً فرعه الخاص (لو موجود).
    """
    if cu["role"] == "superuser":
        return super_branch or None
    return cu.get("branch") or None

def write_audit_log(user_id: int, username: str, role: str,
                    action: str, details: str = "", ip: str = "",
                    cursor=None):
    """Write an audit log entry.
    Pass an existing cursor to reuse the same transaction,
    or leave None to open a new connection.
    """
    try:
        now = datetime.utcnow().isoformat() + "Z"
        sql = """INSERT INTO audit_logs(user_id,username,role,action,details,ip_address,created_at)
                 VALUES(?,?,?,?,?,?,?)"""
        params = (user_id, username, role, action, details, ip, now)
        if cursor is not None:
            cursor.execute(sql, params)
        else:
            with get_db() as conn:
                conn.cursor().execute(sql, params)
    except Exception as e:
        log.warning(f"audit_log write failed: {e}")

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
    branch: Optional[str] = ""

class LoginResp(BaseModel):
    access_token: str; refresh_token: str; token_type: str; user: UserResp

class DriverCreate(BaseModel):
    name: str; phone: str; username: str; password: str
    status: str = "active"
    national_id: Optional[str] = ""
    birth_date: Optional[str] = None
    driver_license_expiry: Optional[str] = None
    vehicle_license_expiry: Optional[str] = None
    branch: Optional[str] = ""
    fixed_number: Optional[str] = ""

class DriverUpdate(BaseModel):
    name: str; phone: str; status: str
    national_id: Optional[str] = ""
    birth_date: Optional[str] = ""
    driver_license_expiry: Optional[str] = ""
    vehicle_license_expiry: Optional[str] = ""
    branch: Optional[str] = ""
    fixed_number: Optional[str] = ""

class DriverResp(BaseModel):
    id: int; name: str; phone: str; status: str
    user_id: Optional[int] = None; username: Optional[str] = None
    national_id: Optional[str] = None; birth_date: Optional[str] = None
    driver_license_expiry: Optional[str] = None; vehicle_license_expiry: Optional[str] = None
    branch: Optional[str] = None; fixed_number: Optional[str] = None

EQUIPMENT_TYPES = {
    "معدات نقل", "معدات رفع", "معدات تحريك تربة",
    "معدات ثابتة", "وسائل انتقال", "معدات إنتاجية",
    "معدات مساعدة", "معدات وآلات ورش", "أخرى"
}

BRANCHES = [
    "القاهرة",
    "مدينة نصر",
    "حلوان",
    "صيناء القصور والآثار",
    "المنشآت المتميزة",
    "مشروع منوريل 6 أكتوبر",
    "مشروع المرحلة الأولى من الخط الرابع مترو الأنفاق القاهرة الكبرى",
    "الأعمال الكهربائية",
    "مشروعات كهروميكانيكية للمباني العامة والمرافق",
    "مشروعات كهروميكانيكية لمياه الشرب والصرف الصحي والمصانع",
    "المصانع والورش المركزية",
    "ترسانة المعصرة",
    "الأعمال الاعتيادية",
    "أعمال الصحي والتشطيبات",
    "الشدات والأعمال التخصصية",
    "الكباري والأعمال التخصصية",
]

WORKSHOP_TYPES = [
    "fuel_solar", "fuel_92", "fuel_95", "fuel_80", "fuel_cng",
    "oil", "filter", "tire", "battery", "belt", "other",
]

class CarCreate(BaseModel):
    plate: str; model: str; status: str = "available"
    car_name: Optional[str] = ""
    car_code: Optional[str] = ""
    chassis: Optional[str] = ""
    engine_number: Optional[str] = ""
    year: Optional[str] = ""
    project: Optional[str] = ""
    branch: Optional[str] = ""
    car_license_expiry: Optional[str] = ""
    equipment_type: Optional[str] = ""
    sector: Optional[str] = ""

class CarUpdate(BaseModel):
    plate: str; model: str; status: str
    car_name: Optional[str] = ""
    car_code: Optional[str] = ""
    chassis: Optional[str] = ""
    engine_number: Optional[str] = ""
    year: Optional[str] = ""
    project: Optional[str] = ""
    branch: Optional[str] = ""
    car_license_expiry: Optional[str] = ""
    equipment_type: Optional[str] = ""
    sector: Optional[str] = ""

class CarResp(BaseModel):
    id: int; plate: str; model: str; status: str
    car_name: Optional[str] = None
    car_code: Optional[str] = None
    chassis: Optional[str] = None
    engine_number: Optional[str] = None
    year: Optional[str] = None
    project: Optional[str] = None
    branch: Optional[str] = None
    car_license_expiry: Optional[str] = None
    equipment_type: Optional[str] = None
    sector: Optional[str] = None

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
    driver_message: str = ""

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

REPORTERS = [
    ("admin1", "24681012"),
    ("admin2", "11214123"),
    ("admin3", "9853247"),
]

SUPERUSERS = [
    ("supre mohamed sayed",    "sup@mosayed3904"),
    ("super abdelrhman sayed", "sup@abdo1414"),
    ("super mohamed mansour",  "sup@momansour84329"),
]

def _sync_accounts(c, accounts: list[tuple[str, str]], role: str):
    """Ensure exactly the given accounts exist for a role.
    Removes accounts no longer in the list, adds missing ones.
    Skips rehashing on restart for accounts that already exist.
    """
    c.execute("SELECT id, username FROM users WHERE role=?", (role,))
    existing = {r["username"]: r["id"] for r in c.fetchall()}
    allowed  = {u for u, _ in accounts}

    for uname, uid in list(existing.items()):
        if uname not in allowed:
            c.execute("DELETE FROM users WHERE id=?", (uid,))
            log.info(f"Removed old {role}: {uname}")

    for uname, pw in accounts:
        if uname not in existing:
            pw_hash = bcrypt.hashpw(pw.encode(), bcrypt.gensalt()).decode()
            c.execute("INSERT OR IGNORE INTO users(username,password,role) VALUES(?,?,?)",
                      (uname, pw_hash, role))
            log.info(f"Created {role}: {uname}")

    log.info(f"✅ {role.capitalize()} accounts synced")


@app.on_event("startup")
async def startup():
    migrate_db()
    with get_db() as conn:
        c = conn.cursor()
        _sync_accounts(c, ADMINS,      "admin")
        _sync_accounts(c, REPORTERS,   "reporter")
        _sync_accounts(c, SUPERUSERS,  "superuser")
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

def validate_coordinates(loc: str) -> bool:
    """
    تتحقق إن الموقع على شكل 'lat,lon' وإن القيم في النطاق الصحيح.
    مثال صحيح: '30.355400,31.304600'
    """
    if not loc or not loc.strip():
        return False
    parts = loc.strip().split(",")
    if len(parts) != 2:
        return False
    try:
        lat = float(parts[0].strip())
        lon = float(parts[1].strip())
        return -90 <= lat <= 90 and -180 <= lon <= 180
    except (ValueError, TypeError):
        return False

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

def _car_fields(car, cid: int) -> dict:
    """Build a consistent car response dict from a CarCreate/CarUpdate model."""
    return {
        "id":                cid,
        "plate":             car.plate,
        "model":             car.model,
        "status":            car.status,
        "car_name":          car.car_name or "",
        "car_code":          car.car_code or "",
        "chassis":           car.chassis or "",
        "engine_number":     car.engine_number or "",
        "year":              car.year or "",
        "project":           car.project or "",
        "branch":            car.branch or "",
        "car_license_expiry":car.car_license_expiry or "",
        "equipment_type":    car.equipment_type or "",
        "sector":            car.sector or "",
    }

# ══════════════════════════════════════════════════════
# 11. AUTH ENDPOINTS
# ══════════════════════════════════════════════════════

@app.post("/login", response_model=LoginResp)
@limiter.limit("10/minute")   # strict: prevent brute-force
async def login(request: Request, data: LoginReq):
    with get_db() as conn:
        c = conn.cursor()
        # username ممكن يتكرر عند السائقين — نجيب كل الحسابات بنفس الاسم
        # ونبحث عن الأول اللي كلمة مروره تطابق
        c.execute("SELECT id,username,password,role,branch FROM users WHERE username=?", (data.username,))
        candidates = c.fetchall()
        u = None
        for candidate in candidates:
            if _verify(data.password, candidate["password"]):
                u = candidate
                break
        if not u:
            log_event("login_failed", username=data.username, ip=request.client.host)
            raise HTTPException(401, "اسم المستخدم أو كلمة المرور غير صحيحة")

        driver_id = None
        if u["role"] == "driver":
            c.execute("SELECT id FROM drivers WHERE user_id=?", (u["id"],))
            drv = c.fetchone()
            driver_id = drv["id"] if drv else None

        # Tokens
        access_token = create_access_token({
            "user_id": u["id"], "role": u["role"], "username": u["username"],
            "branch":  u["branch"] or ""
        })
        refresh_token, refresh_exp = create_refresh_token()

        # Persist refresh token (one per user)
        c.execute("UPDATE users SET refresh_token=?,refresh_exp=?,last_login=? WHERE id=?",
                  (refresh_token, refresh_exp, datetime.utcnow().isoformat() + "Z", u["id"]))

        log_event("login_success", user_id=u["id"], role=u["role"])
        # ── Audit log — pass cursor to avoid nested get_db() conflict ──
        write_audit_log(
            user_id=u["id"], username=u["username"], role=u["role"],
            action="login", details="تسجيل دخول ناجح",
            ip=request.client.host if request.client else "",
            cursor=c
        )
        return LoginResp(
            access_token=access_token,
            refresh_token=refresh_token,
            token_type="bearer",
            user=UserResp(id=u["id"], username=u["username"],
                          role=u["role"], driver_id=driver_id,
                          branch=u["branch"] or "")
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
async def get_drivers(cu: dict = Depends(get_user), branch: Optional[str] = None):
    with get_db() as conn:
        c = conn.cursor()
        branch = _effective_branch(cu, branch)
        if branch:
            c.execute("""SELECT d.*,u.username FROM drivers d
                         LEFT JOIN users u ON d.user_id=u.id
                         WHERE d.branch=?
                         ORDER BY d.id DESC LIMIT 500""", (branch,))
        else:
            c.execute("""SELECT d.*,u.username FROM drivers d
                         LEFT JOIN users u ON d.user_id=u.id
                         ORDER BY d.id DESC LIMIT 500""")
        return [dict(r) for r in c.fetchall()]

@app.post("/drivers", response_model=DriverResp)
async def create_driver(driver: DriverCreate, cu: dict = Depends(require_admin)):
    # لو الأدمن عنده فرع، السائق الجديد لازم يكون نفس الفرع
    if cu["role"] == "admin" and cu.get("branch"):
        driver_branch = cu["branch"]
    else:
        driver_branch = driver.branch or ""
    with get_db() as conn:
        c = conn.cursor()
        # username ممكن يتكرر — اللي مش ممكن يتكرر هو الرقم الثابت
        if driver.fixed_number:
            c.execute("SELECT id FROM drivers WHERE fixed_number=?", (driver.fixed_number,))
            if c.fetchone():
                raise HTTPException(400, f"الرقم الثابت '{driver.fixed_number}' موجود مسبقاً")
        c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                  (driver.username, _hash(driver.password), "driver"))
        uid = c.lastrowid
        c.execute("""INSERT INTO drivers(name,phone,status,user_id,national_id,
                     birth_date,driver_license_expiry,vehicle_license_expiry,branch,fixed_number)
                     VALUES(?,?,?,?,?,?,?,?,?,?)""",
                  (driver.name, driver.phone, driver.status, uid,
                   driver.national_id or "", driver.birth_date or "",
                   driver.driver_license_expiry or "", driver.vehicle_license_expiry or "",
                   driver_branch, driver.fixed_number or ""))
        did = c.lastrowid
        log_event("driver_created", driver_id=did, admin=cu["username"])
        return {"id":did,"name":driver.name,"phone":driver.phone,"status":driver.status,
                "user_id":uid,"username":driver.username,
                "national_id":driver.national_id,"birth_date":driver.birth_date,
                "driver_license_expiry":driver.driver_license_expiry,
                "vehicle_license_expiry":driver.vehicle_license_expiry,
                "branch":driver_branch,"fixed_number":driver.fixed_number or ""}

@app.put("/drivers/{did}", response_model=DriverResp)
async def update_driver(did: int, driver: DriverUpdate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""UPDATE drivers SET name=?,phone=?,status=?,national_id=?,
                     birth_date=?,driver_license_expiry=?,vehicle_license_expiry=?,
                     branch=?,fixed_number=?
                     WHERE id=?""",
                  (driver.name,driver.phone,driver.status,
                   driver.national_id or "",driver.birth_date or "",
                   driver.driver_license_expiry or "",driver.vehicle_license_expiry or "",
                   driver.branch or "",driver.fixed_number or "",did))
        if c.rowcount == 0:
            raise HTTPException(404, "السائق غير موجود")
        c.execute("SELECT user_id FROM drivers WHERE id=?", (did,))
        r = c.fetchone()
        return {"id":did,"name":driver.name,"phone":driver.phone,"status":driver.status,
                "user_id":r["user_id"] if r else None,
                "national_id":driver.national_id,"birth_date":driver.birth_date,
                "driver_license_expiry":driver.driver_license_expiry,
                "vehicle_license_expiry":driver.vehicle_license_expiry,
                "branch":driver.branch or "","fixed_number":driver.fixed_number or ""}

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
        # username مش unique للسائقين — مفيش حاجة نتأكد منها هنا
        if new_password:
            c.execute("UPDATE users SET username=?,password=? WHERE id=?",
                      (new_username, _hash(new_password), uid))
        else:
            c.execute("UPDATE users SET username=? WHERE id=?", (new_username, uid))
        return {"message": "تم التحديث", "username": new_username}

@app.get("/drivers/{did}/primary-car")
async def primary_car(did: int, cu: dict = Depends(get_user)):
    if cu["role"] not in ("admin","superuser") and cu.get("driver_id") != did:
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
    if cu["role"] not in ("admin","superuser") and cu.get("driver_id") != did:
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
async def get_cars(cu: dict = Depends(get_user), branch: Optional[str] = None):
    with get_db() as conn:
        c = conn.cursor()
        branch = _effective_branch(cu, branch)
        base_q = """SELECT id,plate,model,status,car_name,car_code,chassis,engine_number,
                           year,project,branch,car_license_expiry,equipment_type,sector
                    FROM cars"""
        if branch:
            c.execute(base_q + " WHERE branch=? ORDER BY id DESC LIMIT 500", (branch,))
        else:
            c.execute(base_q + " ORDER BY id DESC LIMIT 500")
        rows = []
        for r in c.fetchall():
            d = dict(r)
            for f in ("car_name","car_code","chassis","engine_number","year","project",
                      "branch","car_license_expiry","equipment_type","sector"):
                d.setdefault(f, "")
            rows.append(d)
        return rows

@app.post("/cars", response_model=CarResp)
async def create_car(car: CarCreate, cu: dict = Depends(require_admin)):
    # لو الأدمن عنده فرع، المركبة الجديدة لازم تكون نفس الفرع
    if cu["role"] == "admin" and cu.get("branch"):
        car_branch = cu["branch"]
    else:
        car_branch = car.branch or ""
    with get_db() as conn:
        c = conn.cursor()
        try:
            c.execute("""INSERT INTO cars(plate,model,status,car_name,car_code,chassis,
                                          engine_number,year,project,branch,
                                          car_license_expiry,equipment_type,sector)
                         VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                      (car.plate, car.model, car.status,
                       car.car_name or "", car.car_code or "", car.chassis or "",
                       car.engine_number or "", car.year or "", car.project or "",
                       car_branch, car.car_license_expiry or "",
                       car.equipment_type or "", car.sector or ""))
            cid = c.lastrowid
            return {**_car_fields(car, cid), "branch": car_branch}
        except sqlite3.IntegrityError:
            raise HTTPException(400, "رقم اللوحة موجود مسبقاً")

@app.put("/cars/{cid}", response_model=CarResp)
async def update_car(cid: int, car: CarUpdate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""UPDATE cars SET plate=?,model=?,status=?,car_name=?,car_code=?,
                                     chassis=?,engine_number=?,year=?,project=?,branch=?,
                                     car_license_expiry=?,equipment_type=?,sector=?
                     WHERE id=?""",
                  (car.plate, car.model, car.status,
                   car.car_name or "", car.car_code or "", car.chassis or "",
                   car.engine_number or "", car.year or "", car.project or "",
                   car.branch or "", car.car_license_expiry or "",
                   car.equipment_type or "", car.sector or "", cid))
        if c.rowcount == 0:
            raise HTTPException(404, "المركبة غير موجودة")
        return _car_fields(car, cid)

@app.delete("/cars/{cid}")
async def delete_car(cid: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("DELETE FROM driver_car_permissions WHERE car_id=?", (cid,))
        c.execute("DELETE FROM cars WHERE id=?", (cid,))
        return {"message": "تم الحذف"}

@app.get("/cars/equipment-types")
async def get_equipment_types(cu: dict = Depends(get_user)):
    """قائمة أنواع تصنيف المعدات والمركبات."""
    return {"equipment_types": sorted(EQUIPMENT_TYPES)}

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
async def get_trips(cu: dict = Depends(get_user), branch: Optional[str] = None):
    with get_db() as conn:
        c = conn.cursor()
        branch = _effective_branch(cu, branch)
        if cu["role"] in ("admin", "reporter", "superuser"):
            if branch:
                c.execute("""SELECT t.* FROM trips t
                             JOIN drivers d ON t.driver_id=d.id
                             WHERE d.branch=?
                             ORDER BY t.id DESC LIMIT 500""", (branch,))
            else:
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
        if cu["role"] not in ("admin", "superuser"):
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

        # ── حدد موقع البداية الفعلي ──
        # الأولوية: (1) موقع صحيح من الـ frontend، (2) آخر موقع جراج مسجّل
        raw_start = (trip.start_location or "").strip()
        effective_start_location = raw_start if validate_coordinates(raw_start) else ""

        if not effective_start_location:
            c.execute(
                "SELECT id, location FROM garage_records WHERE driver_id=? ORDER BY id DESC LIMIT 1",
                (trip.driver_id,)
            )
            last_garage = c.fetchone()
            if last_garage:
                effective_start_location = last_garage["location"]
                log_event("trip_start_location_from_garage",
                          driver_id=trip.driver_id,
                          garage_record_id=last_garage["id"],
                          location=effective_start_location)
            else:
                log_event("trip_start_location_missing", driver_id=trip.driver_id,
                          raw_sent=raw_start)
        else:
            log_event("trip_start_location_from_gps", driver_id=trip.driver_id,
                      location=effective_start_location)

        now = datetime.utcnow().isoformat() + "Z"
        c.execute("""INSERT INTO trips(driver_id,car_id,start_time,start_odometer,start_location,garage_location)
                     VALUES(?,?,?,?,?,'')""",
                  (trip.driver_id, trip.car_id, now, trip.start_odometer, effective_start_location))
        tid = c.lastrowid
        log_event("trip_started", trip_id=tid, driver_id=trip.driver_id,
                  car_id=trip.car_id, start_location=effective_start_location)
        return {"id":tid,"driver_id":trip.driver_id,"car_id":trip.car_id,
                "start_time":now,"end_time":None,"start_odometer":trip.start_odometer,
                "end_odometer":None,"start_location":effective_start_location,
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
        if cu["role"] not in ("admin","superuser") and trip["driver_id"] != cu.get("driver_id"):
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
        if cu["role"] not in ("admin","superuser") and trip["driver_id"] != cu.get("driver_id"):
            raise HTTPException(403, "غير مصرح")
        c.execute("UPDATE trips SET garage_location=? WHERE id=?", (loc, tid))
        return {"ok": True}

@app.put("/trips/{tid}/odometer")
async def edit_trip_odometer(tid: int, body: dict, cu: dict = Depends(require_admin)):
    """Admin-only: edit start/end odometer for a trip."""
    start_odo = body.get("start_odometer")
    end_odo   = body.get("end_odometer")
    if start_odo is None and end_odo is None:
        raise HTTPException(400, "يجب إدخال عداد البداية أو النهاية على الأقل")
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, start_odometer, end_odometer FROM trips WHERE id=?", (tid,))
        trip = c.fetchone()
        if not trip:
            raise HTTPException(404, "الرحلة غير موجودة")
        new_start = float(start_odo) if start_odo is not None else trip["start_odometer"]
        new_end   = float(end_odo)   if end_odo   is not None else trip["end_odometer"]
        if new_start is not None and new_start < 0:
            raise HTTPException(400, "عداد البداية يجب أن يكون موجباً")
        if new_end is not None and new_start is not None and new_end <= new_start:
            raise HTTPException(400, "عداد النهاية يجب أن يكون أكبر من عداد البداية")
        c.execute("UPDATE trips SET start_odometer=?, end_odometer=? WHERE id=?",
                  (new_start, new_end, tid))
        log_event("trip_odometer_edited", admin=cu["username"], trip_id=tid,
                  start_odometer=new_start, end_odometer=new_end)
        return {"ok": True, "trip_id": tid, "start_odometer": new_start, "end_odometer": new_end}

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
async def list_users(cu: dict = Depends(require_admin_or_reporter)):
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
    if rec.type not in WORKSHOP_TYPES:
        raise HTTPException(400, "نوع غير صالح")
    if cu["role"] not in ("admin", "superuser") and rec.driver_id != cu.get("driver_id"):
        raise HTTPException(403, "غير مصرح")
    # Enforce admin price for drivers
    final_price = rec.price if cu["role"] in ("admin", "superuser") else 0.0
    if cu["role"] == "driver":
        with get_db() as conn:
            c = conn.cursor()
            # جيب فرع السائق
            c.execute("SELECT branch FROM drivers WHERE id=?", (rec.driver_id,))
            drv_row = c.fetchone()
            drv_branch = drv_row["branch"] if drv_row else ""
            # ابحث عن سعر الفرع أولاً، وإلا الأسعار العامة
            if drv_branch:
                c.execute("SELECT value FROM app_settings WHERE key=?",
                          (f"price_{drv_branch}_{rec.type}",))
                row = c.fetchone()
            else:
                row = None
            if not row:
                c.execute("SELECT value FROM app_settings WHERE key=?", (f"price_{rec.type}",))
                row = c.fetchone()
            unit_price = float(row["value"]) if row else 0.0
    # اخزّن السعر الإجمالي = سعر الوحدة × الكمية
    qty = rec.quantity if rec.quantity and rec.quantity > 0 else 1.0
    final_price = unit_price * qty
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
async def get_workshops(cu: dict = Depends(get_user), branch: Optional[str] = None):
    q = """SELECT w.*,d.name as driver_name,c.plate as vehicle_plate
           FROM workshop_records w
           LEFT JOIN drivers d ON w.driver_id=d.id
           LEFT JOIN cars c ON w.vehicle_id=c.id"""
    with get_db() as conn:
        c = conn.cursor()
        branch = _effective_branch(cu, branch)
        if cu["role"] in ("admin", "reporter", "superuser"):
            if branch:
                c.execute(q + " WHERE d.branch=? ORDER BY w.id DESC LIMIT 500", (branch,))
            else:
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
    # ── التحقق من الموقع — يجب أن يكون إحداثيات صحيحة ──
    loc = (rec.location or "").strip()
    if not loc:
        raise HTTPException(400, "موقع الجراج مطلوب — يجب السماح بالوصول للموقع")
    if not validate_coordinates(loc):
        log_event("garage_invalid_coordinates", driver_id=rec.driver_id, raw=loc)
        raise HTTPException(400, f"إحداثيات الموقع غير صالحة: '{loc}' — تأكد من تفعيل GPS")
    if cu["role"] not in ("admin", "superuser") and rec.driver_id != cu.get("driver_id"):
        raise HTTPException(403, "غير مصرح")

    now = datetime.utcnow().isoformat() + "Z"
    log_event("garage_record_attempt", driver_id=rec.driver_id,
              car_id=rec.car_id, location=loc, odometer=rec.odometer)

    with get_db() as conn:
        c = conn.cursor()

        # ── Anti-duplicate: امنع إدخال نفس الموقع خلال 60 ثانية ──
        # يحمي من double-tap أو إعادة إرسال من الـ frontend
        c.execute("""
            SELECT id, recorded_at FROM garage_records
            WHERE driver_id=? AND location=?
            ORDER BY id DESC LIMIT 1
        """, (rec.driver_id, loc))
        last_same = c.fetchone()
        if last_same:
            try:
                last_time = datetime.fromisoformat(last_same["recorded_at"].replace("Z", "+00:00"))
                diff_seconds = (datetime.now(last_time.tzinfo) - last_time).total_seconds()
                if diff_seconds < 60:
                    log_event("garage_duplicate_blocked", driver_id=rec.driver_id,
                              location=loc, seconds_ago=diff_seconds,
                              existing_id=last_same["id"])
                    raise HTTPException(429, "تم تسجيل نفس الموقع مؤخراً — انتظر دقيقة قبل إعادة المحاولة")
            except HTTPException:
                raise
            except Exception:
                pass  # لو التحليل فشل نكمل عادي

        # ── الجراج يُسجَّل مستقلاً عن انتهاء الرحلة ──
        # السائق ينهي رحلته بزر منفصل، والجراج يُسجَّل مرة واحدة في نهاية اليوم
        log_event("garage_record_attempt", driver_id=rec.driver_id,
                  car_id=rec.car_id, location=loc)

        # ── أدرج سجل الجراج الجديد ──
        c.execute(
            "INSERT INTO garage_records(driver_id,car_id,location,recorded_at,notes) VALUES(?,?,?,?,?)",
            (rec.driver_id, rec.car_id, loc, now, rec.notes or "")
        )
        rid = c.lastrowid
        log_event("garage_record_inserted", garage_record_id=rid,
                  driver_id=rec.driver_id, location=loc)

        plate = dn = None
        if rec.car_id:
            c.execute("SELECT plate FROM cars WHERE id=?", (rec.car_id,))
            r = c.fetchone(); plate = r["plate"] if r else None
        c.execute("SELECT name FROM drivers WHERE id=?", (rec.driver_id,))
        dr = c.fetchone(); dn = dr["name"] if dr else None

        return {"id":rid,"driver_id":rec.driver_id,"car_id":rec.car_id,
                "location":loc,"recorded_at":now,"notes":rec.notes or "",
                "driver_name":dn,"car_plate":plate}

@app.get("/garage/records")
async def get_garage_records(cu: dict = Depends(get_user)):
    with get_db() as conn:
        c = conn.cursor()
        q = """SELECT g.*,d.name as driver_name,c.plate as car_plate
               FROM garage_records g
               LEFT JOIN drivers d ON g.driver_id=d.id
               LEFT JOIN cars c ON g.car_id=c.id"""
        if cu["role"] in ("admin", "reporter", "superuser"):
            c.execute(q + " ORDER BY g.id DESC LIMIT 500")
        else:
            did = cu.get("driver_id")
            if not did: return []
            c.execute(q + " WHERE g.driver_id=? ORDER BY g.id DESC LIMIT 200", (did,))
        return [dict(r) for r in c.fetchall()]

@app.get("/garage/latest")
async def garage_latest(cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        c = conn.cursor()
        # Group by car_id to always return the latest garage record PER CAR
        # (not per driver, so that the map pin reflects the actual current parking spot)
        c.execute("""SELECT g.*,d.name as driver_name,c.plate as car_plate
                     FROM garage_records g
                     INNER JOIN(SELECT car_id, MAX(id) as max_id
                                FROM garage_records
                                WHERE car_id IS NOT NULL
                                GROUP BY car_id) latest
                       ON g.id=latest.max_id
                     LEFT JOIN drivers d ON g.driver_id=d.id
                     LEFT JOIN cars c ON g.car_id=c.id
                     ORDER BY g.recorded_at DESC""")
        rows = [dict(r) for r in c.fetchall()]
        # Also include records without car_id grouped by driver (fallback)
        c.execute("""SELECT g.*,d.name as driver_name,NULL as car_plate
                     FROM garage_records g
                     INNER JOIN(SELECT driver_id, MAX(id) as max_id
                                FROM garage_records
                                WHERE car_id IS NULL
                                GROUP BY driver_id) latest
                       ON g.id=latest.max_id
                     LEFT JOIN drivers d ON g.driver_id=d.id
                     ORDER BY g.recorded_at DESC""")
        rows += [dict(r) for r in c.fetchall()]
        return rows

# ══════════════════════════════════════════════════════
# 20. EMERGENCY — Audio stored as file, not base64
# ══════════════════════════════════════════════════════

@app.post("/emergency/report")
@limiter.limit("5/minute")  # very strict on emergency endpoint
async def create_emergency(request: Request, rep: EmergencyCreate, cu: dict = Depends(get_user)):
    if rep.type not in ("emergency", "accident"):
        raise HTTPException(400, "نوع البلاغ غير صالح")
    if cu["role"] not in ("admin","superuser") and rep.driver_id != cu.get("driver_id"):
        raise HTTPException(403, "غير مصرح")
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        # Save base64 audio to file if provided
        audio_url = ""
        if rep.audio_data and rep.audio_data.startswith("data:audio"):
            try:
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
async def get_emergency_reports(cu: dict = Depends(get_user), branch: Optional[str] = None):
    q = """SELECT e.*,d.name as driver_name,c.plate as car_plate
           FROM emergency_reports e
           LEFT JOIN drivers d ON e.driver_id=d.id
           LEFT JOIN cars c ON e.car_id=c.id"""
    with get_db() as conn:
        c = conn.cursor()
        branch = _effective_branch(cu, branch)
        if cu["role"] in ("admin", "reporter", "superuser"):
            if branch:
                c.execute(q + " WHERE d.branch=? ORDER BY e.id DESC LIMIT 500", (branch,))
            else:
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
            d.setdefault("audio_url", ""); d.setdefault("driver_message", "")
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
                     SET is_handled=?,action_taken=?,handled_by=?,is_read=1,action_time=?,driver_message=?
                     WHERE id=?""",
                  (1 if body.is_handled else 0, body.action_taken, admin_name,
                   action_ts, body.driver_message or "", rid))
        return {"ok": True, "handled_by": admin_name, "action_time": action_ts}

@app.get("/emergency/unread-count")
async def unread_count(cu: dict = Depends(get_user)):
    if cu["role"] not in ("admin", "reporter", "superuser"):
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
    """
    - superuser أو بدون فرع: يرجع كل الأسعار (عامة + فروع)
    - admin/reporter عنده فرع: يرجع أسعار فرعه (يرجع للعام لو مفيش سعر فرع محدد)
    """
    branch = _branch_filter(cu)
    with get_db() as conn:
        c = conn.cursor()
        if branch:
            prices = {}
            for ws_type in WORKSHOP_TYPES:
                # ابحث عن سعر الفرع أولاً
                c.execute("SELECT value FROM app_settings WHERE key=?",
                          (f"price_{branch}_{ws_type}",))
                row = c.fetchone()
                if row:
                    prices[ws_type] = float(row["value"])
                else:
                    # fallback للسعر العام
                    c.execute("SELECT value FROM app_settings WHERE key=?",
                              (f"price_{ws_type}",))
                    row = c.fetchone()
                    prices[ws_type] = float(row["value"]) if row else 0.0
            return prices
        else:
            c.execute("SELECT key,value FROM app_settings WHERE key LIKE 'price_%'")
            return {r["key"].replace("price_",""): float(r["value"]) for r in c.fetchall()}

@app.put("/settings/prices")
async def set_prices(body: dict, cu: dict = Depends(require_superuser)):
    """
    - admin عنده فرع: يحفظ الأسعار لفرعه: key = price_{branch}_{type}
    - superuser أو بدون فرع: يحفظ الأسعار العامة: key = price_{type}
    - superuser ممكن يمرر { branch: 'cairo', prices: {...} } عشان يعدل فرع معين
    """
    branch = _branch_filter(cu)
    # superuser ممكن يحدد فرع معين في الـ body
    if not branch and cu["role"] == "superuser":
        branch = body.pop("branch", None) or None

    valid = WORKSHOP_TYPES
    now   = datetime.utcnow().isoformat() + "Z"
    saved = {}
    with get_db() as conn:
        c = conn.cursor()
        for ws_type, price in body.items():
            if ws_type not in valid:
                continue
            if not isinstance(price, (int, float)) or price < 0:
                continue
            key = f"price_{branch}_{ws_type}" if branch else f"price_{ws_type}"
            c.execute("INSERT OR REPLACE INTO app_settings(key,value,updated_at) VALUES(?,?,?)",
                      (key, str(float(price)), now))
            saved[ws_type] = float(price)
    return {"saved": saved, "branch": branch or "global"}

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
        c.execute("DELETE FROM users WHERE role NOT IN ('admin','superuser','reporter')")
        log_event("data_reset", admin=cu["username"])
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "data_reset", "تم مسح جميع البيانات التشغيلية",
                        request.client.host if request.client else "")
        return {"message": "تم مسح جميع البيانات التشغيلية"}

# ══════════════════════════════════════════════════════
# 23. SUPERUSER — Admin Management + Audit Logs
# ══════════════════════════════════════════════════════

class AdminCreate(BaseModel):
    username: str
    password: str
    branch: Optional[str] = ""   # الفرع التابع له الأدمن/المراقب

class AdminUpdate(BaseModel):
    username: str
    password: Optional[str] = None   # اختياري — لو فارغ مش هيتغير
    branch: Optional[str] = None     # اختياري — لو None مش هيتغير


@app.get("/superuser/admins")
async def superuser_list_admins(cu: dict = Depends(require_superuser)):
    """عرض جميع حسابات الأدمن."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT id, username, role, branch, created_at, last_login
                     FROM users WHERE role='admin' ORDER BY id""")
        return [dict(r) for r in c.fetchall()]


@app.post("/superuser/admins", status_code=201)
async def superuser_create_admin(
    request: Request,
    body: AdminCreate,
    cu: dict = Depends(require_superuser)
):
    """إضافة أدمن جديد."""
    username = body.username.strip()
    if len(username) < 2:
        raise HTTPException(400, "اسم المستخدم قصير جداً")
    validate_password(body.password)
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id FROM users WHERE username=?", (username,))
        if c.fetchone():
            raise HTTPException(400, "اسم المستخدم موجود مسبقاً")
        c.execute("INSERT INTO users(username,password,role,branch) VALUES(?,?,?,?)",
                  (username, _hash(body.password), "admin", body.branch or ""))
        new_id = c.lastrowid
    log_event("admin_created_by_super", new_username=username, superuser=cu["username"])
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "create_admin", f"إنشاء أدمن جديد: {username} (فرع: {body.branch or 'بدون'})",
                    request.client.host if request.client else "")
    return {"id": new_id, "username": username, "role": "admin", "branch": body.branch or ""}


@app.put("/superuser/admins/{uid}")
async def superuser_update_admin(
    request: Request,
    uid: int,
    body: AdminUpdate,
    cu: dict = Depends(require_superuser)
):
    """تعديل اسم المستخدم أو كلمة مرور أدمن."""
    new_username = body.username.strip()
    if len(new_username) < 2:
        raise HTTPException(400, "اسم المستخدم قصير جداً")
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, username, role FROM users WHERE id=?", (uid,))
        target = c.fetchone()
        if not target:
            raise HTTPException(404, "المستخدم غير موجود")
        if target["role"] != "admin":
            raise HTTPException(403, "لا يمكن تعديل إلا حسابات الأدمن من هنا")
        # تأكد من عدم تكرار اسم المستخدم
        c.execute("SELECT id FROM users WHERE username=? AND id!=?", (new_username, uid))
        if c.fetchone():
            raise HTTPException(400, "اسم المستخدم مستخدم من قِبل شخص آخر")
        old_username = target["username"]
        if body.password:
            validate_password(body.password)
            c.execute("UPDATE users SET username=?,password=? WHERE id=?",
                      (new_username, _hash(body.password), uid))
        else:
            c.execute("UPDATE users SET username=? WHERE id=?", (new_username, uid))
        if body.branch is not None:
            c.execute("UPDATE users SET branch=? WHERE id=?", (body.branch, uid))
    log_event("admin_updated_by_super", uid=uid, new_username=new_username, superuser=cu["username"])
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "update_admin",
                    f"تعديل أدمن #{uid} من '{old_username}' إلى '{new_username}'"
                    + (" (مع تغيير كلمة المرور)" if body.password else ""),
                    request.client.host if request.client else "")
    return {"message": "تم التحديث", "id": uid, "username": new_username}


@app.delete("/superuser/admins/{uid}")
async def superuser_delete_admin(
    request: Request,
    uid: int,
    cu: dict = Depends(require_superuser)
):
    """حذف حساب أدمن."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, username, role FROM users WHERE id=?", (uid,))
        target = c.fetchone()
        if not target:
            raise HTTPException(404, "المستخدم غير موجود")
        if target["role"] != "admin":
            raise HTTPException(403, "لا يمكن حذف إلا حسابات الأدمن من هنا")
        deleted_username = target["username"]
        c.execute("DELETE FROM users WHERE id=?", (uid,))
    log_event("admin_deleted_by_super", uid=uid, username=deleted_username, superuser=cu["username"])
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "delete_admin", f"حذف أدمن: {deleted_username} (id={uid})",
                    request.client.host if request.client else "")

    return {"message": f"تم حذف الأدمن: {deleted_username}"}


# ── Reporters management ─────────────────────────────────────────────────────

@app.get("/superuser/reporters")
async def superuser_list_reporters(cu: dict = Depends(require_superuser)):
    """عرض جميع حسابات المراقبين."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT id, username, role, branch, created_at, last_login
                     FROM users WHERE role='reporter' ORDER BY id""")
        return [dict(r) for r in c.fetchall()]


@app.post("/superuser/reporters", status_code=201)
async def superuser_create_reporter(
    request: Request,
    body: AdminCreate,
    cu: dict = Depends(require_superuser)
):
    """إضافة مراقب جديد."""
    username = body.username.strip()
    if len(username) < 2:
        raise HTTPException(400, "اسم المستخدم قصير جداً")
    validate_password(body.password)
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id FROM users WHERE username=?", (username,))
        if c.fetchone():
            raise HTTPException(400, "اسم المستخدم موجود مسبقاً")
        c.execute("INSERT INTO users(username,password,role,branch) VALUES(?,?,?,?)",
                  (username, _hash(body.password), "reporter", body.branch or ""))
        new_id = c.lastrowid
    log_event("reporter_created_by_super", new_username=username, superuser=cu["username"])
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "create_reporter", f"إنشاء مراقب جديد: {username} (فرع: {body.branch or 'بدون'})",
                    request.client.host if request.client else "")
    return {"id": new_id, "username": username, "role": "reporter", "branch": body.branch or ""}


@app.put("/superuser/reporters/{uid}")
async def superuser_update_reporter(
    request: Request,
    uid: int,
    body: AdminUpdate,
    cu: dict = Depends(require_superuser)
):
    """تعديل اسم المستخدم أو كلمة مرور مراقب."""
    new_username = body.username.strip()
    if len(new_username) < 2:
        raise HTTPException(400, "اسم المستخدم قصير جداً")
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, username, role FROM users WHERE id=?", (uid,))
        target = c.fetchone()
        if not target:
            raise HTTPException(404, "المستخدم غير موجود")
        if target["role"] != "reporter":
            raise HTTPException(403, "لا يمكن تعديل إلا حسابات المراقبين من هنا")
        c.execute("SELECT id FROM users WHERE username=? AND id!=?", (new_username, uid))
        if c.fetchone():
            raise HTTPException(400, "اسم المستخدم مستخدم من قِبل شخص آخر")
        old_username = target["username"]
        if body.password:
            validate_password(body.password)
            c.execute("UPDATE users SET username=?,password=? WHERE id=?",
                      (new_username, _hash(body.password), uid))
        else:
            c.execute("UPDATE users SET username=? WHERE id=?", (new_username, uid))
        if body.branch is not None:
            c.execute("UPDATE users SET branch=? WHERE id=?", (body.branch, uid))
    log_event("reporter_updated_by_super", uid=uid, new_username=new_username, superuser=cu["username"])
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "update_reporter",
                    f"تعديل مراقب #{uid} من '{old_username}' إلى '{new_username}'"
                    + (" (مع تغيير كلمة المرور)" if body.password else ""),
                    request.client.host if request.client else "")
    return {"message": "تم التحديث", "id": uid, "username": new_username}


@app.delete("/superuser/reporters/{uid}")
async def superuser_delete_reporter(
    request: Request,
    uid: int,
    cu: dict = Depends(require_superuser)
):
    """حذف حساب مراقب."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, username, role FROM users WHERE id=?", (uid,))
        target = c.fetchone()
        if not target:
            raise HTTPException(404, "المستخدم غير موجود")
        if target["role"] != "reporter":
            raise HTTPException(403, "لا يمكن حذف إلا حسابات المراقبين من هنا")
        deleted_username = target["username"]
        c.execute("DELETE FROM users WHERE id=?", (uid,))
    log_event("reporter_deleted_by_super", uid=uid, username=deleted_username, superuser=cu["username"])
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "delete_reporter", f"حذف مراقب: {deleted_username} (id={uid})",
                    request.client.host if request.client else "")
    return {"message": f"تم حذف المراقب: {deleted_username}"}


@app.get("/superuser/audit-logs")
async def superuser_audit_logs(
    cu: dict = Depends(require_superuser),
    limit: int = 200,
    offset: int = 0,
    role: Optional[str] = None,
    username: Optional[str] = None,
):
    """
    عرض سجل التدقيق الكامل.
    يمكن الفلترة بـ ?role=admin أو ?username=xyz
    """
    limit = min(limit, 500)
    base  = "SELECT * FROM audit_logs"
    conditions: list[str] = []
    params: list           = []
    if role:
        conditions.append("role=?"); params.append(role)
    if username:
        conditions.append("username LIKE ?"); params.append(f"%{username}%")
    if conditions:
        base += " WHERE " + " AND ".join(conditions)
    base += " ORDER BY id DESC LIMIT ? OFFSET ?"
    params += [limit, offset]
    with get_db() as conn:
        c = conn.cursor()
        c.execute(base, params)
        rows = [dict(r) for r in c.fetchall()]
        # total count
        count_q = "SELECT COUNT(*) as cnt FROM audit_logs"
        if conditions:
            count_q += " WHERE " + " AND ".join(conditions)
        c.execute(count_q, params[:-2])
        total = c.fetchone()["cnt"]
    return {"total": total, "limit": limit, "offset": offset, "items": rows}


@app.get("/superuser/audit-logs/summary")
async def audit_logs_summary(cu: dict = Depends(require_superuser)):
    """ملخص إحصائي سريع لسجل التدقيق."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT role, COUNT(*) as count FROM audit_logs GROUP BY role""")
        by_role = {r["role"]: r["count"] for r in c.fetchall()}
        c.execute("""SELECT action, COUNT(*) as count FROM audit_logs
                     GROUP BY action ORDER BY count DESC LIMIT 10""")
        top_actions = [dict(r) for r in c.fetchall()]
        c.execute("""SELECT username, role, action, created_at
                     FROM audit_logs ORDER BY id DESC LIMIT 20""")
        recent = [dict(r) for r in c.fetchall()]
    return {"by_role": by_role, "top_actions": top_actions, "recent": recent}


# ══════════════════════════════════════════════════════
# 23. DRIVER REQUESTS
# ══════════════════════════════════════════════════════

REQUEST_TYPES = {
    "odometer_change": "تغيير قراءة العداد",
    "permission":      "طلب إذن",
    "leave":           "طلب إجازة",
    "other":           "طلب آخر",
}

@app.post("/requests", status_code=201)
async def create_request(body: dict, cu: dict = Depends(get_user)):
    req_type     = (body.get("type") or "").strip()
    notes        = (body.get("notes") or "").strip()
    new_odometer = body.get("new_odometer")
    trip_id      = body.get("trip_id")

    if req_type not in REQUEST_TYPES:
        raise HTTPException(400, "نوع الطلب غير صالح")

    driver_id = cu.get("driver_id")
    if not driver_id:
        raise HTTPException(403, "فقط السائقون يمكنهم تقديم الطلبات")

    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""INSERT INTO driver_requests
                     (driver_id,type,notes,new_odometer,trip_id,status,created_at)
                     VALUES(?,?,?,?,?,'pending',?)""",
                  (driver_id, req_type, notes,
                   float(new_odometer) if new_odometer else None,
                   int(trip_id) if trip_id else None, now))
        rid = c.lastrowid
    log_event("request_created", driver_id=driver_id, request_id=rid, type=req_type)
    return {"id": rid, "status": "pending"}


@app.get("/requests")
async def get_requests(cu: dict = Depends(get_user), branch: Optional[str] = None):
    with get_db() as conn:
        c = conn.cursor()
        if cu["role"] == "driver":
            driver_id = cu.get("driver_id")
            c.execute("""SELECT r.*, d.name as driver_name, d.branch as driver_branch
                         FROM driver_requests r
                         LEFT JOIN drivers d ON r.driver_id=d.id
                         WHERE r.driver_id=?
                         ORDER BY r.id DESC""", (driver_id,))
        else:
            # تحديد فلتر الفرع الفعّال
            eff_branch = _effective_branch(cu, branch)
            if eff_branch:
                c.execute("""SELECT r.*, d.name as driver_name, d.branch as driver_branch
                             FROM driver_requests r
                             LEFT JOIN drivers d ON r.driver_id=d.id
                             WHERE d.branch=?
                             ORDER BY r.id DESC""", (eff_branch,))
            else:
                c.execute("""SELECT r.*, d.name as driver_name, d.branch as driver_branch
                             FROM driver_requests r
                             LEFT JOIN drivers d ON r.driver_id=d.id
                             ORDER BY r.id DESC""")
        return [dict(row) for row in c.fetchall()]


@app.get("/requests/pending-count")
async def requests_pending_count(cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        c = conn.cursor()
        eff_branch = _branch_filter(cu)
        if eff_branch:
            c.execute("""SELECT COUNT(*) as cnt FROM driver_requests r
                         LEFT JOIN drivers d ON r.driver_id=d.id
                         WHERE r.status='pending' AND d.branch=?""", (eff_branch,))
        else:
            c.execute("SELECT COUNT(*) as cnt FROM driver_requests WHERE status='pending'")
        row = c.fetchone()
        return {"count": row["cnt"] if row else 0}


PRIORITY_VALUES = {"urgent", "medium", "deferrable"}

@app.put("/requests/{rid}")
async def handle_request(rid: int, body: dict, cu: dict = Depends(get_user)):
    """
    Admin: يضيف priority + admin_notes + admin_message ويحيل للسوبر.
    Superuser: يمكنه الموافقة/الرفض مباشرة (يتجاوز دور الأدمن).
    """
    if cu["role"] not in ("admin", "superuser"):
        raise HTTPException(403, "غير مصرح")

    admin_notes   = (body.get("admin_notes") or "").strip()
    admin_message = (body.get("admin_message") or "").strip()
    priority      = (body.get("priority") or "medium").strip()
    forward       = bool(body.get("forward_to_super", False))

    if priority not in PRIORITY_VALUES:
        priority = "medium"

    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT * FROM driver_requests WHERE id=?", (rid,))
        req = c.fetchone()
        if not req:
            raise HTTPException(404, "الطلب غير موجود")

        if cu["role"] == "admin":
            # الأدمن يضيف ملاحظاته ويحيل للسوبر — لا يقرر نهائياً
            new_status = "pending_super" if forward else req["status"]
            c.execute("""UPDATE driver_requests
                         SET admin_notes=?, admin_message=?, priority=?,
                             forwarded_to_super=?, handled_by=?, handled_at=?,
                             status=?
                         WHERE id=?""",
                      (admin_notes, admin_message, priority,
                       1 if forward else req["forwarded_to_super"],
                       cu["username"], now, new_status, rid))
            log_event("request_forwarded_by_admin", request_id=rid,
                      priority=priority, forwarded=forward, admin=cu["username"])
            return {"ok": True, "forwarded": forward}

        else:  # superuser — قرار نهائي
            status = (body.get("status") or "").strip()
            if status not in ("approved", "rejected"):
                raise HTTPException(400, "الحالة يجب أن تكون approved أو rejected")
            super_notes = (body.get("super_notes") or "").strip()
            c.execute("""UPDATE driver_requests
                         SET status=?, super_decision=?, super_notes=?,
                             super_handled_by=?, super_handled_at=?
                         WHERE id=?""",
                      (status, status, super_notes, cu["username"], now, rid))
            log_event("request_decided_by_super", request_id=rid,
                      status=status, superuser=cu["username"])
            return {"ok": True, "status": status}


@app.get("/requests/pending-super-count")
async def pending_super_count(cu: dict = Depends(require_superuser)):
    """عدد الطلبات المحالة للسوبر وبانتظار قرار."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT COUNT(*) as cnt FROM driver_requests WHERE status='pending_super'")
        return {"count": c.fetchone()["cnt"]}



# ══════════════════════════════════════════════════════
# 24-B. BRANCHES LIST — لقائمة الفروع المتاحة
# ══════════════════════════════════════════════════════

@app.get("/reports/branches")
async def list_branches(cu: dict = Depends(require_admin_or_reporter)):
    """قائمة الفروع الثابتة للنظام."""
    eff_branch = _branch_filter(cu)
    if eff_branch:
        return {"branches": [eff_branch]}
    return {"branches": BRANCHES}

# ══════════════════════════════════════════════════════
# 25. OPERATIONAL CARD REPORT
# ══════════════════════════════════════════════════════

@app.get("/reports/operational")
async def operational_report(
    car_id:      Optional[int] = None,
    date_from:   Optional[str] = None,
    date_to:     Optional[str] = None,
    report_type: str = "summary",   # summary | detailed
    group_by:    str = "day",       # day | month | year  (used only when detailed)
    cu: dict = Depends(require_admin_or_reporter)
):
    """
    report_type=summary  → one total row for the entire period
    report_type=detailed:
        group_by=day   → one row per day   (with day name)
        group_by=month → one row per month
        group_by=year  → one row per year
    """
    with get_db() as conn:
        c = conn.cursor()

        base_where = [
            "t.end_time IS NOT NULL",
            "t.start_odometer IS NOT NULL",
            "t.end_odometer   IS NOT NULL",
            "t.end_odometer    > t.start_odometer",
            # ✅ تصفية الرحلات الشاذة: المسافة الواحدة يجب أن تكون أقل من 2000 كم
            "(t.end_odometer - t.start_odometer) < 2000",
        ]
        params = []

        if car_id:
            base_where.append("t.car_id = ?")
            params.append(car_id)
        if date_from:
            base_where.append("date(t.start_time) >= ?")
            params.append(date_from)
        if date_to:
            base_where.append("date(t.start_time) <= ?")
            params.append(date_to)

        where_sql = " AND ".join(base_where)

        if report_type == "summary":
            # Single summary row
            c.execute(f"""
                SELECT
                    COUNT(t.id)                                          AS trip_count,
                    ROUND(SUM(t.end_odometer - t.start_odometer), 1)    AS total_km,
                    MIN(t.start_time)                                    AS first_date,
                    (SELECT t2.start_odometer FROM trips t2
                     WHERE {where_sql.replace('t.', 't2.')}
                     ORDER BY t2.start_time ASC  LIMIT 1)               AS first_odo,
                    (SELECT t2.end_odometer   FROM trips t2
                     WHERE {where_sql.replace('t.', 't2.')}
                     ORDER BY t2.end_time   DESC LIMIT 1)               AS last_odo
                FROM trips t
                WHERE {where_sql}
            """, params * 3)
            row = dict(c.fetchone())
            rows = [{
                "period":     "إجمالي الفترة",
                "trip_count": row["trip_count"],
                "total_km":   row["total_km"] or 0,
                "first_date": (row["first_date"] or "")[:10],
                "min_odo":    row["first_odo"],
                "max_odo":    row["last_odo"],
            }]

        else:
            # Detailed — group by day / month / year
            if group_by == "day":
                grp = "strftime('%Y-%m-%d', t.start_time)"
            elif group_by == "month":
                grp = "strftime('%Y-%m', t.start_time)"
            else:  # year
                grp = "strftime('%Y', t.start_time)"

            c.execute(f"""
                SELECT
                    {grp}                                               AS period,
                    COUNT(t.id)                                         AS trip_count,
                    ROUND(SUM(t.end_odometer - t.start_odometer), 1)   AS total_km,
                    MIN(t.start_time)                                   AS first_date,
                    MIN(t.start_odometer)                               AS min_odo,
                    MAX(t.end_odometer)                                 AS max_odo
                FROM trips t
                WHERE {where_sql}
                GROUP BY {grp}
                ORDER BY {grp}
            """, params)
            rows = [dict(r) for r in c.fetchall()]
            # For each row, first_date already reflects the earliest trip in that period
            for r in rows:
                if r.get("first_date"):
                    r["first_date"] = r["first_date"][:10]

        # Car info
        car_info = None
        if car_id:
            c.execute("SELECT * FROM cars WHERE id=?", (car_id,))
            row = c.fetchone()
            if row:
                car_info = dict(row)

        # Grand totals — first_odo = start_odometer of earliest trip,
        #                last_odo  = end_odometer   of latest  trip
        c.execute(f"""
            SELECT
                COUNT(t.id)                                       AS tc,
                ROUND(SUM(t.end_odometer - t.start_odometer), 1) AS tkm,
                (SELECT t2.start_odometer FROM trips t2
                 WHERE {where_sql.replace('t.', 't2.')}
                 ORDER BY t2.start_time ASC  LIMIT 1)             AS first_odo,
                (SELECT t2.end_odometer   FROM trips t2
                 WHERE {where_sql.replace('t.', 't2.')}
                 ORDER BY t2.end_time   DESC LIMIT 1)             AS last_odo
            FROM trips t WHERE {where_sql}
        """, params * 3)
        tot = dict(c.fetchone())

        return {
            "rows":        rows,
            "car_info":    car_info,
            "total_km":    tot["tkm"]    or 0,
            "total_trips": tot["tc"]     or 0,
            "min_odo":     tot["first_odo"],
            "max_odo":     tot["last_odo"],
            "report_type": report_type,
            "group_by":    group_by,
        }



# ══════════════════════════════════════════════════════
# 26-B. OPERATIONAL REPORT — PER DRIVER BREAKDOWN
# ══════════════════════════════════════════════════════

@app.get("/reports/operational-by-driver")
async def operational_report_by_driver(
    car_id:    int,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    cu: dict = Depends(require_admin_or_reporter)
):
    """
    نفس منطق التقرير التشغيلي لكن مقسّم بحسب السائق.
    يُرجع:
      - car_info : بيانات المركبة (تشمل نوع الوقود)
      - fuel_records : سجلات الوقود للمركبة في الفترة
      - drivers : قائمة مرتبة بالسائقين الذين قادوا المركبة
                  كل سائق يحمل: name, trip_count, total_km, min_odo, max_odo, first_date
    """
    with get_db() as conn:
        c = conn.cursor()

        # ── Car info ──
        c.execute("SELECT * FROM cars WHERE id=?", (car_id,))
        row = c.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="المركبة غير موجودة")
        car_info = dict(row)

        # ── Base WHERE filters (بدون driver_id — هنضيفه لاحقاً) ──
        base_conds = [
            "t.car_id = ?",
            "t.end_time IS NOT NULL",
            "t.start_odometer IS NOT NULL",
            "t.end_odometer > t.start_odometer",
            "(t.end_odometer - t.start_odometer) < 2000",
        ]
        base_params = [car_id]
        if date_from:
            base_conds.append("date(t.start_time) >= ?")
            base_params.append(date_from)
        if date_to:
            base_conds.append("date(t.start_time) <= ?")
            base_params.append(date_to)
        where_sql = " AND ".join(base_conds)

        # ── السائقون الذين قادوا المركبة في الفترة ──
        c.execute(f"""
            SELECT
                d.id                                                    AS driver_id,
                d.name                                                  AS driver_name,
                COUNT(t.id)                                             AS trip_count,
                ROUND(SUM(t.end_odometer - t.start_odometer), 1)       AS total_km,
                MIN(t.start_time)                                       AS first_date,
                MIN(t.start_odometer)                                   AS min_odo,
                MAX(t.end_odometer)                                     AS max_odo
            FROM trips t
            JOIN drivers d ON t.driver_id = d.id
            WHERE {where_sql}
            GROUP BY d.id, d.name
            ORDER BY total_km DESC
        """, base_params)
        drivers = []
        for row in c.fetchall():
            r = dict(row)
            r["first_date"] = (r["first_date"] or "")[:10]
            # أول وآخر عداد بالزمن لهذا السائق على هذه المركبة
            c.execute(f"""
                SELECT start_odometer FROM trips t
                WHERE {where_sql} AND t.driver_id = ?
                ORDER BY t.start_time ASC LIMIT 1
            """, base_params + [r["driver_id"]])
            fo = c.fetchone()
            c.execute(f"""
                SELECT end_odometer FROM trips t
                WHERE {where_sql} AND t.driver_id = ?
                ORDER BY t.end_time DESC LIMIT 1
            """, base_params + [r["driver_id"]])
            lo = c.fetchone()
            r["first_odo"] = fo[0] if fo else None
            r["last_odo"]  = lo[0] if lo else None
            drivers.append(r)

        # ── Fuel records للمركبة في الفترة ──
        fuel_conds = ["w.vehicle_id = ?", "w.type LIKE 'fuel_%'"]
        fuel_params = [car_id]
        if date_from:
            fuel_conds.append("date(w.created_at) >= ?")
            fuel_params.append(date_from)
        if date_to:
            fuel_conds.append("date(w.created_at) <= ?")
            fuel_params.append(date_to)
        c.execute(f"""
            SELECT w.type, SUM(w.quantity) AS total_qty, SUM(w.price) AS total_price,
                   COUNT(*) AS count
            FROM workshop_records w
            WHERE {' AND '.join(fuel_conds)}
            GROUP BY w.type
        """, fuel_params)
        fuel_records = [dict(r) for r in c.fetchall()]

        # ── Grand totals ──
        c.execute(f"""
            SELECT COUNT(t.id) AS tc,
                   ROUND(SUM(t.end_odometer - t.start_odometer),1) AS tkm
            FROM trips t WHERE {where_sql}
        """, base_params)
        tot = dict(c.fetchone())

    return {
        "car_info":     car_info,
        "drivers":      drivers,
        "fuel_records": fuel_records,
        "total_km":     tot["tkm"] or 0,
        "total_trips":  tot["tc"]  or 0,
    }


# ══════════════════════════════════════════════════════
# 26. ANOMALOUS TRIPS DIAGNOSTIC
# ══════════════════════════════════════════════════════

@app.get("/reports/anomalous-trips")
async def anomalous_trips(
    car_id:    Optional[int] = None,
    threshold: float = 2000.0,          # كم — أي رحلة أكبر من كده تُعتبر شاذة
    cu: dict = Depends(require_admin_or_reporter)
):
    """
    يُرجع الرحلات التي تجاوزت حدّ المسافة المنطقية (default 2000 كم للرحلة الواحدة).
    يُستخدم لاكتشاف قراءات العداد الخاطئة.
    """
    with get_db() as conn:
        c = conn.cursor()
        q = """
            SELECT
                t.id,
                t.start_time,
                t.end_time,
                t.start_odometer,
                t.end_odometer,
                ROUND(t.end_odometer - t.start_odometer, 1) AS trip_km,
                d.name  AS driver_name,
                ca.plate AS car_plate
            FROM trips t
            LEFT JOIN drivers d  ON t.driver_id = d.id
            LEFT JOIN cars    ca ON t.car_id    = ca.id
            WHERE t.end_time IS NOT NULL
              AND t.start_odometer IS NOT NULL
              AND t.end_odometer   IS NOT NULL
              AND (t.end_odometer - t.start_odometer) >= ?
        """
        p = [threshold]
        if car_id:
            q += " AND t.car_id = ?"
            p.append(car_id)
        q += " ORDER BY trip_km DESC"
        c.execute(q, p)
        rows = [dict(r) for r in c.fetchall()]
    return {"count": len(rows), "threshold_km": threshold, "trips": rows}



# ══════════════════════════════════════════════════════
# 26. BULK IMPORT — CSV / Excel
# ══════════════════════════════════════════════════════

DRIVER_IMPORT_COLS  = ["name","phone","fixed_number","driver_license_expiry","vehicle_license_expiry",
                       "national_id","birth_date","branch","status"]
CAR_IMPORT_COLS     = ["plate","model","car_name","car_code","chassis",
                       "engine_number","year","project","branch",
                       "car_license_expiry","equipment_type","sector","status"]

# خريطة أسماء الأعمدة العربية → الإنجليزية
ARABIC_COL_MAP = {
    "اسم السائق":        "name",
    "الاسم":             "name",
    "name":              "name",
    "رقم ثابت":          "fixed_number",
    "الرقم الثابت":      "fixed_number",
    "fixed_number":      "fixed_number",
    "تليفون السائق":     "phone",
    "الهاتف":            "phone",
    "تليفون":            "phone",
    "phone":             "phone",
    "انتهاء رخصة السائق":  "driver_license_expiry",
    "رخصة السائق":         "driver_license_expiry",
    "driver_license_expiry": "driver_license_expiry",
    "انتهاء رخصة السيارة":  "vehicle_license_expiry",
    "انتهاء رخصة العربة":   "vehicle_license_expiry",
    "رخصة السيارة":          "vehicle_license_expiry",
    "vehicle_license_expiry": "vehicle_license_expiry",
    "الرقم التعريفي":    "national_id",
    "national_id":       "national_id",
    "تاريخ الميلاد":     "birth_date",
    "birth_date":        "birth_date",
    "الفرع":             "branch",
    "branch":            "branch",
    "الحالة":            "status",
    "status":            "status",
    # للمركبات
    "رقم اللوحة":        "plate",
    "plate":             "plate",
    "الموديل":           "model",
    "model":             "model",
    "اسم المركبة":       "car_name",
    "car_name":          "car_name",
    "كود المركبة":       "car_code",
    "car_code":          "car_code",
    "رقم الشاسيه":       "chassis",
    "chassis":           "chassis",
    "رقم الموتور":       "engine_number",
    "engine_number":     "engine_number",
    "سنة الصنع":         "year",
    "year":              "year",
    "المشروع":           "project",
    "project":           "project",
    "انتهاء رخصة المركبة": "car_license_expiry",
    "car_license_expiry": "car_license_expiry",
    "نوع المعدة":        "equipment_type",
    "equipment_type":    "equipment_type",
    "القطاع":            "sector",
    "sector":            "sector",
}

def _normalize_date(val: str) -> str:
    """تحويل تنسيقات التاريخ المختلفة إلى YYYY-MM-DD."""
    if not val or val.strip() in ("", "None", "—"):
        return ""
    val = val.strip().replace("/", "-")
    # handle YYYY-M-D → YYYY-MM-DD
    parts = val.split("-")
    if len(parts) == 3:
        try:
            y, m, d = parts
            # لو السنة مش 4 أرقام تجاهل
            if len(y) != 4:
                return ""
            return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
        except Exception:
            return ""
    return ""

def _parse_upload_file(content: bytes, filename: str) -> list[dict]:
    """
    Parse CSV or Excel file and return list of row dicts with English keys.
    يدعم:
    - الأعمدة العربية والإنجليزية
    - الـ header في أي صف (يبحث عن أول صف فيه بيانات حقيقية)
    - أعمدة فارغة قبل البيانات (زي Book1.xlsx)
    """
    ext = Path(filename).suffix.lower()

    if ext in (".xlsx", ".xls"):
        if not OPENPYXL_AVAILABLE:
            raise HTTPException(400, "ملفات Excel غير مدعومة — ثبّت openpyxl أو استخدم CSV")
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))

        # ابحث عن صف الـ headers: أول صف فيه على الأقل عمودين من أعمدتنا المعروفة
        header_row_idx = None
        headers_raw = []
        for i, row in enumerate(all_rows):
            cells = [str(c).strip() if c is not None else "" for c in row]
            mapped = sum(1 for c in cells if c.lower() in ARABIC_COL_MAP or c in ARABIC_COL_MAP)
            if mapped >= 2:
                header_row_idx = i
                headers_raw = cells
                break

        if header_row_idx is None:
            return []

        # map headers
        headers_mapped = [ARABIC_COL_MAP.get(h, ARABIC_COL_MAP.get(h.lower(), h)) for h in headers_raw]

        result = []
        for row in all_rows[header_row_idx + 1:]:
            cells = [str(v).strip() if v is not None else "" for v in row]
            # تجاهل الصفوف الفارغة كلياً
            if not any(cells):
                continue
            d = {headers_mapped[i]: cells[i] for i in range(min(len(headers_mapped), len(cells)))}
            result.append(d)
        return result

    elif ext == ".csv":
        text = content.decode("utf-8-sig", errors="replace")
        reader = csv.DictReader(io.StringIO(text))
        rows = []
        for row in reader:
            d = {}
            for k, v in row.items():
                k = k.strip() if k else k
                mapped_key = ARABIC_COL_MAP.get(k, ARABIC_COL_MAP.get((k or "").lower(), k))
                d[mapped_key] = (v.strip() if v else "")
            rows.append(d)
        return rows
    else:
        raise HTTPException(400, "نوع الملف غير مدعوم — استخدم CSV أو Excel (.xlsx)")


def _validate_driver_row(row: dict, idx: int, admin_branch: Optional[str]) -> dict:
    """
    التحقق من صف سائق واحد.
    - username = الكلمة الأولى من الاسم (تلقائي)
    - password = الرقم الثابت (تلقائي)
    """
    errors = []
    name         = row.get("name", "").strip()
    fixed_number = row.get("fixed_number", "").strip()
    phone        = row.get("phone", "").strip()

    # توليد username وpassword تلقائياً
    username = row.get("username", "").strip()
    password = row.get("password", "").strip()
    if not username and name:
        username = name.split()[0]   # الكلمة الأولى من الاسم
    if not password and fixed_number:
        password = fixed_number      # الرقم الثابت

    if not name:          errors.append("الاسم مطلوب")
    if not username:      errors.append("اسم المستخدم مطلوب")
    if not password:      errors.append("كلمة المرور مطلوبة (الرقم الثابت فارغ)")
    if not fixed_number:  errors.append("الرقم الثابت مطلوب")

    branch = admin_branch if admin_branch else row.get("branch", "").strip()

    return {
        "row_index":               idx,
        "valid":                   len(errors) == 0,
        "errors":                  errors,
        "name":                    name,
        "phone":                   phone,
        "username":                username,
        "password":                password,
        "national_id":             row.get("national_id", "").strip(),
        "birth_date":              _normalize_date(row.get("birth_date", "")),
        "driver_license_expiry":   _normalize_date(row.get("driver_license_expiry", "")),
        "vehicle_license_expiry":  _normalize_date(row.get("vehicle_license_expiry", "")),
        "branch":                  branch,
        "fixed_number":            fixed_number,
        "status":                  row.get("status", "active").strip() or "active",
    }


def _validate_car_row(row: dict, idx: int, admin_branch: Optional[str]) -> dict:
    """Validate a single car row."""
    errors = []
    plate = row.get("plate", "").strip()
    model = row.get("model", "").strip()
    if not plate: errors.append("رقم اللوحة مطلوب")
    if not model: errors.append("الموديل مطلوب")
    branch = admin_branch if admin_branch else row.get("branch", "").strip()
    return {
        "row_index":       idx,
        "valid":           len(errors) == 0,
        "errors":          errors,
        "plate":           plate,
        "model":           model,
        "car_name":        row.get("car_name", "").strip(),
        "car_code":        row.get("car_code", "").strip(),
        "chassis":         row.get("chassis", "").strip(),
        "engine_number":   row.get("engine_number", "").strip(),
        "year":            row.get("year", "").strip(),
        "project":         row.get("project", "").strip(),
        "branch":          branch,
        "car_license_expiry": _normalize_date(row.get("car_license_expiry", "")),
        "equipment_type":  row.get("equipment_type", "").strip(),
        "sector":          row.get("sector", "").strip(),
        "status":          row.get("status", "available").strip() or "available",
    }


@app.post("/import/preview")
async def import_preview(
    request: Request,
    file: UploadFile = File(...),
    import_type: str = "drivers",   # drivers | cars
    cu: dict = Depends(require_admin)
):
    """
    استعراض البيانات من ملف CSV/Excel قبل الرفع الفعلي.
    يرجع:
      - rows: قائمة الصفوف مع حالة التحقق (valid/errors)
      - summary: { total, valid, invalid }
    """
    if import_type not in ("drivers", "cars"):
        raise HTTPException(400, "import_type يجب أن يكون drivers أو cars")

    content = await file.read()
    if len(content) > 5 * 1024 * 1024:  # 5 MB max
        raise HTTPException(400, "حجم الملف يتجاوز 5 MB")

    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows:
        raise HTTPException(400, "الملف فارغ أو لا يحتوي على بيانات")

    admin_branch = _branch_filter(cu)

    if import_type == "drivers":
        rows = [_validate_driver_row(r, i+1, admin_branch) for i, r in enumerate(raw_rows)]
    else:
        rows = [_validate_car_row(r, i+1, admin_branch) for i, r in enumerate(raw_rows)]

    valid_count   = sum(1 for r in rows if r["valid"])
    invalid_count = len(rows) - valid_count

    log_event("import_preview", import_type=import_type,
              total=len(rows), valid=valid_count, admin=cu["username"])
    return {
        "import_type": import_type,
        "rows": rows,
        "summary": {"total": len(rows), "valid": valid_count, "invalid": invalid_count},
    }


@app.post("/import/confirm")
async def import_confirm(
    request: Request,
    file: UploadFile = File(...),
    import_type: str = "drivers",
    skip_errors: bool = False,    # لو True يكمل ويتجاهل الصفوف الغلط
    cu: dict = Depends(require_admin)
):
    """
    تأكيد رفع البيانات من ملف CSV/Excel.
    - skip_errors=false (الافتراضي): يوقف لو في أي صف غلط
    - skip_errors=true: يرفع الصفوف الصحيحة ويتجاهل الغلط
    يرجع: { inserted, skipped, errors }
    """
    if import_type not in ("drivers", "cars"):
        raise HTTPException(400, "import_type يجب أن يكون drivers أو cars")

    content = await file.read()
    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows:
        raise HTTPException(400, "الملف فارغ")

    admin_branch = _branch_filter(cu)

    if import_type == "drivers":
        rows = [_validate_driver_row(r, i+1, admin_branch) for i, r in enumerate(raw_rows)]
    else:
        rows = [_validate_car_row(r, i+1, admin_branch) for i, r in enumerate(raw_rows)]

    invalid = [r for r in rows if not r["valid"]]
    if invalid and not skip_errors:
        raise HTTPException(422, {
            "message": f"يوجد {len(invalid)} صف بأخطاء — استخدم skip_errors=true لتجاهلها",
            "invalid_rows": invalid
        })

    valid_rows = [r for r in rows if r["valid"]]
    inserted, skipped_rows = [], []
    now = datetime.utcnow().isoformat() + "Z"

    with get_db() as conn:
        c = conn.cursor()
        if import_type == "drivers":
            for row in valid_rows:
                try:
                    # اللي مش ممكن يتكرر هو الرقم الثابت — مش اسم المستخدم
                    if row["fixed_number"]:
                        c.execute("SELECT id FROM drivers WHERE fixed_number=?", (row["fixed_number"],))
                        if c.fetchone():
                            row["errors"] = [f"الرقم الثابت '{row['fixed_number']}' موجود مسبقاً"]
                            skipped_rows.append(row)
                            continue
                    c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                              (row["username"], _hash(row["password"]), "driver"))
                    uid = c.lastrowid
                    c.execute("""INSERT INTO drivers(name,phone,status,user_id,national_id,
                                 birth_date,driver_license_expiry,vehicle_license_expiry,
                                 branch,fixed_number)
                                 VALUES(?,?,?,?,?,?,?,?,?,?)""",
                              (row["name"], row["phone"], row["status"], uid,
                               row["national_id"], row["birth_date"],
                               row["driver_license_expiry"], row["vehicle_license_expiry"],
                               row["branch"], row["fixed_number"]))
                    inserted.append({"row_index": row["row_index"], "name": row["name"],
                                     "fixed_number": row["fixed_number"]})
                except Exception as e:
                    row["errors"] = [str(e)]
                    skipped_rows.append(row)

        else:  # cars
            for row in valid_rows:
                try:
                    c.execute("SELECT id FROM cars WHERE plate=?", (row["plate"],))
                    if c.fetchone():
                        row["errors"] = [f"رقم اللوحة '{row['plate']}' موجود مسبقاً"]
                        skipped_rows.append(row)
                        continue
                    c.execute("""INSERT INTO cars(plate,model,status,car_name,car_code,chassis,
                                 engine_number,year,project,branch,
                                 car_license_expiry,equipment_type,sector)
                                 VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                              (row["plate"], row["model"], row["status"],
                               row["car_name"], row["car_code"], row["chassis"],
                               row["engine_number"], row["year"], row["project"], row["branch"],
                               row["car_license_expiry"], row["equipment_type"], row["sector"]))
                    inserted.append({"row_index": row["row_index"], "plate": row["plate"]})
                except Exception as e:
                    row["errors"] = [str(e)]
                    skipped_rows.append(row)

    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    f"bulk_import_{import_type}",
                    f"استيراد {len(inserted)} {import_type} من ملف",
                    request.client.host if request.client else "")
    log_event("import_confirmed", import_type=import_type,
              inserted=len(inserted), skipped=len(skipped_rows), admin=cu["username"])
    return {
        "import_type": import_type,
        "inserted": len(inserted),
        "skipped":  len(skipped_rows),
        "inserted_items": inserted,
        "skipped_items":  skipped_rows,
    }


@app.get("/import/template/{import_type}")
async def import_template(import_type: str, cu: dict = Depends(require_admin)):
    """تحميل قالب CSV فارغ للسائقين أو المركبات."""
    if import_type == "drivers":
        cols = ["اسم السائق","رقم ثابت","تليفون السائق","انتهاء رخصة السائق","انتهاء رخصة السيارة","الفرع","الحالة"]
        sample = ["محمد أحمد","123456","01012345678","2026-12-31","2026-12-31","القاهرة","active"]
    elif import_type == "cars":
        cols = CAR_IMPORT_COLS
        sample = ["أ-ب-1234","تويوتا لاندكروزر","هيلاكس","T-001","CH123456",
                  "ENG789","2022","مشروع النيل","القاهرة","2026-06-30",
                  "وسائل انتقال","المنطقة أ","available"]
    else:
        raise HTTPException(400, "import_type يجب أن يكون drivers أو cars")

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(cols)
    writer.writerow(sample)
    csv_content = output.getvalue()

    from fastapi.responses import Response
    return Response(
        content=csv_content.encode("utf-8-sig"),   # utf-8-sig for Excel compatibility
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={import_type}_template.csv"}
    )


# ══════════════════════════════════════════════════════
# 24. ENTRY POINT
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

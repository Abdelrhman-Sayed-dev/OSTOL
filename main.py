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
            role TEXT NOT NULL CHECK(role IN('superuser','admin','driver','reporter','supervisor_workshop','supervisor_field')),
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

        # VOICE NOTES (سوبر يوزر → أدمنز أو سائقين)
        c.execute("""CREATE TABLE IF NOT EXISTS voice_notes(
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            sender_id    INTEGER NOT NULL,
            target_group TEXT NOT NULL CHECK(target_group IN('admins','drivers')),
            audio_data   TEXT NOT NULL,
            duration_sec REAL DEFAULT 0,
            created_at   TEXT NOT NULL,
            expires_at   TEXT NOT NULL,
            play_count   INTEGER DEFAULT 0,
            max_plays    INTEGER DEFAULT 2,
            is_deleted   INTEGER DEFAULT 0
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_vnotes_group   ON voice_notes(target_group)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_vnotes_expires ON voice_notes(expires_at)")

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
                role TEXT NOT NULL CHECK(role IN('superuser','admin','driver','reporter','supervisor_workshop','supervisor_field')),
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

    # ── جدول الصيانة الدورية ──
    c.execute("""CREATE TABLE IF NOT EXISTS maintenance_schedule (
        id                INTEGER PRIMARY KEY AUTOINCREMENT,
        car_id            INTEGER NOT NULL,
        maintenance_type  TEXT    NOT NULL,
        interval_km       REAL,
        interval_days     INTEGER,
        last_done_km      REAL,
        last_done_date    TEXT,
        alert_km_before   REAL    DEFAULT 500,
        alert_days_before INTEGER DEFAULT 7,
        notes             TEXT    DEFAULT '',
        created_at        TEXT    NOT NULL,
        updated_at        TEXT    NOT NULL,
        FOREIGN KEY (car_id) REFERENCES cars(id) ON DELETE CASCADE
    )""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_maint_car ON maintenance_schedule(car_id)")

    # ── تصحيح السجلات القديمة: price كان سعر/وحدة، دلوقتي يبقى إجمالي ──
    # نضرب السجلات اللي price فيها صغيرة (< 200) وعندها كمية > 0
    FUEL_TYPES_FOR_MIGRATION = [t for t in WORKSHOP_TYPES if t.startswith("fuel_")] + ["oil"]
    for ws_type in FUEL_TYPES_FOR_MIGRATION:
        try:
            # فقط السجلات اللي price فيها أصغر من 200 (سعر/لتر عادةً) وعندها كمية
            c.execute("""UPDATE workshop_records
                         SET price = price * COALESCE(quantity,1)
                         WHERE type=?
                           AND quantity IS NOT NULL AND quantity > 0
                           AND price > 0 AND price < 200
                           AND quantity != 1""",
                      (ws_type,))
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

ADMIN_LIKE_ROLES = ("superuser","admin","reporter","supervisor_workshop","supervisor_field")

def require_admin_or_reporter(cu: dict = Depends(get_user)):
    """Allows superuser, admin, reporter and supervisors (read-only) roles."""
    if cu["role"] not in ADMIN_LIKE_ROLES:
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
    username: str
    password: str
    role:     str
    branch:   Optional[str] = ""

    @validator('role')
    def validate_role(cls, v):
        allowed = ("admin","superuser","driver","reporter","supervisor_workshop","supervisor_field")
        if v not in allowed:
            raise ValueError(f"دور غير صالح. المتاح: {allowed}")
        return v

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
    """Ensure the hardcoded accounts exist for a role.
    Only removes accounts that WERE previously hardcoded but removed from the list.
    Dynamically added accounts (via API) are preserved.
    Skips rehashing on restart for accounts that already exist.
    """
    c.execute("SELECT id, username FROM users WHERE role=?", (role,))
    existing = {r["username"]: r["id"] for r in c.fetchall()}
    hardcoded_names = {u for u, _ in accounts}

    # Only delete accounts that are NOT in hardcoded list AND were hardcoded before
    # We track this via a special marker: accounts seeded by code get flag in a settings table
    c.execute("""CREATE TABLE IF NOT EXISTS _seeded_accounts (
        username TEXT PRIMARY KEY,
        role TEXT NOT NULL
    )""")
    c.execute("SELECT username FROM _seeded_accounts WHERE role=?", (role,))
    previously_seeded = {r["username"] for r in c.fetchall()}

    for uname, uid in list(existing.items()):
        # Only delete if it was previously seeded AND is no longer in hardcoded list
        if uname in previously_seeded and uname not in hardcoded_names:
            c.execute("DELETE FROM users WHERE id=?", (uid,))
            c.execute("DELETE FROM _seeded_accounts WHERE username=?", (uname,))
            log.info(f"Removed old seeded {role}: {uname}")

    for uname, pw in accounts:
        if uname not in existing:
            pw_hash = bcrypt.hashpw(pw.encode(), bcrypt.gensalt()).decode()
            c.execute("INSERT OR IGNORE INTO users(username,password,role) VALUES(?,?,?)",
                      (uname, pw_hash, role))
            log.info(f"Created {role}: {uname}")
        # Mark as seeded (or update marker if already exists)
        c.execute("INSERT OR REPLACE INTO _seeded_accounts(username,role) VALUES(?,?)",
                  (uname, role))

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

@app.put("/drivers/{did}/license")
async def update_driver_license(did: int, body: dict, cu: dict = Depends(require_admin)):
    """
    تحديث بيانات الرخصة من الأدمن.
    Input: { license_type: 'driver'|'vehicle', renewal_date: str, new_expiry_date: str }
    """
    license_type    = (body.get("license_type") or "").strip()
    renewal_date    = (body.get("renewal_date") or "").strip()
    new_expiry_date = (body.get("new_expiry_date") or "").strip()

    if license_type not in ("driver", "vehicle"):
        raise HTTPException(400, "نوع الرخصة غير صالح — driver أو vehicle")
    if not new_expiry_date:
        raise HTTPException(400, "تاريخ انتهاء الرخصة الجديد مطلوب")

    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id FROM drivers WHERE id=?", (did,))
        if not c.fetchone():
            raise HTTPException(404, "السائق غير موجود")
        if license_type == "driver":
            c.execute("UPDATE drivers SET driver_license_expiry=? WHERE id=?",
                      (new_expiry_date, did))
        else:
            c.execute("UPDATE drivers SET vehicle_license_expiry=? WHERE id=?",
                      (new_expiry_date, did))
        log_event("license_renewed", driver_id=did, type=license_type,
                  renewal_date=renewal_date, new_expiry=new_expiry_date, admin=cu["username"])
        return {"ok": True, "license_type": license_type, "new_expiry_date": new_expiry_date}

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
                     WHERE p.driver_id=? ORDER BY p.id DESC LIMIT 1""", (did,))
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
            # Parse category|subtype from equipment_type
            eq = d.get("equipment_type","") or ""
            parts = eq.split("|",1)
            d["vehicle_category"] = parts[0] if parts else ""
            d["vehicle_subtype"]  = parts[1] if len(parts)>1 else ""
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

@app.post("/permissions/reassign")
async def reassign_permission(body: dict, cu: dict = Depends(require_superuser)):
    """
    إعادة تخصيص مركبة من سائق لآخر.
    Input: { fixed_number: str, plate: str }
    - يجيب بيانات السائق الحالي والمركبة
    - بعد التأكيد يُلغي الصلاحية القديمة ويضيف الجديدة
    """
    fixed_number = (body.get("fixed_number") or "").strip()
    plate        = (body.get("plate") or "").strip()
    confirm      = body.get("confirm", False)

    if not fixed_number or not plate:
        raise HTTPException(400, "الرقم الثابت ورقم اللوحة مطلوبان")

    with get_db() as conn:
        c = conn.cursor()
        # جيب السائق الجديد
        c.execute("SELECT id, name, branch FROM drivers WHERE fixed_number=?", (fixed_number,))
        new_driver = c.fetchone()
        if not new_driver:
            raise HTTPException(404, f"لا يوجد سائق بالرقم الثابت '{fixed_number}'")

        # جيب المركبة
        c.execute("SELECT id, plate, model, car_name FROM cars WHERE plate=?", (plate,))
        car = c.fetchone()
        if not car:
            raise HTTPException(404, f"لا توجد مركبة بلوحة '{plate}'")

        # جيب السائق الحالي المرتبط بالمركبة
        c.execute("""SELECT p.id as perm_id, d.id as driver_id, d.name as driver_name,
                            d.fixed_number as driver_fixed
                     FROM driver_car_permissions p
                     JOIN drivers d ON d.id=p.driver_id
                     WHERE p.car_id=?
                     ORDER BY p.id DESC""", (car["id"],))
        current_perms = [dict(r) for r in c.fetchall()]

        if not confirm:
            # مرحلة الاستعراض فقط
            return {
                "car": dict(car),
                "new_driver": dict(new_driver),
                "current_permissions": current_perms,
                "message": "راجع البيانات ثم أرسل مع confirm=true للتنفيذ"
            }

        # مرحلة التنفيذ
        # 1. احذف كل الصلاحيات القديمة للمركبة (أي سائق كان معها)
        c.execute("DELETE FROM driver_car_permissions WHERE car_id=?", (car["id"],))
        deleted_car_perms = conn.execute("SELECT changes()").fetchone()[0]

        # 2. احذف كل صلاحيات السائق الجديد للعربيات التانية (عشان يفضل معاه العربية الجديدة بس)
        c.execute("DELETE FROM driver_car_permissions WHERE driver_id=?", (new_driver["id"],))
        deleted_driver_perms = conn.execute("SELECT changes()").fetchone()[0]

        deleted_count = deleted_car_perms + deleted_driver_perms

        # 3. أضف الصلاحية الجديدة
        c.execute("INSERT INTO driver_car_permissions(driver_id, car_id) VALUES(?,?)",
                  (new_driver["id"], car["id"]))

        log_event("permission_reassigned",
                  car_id=car["id"], plate=plate,
                  new_driver_id=new_driver["id"], new_driver_name=new_driver["name"],
                  deleted_old_perms=deleted_count,
                  superuser=cu["username"])

        return {
            "ok": True,
            "car": dict(car),
            "new_driver": dict(new_driver),
            "deleted_old_permissions": deleted_count,
            "message": f"✅ تم إعادة تخصيص '{plate}' للسائق '{new_driver['name']}' بنجاح"
        }

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
                             ORDER BY t.id DESC LIMIT 2000""", (branch,))
            else:
                c.execute("SELECT * FROM trips ORDER BY id DESC LIMIT 2000")
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
        MAX_TRIP_KM = 1500
        if (te.end_odometer - trip["start_odometer"]) > MAX_TRIP_KM:
            raise HTTPException(400, f"المسافة المدخلة ({round(te.end_odometer - trip['start_odometer'], 1)} كم) تتجاوز الحد الأقصى المسموح به ({MAX_TRIP_KM} كم للرحلة الواحدة) — يرجى التحقق من قراءة العداد")
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
        if (end_odo - active["start_odometer"]) > 1500:
            raise HTTPException(400, f"المسافة ({round(end_odo - active['start_odometer'], 1)} كم) تتجاوز الحد الأقصى 1500 كم — تحقق من القراءة")
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
            c.execute("INSERT INTO users(username,password,role,branch) VALUES(?,?,?,?)",
                      (user.username, _hash(user.password), user.role, user.branch or ""))
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
        c.execute("SELECT id,username,role,branch,last_login FROM users ORDER BY id DESC LIMIT 200")
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
    # ── احسب السعر الإجمالي بشكل موحّد لكل الأدوار ──
    # الأنواع اللي ليها كمية: الوقود بكل أنواعه + الزيت + الكوتش
    QUANTITY_TYPES = {t for t in WORKSHOP_TYPES if t.startswith("fuel_") or t in ("oil","tire")}
    HAS_QUANTITY   = rec.type in QUANTITY_TYPES

    with get_db() as conn:
        c = conn.cursor()
        # جيب فرع السائق/المركبة لتحديد السعر الصحيح
        c.execute("SELECT branch FROM drivers WHERE id=?", (rec.driver_id,))
        drv_row = c.fetchone()
        drv_branch = drv_row["branch"] if drv_row else ""

        if drv_branch:
            c.execute("SELECT value FROM app_settings WHERE key=?",
                      (f"price_{drv_branch}_{rec.type}",))
            srow = c.fetchone()
        else:
            srow = None
        if not srow:
            c.execute("SELECT value FROM app_settings WHERE key=?", (f"price_{rec.type}",))
            srow = c.fetchone()
        unit_price_cfg = float(srow["value"]) if srow else 0.0

    if cu["role"] in ("admin", "superuser"):
        # الأدمن يرسل price من الإعدادات (سعر/وحدة) — نضرب في الكمية
        unit_price = rec.price if rec.price and rec.price > 0 else unit_price_cfg
    else:
        # السائق: دايماً سعر الإعدادات
        unit_price = unit_price_cfg

    qty = rec.quantity if rec.quantity and rec.quantity > 0 else 1.0
    # الإجمالي = سعر الوحدة × الكمية (للأنواع ذات الكمية) أو سعر الوحدة مباشرة
    final_price = unit_price * qty if HAS_QUANTITY else unit_price
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
        # ── تحديث جدول الصيانة الدورية تلقائياً ──
        if rec.vehicle_id and rec.odometer_reading:
            MAINT_MAP = {
                "oil":    ["تغيير زيت","فلتر زيت"],
                "filter": ["فلتر هواء","فلتر وقود","فلتر مكيف"],
                "tire":   ["إطارات"],
                "battery":["بطارية"],
                "belt":   ["سير توزيع"],
            }
            related_types = MAINT_MAP.get(rec.type, [])
            if related_types:
                placeholders_m = ",".join("?" * len(related_types))
                c.execute(f"""UPDATE maintenance_schedule
                              SET last_done_km=?, last_done_date=?, updated_at=?
                              WHERE car_id=? AND maintenance_type IN ({placeholders_m})""",
                          [rec.odometer_reading, now[:10], now, rec.vehicle_id] + related_types)
                updated_m = conn.execute("SELECT changes()").fetchone()[0]
                if updated_m:
                    log_event("maintenance_auto_updated",
                              car_id=rec.vehicle_id, types=related_types, km=rec.odometer_reading)

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
async def get_prices(branch: Optional[str] = None, cu: dict = Depends(get_user)):
    """
    - superuser يمرر ?branch=xxx لأسعار فرع محدد
    - admin/reporter عنده فرع: يرجع أسعار فرعه تلقائياً
    - بدون فرع: الأسعار العامة
    """
    # السوبر يوزر يقدر يحدد فرع عبر query param
    eff_branch = branch if (cu["role"] == "superuser" and branch) else _branch_filter(cu)

    def _prices_for_branch(c, br):
        prices = {}
        for ws_type in WORKSHOP_TYPES:
            c.execute("SELECT value FROM app_settings WHERE key=?", (f"price_{br}_{ws_type}",))
            row = c.fetchone()
            if row:
                prices[ws_type] = float(row["value"])
            else:
                c.execute("SELECT value FROM app_settings WHERE key=?", (f"price_{ws_type}",))
                row = c.fetchone()
                prices[ws_type] = float(row["value"]) if row else 0.0
        return prices

    with get_db() as conn:
        c = conn.cursor()
        if eff_branch:
            return _prices_for_branch(c, eff_branch)
        else:
            c.execute("SELECT key,value FROM app_settings WHERE key LIKE 'price_%'")
            all_rows = {r["key"]: float(r["value"]) for r in c.fetchall()}
            # رجّع بس الأسعار العامة (مش الفروع) — key بدون underscore زيادة
            return {
                k.replace("price_",""): v
                for k, v in all_rows.items()
                if k.count('_') == 1  # price_oil, price_fuel_solar لكن مش price_branch_oil
                   or k.startswith('price_fuel_')  # fuel subtypes
            }

@app.put("/settings/prices")
async def set_prices(body: dict, cu: dict = Depends(require_admin)):
    """
    - admin عنده فرع: يحفظ الأسعار لفرعه (ماعدا الوقود — ده للسوبر فقط)
    - superuser أو بدون فرع: يحفظ كل الأسعار العامة
    - superuser ممكن يمرر { branch: 'cairo', prices: {...} } عشان يعدل فرع معين
    """
    branch = _branch_filter(cu)
    is_super = cu["role"] == "superuser"

    # superuser ممكن يحدد فرع معين في الـ body
    if not branch and is_super:
        branch = body.pop("branch", None) or None

    FUEL_TYPES_SET = {"fuel_solar", "fuel_92", "fuel_95", "fuel_80", "fuel_cng"}
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
            # admin لا يستطيع تعديل أسعار الوقود — فقط السوبر يوزر
            if ws_type in FUEL_TYPES_SET and not is_super:
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
    "license_renew":   "تجديد رخصة",
    "other":           "طلب آخر",
}

@app.post("/requests", status_code=201)
async def create_request(body: dict, cu: dict = Depends(get_user)):
    req_type       = (body.get("type") or "").strip()
    notes          = (body.get("notes") or "").strip()
    new_odometer   = body.get("new_odometer")
    trip_id        = body.get("trip_id")
    renewal_date   = (body.get("renewal_date") or "").strip()     # تاريخ التجديد الفعلي
    new_expiry_date= (body.get("new_expiry_date") or "").strip()  # تاريخ الانتهاء الجديد
    license_type   = (body.get("license_type") or "").strip()     # driver | vehicle

    if req_type not in REQUEST_TYPES:
        raise HTTPException(400, "نوع الطلب غير صالح")

    driver_id = cu.get("driver_id")
    if not driver_id:
        raise HTTPException(403, "فقط السائقون يمكنهم تقديم الطلبات")

    # حقق من صحة تواريخ التجديد
    if req_type == "license_renew":
        if not renewal_date:
            raise HTTPException(400, "تاريخ التجديد مطلوب")
        if not new_expiry_date:
            raise HTTPException(400, "تاريخ انتهاء الرخصة الجديد مطلوب")
        if not license_type:
            raise HTTPException(400, "نوع الرخصة مطلوب (driver أو vehicle)")
        # ضيف المعلومات في الملاحظات
        notes = f"نوع الرخصة: {'رخصة القيادة' if license_type=='driver' else 'رخصة العربة'} | تاريخ التجديد: {renewal_date} | تاريخ الانتهاء الجديد: {new_expiry_date}"

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
            # الأدمن يقدر يقرر مباشرة في الإجازات والأذونات وتجديد الرخصة
            # بس الـ odometer_change والأخرى بتتحال للسوبر
            direct_types = {"leave", "permission", "license_renew"}
            req_type = req.get("type", "") if hasattr(req, "get") else dict(req).get("type", "")

            status_direct = (body.get("status") or "").strip()
            if status_direct in ("approved", "rejected") and req_type in direct_types:
                # قرار مباشر من الأدمن
                admin_msg = (body.get("admin_message") or "").strip()

                # لو تجديد رخصة وتمت الموافقة — حدّث بيانات السائق
                if req_type == "license_renew" and status_direct == "approved":
                    notes_text = dict(req).get("notes", "")
                    # استخرج تاريخ الانتهاء الجديد من الملاحظات
                    import re
                    exp_match = re.search(r'تاريخ الانتهاء الجديد: (\d{4}-\d{2}-\d{2})', notes_text)
                    type_match = re.search(r'نوع الرخصة: (.+?) \|', notes_text)
                    if exp_match and type_match:
                        new_exp = exp_match.group(1)
                        lic_type_label = type_match.group(1)
                        driver_id_req = dict(req).get("driver_id")
                        if "القيادة" in lic_type_label:
                            c.execute("UPDATE drivers SET driver_license_expiry=? WHERE id=?",
                                      (new_exp, driver_id_req))
                        else:
                            c.execute("UPDATE drivers SET vehicle_license_expiry=? WHERE id=?",
                                      (new_exp, driver_id_req))

                c.execute("""UPDATE driver_requests
                             SET status=?, admin_notes=?, admin_message=?,
                                 handled_by=?, handled_at=?
                             WHERE id=?""",
                          (status_direct, admin_notes, admin_msg, cu["username"], now, rid))
                log_event("request_decided_by_admin", request_id=rid,
                          status=status_direct, admin=cu["username"], type=req_type)
                return {"ok": True, "status": status_direct, "direct": True}

            # الإحالة للسوبر (لأنواع أخرى أو لو الأدمن اختار الإحالة)
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
# 27. MAINTENANCE SCHEDULE (جدول الصيانة الدورية)
# ══════════════════════════════════════════════════════

MAINTENANCE_TYPES = [
    "تغيير زيت", "فلتر زيت", "فلتر هواء", "فلتر وقود",
    "فلتر مكيف", "إطارات", "بطارية", "فرامل أمامية",
    "فرامل خلفية", "سير توزيع", "ترمس", "ماء تبريد",
    "سائل فرامل", "فحص دوري عام", "أخرى",
]

class MaintenanceScheduleCreate(BaseModel):
    car_id:           int
    maintenance_type: str
    interval_km:      Optional[float] = None   # كل كام كم
    interval_days:    Optional[int]   = None   # أو كل كام يوم
    last_done_km:     Optional[float] = None   # آخر عداد عند الصيانة
    last_done_date:   Optional[str]   = None   # تاريخ آخر صيانة (YYYY-MM-DD)
    alert_km_before:  Optional[float] = 500.0  # إنذار قبل كام كم
    alert_days_before:Optional[int]   = 7      # إنذار قبل كام يوم
    notes:            Optional[str]   = ""

class MaintenanceScheduleUpdate(BaseModel):
    maintenance_type: Optional[str]   = None
    interval_km:      Optional[float] = None
    interval_days:    Optional[int]   = None
    last_done_km:     Optional[float] = None
    last_done_date:   Optional[str]   = None
    alert_km_before:  Optional[float] = None
    alert_days_before:Optional[int]   = None
    notes:            Optional[str]   = None

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
            # ✅ تصفية الرحلات الشاذة: المسافة الواحدة يجب أن تكون أقل من 1500 كم
            "(t.end_odometer - t.start_odometer) < 1500",
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
            "(t.end_odometer - t.start_odometer) < 1500",
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
    threshold: float = 1500.0,          # كم — أي رحلة أكبر من كده تُعتبر شاذة
    cu: dict = Depends(require_admin_or_reporter)
):
    """
    يُرجع الرحلات التي تجاوزت حدّ المسافة المنطقية (default 1500 كم للرحلة الواحدة).
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
# 30. PERMISSIONS BULK IMPORT — CSV / Excel
#     الأعمدة: fixed_number + plate أو car_code
# ══════════════════════════════════════════════════════

def _validate_permission_row(row: dict, idx: int, drivers_map: dict,
                              cars_map_plate: dict, cars_map_code: dict) -> dict:
    errors = []
    fixed = str(row.get("fixed_number") or row.get("رقم ثابت") or "").strip()
    plate = str(row.get("plate") or row.get("رقم اللوحة") or "").strip()
    code  = str(row.get("car_code") or row.get("كود المركبة") or "").strip()

    driver_id = drivers_map.get(fixed)
    if not fixed:
        errors.append("الرقم الثابت مطلوب")
    elif driver_id is None:
        errors.append(f"لا يوجد سائق بالرقم الثابت '{fixed}'")

    car_id = cars_map_plate.get(plate) or cars_map_code.get(code)
    if not plate and not code:
        errors.append("رقم اللوحة أو كود المركبة مطلوب")
    elif car_id is None:
        errors.append(f"لا توجد مركبة بـ '{plate or code}'")

    return {
        "row_index":    idx,
        "fixed_number": fixed,
        "plate":        plate or code,
        "driver_id":    driver_id,
        "car_id":       car_id,
        "valid":        len(errors) == 0,
        "errors":       errors,
    }


def _build_perm_maps(c):
    c.execute("SELECT id, fixed_number FROM drivers WHERE fixed_number != ''")
    drivers_map = {r["fixed_number"]: r["id"] for r in c.fetchall()}
    c.execute("SELECT id, plate, car_code FROM cars")
    cars_map_plate, cars_map_code = {}, {}
    for r in c.fetchall():
        if r["plate"]:    cars_map_plate[r["plate"]]   = r["id"]
        if r["car_code"]: cars_map_code[r["car_code"]] = r["id"]
    c.execute("SELECT driver_id, car_id FROM driver_car_permissions")
    existing = {(r["driver_id"], r["car_id"]) for r in c.fetchall()}
    return drivers_map, cars_map_plate, cars_map_code, existing


@app.post("/import/permissions/preview")
async def import_permissions_preview(
    file: UploadFile = File(...),
    cu: dict = Depends(require_admin)
):
    content = await file.read()
    if len(content) > 5 * 1024 * 1024:
        raise HTTPException(400, "حجم الملف يتجاوز 5 MB")
    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows:
        raise HTTPException(400, "الملف فارغ أو لا يحتوي على بيانات")

    with get_db() as conn:
        dm, cmp, cmc, existing = _build_perm_maps(conn.cursor())

    rows = [_validate_permission_row(r, i+1, dm, cmp, cmc) for i, r in enumerate(raw_rows)]
    for r in rows:
        if r["valid"] and (r["driver_id"], r["car_id"]) in existing:
            r["valid"] = False; r["errors"] = ["الصلاحية موجودة مسبقاً"]

    valid = sum(1 for r in rows if r["valid"])
    return {
        "import_type": "permissions",
        "rows": rows,
        "summary": {"total": len(rows), "valid": valid, "invalid": len(rows) - valid},
    }


@app.post("/import/permissions/confirm")
async def import_permissions_confirm(
    file: UploadFile = File(...),
    skip_errors: bool = False,
    cu: dict = Depends(require_admin)
):
    content = await file.read()
    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows:
        raise HTTPException(400, "الملف فارغ")

    with get_db() as conn:
        dm, cmp, cmc, existing = _build_perm_maps(conn.cursor())

    rows = [_validate_permission_row(r, i+1, dm, cmp, cmc) for i, r in enumerate(raw_rows)]
    for r in rows:
        if r["valid"] and (r["driver_id"], r["car_id"]) in existing:
            r["valid"] = False; r["errors"] = ["الصلاحية موجودة مسبقاً"]

    invalid = [r for r in rows if not r["valid"]]
    if invalid and not skip_errors:
        raise HTTPException(422, {"message": f"يوجد {len(invalid)} صف بأخطاء", "invalid_rows": invalid})

    inserted, skipped_rows = [], []
    with get_db() as conn:
        c = conn.cursor()
        for row in [r for r in rows if r["valid"]]:
            try:
                c.execute("INSERT INTO driver_car_permissions(driver_id,car_id) VALUES(?,?)",
                          (row["driver_id"], row["car_id"]))
                inserted.append({"row_index": row["row_index"],
                                  "fixed_number": row["fixed_number"], "plate": row["plate"]})
            except Exception as e:
                row["errors"] = [str(e)]; skipped_rows.append(row)

    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "bulk_import_permissions", f"استيراد {len(inserted)} صلاحية من ملف")
    return {
        "import_type": "permissions", "inserted": len(inserted), "skipped": len(skipped_rows),
        "inserted_items": inserted, "skipped_items": skipped_rows,
    }


@app.get("/import/template/permissions")
async def import_permissions_template(cu: dict = Depends(require_admin)):
    cols   = ["رقم ثابت", "رقم اللوحة", "كود المركبة"]
    sample = ["123456", "أ-ب-1234", ""]
    output = io.StringIO()
    csv.writer(output).writerows([cols, sample])
    from fastapi.responses import Response
    return Response(content=output.getvalue().encode("utf-8-sig"), media_type="text/csv",
                    headers={"Content-Disposition": "attachment; filename=permissions_template.csv"})


# ══════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════
# 31. LAST ODOMETER — آخر عداد لمركبة (للسائق عند بدء الرحلة)
# ══════════════════════════════════════════════════════
@app.get("/cars/{car_id}/last-odometer")
async def get_last_odometer(car_id: int, cu: dict = Depends(get_user)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT end_odometer FROM trips
                     WHERE car_id=? AND end_odometer IS NOT NULL AND end_odometer > 0
                     ORDER BY end_time DESC LIMIT 1""", (car_id,))
        row = c.fetchone()
    return {"car_id": car_id, "last_odometer": row["end_odometer"] if row else None}


# ══════════════════════════════════════════════════════
# 32. FRAUD DETECTION — كشف التلاعب
# ══════════════════════════════════════════════════════
@app.get("/reports/fraud-detection")
async def fraud_detection(cu: dict = Depends(require_superuser)):
    """تحليل شامل لكشف التلاعب المحتمل في البيانات."""
    with get_db() as conn:
        c = conn.cursor()

        # 1. عداد النهاية أقل من عداد البداية (تلاعب صريح)
        c.execute("""SELECT t.id, t.start_odometer, t.end_odometer,
                            t.start_time, t.end_time,
                            d.name driver_name, d.fixed_number,
                            ca.plate
                     FROM trips t
                     JOIN drivers d ON d.id=t.driver_id
                     JOIN cars ca ON ca.id=t.car_id
                     WHERE t.end_odometer IS NOT NULL
                       AND t.start_odometer IS NOT NULL
                       AND t.end_odometer < t.start_odometer""")
        odometer_reversed = [dict(r) for r in c.fetchall()]

        # 2. رحلات بمسافة مشبوهة (أكثر من 500 كم في رحلة واحدة)
        c.execute("""SELECT t.id, t.start_odometer, t.end_odometer,
                            (t.end_odometer - t.start_odometer) AS distance,
                            t.start_time, t.end_time,
                            d.name driver_name, d.fixed_number, ca.plate
                     FROM trips t
                     JOIN drivers d ON d.id=t.driver_id
                     JOIN cars ca ON ca.id=t.car_id
                     WHERE t.end_odometer IS NOT NULL AND t.start_odometer IS NOT NULL
                       AND (t.end_odometer - t.start_odometer) > 500
                     ORDER BY distance DESC""")
        high_distance = [dict(r) for r in c.fetchall()]

        # 3. رحلتان متداخلتان لنفس السائق
        c.execute("""SELECT a.id id1, b.id id2,
                            a.start_time, a.end_time,
                            b.start_time start2, b.end_time end2,
                            d.name driver_name, d.fixed_number
                     FROM trips a JOIN trips b ON a.driver_id=b.driver_id AND a.id<b.id
                     JOIN drivers d ON d.id=a.driver_id
                     WHERE a.end_time IS NOT NULL AND b.end_time IS NOT NULL
                       AND a.start_time < b.end_time AND a.end_time > b.start_time""")
        overlapping = [dict(r) for r in c.fetchall()]

        # 4. رحلتان متداخلتان لنفس المركبة (مع سائقين مختلفين)
        c.execute("""SELECT a.id id1, b.id id2, ca.plate,
                            da.name driver1, db.name driver2,
                            a.start_time, a.end_time,
                            b.start_time start2, b.end_time end2
                     FROM trips a JOIN trips b ON a.car_id=b.car_id AND a.id<b.id
                     JOIN cars ca ON ca.id=a.car_id
                     JOIN drivers da ON da.id=a.driver_id
                     JOIN drivers db ON db.id=b.driver_id
                     WHERE a.end_time IS NOT NULL AND b.end_time IS NOT NULL
                       AND a.start_time < b.end_time AND a.end_time > b.start_time
                       AND a.driver_id != b.driver_id""")
        car_overlap = [dict(r) for r in c.fetchall()]

        # 5. رحلة وقتها قصير جداً (<3 دقائق) لكن مسافة كبيرة (>10 كم)
        c.execute("""SELECT t.id,
                            (julianday(t.end_time)-julianday(t.start_time))*1440 AS minutes,
                            (t.end_odometer - t.start_odometer) AS distance,
                            d.name driver_name, ca.plate,
                            t.start_time, t.end_time
                     FROM trips t
                     JOIN drivers d ON d.id=t.driver_id
                     JOIN cars ca ON ca.id=t.car_id
                     WHERE t.end_time IS NOT NULL AND t.end_odometer IS NOT NULL
                       AND t.start_odometer IS NOT NULL
                       AND (julianday(t.end_time)-julianday(t.start_time))*1440 < 3
                       AND (t.end_odometer - t.start_odometer) > 10""")
        impossible_speed = [dict(r) for r in c.fetchall()]

        # 6. وقود كثير في فترة قصيرة (أكثر من 200 لتر في يوم واحد لنفس المركبة)
        c.execute("""SELECT vehicle_id, DATE(created_at) AS day,
                            SUM(quantity) AS total_liters, COUNT(*) AS refills,
                            ca.plate
                     FROM workshop_records w
                     JOIN cars ca ON ca.id=w.vehicle_id
                     WHERE type LIKE 'fuel_%' AND quantity IS NOT NULL
                     GROUP BY vehicle_id, day
                     HAVING total_liters > 200
                     ORDER BY total_liters DESC""")
        excess_fuel = [dict(r) for r in c.fetchall()]

        # 7. سائق عدّل طلب عداد أكثر من 3 مرات في أسبوع
        c.execute("""SELECT driver_id, COUNT(*) AS requests,
                            d.name driver_name, d.fixed_number,
                            MIN(created_at) AS first_req
                     FROM driver_requests
                     JOIN drivers d ON d.id=driver_requests.driver_id
                     WHERE type='odometer_change'
                       AND created_at >= datetime('now','-7 days')
                     GROUP BY driver_id HAVING requests > 3
                     ORDER BY requests DESC""")
        repeated_odometer = [dict(r) for r in c.fetchall()]

    total_alerts = (len(odometer_reversed) + len(overlapping) + len(car_overlap) +
                    len(impossible_speed) + len(excess_fuel) + len(repeated_odometer))
    high_risk   = len(odometer_reversed) + len(impossible_speed) + len(car_overlap)

    return {
        "summary": {
            "total_alerts": total_alerts,
            "high_risk":    high_risk,
            "medium_risk":  len(overlapping) + len(excess_fuel),
            "low_risk":     len(repeated_odometer) + len(high_distance),
        },
        "odometer_reversed":  odometer_reversed,
        "high_distance":      high_distance,
        "overlapping_driver": overlapping,
        "car_overlap":        car_overlap,
        "impossible_speed":   impossible_speed,
        "excess_fuel":        excess_fuel,
        "repeated_odometer":  repeated_odometer,
    }


# ══════════════════════════════════════════════════════
# 33. VOICE NOTES — رسائل صوتية من السوبر
# ══════════════════════════════════════════════════════
class VoiceNoteCreate(BaseModel):
    target_group: str          # 'admins' | 'drivers'
    audio_data:   str          # base64
    duration_sec: float = 0
    max_plays:    int   = 2    # مرتين افتراضياً
    hours_ttl:    int   = 24   # يتمسح بعد 24 ساعة

@app.post("/voice-notes")
async def create_voice_note(body: VoiceNoteCreate, cu: dict = Depends(require_superuser)):
    if body.target_group not in ("admins", "drivers"):
        raise HTTPException(400, "target_group يجب أن يكون admins أو drivers")
    now      = datetime.utcnow()
    expires  = (now + timedelta(hours=body.hours_ttl)).isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""INSERT INTO voice_notes(sender_id,target_group,audio_data,
                     duration_sec,created_at,expires_at,max_plays)
                     VALUES(?,?,?,?,?,?,?)""",
                  (cu["user_id"], body.target_group, body.audio_data,
                   body.duration_sec, now.isoformat()+"Z", expires, body.max_plays))
        note_id = c.lastrowid
    return {"id": note_id, "expires_at": expires}

@app.get("/voice-notes")
async def list_voice_notes(cu: dict = Depends(get_user)):
    role   = cu["role"]
    now    = datetime.utcnow().isoformat() + "Z"

    # السوبر يشوف اللي أرسلها هو
    if role == "superuser":
        with get_db() as conn:
            c = conn.cursor()
            c.execute("""SELECT id, target_group, duration_sec, created_at,
                                expires_at, play_count, max_plays, is_deleted
                         FROM voice_notes WHERE sender_id=?
                         ORDER BY created_at DESC""", (cu["user_id"],))
            return [dict(r) for r in c.fetchall()]

    # أدمن → يشوف اللي للـ admins
    if role in ("admin","reporter","superuser","supervisor_workshop","supervisor_field"):
        group = "admins"
    elif role == "driver":
        group = "drivers"
    else:
        return []

    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT id, duration_sec, created_at, expires_at,
                            play_count, max_plays, audio_data
                     FROM voice_notes
                     WHERE target_group=? AND is_deleted=0
                       AND expires_at > ? AND play_count < max_plays
                     ORDER BY created_at DESC""", (group, now))
        return [dict(r) for r in c.fetchall()]

@app.post("/voice-notes/{note_id}/play")
async def play_voice_note(note_id: int, cu: dict = Depends(get_user)):
    """تسجّل تشغيل — وتمسح لو وصلت الحد."""
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT play_count, max_plays, is_deleted, expires_at FROM voice_notes WHERE id=?", (note_id,))
        row = c.fetchone()
        if not row:                    raise HTTPException(404, "الرسالة غير موجودة")
        if row["is_deleted"]:          raise HTTPException(410, "الرسالة تم حذفها")
        if row["expires_at"] < now:    raise HTTPException(410, "انتهت صلاحية الرسالة")
        new_count = row["play_count"] + 1
        if new_count >= row["max_plays"]:
            c.execute("UPDATE voice_notes SET play_count=?, is_deleted=1 WHERE id=?", (new_count, note_id))
        else:
            c.execute("UPDATE voice_notes SET play_count=? WHERE id=?", (new_count, note_id))
    return {"play_count": new_count, "deleted": new_count >= row["max_plays"]}

@app.delete("/voice-notes/{note_id}")
async def delete_voice_note(note_id: int, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        conn.cursor().execute("UPDATE voice_notes SET is_deleted=1 WHERE id=? AND sender_id=?",
                              (note_id, cu["user_id"]))
    return {"deleted": True}

@app.delete("/voice-notes/cleanup/expired")
async def cleanup_expired_notes(cu: dict = Depends(require_superuser)):
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        c.execute("DELETE FROM voice_notes WHERE expires_at < ? OR is_deleted=1", (now,))
        return {"deleted": c.rowcount}


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
# ══════════════════════════════════════════════════════
# 27. MAINTENANCE SCHEDULE — CRUD + ALERTS
# ══════════════════════════════════════════════════════

def _maintenance_row(row: dict, cars_map: dict) -> dict:
    """Enrich a maintenance_schedule row with computed fields."""
    d = dict(row)
    d.setdefault("notes", "")
    car = cars_map.get(d.get("car_id"), {})
    d["car_plate"]  = car.get("plate", "—")
    d["car_model"]  = car.get("model", "—")
    d["car_branch"] = car.get("branch", "")

    # Last known odometer from trips
    d["current_km"] = d.get("_current_km", None)

    # Compute next due
    next_km = None
    if d.get("last_done_km") is not None and d.get("interval_km"):
        next_km = float(d["last_done_km"]) + float(d["interval_km"])
    d["next_due_km"] = next_km

    next_date = None
    if d.get("last_done_date") and d.get("interval_days"):
        from datetime import timedelta
        try:
            ld = datetime.fromisoformat(d["last_done_date"])
            next_date = (ld + timedelta(days=int(d["interval_days"]))).strftime("%Y-%m-%d")
        except Exception:
            pass
    d["next_due_date"] = next_date

    # Alert status
    today = datetime.utcnow().strftime("%Y-%m-%d")
    alert_km_before   = float(d.get("alert_km_before") or 500)
    alert_days_before = int(d.get("alert_days_before") or 7)
    cur_km = d.get("current_km")

    status = "ok"
    if next_km is not None and cur_km is not None:
        remaining_km = next_km - float(cur_km)
        if remaining_km <= 0:
            status = "overdue"
        elif remaining_km <= alert_km_before:
            status = "warning"

    if next_date and status == "ok":
        from datetime import timedelta
        try:
            nd = datetime.fromisoformat(next_date)
            days_left = (nd - datetime.utcnow()).days
            if days_left <= 0:
                status = "overdue"
            elif days_left <= alert_days_before:
                status = "warning"
        except Exception:
            pass

    d["alert_status"] = status
    d.pop("_current_km", None)
    return d


@app.get("/maintenance/schedule")
async def get_maintenance_schedule(cu: dict = Depends(require_admin_or_reporter)):
    branch = _branch_filter(cu)
    with get_db() as conn:
        c = conn.cursor()
        # جيب أحدث عداد لكل عربية من الرحلات المنتهية
        c.execute("""SELECT car_id, MAX(end_odometer) as max_odo
                     FROM trips WHERE end_odometer IS NOT NULL AND end_time IS NOT NULL
                     GROUP BY car_id""")
        odo_map = {r["car_id"]: r["max_odo"] for r in c.fetchall()}

        # جيب العربيات
        if branch:
            c.execute("SELECT id,plate,model,branch FROM cars WHERE branch=?", (branch,))
        else:
            c.execute("SELECT id,plate,model,branch FROM cars")
        cars_map = {r["id"]: dict(r) for r in c.fetchall()}

        # جيب الجدول
        car_ids = list(cars_map.keys())
        if not car_ids:
            return []
        placeholders = ",".join("?" * len(car_ids))
        c.execute(f"""SELECT * FROM maintenance_schedule
                      WHERE car_id IN ({placeholders})
                      ORDER BY car_id, maintenance_type""", car_ids)
        rows = []
        for r in c.fetchall():
            d = dict(r)
            d["_current_km"] = odo_map.get(d["car_id"])
            rows.append(_maintenance_row(d, cars_map))
        return rows


@app.get("/maintenance/alerts")
async def get_maintenance_alerts(cu: dict = Depends(require_admin_or_reporter)):
    """يرجع فقط السجلات اللي status فيها warning أو overdue."""
    all_rows = await get_maintenance_schedule(cu)
    return [r for r in all_rows if r["alert_status"] in ("warning", "overdue")]


@app.post("/maintenance/schedule", status_code=201)
async def create_maintenance(rec: MaintenanceScheduleCreate, cu: dict = Depends(require_admin)):
    if rec.maintenance_type not in MAINTENANCE_TYPES:
        raise HTTPException(400, f"نوع صيانة غير صالح. الأنواع المتاحة: {MAINTENANCE_TYPES}")
    if not rec.interval_km and not rec.interval_days:
        raise HTTPException(400, "يجب تحديد interval_km أو interval_days على الأقل")
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id FROM cars WHERE id=?", (rec.car_id,))
        if not c.fetchone():
            raise HTTPException(404, "المركبة غير موجودة")
        # منع تكرار نفس النوع لنفس العربية
        c.execute("SELECT id FROM maintenance_schedule WHERE car_id=? AND maintenance_type=?",
                  (rec.car_id, rec.maintenance_type))
        if c.fetchone():
            raise HTTPException(400, "يوجد جدول صيانة لهذا النوع لهذه المركبة بالفعل")
        c.execute("""INSERT INTO maintenance_schedule
                     (car_id,maintenance_type,interval_km,interval_days,
                      last_done_km,last_done_date,alert_km_before,alert_days_before,notes,created_at,updated_at)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                  (rec.car_id, rec.maintenance_type, rec.interval_km, rec.interval_days,
                   rec.last_done_km, rec.last_done_date,
                   rec.alert_km_before or 500, rec.alert_days_before or 7,
                   rec.notes or "", now, now))
        rid = c.lastrowid
        log_event("maintenance_created", id=rid, car_id=rec.car_id, type=rec.maintenance_type)
        return {"id": rid, "ok": True}


@app.put("/maintenance/schedule/{mid}")
async def update_maintenance(mid: int, rec: MaintenanceScheduleUpdate, cu: dict = Depends(require_admin)):
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT * FROM maintenance_schedule WHERE id=?", (mid,))
        row = c.fetchone()
        if not row:
            raise HTTPException(404, "السجل غير موجود")
        updates = {k: v for k, v in rec.dict(exclude_none=True).items()}
        if not updates:
            raise HTTPException(400, "لا توجد بيانات للتحديث")
        updates["updated_at"] = now
        set_clause = ", ".join(f"{k}=?" for k in updates)
        c.execute(f"UPDATE maintenance_schedule SET {set_clause} WHERE id=?",
                  list(updates.values()) + [mid])
        log_event("maintenance_updated", id=mid, changes=list(updates.keys()))
        return {"ok": True}


@app.delete("/maintenance/schedule/{mid}")
async def delete_maintenance(mid: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("DELETE FROM maintenance_schedule WHERE id=?", (mid,))
        if conn.execute("SELECT changes()").fetchone()[0] == 0:
            raise HTTPException(404, "السجل غير موجود")
        log_event("maintenance_deleted", id=mid)
        return {"ok": True}


# ══════════════════════════════════════════════════════
# 29. FUEL EFFICIENCY REPORT (تقرير كفاءة الوقود)
#     — للسوبر يوزر فقط —
#     يستخدم منهجية Fill-to-Fill الصحيحة:
#       • التفويلة الأولى = نقطة مرجعية فقط، لا تُحسب
#       • الكفاءة بين كل تفويلتين = km_diff / liters_of_new_fill
#       • يجب وجود ≥ 2 تفويلات حتى تظهر النتيجة
# ══════════════════════════════════════════════════════

def _fill_to_fill_efficiency(fills: list) -> dict:
    """
    fills: list of dicts مرتبة تصاعدياً بـ odometer_reading
           كل dict: {odometer_reading, quantity, price, created_at}

    يرجع:
        total_km      : مجموع المسافات بين التفويلات
        total_liters  : مجموع اللترات (من التفويلة الثانية فصاعداً)
        total_cost    : مجموع التكلفة  (من التفويلة الثانية فصاعداً)
        refills       : عدد التفويلات اللي دخلت في الحساب (الأولى مش محسوبة)
        efficiency    : كم/لتر  (None لو أقل من 2 تفويلات أو لترات=0)
        cost_per_km   : جنيه/كم (None لو لا يوجد)
        segments      : تفاصيل كل فترة بين تفويلتين
    """
    if len(fills) < 2:
        return {
            "total_km": 0, "total_liters": 0, "total_cost": 0,
            "refills": 0, "efficiency": None, "cost_per_km": None,
            "segments": [], "note": "بيانات غير كافية — تحتاج ≥ 2 تفويلات"
        }

    total_km     = 0.0
    total_liters = 0.0
    total_cost   = 0.0
    segments     = []

    for i in range(1, len(fills)):
        prev = fills[i - 1]
        curr = fills[i]
        odo_prev = prev.get("odometer_reading") or 0
        odo_curr = curr.get("odometer_reading") or 0
        km_diff  = max(odo_curr - odo_prev, 0)
        liters   = float(curr.get("quantity") or 0)
        cost     = float(curr.get("price")    or 0)

        if km_diff > 0 and liters > 0:
            seg_eff = round(km_diff / liters, 2)
        else:
            seg_eff = None

        total_km     += km_diff
        total_liters += liters
        total_cost   += cost

        segments.append({
            "from_odometer": odo_prev,
            "to_odometer":   odo_curr,
            "km":            round(km_diff, 1),
            "liters":        round(liters, 2),
            "cost":          round(cost, 2),
            "efficiency_km_per_liter": seg_eff,
            "date": curr.get("created_at", "")[:10],
        })

    efficiency  = round(total_km / total_liters, 2) if total_liters > 0 else None
    cost_per_km = round(total_cost / total_km, 2)   if total_km    > 0 else None

    return {
        "total_km":     round(total_km, 1),
        "total_liters": round(total_liters, 2),
        "total_cost":   round(total_cost, 2),
        "refills":      len(fills) - 1,   # الأولى reference فقط
        "efficiency":   efficiency,
        "cost_per_km":  cost_per_km,
        "segments":     segments,
    }


@app.get("/reports/fuel-efficiency")
async def fuel_efficiency_report(
    date_from: Optional[str] = None,   # YYYY-MM-DD
    date_to:   Optional[str] = None,   # YYYY-MM-DD
    car_id:    Optional[int] = None,
    cu: dict = Depends(require_superuser)   # ← سوبر يوزر فقط
):
    """
    تقرير كفاءة الوقود — Fill-to-Fill Method
    ─────────────────────────────────────────
    المنهجية:
      1. نجيب كل سجلات الوقود للعربية مرتبة بعداد المسافة
      2. نتجاهل التفويلة الأولى (reference point)
      3. من التفويلة الثانية: كفاءة_القطعة = km_diff / liters_هذه_التفويلة
      4. الكفاءة الإجمالية = مجموع_الكم / مجموع_اللترات (من الثانية فصاعداً)

    ملاحظة: العربية اللي عندها تفويلة واحدة فقط لن تظهر لها كفاءة.
    """
    today = datetime.utcnow().strftime("%Y-%m-%d")
    if not date_from:
        date_from = (datetime.utcnow() - timedelta(days=30)).strftime("%Y-%m-%d")
    if not date_to:
        date_to = today

    fuel_types = [t for t in WORKSHOP_TYPES if t.startswith("fuel_")]
    fuel_types_ph = ",".join("?" * len(fuel_types))

    with get_db() as conn:
        c = conn.cursor()

        # ── العربيات ──
        if car_id:
            c.execute("SELECT id,plate,model,car_name,branch,equipment_type FROM cars WHERE id=?", (car_id,))
        else:
            c.execute("SELECT id,plate,model,car_name,branch,equipment_type FROM cars ORDER BY plate")
        cars_list = [dict(r) for r in c.fetchall()]
        cars_map  = {car["id"]: car for car in cars_list}

        if not cars_map:
            return {"date_from": date_from, "date_to": date_to, "cars": []}

        car_ids      = list(cars_map.keys())
        placeholders = ",".join("?" * len(car_ids))

        # ── جيب كل تفويلات الوقود للعربيات في الفترة
        #    مرتبة بـ odometer_reading تصاعدياً عشان fill-to-fill يشتغل صح
        #    لو مفيش عداد نرتب بـ created_at
        c.execute(f"""
            SELECT vehicle_id, odometer_reading, quantity, price, created_at
            FROM   workshop_records
            WHERE  vehicle_id IN ({placeholders})
              AND  type IN ({fuel_types_ph})
              AND  created_at >= ?
              AND  created_at <= ?
              AND  quantity IS NOT NULL AND quantity > 0
            ORDER BY vehicle_id,
                     COALESCE(odometer_reading, 0),
                     created_at
        """, car_ids + fuel_types + [date_from + "T00:00:00", date_to + "T23:59:59"])

        # نجمع التفويلات لكل عربية
        fills_by_car: dict = {}
        for row in c.fetchall():
            vid = row["vehicle_id"]
            fills_by_car.setdefault(vid, []).append(dict(row))

    result = []
    for cid, car in cars_map.items():
        fills = fills_by_car.get(cid, [])
        stats = _fill_to_fill_efficiency(fills)
        eq    = (car.get("equipment_type") or "").split("|")

        result.append({
            "car_id":       cid,
            "plate":        car["plate"],
            "model":        car["model"],
            "car_name":     car.get("car_name", ""),
            "branch":       car.get("branch", ""),
            "category":     eq[0] if eq else "",
            "total_km":     stats["total_km"],
            "total_liters": stats["total_liters"],
            "total_cost":   stats["total_cost"],
            "refills":      stats["refills"],
            "efficiency_km_per_liter": stats["efficiency"],
            "cost_per_km":  stats["cost_per_km"],
            "segments":     stats["segments"],
            "note":         stats.get("note", ""),
        })

    # رتّب: الأعلى كفاءة أولاً، العربيات بدون بيانات في الآخر
    result.sort(key=lambda x: (
        x["efficiency_km_per_liter"] is None,
        -(x["efficiency_km_per_liter"] or 0)
    ))
    return {
        "date_from":  date_from,
        "date_to":    date_to,
        "method":     "fill_to_fill",
        "note":       "التفويلة الأولى لكل عربية تُعتبر نقطة مرجعية فقط ولا تدخل في حساب الكفاءة",
        "cars":       result,
    }

# ══════════════════════════════════════════════════════
# 30. MONTHLY REPORT — تقرير شهري شامل
# ══════════════════════════════════════════════════════

@app.get("/reports/monthly")
async def monthly_report(
    month:  Optional[str] = None,   # YYYY-MM
    branch: Optional[str] = None,   # فرع محدد أو None للكل
    cu: dict = Depends(require_admin_or_reporter)
):
    if not month:
        month = datetime.utcnow().strftime("%Y-%m")
    try:
        datetime.strptime(month, "%Y-%m")
    except ValueError:
        raise HTTPException(400, "صيغة الشهر: YYYY-MM")

    # الأدمن المقيّد بفرع لا يقدر يشوف فروع تانية
    eff_branch = _effective_branch(cu, branch)
    like = month + "%"

    with get_db() as conn:
        c = conn.cursor()

        # ── السائقون ──
        if eff_branch:
            c.execute("SELECT id,name,fixed_number,branch FROM drivers WHERE branch=?", (eff_branch,))
        else:
            c.execute("SELECT id,name,fixed_number,branch FROM drivers")
        drivers_list = [dict(r) for r in c.fetchall()]
        driver_ids = [d["id"] for d in drivers_list]

        # ── المركبات ──
        if eff_branch:
            c.execute("SELECT id,plate,model,car_name,branch FROM cars WHERE branch=?", (eff_branch,))
        else:
            c.execute("SELECT id,plate,model,car_name,branch FROM cars")
        cars_list = [dict(r) for r in c.fetchall()]
        car_ids = [c2["id"] for c2 in cars_list]

        if not driver_ids:
            return {"month": month, "branch": eff_branch, "drivers": [], "cars": [],
                    "summary": {}, "workshops": []}

        drv_ph  = ",".join("?" * len(driver_ids))
        car_ph  = ",".join("?" * len(car_ids)) if car_ids else "NULL"

        # ── الرحلات ──
        c.execute(f"""SELECT driver_id, car_id,
                             COUNT(*) as trip_count,
                             SUM(CASE WHEN end_time IS NOT NULL THEN 1 ELSE 0 END) as completed,
                             SUM(CASE WHEN end_odometer > start_odometer
                                      THEN end_odometer - start_odometer ELSE 0 END) as total_km
                      FROM trips
                      WHERE driver_id IN ({drv_ph}) AND start_time LIKE ?
                      GROUP BY driver_id""", driver_ids + [like])
        trip_stats = {r["driver_id"]: dict(r) for r in c.fetchall()}

        # ── الورش ──
        if car_ids:
            c.execute(f"""SELECT w.vehicle_id, w.type,
                                 SUM(w.price)    as total_cost,
                                 SUM(w.quantity) as total_qty,
                                 COUNT(*)        as count
                          FROM workshop_records w
                          WHERE w.vehicle_id IN ({car_ph}) AND w.created_at LIKE ?
                          GROUP BY w.vehicle_id, w.type""", car_ids + [like])
            ws_rows = [dict(r) for r in c.fetchall()]
        else:
            ws_rows = []

        # ── تجميع الورش لكل عربية ──
        car_ws_map: dict = {}
        for w in ws_rows:
            vid = w["vehicle_id"]
            if vid not in car_ws_map:
                car_ws_map[vid] = []
            car_ws_map[vid].append(w)

        # ── إجمالي تكاليف الوقود ──
        total_fuel_cost  = sum(w["total_cost"] for w in ws_rows if w["type"].startswith("fuel_"))
        total_fuel_liters= sum(w["total_qty"]  for w in ws_rows if w["type"].startswith("fuel_"))
        total_maint_cost = sum(w["total_cost"] for w in ws_rows if not w["type"].startswith("fuel_"))
        total_km_all     = sum(ts["total_km"] for ts in trip_stats.values())
        total_trips_all  = sum(ts["trip_count"] for ts in trip_stats.values())

        # ── بيانات السائقين ──
        drivers_report = []
        for d in drivers_list:
            ts = trip_stats.get(d["id"], {})
            drivers_report.append({
                "driver_id":   d["id"],
                "name":        d["name"],
                "fixed_number":d.get("fixed_number",""),
                "branch":      d.get("branch",""),
                "trips":       ts.get("trip_count",0),
                "completed":   ts.get("completed",0),
                "km":          round(ts.get("total_km",0),1),
            })
        drivers_report.sort(key=lambda x: -x["km"])

        # ── بيانات العربيات ──
        cars_report = []
        for car in cars_list:
            ws = car_ws_map.get(car["id"],[])
            fuel_l = sum(w["total_qty"] or 0 for w in ws if w["type"].startswith("fuel_"))
            fuel_c = sum(w["total_cost"] or 0 for w in ws if w["type"].startswith("fuel_"))
            maint_c= sum(w["total_cost"] or 0 for w in ws if not w["type"].startswith("fuel_"))
            cars_report.append({
                "car_id":      car["id"],
                "plate":       car["plate"],
                "model":       car["model"],
                "car_name":    car.get("car_name",""),
                "branch":      car.get("branch",""),
                "fuel_liters": round(fuel_l,1),
                "fuel_cost":   round(fuel_c,2),
                "maint_cost":  round(maint_c,2),
                "total_cost":  round(fuel_c+maint_c,2),
                "workshops":   ws,
            })
        cars_report.sort(key=lambda x: -x["total_cost"])

        # ── سائقو كل مركبة خلال الشهر ──
        drivers_per_car: dict = {}
        if car_ids:
            c.execute(f"""SELECT t.car_id, t.driver_id, d.name, d.fixed_number,
                                 MIN(t.start_time) as first_trip,
                                 MAX(COALESCE(t.end_time, t.start_time)) as last_trip
                          FROM trips t
                          JOIN drivers d ON d.id = t.driver_id
                          WHERE t.car_id IN ({car_ph}) AND t.start_time LIKE ?
                          GROUP BY t.car_id, t.driver_id
                          ORDER BY t.car_id, first_trip""", car_ids + [like])
            for row in c.fetchall():
                cid = row["car_id"]
                if cid not in drivers_per_car:
                    drivers_per_car[cid] = []
                # تنسيق الفترة
                ft = row["first_trip"][:10] if row["first_trip"] else ""
                lt = row["last_trip"][:10]  if row["last_trip"]  else ""
                period = ft if ft == lt else f"{ft} → {lt}"
                drivers_per_car[cid].append({
                    "driver_id":   row["driver_id"],
                    "name":        row["name"],
                    "fixed_number":row["fixed_number"] or "",
                    "period":      period,
                })

        # أيضاً نضيف trips و km لكل مركبة من trips
        if car_ids:
            c.execute(f"""SELECT car_id,
                                 COUNT(*) as trip_count,
                                 SUM(CASE WHEN end_odometer > start_odometer
                                          THEN end_odometer - start_odometer ELSE 0 END) as total_km
                          FROM trips
                          WHERE car_id IN ({car_ph}) AND start_time LIKE ?
                          GROUP BY car_id""", car_ids + [like])
            car_trip_stats = {r["car_id"]: dict(r) for r in c.fetchall()}
            for car in cars_report:
                cts = car_trip_stats.get(car["car_id"], {})
                car["trips"] = cts.get("trip_count", 0)
                car["km"]    = round(cts.get("total_km", 0), 1)
        else:
            for car in cars_report:
                car["trips"] = 0
                car["km"] = 0.0

    return {
        "month":         month,
        "branch":        eff_branch or "كل الفروع",
        "generated_at":  datetime.utcnow().isoformat() + "Z",
        "summary": {
            "total_drivers":    len(drivers_list),
            "total_cars":       len(cars_list),
            "total_trips":      total_trips_all,
            "total_km":         round(total_km_all, 1),
            "total_fuel_cost":  round(total_fuel_cost, 2),
            "total_fuel_liters":round(total_fuel_liters, 1),
            "total_maint_cost": round(total_maint_cost, 2),
            "total_cost":       round(total_fuel_cost + total_maint_cost, 2),
        },
        "drivers": drivers_report,
        "cars":    cars_report,
        "drivers_per_car": drivers_per_car,
    }

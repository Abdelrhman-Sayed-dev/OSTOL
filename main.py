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
    UploadFile, File, status, Query
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
            role TEXT NOT NULL CHECK(role IN('superuser','admin','driver','reporter','supervisor_workshop','supervisor_field','operator')),
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

        # EQUIPMENT (جدول معدات منفصل - بدون لوحة، بالكود فقط)
        c.execute("""CREATE TABLE IF NOT EXISTS equipment (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            car_code       TEXT    NOT NULL UNIQUE,
            equipment_name TEXT    NOT NULL DEFAULT '',
            model          TEXT    DEFAULT '',
            brand          TEXT    DEFAULT '',
            status         TEXT    DEFAULT 'active',
            branch         TEXT    DEFAULT '',
            sector         TEXT    DEFAULT '',
            project        TEXT    DEFAULT '',
            year           TEXT    DEFAULT '',
            chassis        TEXT    DEFAULT '',
            engine_number  TEXT    DEFAULT '',
            serial_no      TEXT    DEFAULT '',
            capacity       TEXT    DEFAULT '',
            equipment_type TEXT    DEFAULT 'معدات تحريك تربة',
            license_expiry TEXT    DEFAULT '',
            notes          TEXT    DEFAULT '',
            created_at     TEXT    DEFAULT (datetime('now'))
        )""")
        # Migration: أضف أعمدة equipment الجديدة لو الجدول قديم
        _equip_cols = [
            ("equipment_name",  "TEXT NOT NULL DEFAULT ''"),
            ("brand",           "TEXT DEFAULT ''"),
            ("equipment_type",  "TEXT DEFAULT 'معدات تحريك تربة'"),
            ("license_expiry",  "TEXT DEFAULT ''"),
            ("serial_no",       "TEXT DEFAULT ''"),
            ("capacity",        "TEXT DEFAULT ''"),
            ("notes",           "TEXT DEFAULT ''"),
            ("created_at",      "TEXT DEFAULT ''"),
        ]
        for _col, _def in _equip_cols:
            try: c.execute(f"ALTER TABLE equipment ADD COLUMN {_col} {_def}")
            except Exception: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_equip_code   ON equipment(car_code)")
        except Exception: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_equip_branch ON equipment(branch)")
        except Exception: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_equip_type   ON equipment(equipment_type)")
        except Exception: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_equip_status ON equipment(status)")
        except Exception: pass

        # RENTAL EQUIPMENT
        c.execute("""CREATE TABLE IF NOT EXISTS rental_equipment (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            branch           TEXT DEFAULT '',
            equipment_name   TEXT NOT NULL,
            code             TEXT DEFAULT '',
            rental_source    TEXT DEFAULT '',
            work_days        TEXT DEFAULT '',
            daily_rate       TEXT DEFAULT '',
            monthly_rate     TEXT DEFAULT '',
            fuel_liters      TEXT DEFAULT '',
            hours_start      TEXT DEFAULT '',
            hours_end        TEXT DEFAULT '',
            hours_diff       TEXT DEFAULT '',
            utilization_pct  TEXT DEFAULT '',
            consumption_rate TEXT DEFAULT '',
            project          TEXT DEFAULT '',
            notes            TEXT DEFAULT '',
            month            TEXT DEFAULT '',
            sector           TEXT DEFAULT '',
            created_at       TEXT DEFAULT (datetime('now'))
        )""")
        # ── Migration: أضف الأعمدة الجديدة لو مش موجودة ──
        _rental_new_cols = [
            ("equipment_name",   "TEXT NOT NULL DEFAULT ''"),
            ("code",             "TEXT DEFAULT ''"),
            ("month",            "TEXT DEFAULT ''"),
            ("sector",           "TEXT DEFAULT ''"),
            ("daily_rate",       "TEXT DEFAULT ''"),
            ("consumption_rate", "TEXT DEFAULT ''"),
            ("project",          "TEXT DEFAULT ''"),
            ("notes",            "TEXT DEFAULT ''"),
            ("created_at",       "TEXT DEFAULT ''"),
            ("rental_source",    "TEXT DEFAULT ''"),
            ("work_days",        "TEXT DEFAULT ''"),
            ("monthly_rate",     "TEXT DEFAULT ''"),
            ("fuel_liters",      "TEXT DEFAULT ''"),
            ("hours_start",      "TEXT DEFAULT ''"),
            ("hours_end",        "TEXT DEFAULT ''"),
            ("hours_diff",       "TEXT DEFAULT ''"),
            ("utilization_pct",  "TEXT DEFAULT ''"),
            ("start_date",       "TEXT DEFAULT ''"),
        ]
        for _col, _def in _rental_new_cols:
            try:
                c.execute(f"ALTER TABLE rental_equipment ADD COLUMN {_col} {_def}")
            except Exception: pass
        # ── إصلاح قيد start_date NOT NULL: إذا كان الجدول القديم يحتوي على NOT NULL بدون DEFAULT ──
        try:
            c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='rental_equipment'")
            _rental_row = c.fetchone()
            _rental_sql = _rental_row["sql"] if _rental_row else ""
            if "start_date" in _rental_sql and "NOT NULL" in _rental_sql and "DEFAULT" not in _rental_sql.split("start_date")[1].split("\n")[0]:
                # أعد بناء الجدول بدون قيد NOT NULL على start_date
                c.execute("PRAGMA foreign_keys=OFF")
                c.execute("""CREATE TABLE IF NOT EXISTS rental_equipment_v2 (
                    id               INTEGER PRIMARY KEY AUTOINCREMENT,
                    branch           TEXT DEFAULT '',
                    equipment_name   TEXT NOT NULL DEFAULT '',
                    code             TEXT DEFAULT '',
                    rental_source    TEXT DEFAULT '',
                    work_days        TEXT DEFAULT '',
                    daily_rate       TEXT DEFAULT '',
                    monthly_rate     TEXT DEFAULT '',
                    fuel_liters      TEXT DEFAULT '',
                    hours_start      TEXT DEFAULT '',
                    hours_end        TEXT DEFAULT '',
                    hours_diff       TEXT DEFAULT '',
                    utilization_pct  TEXT DEFAULT '',
                    consumption_rate TEXT DEFAULT '',
                    project          TEXT DEFAULT '',
                    notes            TEXT DEFAULT '',
                    month            TEXT DEFAULT '',
                    sector           TEXT DEFAULT '',
                    start_date       TEXT DEFAULT '',
                    created_at       TEXT DEFAULT (datetime('now'))
                )""")
                c.execute("""INSERT OR IGNORE INTO rental_equipment_v2
                    SELECT id,branch,
                           COALESCE(equipment_name,''),code,rental_source,work_days,daily_rate,
                           monthly_rate,fuel_liters,hours_start,hours_end,hours_diff,
                           utilization_pct,consumption_rate,project,notes,month,sector,
                           COALESCE(start_date,''),COALESCE(created_at,'')
                    FROM rental_equipment""")
                c.execute("DROP TABLE rental_equipment")
                c.execute("ALTER TABLE rental_equipment_v2 RENAME TO rental_equipment")
                c.execute("PRAGMA foreign_keys=ON")
                log.info("✅ rental_equipment: removed NOT NULL from start_date")
        except Exception as _e:
            log.warning(f"rental_equipment migration (start_date fix): {_e}")
        # ── Indexes بعد التأكد إن الأعمدة موجودة ──
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_rental_branch ON rental_equipment(branch)")
        except Exception: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_rental_month  ON rental_equipment(month)")
        except Exception: pass

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

        # VOICE NOTES (سوبر يوزر → أدمنز أو سائقين أو شخص محدد)
        c.execute("""CREATE TABLE IF NOT EXISTS voice_notes(
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            sender_id      INTEGER NOT NULL,
            target_group   TEXT    NOT NULL DEFAULT '',
            target_user_id INTEGER DEFAULT NULL,
            audio_data     TEXT NOT NULL,
            duration_sec   REAL DEFAULT 0,
            created_at     TEXT NOT NULL,
            expires_at     TEXT NOT NULL,
            play_count     INTEGER DEFAULT 0,
            max_plays      INTEGER DEFAULT 2,
            is_deleted     INTEGER DEFAULT 0
        )""")
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_vnotes_group   ON voice_notes(target_group)")
        except Exception: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_vnotes_target  ON voice_notes(target_user_id)")
        except Exception: pass
        try: c.execute("CREATE INDEX IF NOT EXISTS idx_vnotes_expires ON voice_notes(expires_at)")
        except Exception: pass

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
        needs_role_fix   = 'superuser' not in tbl_sql or "'operator'" not in tbl_sql
        needs_unique_fix = 'username TEXT UNIQUE' in tbl_sql  # نشيل UNIQUE من username
        if needs_role_fix or needs_unique_fix:
            c.execute("PRAGMA foreign_keys=OFF")
            c.execute("""CREATE TABLE IF NOT EXISTS users_migrated(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL CHECK(role IN('superuser','admin','driver','reporter','supervisor_workshop','supervisor_field','operator')),
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
            log.info("✅ Migrated users: added operator role, removed UNIQUE on username")
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
                               ("description","TEXT DEFAULT ''"),("tire_action","TEXT DEFAULT ''"),
                               ("doc_number","TEXT DEFAULT ''"),
                               ("engine_hours","REAL"),
                               ("supply_source","TEXT DEFAULT ''"),
                               ("item_name","TEXT DEFAULT ''"),
                               ("item_spec","TEXT DEFAULT ''"),
                               ("receiver_name","TEXT DEFAULT ''")],
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

    # ── Migration: إضافة operator role في users CHECK constraint ──
    try:
        c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='users'")
        row = c.fetchone()
        if row:
            existing_sql = row['sql'] or ''
            if "'operator'" not in existing_sql:
                log.info("🔄 Adding operator role to users table...")
                c.execute("PRAGMA foreign_keys=OFF")
                c.execute("""CREATE TABLE IF NOT EXISTS _users_op_fix(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL,
                    password TEXT NOT NULL,
                    role TEXT NOT NULL CHECK(role IN('superuser','admin','driver','reporter','supervisor_workshop','supervisor_field','operator')),
                    branch TEXT DEFAULT '',
                    created_at TEXT DEFAULT(datetime('now')),
                    last_login TEXT,
                    refresh_token TEXT,
                    refresh_exp TEXT,
                    avatar_url TEXT DEFAULT ''
                )""")
                c.execute("INSERT OR IGNORE INTO _users_op_fix SELECT id,username,password,role,COALESCE(branch,''),created_at,last_login,refresh_token,refresh_exp,COALESCE(avatar_url,'') FROM users")
                c.execute("DROP TABLE users")
                c.execute("ALTER TABLE _users_op_fix RENAME TO users")
                c.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")
                c.execute("PRAGMA foreign_keys=ON")
                log.info("✅ operator role added to users table")
    except Exception as _e:
        log.warning(f"operator role migration: {_e}")

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

    # ── migration: avatar_url للمستخدمين ──
    try:
        c.execute("ALTER TABLE users ADD COLUMN avatar_url TEXT DEFAULT ''")
    except Exception:
        pass  # العمود موجود بالفعل

    # ── جدول صور السائقين ──
    c.execute("""CREATE TABLE IF NOT EXISTS driver_photos (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        driver_id   INTEGER NOT NULL,
        photo_url   TEXT    NOT NULL,
        caption     TEXT    DEFAULT '',
        created_at  TEXT    NOT NULL,
        FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE CASCADE
    )""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_photos_driver ON driver_photos(driver_id)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_photos_created ON driver_photos(created_at)")

    # ── جدول المواقع الحية للسائقين ──
    c.execute("""CREATE TABLE IF NOT EXISTS live_locations (
        driver_id   INTEGER PRIMARY KEY,
        car_id      INTEGER,
        lat         REAL    NOT NULL,
        lng         REAL    NOT NULL,
        accuracy    REAL    DEFAULT 0,
        heading     REAL    DEFAULT 0,
        speed       REAL    DEFAULT 0,
        updated_at  TEXT    NOT NULL,
        trip_id     INTEGER,
        FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE CASCADE
    )""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_live_updated ON live_locations(updated_at)")

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

    # ── Migration: voice_notes — أزل CHECK constraint القديم وأضف target_user_id ──
    try:
        c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='voice_notes'")
        row = c.fetchone()
        vn_sql = row['sql'] if row else ''
        if vn_sql and "CHECK" in vn_sql and "target_user_id" not in vn_sql:
            # إعادة بناء الجدول بدون CHECK وبإضافة target_user_id
            c.execute("PRAGMA foreign_keys=OFF")
            c.execute("""CREATE TABLE IF NOT EXISTS voice_notes_v2(
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                sender_id      INTEGER NOT NULL,
                target_group   TEXT    NOT NULL DEFAULT '',
                target_user_id INTEGER DEFAULT NULL,
                audio_data     TEXT NOT NULL,
                duration_sec   REAL DEFAULT 0,
                created_at     TEXT NOT NULL,
                expires_at     TEXT NOT NULL,
                play_count     INTEGER DEFAULT 0,
                max_plays      INTEGER DEFAULT 2,
                is_deleted     INTEGER DEFAULT 0
            )""")
            c.execute("""INSERT OR IGNORE INTO voice_notes_v2
                (id,sender_id,target_group,audio_data,duration_sec,
                 created_at,expires_at,play_count,max_plays,is_deleted)
                SELECT id,sender_id,target_group,audio_data,duration_sec,
                       created_at,expires_at,play_count,max_plays,is_deleted
                FROM voice_notes""")
            c.execute("DROP TABLE voice_notes")
            c.execute("ALTER TABLE voice_notes_v2 RENAME TO voice_notes")
            c.execute("PRAGMA foreign_keys=ON")
            log.info("✅ voice_notes: removed CHECK constraint, added target_user_id")
    except Exception as e:
        log.warning(f"voice_notes migration: {e}")
    try: c.execute("ALTER TABLE voice_notes ADD COLUMN target_user_id INTEGER DEFAULT NULL")
    except Exception: pass
    try: c.execute("CREATE INDEX IF NOT EXISTS idx_vnotes_target ON voice_notes(target_user_id)")
    except Exception: pass

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
    if result["role"] == "operator":
        with get_db() as conn:
            c = conn.cursor()
            c.execute("SELECT id FROM equipment_operators WHERE user_id=?", (result["user_id"],))
            op = c.fetchone()
            result["operator_id"] = op["id"] if op else None
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
    avatar_url: Optional[str] = ""

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
    "معدات تحريك تربة", "معدات رفع", "معدات ثابتة", "معدات نقل ثقيل",
    "معدات نقل", "وسائل انتقال", "معدات إنتاجية",
    "معدات مساعدة", "معدات وآلات ورش", "أخرى"
}

BRANCHES = [
    "القاهرة",
    "مدينة نصر",
    "حلوان",
    "صيانة القصور والآثار",
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
    "oil",           # إجمالي الزيوت (legacy)
    "oil_motor",     # زيت المحرك
    "oil_gear",      # زيت التروس
    "oil_brake",     # زيت الفرامل
    "oil_hydro",     # زيت الهيدروليك
    "oil_grease",    # شحم
    "filter", "tire", "battery", "belt", "other",
]

# أنواع الزيت التفصيلية
OIL_SUBTYPES = ["oil_motor","oil_gear","oil_brake","oil_hydro","oil_grease"]

OIL_SUBTYPE_LABELS = {
    "oil_motor":  "زيت المحرك",
    "oil_gear":   "زيت التروس",
    "oil_brake":  "زيت الفرامل",
    "oil_hydro":  "زيت الهيدروليك",
    "oil_grease": "شحم",
}

# col name for page2 mapping
OIL_SUBTYPE_COL = {
    "oil_motor":  "motor",
    "oil_gear":   "gear",
    "oil_brake":  "brake",
    "oil_hydro":  "hydraulic",
    "oil_grease": "grease",
}

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
        allowed = ("admin","superuser","driver","reporter","supervisor_workshop","supervisor_field","operator")
        if v not in allowed:
            raise ValueError(f"دور غير صالح. المتاح: {allowed}")
        return v

class WorkshopCreate(BaseModel):
    driver_id: int; type: str; quantity: Optional[float] = None
    price: float = 0; notes: Optional[str] = ""
    operation_type: Optional[str] = ""; vehicle_id: Optional[int] = None
    odometer_reading: Optional[float] = None; description: Optional[str] = ""
    tire_action: Optional[str] = ""; location: Optional[str] = ""
    # حقول البطاقة التشغيلية
    doc_number:    Optional[str]   = ""    # رقم مستند الصرف
    engine_hours:  Optional[float] = None  # ساعات المحرك
    supply_source: Optional[str]   = ""    # جهة الصرف
    item_name:     Optional[str]   = ""    # اسم الصنف (كاوتش/بطارية)
    item_spec:     Optional[str]   = ""    # المقاس
    receiver_name: Optional[str]   = ""    # اسم المتسلم

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
        c.execute("SELECT id,username,password,role,branch,avatar_url FROM users WHERE username=?", (data.username,))
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
            "user_id":   u["id"], "role": u["role"], "username": u["username"],
            "branch":    u["branch"] or "",
            "driver_id": driver_id  # مضاف عشان update_location يشتغل
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
                          branch=u["branch"] or "",
                          avatar_url=u["avatar_url"] or "")
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

        # جيب driver_id لو الـ role = driver
        drv_id_refresh = None
        if u["role"] == "driver":
            c.execute("SELECT id FROM drivers WHERE user_id=?", (u["id"],))
            drv_row = c.fetchone()
            drv_id_refresh = drv_row["id"] if drv_row else None

        new_access = create_access_token({
            "user_id": u["id"], "role": u["role"], "username": u["username"],
            "driver_id": drv_id_refresh
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
# USER AVATAR — رفع وجلب صورة البروفايل للمستخدم
# ══════════════════════════════════════════════════════

@app.post("/user/avatar")
async def upload_user_avatar(body: dict, cu: dict = Depends(get_user)):
    """
    يرفع صورة بروفايل للمستخدم الحالي (أدمن أو سوبر يوزر).
    body: { image: base64string }
    """
    import base64, re as _re
    raw_b64 = body.get("image", "")
    if not raw_b64:
        raise HTTPException(400, "لم يتم إرسال صورة")
    # اقبل data:image/...;base64,... أو raw base64
    m = _re.match(r"data:image/(\w+);base64,(.+)", raw_b64, _re.DOTALL)
    if m:
        ext, raw_b64 = m.group(1).lower(), m.group(2)
    else:
        ext = "jpg"
    allowed_exts = {"jpg","jpeg","png","webp","gif"}
    if ext not in allowed_exts:
        raise HTTPException(400, "نوع الصورة غير مدعوم")
    try:
        raw = base64.b64decode(raw_b64)
    except Exception:
        raise HTTPException(400, "صورة غير صالحة")
    if len(raw) > 5 * 1024 * 1024:
        raise HTTPException(400, "حجم الصورة يتجاوز 5 MB")

    fname = f"avatar_user_{cu['user_id']}_{int(datetime.utcnow().timestamp()*1000)}.{ext}"
    fpath = UPLOAD_DIR / fname
    fpath.write_bytes(raw)
    avatar_url = f"/uploads/{fname}"

    with get_db() as conn:
        # احذف الصورة القديمة لو موجودة
        old = conn.execute("SELECT avatar_url FROM users WHERE id=?", (cu["user_id"],)).fetchone()
        if old and old["avatar_url"]:
            old_path = UPLOAD_DIR / old["avatar_url"].replace("/uploads/","")
            try:
                if old_path.exists() and "avatar_user_" in old_path.name:
                    old_path.unlink()
            except Exception:
                pass
        conn.execute("UPDATE users SET avatar_url=? WHERE id=?", (avatar_url, cu["user_id"]))

    log_event("avatar_uploaded", user_id=cu["user_id"])
    return {"ok": True, "avatar_url": avatar_url}

@app.get("/user/avatar")
async def get_user_avatar(cu: dict = Depends(get_user)):
    """يرجع avatar_url للمستخدم الحالي."""
    with get_db() as conn:
        row = conn.execute("SELECT avatar_url FROM users WHERE id=?", (cu["user_id"],)).fetchone()
    return {"avatar_url": row["avatar_url"] if row else ""}

@app.delete("/user/avatar")
async def delete_user_avatar(cu: dict = Depends(get_user)):
    """يحذف صورة البروفايل."""
    with get_db() as conn:
        old = conn.execute("SELECT avatar_url FROM users WHERE id=?", (cu["user_id"],)).fetchone()
        if old and old["avatar_url"]:
            old_path = UPLOAD_DIR / old["avatar_url"].replace("/uploads/","")
            try:
                if old_path.exists() and "avatar_user_" in old_path.name:
                    old_path.unlink()
            except Exception:
                pass
        conn.execute("UPDATE users SET avatar_url='' WHERE id=?", (cu["user_id"],))
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
# 15a. DRIVER PHOTOS — رفع وعرض صور السائقين
# ══════════════════════════════════════════════════════

@app.post("/driver/photo/upload")
async def upload_driver_photo(body: dict, cu: dict = Depends(get_user)):
    """السائق يرفع صورة (base64) — تُحفظ وتُعرض عند السوبر يوزر."""
    if cu["role"] != "driver":
        raise HTTPException(403, "للسائقين فقط")

    driver_id = cu.get("driver_id")
    if not driver_id:
        with get_db() as _c:
            row = _c.execute("SELECT id FROM drivers WHERE user_id=?", (cu["user_id"],)).fetchone()
            driver_id = row["id"] if row else None
    if not driver_id:
        raise HTTPException(400, "لا يوجد سائق مرتبط بهذا الحساب")

    b64 = body.get("image")
    caption = body.get("caption", "")[:200]
    if not b64:
        raise HTTPException(400, "image مطلوبة")

    # Decode & validate image
    try:
        raw = base64.b64decode(b64)
    except Exception:
        raise HTTPException(400, "صورة غير صالحة")

    # Detect type from magic bytes
    # Detect image type from first bytes
    ext = "jpg"
    if len(raw) >= 2 and raw[0] == 0xFF and raw[1] == 0xD8:
        ext = "jpg"
    elif len(raw) >= 4 and raw[0] == 0x89 and raw[1:4] == b"PNG":
        ext = "png"
    elif len(raw) >= 3 and raw[:3] == b"GIF":
        ext = "gif"
    elif len(raw) >= 12 and raw[:4] == b"RIFF" and raw[8:12] == b"WEBP":
        ext = "webp"

    now = datetime.utcnow()
    fname = f"photo_{driver_id}_{int(now.timestamp()*1000)}.{ext}"
    dest  = UPLOAD_DIR / fname
    dest.write_bytes(raw)

    photo_url = f"/uploads/{fname}"
    created_at = now.isoformat() + "Z"

    with get_db() as conn:
        conn.execute(
            "INSERT INTO driver_photos(driver_id,photo_url,caption,created_at) VALUES(?,?,?,?)",
            (driver_id, photo_url, caption, created_at)
        )
    log_event("photo_uploaded", user_id=cu["user_id"], size=len(raw))
    return {"ok": True, "photo_url": photo_url, "created_at": created_at}


@app.get("/driver/photos")
async def get_driver_photos(
    driver_id: Optional[int] = None,
    limit: int = 50,
    cu: dict = Depends(require_admin_or_reporter)
):
    """الأدمن والسوبر يوزر يشوفوا صور السائقين."""
    with get_db() as conn:
        c = conn.cursor()
        if driver_id:
            c.execute("""SELECT dp.*, d.name as driver_name, d.fixed_number
                         FROM driver_photos dp
                         JOIN drivers d ON d.id = dp.driver_id
                         WHERE dp.driver_id=?
                         ORDER BY dp.created_at DESC LIMIT ?""",
                      (driver_id, limit))
        else:
            branch = _branch_filter(cu)
            if branch:
                c.execute("""SELECT dp.*, d.name as driver_name, d.fixed_number
                             FROM driver_photos dp
                             JOIN drivers d ON d.id = dp.driver_id
                             WHERE d.branch=?
                             ORDER BY dp.created_at DESC LIMIT ?""",
                          (branch, limit))
            else:
                c.execute("""SELECT dp.*, d.name as driver_name, d.fixed_number
                             FROM driver_photos dp
                             JOIN drivers d ON d.id = dp.driver_id
                             ORDER BY dp.created_at DESC LIMIT ?""",
                          (limit,))
        return [dict(r) for r in c.fetchall()]


@app.delete("/driver/photos/{photo_id}")
async def delete_driver_photo(photo_id: int, cu: dict = Depends(require_superuser)):
    """السوبر يوزر يحذف صورة."""
    with get_db() as conn:
        row = conn.execute("SELECT photo_url FROM driver_photos WHERE id=?", (photo_id,)).fetchone()
        if not row:
            raise HTTPException(404, "الصورة غير موجودة")
        # Delete file
        try:
            fp = Path(str(UPLOAD_DIR) + row["photo_url"].replace("/uploads",""))
            if fp.exists(): fp.unlink()
        except Exception: pass
        conn.execute("DELETE FROM driver_photos WHERE id=?", (photo_id,))
    return {"ok": True}


# ══════════════════════════════════════════════════════
# 15b. LIVE LOCATION — real-time driver tracking
# ══════════════════════════════════════════════════════

@app.post("/location/update")
async def update_location(body: dict, cu: dict = Depends(get_user)):
    """السائق يبعت موقعه الحالي (كل 5 ثواني)."""
    if cu["role"] != "driver":
        raise HTTPException(403, "للسائقين فقط")

    driver_id = cu.get("driver_id")
    # Fallback: ابحث عن driver_id من الـ DB لو مش موجود في الـ token
    if not driver_id:
        with get_db() as _conn:
            _row = _conn.execute("SELECT id FROM drivers WHERE user_id=?", (cu["user_id"],)).fetchone()
            driver_id = _row["id"] if _row else None
    if not driver_id:
        raise HTTPException(400, "لا يوجد سائق مرتبط بهذا الحساب")

    lat = body.get("lat")
    lng = body.get("lng")
    if lat is None or lng is None:
        raise HTTPException(400, "lat و lng مطلوبان")

    car_id   = body.get("car_id")
    accuracy = body.get("accuracy", 0)
    heading  = body.get("heading", 0)
    speed    = body.get("speed", 0)
    trip_id  = body.get("trip_id")
    now      = datetime.utcnow().isoformat() + "Z"

    with get_db() as conn:
        conn.execute("""INSERT INTO live_locations
            (driver_id,car_id,lat,lng,accuracy,heading,speed,updated_at,trip_id)
            VALUES(?,?,?,?,?,?,?,?,?)
            ON CONFLICT(driver_id) DO UPDATE SET
                car_id=excluded.car_id, lat=excluded.lat, lng=excluded.lng,
                accuracy=excluded.accuracy, heading=excluded.heading,
                speed=excluded.speed, updated_at=excluded.updated_at,
                trip_id=excluded.trip_id""",
            (driver_id, car_id, lat, lng, accuracy, heading, speed, now, trip_id))
    return {"ok": True}


@app.get("/location/live")
async def get_live_locations(cu: dict = Depends(require_admin_or_reporter)):
    """الأدمن يجيب آخر مواقع كل السائقين النشطين (آخر 2 دقيقة)."""
    stale_cutoff = (datetime.utcnow() - timedelta(seconds=30)).isoformat() + "Z"
    branch = _branch_filter(cu)
    with get_db() as conn:
        c = conn.cursor()
        if branch:
            c.execute("""SELECT ll.driver_id, ll.car_id, ll.lat, ll.lng,
                                ll.accuracy, ll.heading, ll.speed,
                                ll.updated_at, ll.trip_id,
                                d.name as driver_name,
                                c.plate as car_plate, c.model as car_model
                         FROM live_locations ll
                         LEFT JOIN drivers d ON d.id = ll.driver_id
                         LEFT JOIN cars    c ON c.id = ll.car_id
                         WHERE ll.updated_at >= ? AND d.branch = ?
                         ORDER BY ll.updated_at DESC""", (stale_cutoff, branch))
        else:
            c.execute("""SELECT ll.driver_id, ll.car_id, ll.lat, ll.lng,
                                ll.accuracy, ll.heading, ll.speed,
                                ll.updated_at, ll.trip_id,
                                d.name as driver_name,
                                c.plate as car_plate, c.model as car_model
                         FROM live_locations ll
                         LEFT JOIN drivers d ON d.id = ll.driver_id
                         LEFT JOIN cars    c ON c.id = ll.car_id
                         WHERE ll.updated_at >= ?
                         ORDER BY ll.updated_at DESC""", (stale_cutoff,))
        return [dict(r) for r in c.fetchall()]


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
    QUANTITY_TYPES = {t for t in WORKSHOP_TYPES if t.startswith("fuel_") or t.startswith("oil") or t in ("tire",)}
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
                      operation_type,vehicle_id,odometer_reading,description,tire_action,location,
                      doc_number,engine_hours,supply_source,item_name,item_spec,receiver_name)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (rec.driver_id, rec.type, rec.quantity, final_price, rec.notes or "", now,
                   rec.operation_type or "", rec.vehicle_id, rec.odometer_reading,
                   rec.description or "", rec.tire_action or "", rec.location or "",
                   rec.doc_number or "", rec.engine_hours, rec.supply_source or "",
                   rec.item_name or "", rec.item_spec or "", rec.receiver_name or ""))
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
async def get_workshops(cu: dict = Depends(get_user), branch: Optional[str] = None, month: Optional[str] = None):
    q = """SELECT w.*,d.name as driver_name,c.plate as vehicle_plate
           FROM workshop_records w
           LEFT JOIN drivers d ON w.driver_id=d.id
           LEFT JOIN cars c ON w.vehicle_id=c.id"""
    with get_db() as conn:
        c = conn.cursor()
        branch = _effective_branch(cu, branch)
        if cu["role"] in ("admin", "reporter", "superuser"):
            conditions = []
            params = []
            if branch:
                conditions.append("d.branch=?")
                params.append(branch)
            if month:
                conditions.append("w.created_at LIKE ?")
                params.append(month + "%")
            where = (" WHERE " + " AND ".join(conditions)) if conditions else ""
            c.execute(q + where + " ORDER BY w.id DESC LIMIT 500", params)
        else:
            did = cu.get("driver_id")
            if not did: return []
            month_cond = " AND w.created_at LIKE ?" if month else ""
            month_params = [month + "%"] if month else []
            c.execute(q + " WHERE w.driver_id=?" + month_cond + " ORDER BY w.id DESC LIMIT 200", [did] + month_params)
        rows = []
        for r in c.fetchall():
            d = dict(r)
            for k in ["operation_type","vehicle_id","odometer_reading",
                      "description","vehicle_plate","tire_action","location",
                      "doc_number","engine_hours","supply_source",
                      "item_name","item_spec","receiver_name"]:
                d.setdefault(k, None)
            rows.append(d)
        return rows

# ══════════════════════════════════════════════════════
# 18b. OPERATIONAL CARD PAGE 2 — زيوت وكاوتش وبطاريات
# ══════════════════════════════════════════════════════

@app.get("/reports/operational/page2")
async def get_operational_page2(
    car_id: Optional[int] = None,
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    cu: dict = Depends(require_admin_or_reporter)
):
    """
    يجيب سجلات الصيانة لمركبة في فترة زمنية:
    - oils: تغيير زيت / فلاتر / شحم / هيدروليك / فرامل
    - tires: إطارات
    - batteries: بطاريات
    """
    with get_db() as conn:
        c = conn.cursor()

        conditions = []
        params = []

        if car_id:
            conditions.append("w.vehicle_id = ?")
            params.append(car_id)
        if date_from:
            conditions.append("date(w.created_at) >= ?")
            params.append(date_from)
        if date_to:
            conditions.append("date(w.created_at) <= ?")
            params.append(date_to)

        # Branch filter
        branch = _branch_filter(cu)
        if branch:
            conditions.append("d.branch = ?")
            params.append(branch)

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        q = f"""
            SELECT w.id, w.created_at, w.type, w.operation_type,
                   w.quantity, w.price, w.notes, w.odometer_reading,
                   w.tire_action, w.location, w.description,
                   d.name as driver_name,
                   c.plate as vehicle_plate, c.model as vehicle_model
            FROM workshop_records w
            LEFT JOIN drivers d ON w.driver_id = d.id
            LEFT JOIN cars c ON w.vehicle_id = c.id
            {where}
            ORDER BY w.created_at ASC
        """
        c.execute(q, params)
        rows = [dict(r) for r in c.fetchall()]

    # Categorize rows
    oils      = []
    tires     = []
    batteries = []

    OIL_TYPES     = {"oil", "oil_motor", "oil_gear", "oil_brake", "oil_hydro", "oil_grease", "filter"}
    TIRE_TYPES    = {"tire"}
    BATTERY_TYPES = {"battery"}

    OIL_OP_MAP = {
        "تغيير زيت":    "زيت_محرك",
        "فلتر زيت":     "زيت_محرك",
        "فلتر هواء":    "أخرى",
        "فلتر وقود":    "أخرى",
        "فلتر مكيف":    "أخرى",
    }

    # Map oil subtype → column
    OIL_TYPE_COL = {
        "oil_motor":  "motor",
        "oil_gear":   "gear",
        "oil_brake":  "brake",
        "oil_hydro":  "hydraulic",
        "oil_grease": "grease",
        "oil":        "other",   # legacy
    }
    OIL_OP_LABELS = {
        "زيت محرك":    "motor",
        "زيت تروس":    "gear",
        "زيت فرامل":   "brake",
        "زيت هيدروليك":"hydraulic",
        "شحم":          "grease",
    }

    for r in rows:
        t   = r.get("type","")
        op  = r.get("operation_type","") or ""
        date_str = (r.get("created_at","") or "")[:10]
        qty      = r.get("quantity")
        odo      = r.get("odometer_reading")
        notes    = r.get("notes","") or ""
        price    = r.get("price") or ""
        doc_num  = r.get("doc_number","") or ""
        eng_h    = r.get("engine_hours")
        supply   = r.get("supply_source","") or ""
        item_n   = r.get("item_name","") or ""
        item_s   = r.get("item_spec","") or ""
        recv     = r.get("receiver_name","") or ""

        if t in OIL_TYPES:
            # If it's a subtype, use type directly; else fall back to operation_type
            if t in OIL_TYPE_COL:
                oil_col = OIL_TYPE_COL[t]
            else:
                oil_col = OIL_OP_LABELS.get(op, "other")
            oils.append({
                "date":         date_str,
                "oil_col":      oil_col,      # motor/gear/brake/hydraulic/grease/other
                "operation":    op or t,
                "quantity":     qty or "",
                "price":        price,
                "odometer":     odo or "",
                "engine_hours": eng_h or "",
                "doc_number":   doc_num,
                "supply_source":supply,
                "notes":        notes,
                "driver":       r.get("driver_name",""),
                "vehicle":      r.get("vehicle_plate",""),
            })
        elif t in TIRE_TYPES:
            tires.append({
                "date":          date_str,
                "item":          item_n or "إطارات",
                "spec":          item_s or r.get("tire_action","") or "",
                "quantity":      qty or "",
                "price":         price,
                "doc_number":    doc_num,
                "receiver_name": recv,
                "notes":         notes,
                "driver":        r.get("driver_name",""),
                "vehicle":       r.get("vehicle_plate",""),
            })
        elif t in BATTERY_TYPES:
            tires.append({
                "date":          date_str,
                "item":          item_n or "بطارية",
                "spec":          item_s or op or "",
                "quantity":      qty or "",
                "price":         price,
                "doc_number":    doc_num,
                "receiver_name": recv,
                "notes":         notes,
                "driver":        r.get("driver_name",""),
                "vehicle":       r.get("vehicle_plate",""),
            })

    return {
        "oils":      oils,
        "tires":     tires,
        "total_oils":     len(oils),
        "total_tires":    len(tires),
    }


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
    branch = _branch_filter(cu)
    with get_db() as conn:
        c = conn.cursor()
        branch_join = "LEFT JOIN cars c ON g.car_id=c.id" if not branch else "INNER JOIN cars c ON g.car_id=c.id AND c.branch=?"
        branch_params_car = (branch,) if branch else ()
        branch_params_nocar = ()

        c.execute(f"""SELECT g.*,d.name as driver_name,c.plate as car_plate
                     FROM garage_records g
                     INNER JOIN(SELECT car_id, MAX(id) as max_id
                                FROM garage_records
                                WHERE car_id IS NOT NULL
                                GROUP BY car_id) latest
                       ON g.id=latest.max_id
                     LEFT JOIN drivers d ON g.driver_id=d.id
                     {branch_join}
                     {"WHERE c.id IS NOT NULL" if branch else ""}
                     ORDER BY g.recorded_at DESC""", branch_params_car)
        rows = [dict(r) for r in c.fetchall()]

        if not branch:
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
    group_by:    str = "month",     # day | week | month | year  (used only when detailed)
    granularity: str = "monthly",   # daily | weekly | monthly (alias, mapped by frontend)
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
            elif group_by == "week":
                # Return the Monday (start) of the week as a date string
                grp = "date(t.start_time, 'weekday 1', '-6 days')"
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
        # ملخص حسب النوع
        c.execute(f"""
            SELECT w.type, SUM(w.quantity) AS total_qty, SUM(w.price) AS total_price,
                   COUNT(*) AS count
            FROM workshop_records w
            WHERE {' AND '.join(fuel_conds)}
            GROUP BY w.type
        """, fuel_params)
        fuel_records = [dict(r) for r in c.fetchall()]

        # تفاصيل كل عملية تمويل يوم بيوم (للسولار تحديداً)
        c.execute(f"""
            SELECT date(w.created_at) AS fuel_date,
                   w.type,
                   w.quantity,
                   w.price,
                   w.odometer_reading,
                   w.notes,
                   d.name AS driver_name
            FROM workshop_records w
            LEFT JOIN drivers d ON d.id = w.driver_id
            WHERE {' AND '.join(fuel_conds)}
            ORDER BY w.created_at ASC
        """, fuel_params)
        fuel_detail = [dict(r) for r in c.fetchall()]

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
        "fuel_detail":  fuel_detail,
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

EQUIPMENT_IMPORT_COLS = ["car_code","equipment_name","brand","model","year","chassis",
                          "equipment_type","branch","sector","project","status","notes"]

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
async def fraud_detection(
    date_from: Optional[str] = None,
    date_to:   Optional[str] = None,
    cu: dict = Depends(require_superuser)
):
    """تحليل شامل لكشف التلاعب المحتمل في البيانات — يدعم تصفية بالتاريخ."""

    # ── بناء شرط التاريخ للرحلات ──
    trip_date_conds = []
    trip_date_params: list = []
    if date_from:
        trip_date_conds.append("date(t.start_time) >= ?")
        trip_date_params.append(date_from)
    if date_to:
        trip_date_conds.append("date(t.start_time) <= ?")
        trip_date_params.append(date_to)
    trip_date_sql = (" AND " + " AND ".join(trip_date_conds)) if trip_date_conds else ""

    # ── بناء شرط التاريخ لسجلات الورشة ──
    ws_date_conds = []
    ws_date_params: list = []
    if date_from:
        ws_date_conds.append("date(w.created_at) >= ?")
        ws_date_params.append(date_from)
    if date_to:
        ws_date_conds.append("date(w.created_at) <= ?")
        ws_date_params.append(date_to)
    ws_date_sql = (" AND " + " AND ".join(ws_date_conds)) if ws_date_conds else ""

    # ── شرط تاريخ لطلبات السائقين (بدون JOIN رحلات) ──
    req_date_conds = []
    req_date_params: list = []
    if date_from:
        req_date_conds.append("date(dr.created_at) >= ?")
        req_date_params.append(date_from)
    if date_to:
        req_date_conds.append("date(dr.created_at) <= ?")
        req_date_params.append(date_to)
    req_date_sql = (" AND " + " AND ".join(req_date_conds)) if req_date_conds else ""

    with get_db() as conn:
        c = conn.cursor()

        # 1. عداد النهاية أقل من عداد البداية (تلاعب صريح)
        c.execute(f"""SELECT t.id, t.start_odometer, t.end_odometer,
                            t.start_time, t.end_time,
                            d.name driver_name, d.fixed_number,
                            ca.plate, ca.model
                     FROM trips t
                     JOIN drivers d ON d.id=t.driver_id
                     JOIN cars ca ON ca.id=t.car_id
                     WHERE t.end_odometer IS NOT NULL
                       AND t.start_odometer IS NOT NULL
                       AND t.end_odometer < t.start_odometer
                       {trip_date_sql}""", trip_date_params)
        odometer_reversed = [dict(r) for r in c.fetchall()]

        # 2. رحلات بمسافة مشبوهة (أكثر من 500 كم في رحلة واحدة)
        c.execute(f"""SELECT t.id, t.start_odometer, t.end_odometer,
                            (t.end_odometer - t.start_odometer) AS distance,
                            t.start_time, t.end_time,
                            d.name driver_name, d.fixed_number,
                            ca.plate, ca.model
                     FROM trips t
                     JOIN drivers d ON d.id=t.driver_id
                     JOIN cars ca ON ca.id=t.car_id
                     WHERE t.end_odometer IS NOT NULL AND t.start_odometer IS NOT NULL
                       AND (t.end_odometer - t.start_odometer) > 500
                       {trip_date_sql}
                     ORDER BY distance DESC""", trip_date_params)
        high_distance = [dict(r) for r in c.fetchall()]

        # 3. رحلتان متداخلتان لنفس السائق
        c.execute(f"""SELECT a.id id1, b.id id2,
                            a.start_time, a.end_time,
                            b.start_time start2, b.end_time end2,
                            d.name driver_name, d.fixed_number,
                            '' AS plate, '' AS model
                     FROM trips a JOIN trips b ON a.driver_id=b.driver_id AND a.id<b.id
                     JOIN drivers d ON d.id=a.driver_id
                     WHERE a.end_time IS NOT NULL AND b.end_time IS NOT NULL
                       AND a.start_time < b.end_time AND a.end_time > b.start_time
                       {trip_date_sql.replace('t.start_time', 'a.start_time')}""",
                  trip_date_params)
        overlapping = [dict(r) for r in c.fetchall()]

        # 4. رحلتان متداخلتان لنفس المركبة (مع سائقين مختلفين)
        c.execute(f"""SELECT a.id id1, b.id id2, ca.plate, ca.model,
                            da.name driver1, db.name driver2,
                            a.start_time, a.end_time,
                            b.start_time start2, b.end_time end2
                     FROM trips a JOIN trips b ON a.car_id=b.car_id AND a.id<b.id
                     JOIN cars ca ON ca.id=a.car_id
                     JOIN drivers da ON da.id=a.driver_id
                     JOIN drivers db ON db.id=b.driver_id
                     WHERE a.end_time IS NOT NULL AND b.end_time IS NOT NULL
                       AND a.start_time < b.end_time AND a.end_time > b.start_time
                       AND a.driver_id != b.driver_id
                       {trip_date_sql.replace('t.start_time', 'a.start_time')}""",
                  trip_date_params)
        car_overlap = [dict(r) for r in c.fetchall()]

        # 5. رحلة وقتها قصير جداً (<3 دقائق) لكن مسافة كبيرة (>10 كم)
        c.execute(f"""SELECT t.id,
                            (julianday(t.end_time)-julianday(t.start_time))*1440 AS minutes,
                            (t.end_odometer - t.start_odometer) AS distance,
                            d.name driver_name, ca.plate, ca.model,
                            t.start_time, t.end_time
                     FROM trips t
                     JOIN drivers d ON d.id=t.driver_id
                     JOIN cars ca ON ca.id=t.car_id
                     WHERE t.end_time IS NOT NULL AND t.end_odometer IS NOT NULL
                       AND t.start_odometer IS NOT NULL
                       AND (julianday(t.end_time)-julianday(t.start_time))*1440 < 3
                       AND (t.end_odometer - t.start_odometer) > 10
                       {trip_date_sql}""", trip_date_params)
        impossible_speed = [dict(r) for r in c.fetchall()]

        # 6. وقود كثير في يوم واحد (>200 لتر)
        c.execute(f"""SELECT w.vehicle_id, DATE(w.created_at) AS day,
                            SUM(w.quantity) AS total_liters, COUNT(*) AS refills,
                            ca.plate, ca.model
                     FROM workshop_records w
                     JOIN cars ca ON ca.id=w.vehicle_id
                     WHERE w.type LIKE 'fuel_%' AND w.quantity IS NOT NULL
                       {ws_date_sql}
                     GROUP BY w.vehicle_id, day
                     HAVING total_liters > 200
                     ORDER BY total_liters DESC""", ws_date_params)
        excess_fuel = [dict(r) for r in c.fetchall()]

        # 7. سائق عدّل طلب عداد أكثر من 3 مرات في الفترة
        req_week_cond = "AND dr.created_at >= datetime('now','-7 days')" if not req_date_conds else ""
        c.execute(f"""SELECT dr.driver_id, COUNT(*) AS requests,
                            d.name driver_name, d.fixed_number,
                            MIN(dr.created_at) AS first_req
                     FROM driver_requests dr
                     JOIN drivers d ON d.id=dr.driver_id
                     WHERE dr.type='odometer_change'
                       {req_week_cond}
                       {req_date_sql}
                     GROUP BY dr.driver_id HAVING requests > 3
                     ORDER BY requests DESC""", req_date_params)
        repeated_odometer = [dict(r) for r in c.fetchall()]

        # 8. رحلات مفتوحة (بدون end_time) منذ أكثر من 3 ساعات
        c.execute(f"""SELECT t.id, t.start_time, t.start_odometer,
                            d.name driver_name, d.fixed_number,
                            ca.plate, ca.model,
                            ROUND((julianday('now') - julianday(t.start_time)) * 24, 1) AS hours_open
                     FROM trips t
                     JOIN drivers d ON d.id=t.driver_id
                     JOIN cars ca ON ca.id=t.car_id
                     WHERE t.end_time IS NULL
                       AND t.start_time IS NOT NULL
                       AND (julianday('now') - julianday(t.start_time)) * 24 > 3
                       {trip_date_sql}
                     ORDER BY hours_open DESC""", trip_date_params)
        open_trips = [dict(r) for r in c.fetchall()]

    # ── بناء قائمة موحدة من التنبيهات (alerts) للـ frontend ──
    alerts: list = []

    for r in odometer_reversed:
        alerts.append({
            "type":     "odometer_reverse",
            "severity": "high",
            "label":    "عداد متناقص",
            "plate":    r.get("plate", "—"),
            "model":    r.get("model", ""),
            "date":     (r.get("start_time") or "")[:10],
            "detail":   f"عداد بداية {r['start_odometer']} → نهاية {r['end_odometer']} | السائق: {r['driver_name']}",
        })

    for r in impossible_speed:
        alerts.append({
            "type":     "impossible_speed",
            "severity": "high",
            "label":    "سرعة مستحيلة",
            "plate":    r.get("plate", "—"),
            "model":    r.get("model", ""),
            "date":     (r.get("start_time") or "")[:10],
            "detail":   f"{round(r['distance'],1)} كم في {round(r['minutes'],1)} دقيقة | السائق: {r['driver_name']}",
        })

    for r in car_overlap:
        alerts.append({
            "type":     "car_overlap",
            "severity": "high",
            "label":    "مركبة بسائقين متزامنين",
            "plate":    r.get("plate", "—"),
            "model":    r.get("model", ""),
            "date":     (r.get("start_time") or "")[:10],
            "detail":   f"{r['driver1']} و {r['driver2']} في نفس الوقت",
        })

    for r in overlapping:
        alerts.append({
            "type":     "driver_overlap",
            "severity": "medium",
            "label":    "رحلتان متداخلتان لسائق",
            "plate":    "—",
            "model":    "",
            "date":     (r.get("start_time") or "")[:10],
            "detail":   f"السائق: {r['driver_name']} | رحلة {r['id1']} ورحلة {r['id2']}",
        })

    for r in excess_fuel:
        alerts.append({
            "type":     "duplicate_fuel",
            "severity": "medium",
            "label":    "وقود زائد في يوم",
            "plate":    r.get("plate", "—"),
            "model":    r.get("model", ""),
            "date":     r.get("day", ""),
            "detail":   f"{round(r['total_liters'],1)} لتر في {r['refills']} عمليات تزويد",
        })

    for r in high_distance:
        alerts.append({
            "type":     "high_distance",
            "severity": "low",
            "label":    "مسافة مرتفعة",
            "plate":    r.get("plate", "—"),
            "model":    r.get("model", ""),
            "date":     (r.get("start_time") or "")[:10],
            "detail":   f"{round(r['distance'],1)} كم في رحلة واحدة | السائق: {r['driver_name']}",
        })

    for r in repeated_odometer:
        alerts.append({
            "type":     "repeated_odometer",
            "severity": "low",
            "label":    "طلبات تعديل عداد متكررة",
            "plate":    "—",
            "model":    "",
            "date":     (r.get("first_req") or "")[:10],
            "detail":   f"السائق: {r['driver_name']} | {r['requests']} طلبات",
        })

    for r in open_trips:
        hours = r.get("hours_open", 0)
        severity = "high" if hours > 24 else "medium" if hours > 8 else "low"
        alerts.append({
            "type":     "open_trip",
            "severity": severity,
            "label":    "رحلة مفتوحة",
            "plate":    r.get("plate", "—"),
            "model":    r.get("model", ""),
            "date":     (r.get("start_time") or "")[:10],
            "detail":   f"السائق: {r['driver_name']} | مفتوحة منذ {hours} ساعة | بدأت {(r.get('start_time') or '')[:16]}",
        })

    high_count   = sum(1 for a in alerts if a["severity"] == "high")
    medium_count = sum(1 for a in alerts if a["severity"] == "medium")
    low_count    = sum(1 for a in alerts if a["severity"] == "low")

    return {
        "summary": {
            "total":  len(alerts),
            "high":   high_count,
            "medium": medium_count,
            "low":    low_count,
            # backward compat keys
            "total_alerts": len(alerts),
            "high_risk":    high_count,
            "medium_risk":  medium_count,
            "low_risk":     low_count,
        },
        "alerts": alerts,
        # raw lists kept for backward compat
        "odometer_reversed":  odometer_reversed,
        "high_distance":      high_distance,
        "overlapping_driver": overlapping,
        "car_overlap":        car_overlap,
        "impossible_speed":   impossible_speed,
        "excess_fuel":        excess_fuel,
        "repeated_odometer":  repeated_odometer,
        "open_trips":         open_trips,
    }


# ══════════════════════════════════════════════════════
# 33. VOICE NOTES — رسائل صوتية من السوبر
# ══════════════════════════════════════════════════════
class VoiceNoteCreate(BaseModel):
    target_group:   str           = ''   # 'admins' | 'drivers' | '' لو target_user_id محدد
    target_user_id: Optional[int] = None  # id المستخدم لو هترسل لشخص محدد
    audio_data:     str                   # base64
    duration_sec:   float = 0
    max_plays:      int   = 2    # مرتين افتراضياً
    hours_ttl:      int   = 24   # يتمسح بعد 24 ساعة

@app.post("/voice-notes")
async def create_voice_note(body: VoiceNoteCreate, cu: dict = Depends(get_user)):
    if cu["role"] not in ("superuser", "admin"):
        raise HTTPException(403, "غير مصرح")
    # يجب أن يكون target_group صالح أو target_user_id محدد
    if body.target_user_id is None and body.target_group not in ("admins", "drivers"):
        raise HTTPException(400, "حدد مجموعة مستلمين (admins/drivers) أو شخصاً محدداً")
    now      = datetime.utcnow()
    expires  = (now + timedelta(hours=body.hours_ttl)).isoformat() + "Z"
    # لو target_user_id محدد، نستخدم 'admins' كـ target_group حتى تمر CHECK constraint القديمة
    # الـ routing الفعلي بيتم بالـ target_user_id وليس target_group
    tg = body.target_group if body.target_user_id is None else (body.target_group or 'admins')
    if tg not in ('admins', 'drivers'):
        tg = 'admins'
    with get_db() as conn:
        c = conn.cursor()
        # جرب INSERT مع target_user_id أولاً، لو فشل جرب بدونه (جدول قديم)
        try:
            c.execute("""INSERT INTO voice_notes(sender_id,target_group,target_user_id,audio_data,
                         duration_sec,created_at,expires_at,max_plays)
                         VALUES(?,?,?,?,?,?,?,?)""",
                      (cu["user_id"], tg, body.target_user_id, body.audio_data,
                       body.duration_sec, now.isoformat()+"Z", expires, body.max_plays))
        except Exception as e:
            if 'target_user_id' in str(e):
                # عمود target_user_id مش موجود بعد — اعمل migration وإعادة المحاولة
                c.execute("ALTER TABLE voice_notes ADD COLUMN target_user_id INTEGER DEFAULT NULL")
                c.execute("""INSERT INTO voice_notes(sender_id,target_group,target_user_id,audio_data,
                             duration_sec,created_at,expires_at,max_plays)
                             VALUES(?,?,?,?,?,?,?,?)""",
                          (cu["user_id"], tg, body.target_user_id, body.audio_data,
                           body.duration_sec, now.isoformat()+"Z", expires, body.max_plays))
            else:
                raise
        note_id = c.lastrowid
    return {"id": note_id, "expires_at": expires}

@app.get("/voice-notes")
async def list_voice_notes(cu: dict = Depends(get_user)):
    """
    - superuser/admin بيبعتوا → بيشوفوا الرسائل اللي أرسلوها هم (صفحة الإدارة)
    - driver/admin (كمستقبل) → بيشوف الرسائل اللي مرسلة له بالاسم أو لمجموعته
    الفصل: نستخدم query param ?sent=1 لعرض المُرسَل
    """
    role = cu["role"]
    uid  = cu["user_id"]
    now  = datetime.utcnow().isoformat() + "Z"

    # ── صفحة إدارة الرسائل (superuser أو admin) ──
    # GET /voice-notes يُرجع الرسائل المُرسَلة من الشخص ده
    if role in ("superuser", "admin"):
        with get_db() as conn:
            c = conn.cursor()
            try:
                c.execute("""
                    SELECT id, target_group,
                           COALESCE(target_user_id, NULL) AS target_user_id,
                           duration_sec, created_at,
                           expires_at, play_count, max_plays, is_deleted
                    FROM voice_notes
                    WHERE sender_id = ?
                    ORDER BY created_at DESC
                """, (uid,))
            except Exception:
                c.execute("""
                    SELECT id, target_group, NULL AS target_user_id,
                           duration_sec, created_at,
                           expires_at, play_count, max_plays, is_deleted
                    FROM voice_notes
                    WHERE sender_id = ?
                    ORDER BY created_at DESC
                """, (uid,))
            rows = [dict(r) for r in c.fetchall()]
            # أضف اسم المستلم
            for row in rows:
                tid = row.get("target_user_id")
                if tid:
                    c.execute("SELECT username FROM users WHERE id=?", (tid,))
                    u = c.fetchone()
                    if u:
                        row["target_name"] = u["username"]
                    else:
                        c.execute("SELECT name FROM drivers WHERE user_id=?", (tid,))
                        d = c.fetchone()
                        row["target_name"] = d["name"] if d else "مستخدم"
            return rows

    # ── استقبال الرسائل (driver, reporter …) ──
    if role == "driver":
        group = "drivers"
    elif role in ("reporter", "supervisor_workshop", "supervisor_field"):
        group = "admins"
    else:
        return []

    with get_db() as conn:
        c = conn.cursor()
        # الرسائل المرسلة للمجموعة أو للشخص تحديداً — مع استثناء الرسائل اللي بعتها هو نفسه
        c.execute("""
            SELECT id, duration_sec, created_at, expires_at,
                   play_count, max_plays, audio_data
            FROM voice_notes
            WHERE (target_group = ? OR COALESCE(target_user_id, 0) = ?)
              AND sender_id != ?
              AND is_deleted = 0
              AND expires_at > ?
              AND play_count < max_plays
            ORDER BY created_at DESC
        """, (group, uid, uid, now))
        return [dict(r) for r in c.fetchall()]


@app.get("/voice-notes/sent")
async def list_sent_voice_notes(cu: dict = Depends(get_user)):
    """رسائل صوتية أرسلها هذا المستخدم — للسوبر والأدمن فقط"""
    if cu["role"] not in ("superuser", "admin"):
        raise HTTPException(403, "غير مصرح")
    uid = cu["user_id"]
    with get_db() as conn:
        c = conn.cursor()
        try:
            c.execute("""
                SELECT id, target_group,
                       COALESCE(target_user_id, NULL) AS target_user_id,
                       duration_sec, created_at,
                       expires_at, play_count, max_plays, is_deleted
                FROM voice_notes
                WHERE sender_id = ?
                ORDER BY created_at DESC
            """, (uid,))
        except Exception:
            c.execute("""
                SELECT id, target_group, NULL AS target_user_id,
                       duration_sec, created_at,
                       expires_at, play_count, max_plays, is_deleted
                FROM voice_notes WHERE sender_id = ?
                ORDER BY created_at DESC
            """, (uid,))
        rows = [dict(r) for r in c.fetchall()]
        for row in rows:
            tid = row.get("target_user_id")
            if tid:
                c.execute("SELECT username FROM users WHERE id=?", (tid,))
                u = c.fetchone()
                row["target_name"] = u["username"] if u else None
                if not row["target_name"]:
                    c.execute("SELECT name FROM drivers WHERE user_id=?", (tid,))
                    d = c.fetchone()
                    row["target_name"] = d["name"] if d else "مستخدم"
        return rows


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


@app.get("/voice-notes/inbox")
async def get_voice_notes_inbox(cu: dict = Depends(get_user)):
    """
    رسائل صوتية مستلمة — للأدمن والسائقين:
    - admin: يستلم الرسائل المرسلة لـ 'admins' أو لشخصه تحديداً
    - driver: يستلم الرسائل المرسلة لـ 'drivers' أو لشخصه تحديداً
    """
    role = cu["role"]
    uid  = cu["user_id"]
    now  = datetime.utcnow().isoformat() + "Z"

    if role in ("admin", "reporter", "supervisor_workshop", "supervisor_field"):
        group = "admins"
    elif role == "driver":
        group = "drivers"
    else:
        return []

    with get_db() as conn:
        c = conn.cursor()
        c.execute("""
            SELECT id, duration_sec, created_at, expires_at,
                   play_count, max_plays, audio_data, target_group, target_user_id
            FROM voice_notes
            WHERE (target_group = ? OR COALESCE(target_user_id, 0) = ?)
              AND sender_id != ?
              AND is_deleted = 0
              AND expires_at > ?
              AND play_count < max_plays
            ORDER BY created_at DESC
        """, (group, uid, uid, now))
        return [dict(r) for r in c.fetchall()]

@app.delete("/voice-notes/{note_id}")
async def delete_voice_note(note_id: int, cu: dict = Depends(get_user)):
    if cu["role"] not in ("superuser", "admin"):
        raise HTTPException(403, "غير مصرح")
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



@app.get("/voice-notes/targets")
async def get_voice_note_targets(cu: dict = Depends(get_user)):
    if cu["role"] not in ("superuser", "admin"):
        raise HTTPException(403, "غير مصرح")
    """يرجع قائمة المجموعات المتاحة لإرسال الرسائل الصوتية إليها."""
    targets = [
        {"id": "admins",  "label": "المشرفون والإداريون"},
        {"id": "drivers", "label": "السائقون"},
    ]
    return {"targets": targets}


@app.get("/voice-notes/users-search")
async def search_voice_note_users(
    q: str = Query(""),
    cu: dict = Depends(get_user),
):
    """بحث عن مستخدمين/سائقين - متاح للسوبر والأدمن."""
    if cu["role"] not in ("superuser", "admin"):
        raise HTTPException(403, "غير مصرح")
    like = f"%{q}%"
    eff_branch = _branch_filter(cu)
    with get_db() as conn:
        c = conn.cursor()
        # السائقون (يُبحث بالاسم الحقيقي)
        drv_q = """
            SELECT d.user_id AS id, d.name,
                   'سائق' AS role, 'driver' AS type,
                   COALESCE(u.avatar_url,'') AS avatar_url,
                   d.branch
            FROM drivers d
            LEFT JOIN users u ON u.id = d.user_id
            WHERE d.status='active' AND d.user_id IS NOT NULL
              AND d.name LIKE ?
        """
        params = [like]
        if eff_branch:
            drv_q += " AND d.branch = ?"
            params.append(eff_branch)
        drv_q += " ORDER BY d.name LIMIT 40"
        c.execute(drv_q, params)
        drivers = [{"id": r["id"], "name": r["name"], "role": r["role"],
                    "type": r["type"], "avatar": r["avatar_url"],
                    "branch": r["branch"]} for r in c.fetchall()]

        # الأدمنز والمشرفون (يُبحث بالـ username)
        admins = []
        if cu["role"] == "superuser":
            adm_q = """
                SELECT id, username AS name, role,
                       COALESCE(avatar_url,'') AS avatar_url, branch
                FROM users
                WHERE role IN('admin','superuser')
                  AND username LIKE ?
                ORDER BY username LIMIT 20
            """
            c.execute(adm_q, (like,))
            admins = [{"id": r["id"], "name": r["name"], "role": r["role"],
                       "type": "admin", "avatar": r["avatar_url"],
                       "branch": r["branch"]} for r in c.fetchall()]

    return {"users": drivers + admins}


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
        # جيب عداد آخر رحلة منتهية لكل عربية (بالتاريخ)
        c.execute("""SELECT t.car_id, t.end_odometer as max_odo
                     FROM trips t
                     INNER JOIN (
                         SELECT car_id, MAX(end_time) as latest
                         FROM trips
                         WHERE end_odometer IS NOT NULL AND end_time IS NOT NULL
                         GROUP BY car_id
                     ) lx ON t.car_id = lx.car_id AND t.end_time = lx.latest
                     WHERE t.end_odometer IS NOT NULL""")
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

# ══════════════════════════════════════════════════════
# 35. DRIVER KPI — مؤشرات أداء السائقين
# ══════════════════════════════════════════════════════

def _calc_kpi_score(
    trips: int, km: float, fuel_liters: float, fuel_consumption: float,
    emergencies: int, open_requests: int,
    target_km: float = 3000.0,
    oil_liters: float = 0.0, avg_km_all_drivers: float = 0.0,
    fuel_dist: Optional[float] = None,   # المسافة بين أول وآخر تفويلة (بدون آخر تفويلة)
) -> dict:
    """
    حساب نقاط KPI لسائق واحد بالأوزان الجديدة:
      عدد الرحلات         → 10%
      كمية الكيلومترات    → 40%
      استهلاك الوقود      → 40%  (لو مفيش بيانات وقود يُعوَّض من الكم)
      استهلاك الزيوت      → 5%   (لو مفيش بيانات يُعوَّض من الكم)
      الحوادث والطوارئ    → 5%
    """
    # ── 1. عدد الرحلات (10%) ──
    if trips >= 20:
        trips_score = 100.0
    elif trips >= 15:
        trips_score = 85.0
    elif trips >= 10:
        trips_score = 65.0
    elif trips >= 5:
        trips_score = 40.0
    elif trips >= 1:
        trips_score = 20.0
    else:
        trips_score = 0.0

    # ── 2. الكيلومترات (40%) ──
    # المرجع: متوسط كيلومترات كل السائقين في نفس الفترة
    ref_km = avg_km_all_drivers if avg_km_all_drivers > 0 else target_km
    if ref_km > 0:
        km_ratio = km / ref_km
        if km_ratio >= 1.2:
            km_score = 100.0
        elif km_ratio >= 1.0:
            km_score = 90.0
        elif km_ratio >= 0.8:
            km_score = 75.0
        elif km_ratio >= 0.6:
            km_score = 55.0
        elif km_ratio >= 0.4:
            km_score = 35.0
        else:
            km_score = 10.0
    else:
        km_score = 0.0

    # ── 3. استهلاك الوقود (40%) ──
    # المسافة المستخدمة: fuel_dist (بين أول وآخر تفويلة) إذا متاحة،
    # وإلا km الكلية. الوقود بدون آخر تفويلة.
    effective_dist = fuel_dist if (fuel_dist and fuel_dist > 50) else km
    if effective_dist > 50 and fuel_liters > 0:
        consumption = (fuel_liters / effective_dist) * 100  # لتر / 100كم
        if consumption <= 20:
            fuel_score = 100.0
        elif consumption <= 25:
            fuel_score = 90.0
        elif consumption <= 30:
            fuel_score = 75.0
        elif consumption <= 35:
            fuel_score = 55.0
        elif consumption <= 40:
            fuel_score = 35.0
        else:
            fuel_score = 15.0
    else:
        fuel_score = None  # بيانات وقود غير كافية

    # ── 4. استهلاك الزيوت (5%) ──
    # معدل: لتر زيت / 1000 كم — كلما أقل أفضل
    if km > 0 and oil_liters > 0:
        oil_per_1000km = (oil_liters / km) * 1000
        if oil_per_1000km <= 0.5:
            oil_score = 100.0
        elif oil_per_1000km <= 1.0:
            oil_score = 80.0
        elif oil_per_1000km <= 2.0:
            oil_score = 55.0
        elif oil_per_1000km <= 3.0:
            oil_score = 30.0
        else:
            oil_score = 10.0
    else:
        oil_score = None  # مفيش بيانات زيوت

    # ── 5. الحوادث والطوارئ (5%) ──
    if emergencies == 0:
        emg_score = 100.0
    elif emergencies == 1:
        emg_score = 60.0
    elif emergencies == 2:
        emg_score = 20.0
    else:
        emg_score = 0.0

    # ── حساب المجموع الموزون ──
    # الأوزان الأساسية: رحلات 10% | كم 40% | وقود 40% | زيوت 5% | طوارئ 5%
    W_TRIPS = 0.10
    W_KM    = 0.40
    W_FUEL  = 0.40
    W_OIL   = 0.05
    W_EMG   = 0.05

    weighted = trips_score * W_TRIPS + km_score * W_KM + emg_score * W_EMG

    # لو مفيش وقود: وزنه ينتقل للكيلومترات
    if fuel_score is not None:
        weighted += fuel_score * W_FUEL
        avail_w   = W_TRIPS + W_KM + W_FUEL + W_EMG
    else:
        weighted += km_score * W_FUEL   # عوّض الوقود بالكيلومترات
        avail_w   = W_TRIPS + W_KM + W_FUEL + W_EMG

    # لو مفيش زيوت: وزنه ينتقل للكيلومترات
    if oil_score is not None:
        weighted += oil_score * W_OIL
        avail_w  += W_OIL
    else:
        weighted += km_score * W_OIL    # عوّض الزيوت بالكيلومترات
        avail_w  += W_OIL

    total = round(weighted, 1)

    if total >= 85:
        grade, grade_color = "ممتاز",       "#10b981"
    elif total >= 70:
        grade, grade_color = "جيد",         "#3b82f6"
    elif total >= 50:
        grade, grade_color = "مقبول",       "#f97316"
    else:
        grade, grade_color = "يحتاج تحسين", "#ef4444"

    return {
        "total_score":  total,
        "grade":        grade,
        "grade_color":  grade_color,
        "trips_score":  trips_score,
        "km_score":     km_score,
        "fuel_score":   fuel_score,
        "oil_score":    oil_score,
        "emg_score":    emg_score,
        "consumption":  round((fuel_liters / effective_dist) * 100, 1) if (effective_dist > 50 and fuel_liters > 0) else None,
        "oil_per_1000": round((oil_liters / km) * 1000, 2) if km > 0 and oil_liters > 0 else None,
    }


@app.get("/reports/driver-kpi")
async def driver_kpi_report(
    month:     str = Query(None, description="YYYY-MM"),
    driver_id: int = Query(None),
    branch:    str = Query(None),
    cu: dict = Depends(require_admin_or_reporter),
):
    """تقرير مؤشرات أداء السائقين KPI."""
    if not month:
        month = datetime.utcnow().strftime("%Y-%m")

    try:
        y, m = month.split("-")
        month_start = f"{y}-{m}-01"
        next_m = int(m) + 1
        next_y = int(y)
        if next_m > 12:
            next_m = 1
            next_y += 1
        month_end = f"{next_y}-{next_m:02d}-01"
    except Exception:
        raise HTTPException(400, "صيغة الشهر خاطئة — استخدم YYYY-MM")

    eff_branch = _branch_filter(cu) or branch

    with get_db() as conn:
        c = conn.cursor()
        q = "SELECT d.id, d.name, d.branch, d.status FROM drivers d WHERE d.status='active'"
        params: list = []
        if eff_branch:
            q += " AND d.branch=?"
            params.append(eff_branch)
        if driver_id:
            q += " AND d.id=?"
            params.append(driver_id)
        q += " ORDER BY d.name"
        c.execute(q, params)
        drivers = [dict(r) for r in c.fetchall()]

        results  = []
        raw_data = []
        for d in drivers:
            did = d["id"]

            c.execute("""
                SELECT COUNT(*) as cnt,
                       COALESCE(SUM(end_odometer - start_odometer),0) as km,
                       MIN(start_time) as first_trip,
                       MAX(end_time)   as last_trip
                FROM trips
                WHERE driver_id=?
                  AND end_time IS NOT NULL AND end_odometer IS NOT NULL
                  AND date(start_time) >= ? AND date(start_time) < ?
            """, (did, month_start, month_end))
            tr = dict(c.fetchone())
            trips_count = tr["cnt"] or 0
            km_total    = max(0.0, float(tr["km"] or 0))

            # ── وقود — بيانات كل تفويلة على حدة ──
            c.execute("""
                SELECT quantity, odometer_reading, created_at, price
                FROM workshop_records
                WHERE driver_id=? AND type LIKE 'fuel_%'
                  AND date(created_at) >= ? AND date(created_at) < ?
                ORDER BY created_at ASC
            """, (did, month_start, month_end))
            fuel_rows = [dict(r) for r in c.fetchall()]
            # إجمالي الوقود الفعلي للعرض (كل التفويلات)
            fuel_liters_total = sum(float(r.get("quantity") or 0) for r in fuel_rows)
            fuel_cost = sum(float(r.get("price") or 0) for r in fuel_rows)

            # معادلة الاستهلاك الصحيحة: بدون آخر تفويلة
            # المسافة = عداد آخر تفويلة − عداد أول تفويلة
            # الوقود  = مجموع الكميات بدون آخر تفويلة
            if len(fuel_rows) >= 2:
                rows_with_odo = [r for r in fuel_rows if r.get("odometer_reading")]
                if len(rows_with_odo) >= 2:
                    first_odo = float(rows_with_odo[0]["odometer_reading"])
                    last_odo  = float(rows_with_odo[-1]["odometer_reading"])
                    dist_km   = last_odo - first_odo
                    liters_ex_last = sum(
                        float(r.get("quantity") or 0)
                        for r in fuel_rows[:-1]
                    )
                    if dist_km > 0 and liters_ex_last > 0:
                        fuel_liters_calc = liters_ex_last   # للمعدل فقط
                        _fuel_dist       = dist_km
                    else:
                        fuel_liters_calc = fuel_liters_total
                        _fuel_dist       = None
                else:
                    fuel_liters_calc = fuel_liters_total
                    _fuel_dist       = None
            else:
                # صفر أو تفويلة واحدة — مش كافي نحسب معدل
                fuel_liters_calc = 0.0
                _fuel_dist       = None

            # ── زيوت (oil, filter) ──
            c.execute("""
                SELECT COALESCE(SUM(quantity),0) as liters
                FROM workshop_records
                WHERE driver_id=? AND type IN ('oil','filter')
                  AND date(created_at) >= ? AND date(created_at) < ?
            """, (did, month_start, month_end))
            oil_liters = float(c.fetchone()[0] or 0)

            # ── طوارئ ──
            c.execute("""
                SELECT COUNT(*) FROM emergency_reports
                WHERE driver_id=?
                  AND date(created_at) >= ? AND date(created_at) < ?
            """, (did, month_start, month_end))
            emg_count = c.fetchone()[0] or 0

            # ── طلبات مفتوحة ──
            c.execute("""
                SELECT COUNT(*) FROM driver_requests
                WHERE driver_id=? AND status='pending'
            """, (did,))
            open_req = c.fetchone()[0] or 0

            # ── صيانة ──
            c.execute("""
                SELECT COALESCE(SUM(price),0) as mc, COUNT(*) as mcc
                FROM workshop_records
                WHERE driver_id=? AND type NOT LIKE 'fuel_%'
                  AND date(created_at) >= ? AND date(created_at) < ?
            """, (did, month_start, month_end))
            wr = dict(c.fetchone())
            maint_cost  = float(wr["mc"] or 0)
            maint_count = int(wr["mcc"] or 0)

            raw_data.append({
                "driver_id":        did,
                "driver_name":      d["name"],
                "branch":           d["branch"] or "",
                "trips":            trips_count,
                "km":               km_total,
                "fuel_liters":      fuel_liters_total,   # للعرض — كل التفويلات
                "fuel_liters_calc": fuel_liters_calc,    # للمعدل — بدون آخر تفويلة
                "fuel_dist":        _fuel_dist,
                "fuel_cost":        fuel_cost,
                "oil_liters":       oil_liters,
                "emergencies":      emg_count,
                "open_req":         open_req,
                "maint_cost":       maint_cost,
                "maint_count":      maint_count,
                "first_trip":       tr["first_trip"],
                "last_trip":        tr["last_trip"],
            })

        # ── متوسط الكيلومترات لكل السائقين (للمقارنة النسبية) ──
        km_values = [d["km"] for d in raw_data if d["km"] > 0]
        avg_km = (sum(km_values) / len(km_values)) if km_values else 0.0

        # ── حساب الـ KPI لكل سائق ──
        for d in raw_data:
            scores = _calc_kpi_score(
                trips=d["trips"],
                km=d["km"],
                fuel_liters=d["fuel_liters_calc"],   # بدون آخر تفويلة للمعدل
                fuel_consumption=0,
                emergencies=d["emergencies"],
                open_requests=d["open_req"],
                oil_liters=d["oil_liters"],
                avg_km_all_drivers=avg_km,
                fuel_dist=d.get("fuel_dist"),
            )
            consumption = scores.pop("consumption", None)

            results.append({
                "driver_id":     d["driver_id"],
                "driver_name":   d["driver_name"],
                "branch":        d["branch"],
                "month":         month,
                "trips":         d["trips"],
                "km":            round(d["km"], 1),
                "fuel_liters":   round(d["fuel_liters"], 1),      # للعرض
                "fuel_cost":     round(d["fuel_cost"], 2),
                "oil_liters":    round(d["oil_liters"], 2),
                "consumption":   consumption,
                "oil_per_1000":  scores.pop("oil_per_1000", None),
                "emergencies":   d["emergencies"],
                "open_requests": d["open_req"],
                "maint_cost":    round(d["maint_cost"], 2),
                "maint_count":   d["maint_count"],
                "first_trip":    d["first_trip"],
                "last_trip":     d["last_trip"],
                "avg_km_peers":  round(avg_km, 1),
                **scores,
            })

    results.sort(key=lambda x: x["total_score"], reverse=True)
    for i, r in enumerate(results, 1):
        r["rank"] = i

    return {
        "month":         month,
        "branch":        eff_branch or "كل الفروع",
        "generated_at":  datetime.utcnow().isoformat() + "Z",
        "total_drivers": len(results),
        "avg_km":        round(avg_km, 1),
        "drivers":       results,
    }


# ══════════════════════════════════════════════════════
# RENTAL EQUIPMENT — المعدات المستأجرة
# ══════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════
# RENTAL EQUIPMENT — معدات مستأجرة (من الداتابيز)
# ══════════════════════════════════════════════════════════════════

RENTAL_COLS = ["branch","equipment_name","code","rental_source","work_days",
               "daily_rate","monthly_rate","fuel_liters","hours_start","hours_end",
               "hours_diff","utilization_pct","consumption_rate","project","notes","month","sector"]

def _validate_rental_row(row: dict, idx: int) -> dict:
    """التحقق من صف معدة مستأجرة"""
    errors = []
    name = str(row.get("equipment_name") or row.get("اسم المعدة") or row.get("نوع المعدة") or "").strip()
    if not name:
        errors.append("اسم المعدة مطلوب")
    branch = str(row.get("branch") or row.get("الفرع") or "").strip()
    if not branch:
        errors.append("الفرع مطلوب")

    def _s(key, ar_key):
        return str(row.get(key) or row.get(ar_key) or "").strip()

    return {
        "row_index":       idx,
        "valid":           len(errors) == 0,
        "errors":          errors,
        "branch":          branch,
        "equipment_name":  name,
        "code":            _s("code","الكود"),
        "rental_source":   _s("rental_source","جهة الايجار"),
        "work_days":       _s("work_days","ايام العمل"),
        "daily_rate":      _s("daily_rate","القيمة اليومية"),
        "monthly_rate":    _s("monthly_rate","القيمة الشهرية"),
        "fuel_liters":     _s("fuel_liters","استهلاك السولار"),
        "hours_start":     _s("hours_start","عداد اول الشهر"),
        "hours_end":       _s("hours_end","عداد اخر الشهر"),
        "hours_diff":      _s("hours_diff","الفرق"),
        "utilization_pct": _s("utilization_pct","نسبة الاستخدام"),
        "consumption_rate":_s("consumption_rate","معدل الاستهلاك"),
        "project":         _s("project","المشروع"),
        "notes":           _s("notes","ملاحظات"),
        "month":           _s("month","الشهر") or datetime.utcnow().strftime("%Y-%m"),
        "sector":          _s("sector","القطاع"),
    }


@app.get("/rental-equipment")
async def get_rental_equipment(
    month:  str = Query(None),
    branch: str = Query(None),
    source: str = Query(None),
    search: str = Query(None),
    cu: dict = Depends(require_admin_or_reporter),
):
    """قائمة المعدات المستأجرة من الداتابيز"""
    eff_branch = _branch_filter(cu) or branch
    # تطبيع الشهر: 2026-05-01 → 2026-05
    if month:
        month = month.strip()[:7]

    with get_db() as conn:
        q = "SELECT * FROM rental_equipment WHERE 1=1"
        params = []
        if eff_branch:
            q += " AND branch=?"; params.append(eff_branch)
        if source:
            q += " AND rental_source LIKE ?"; params.append(f"%{source}%")
        if month:
            q += " AND (month=? OR month LIKE ?)"; params.extend([month, month + "-%"])
        if search:
            q += " AND (equipment_name LIKE ? OR code LIKE ?)"; s=f"%{search}%"; params+=[s,s]
        q += " ORDER BY branch, equipment_name"
        try:
            records = [dict(r) for r in conn.execute(q, params).fetchall()]
        except Exception as e:
            log.error(f"rental_equipment query error: {e}")
            raise HTTPException(500, f"خطأ في جلب البيانات: {e}")

    def _to_f(v):
        try: return float(str(v).replace(',',''))
        except: return 0.0

    total_monthly = sum(_to_f(r.get('monthly_rate')) for r in records)
    total_fuel    = sum(_to_f(r.get('fuel_liters'))  for r in records)
    total_days    = sum(_to_f(r.get('work_days'))    for r in records)

    return {
        "month":              month or "كل الشهور",
        "branch":             eff_branch or "كل الفروع",
        "total":              len(records),
        "total_monthly_cost": round(total_monthly, 2),
        "total_fuel_liters":  round(total_fuel, 1),
        "total_work_days":    round(total_days, 0),
        "records":            records,
    }


@app.get("/rental-equipment/{rec_id}")
async def get_rental_record(rec_id: int, cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        row = conn.execute("SELECT * FROM rental_equipment WHERE id=?", (rec_id,)).fetchone()
    if not row: raise HTTPException(404, "السجل غير موجود")
    return dict(row)


class RentalCreate(BaseModel):
    branch:          str
    equipment_name:  str
    code:            str = ""
    rental_source:   str = ""
    work_days:       str = ""
    daily_rate:      str = ""
    monthly_rate:    str = ""
    fuel_liters:     str = ""
    hours_start:     str = ""
    hours_end:       str = ""
    hours_diff:      str = ""
    utilization_pct: str = ""
    consumption_rate:str = ""
    project:         str = ""
    notes:           str = ""
    month:           str = ""
    sector:          str = ""
    start_date:      str = ""


@app.post("/rental-equipment")
async def create_rental(body: RentalCreate, cu: dict = Depends(require_admin)):
    if not body.equipment_name.strip(): raise HTTPException(400, "اسم المعدة مطلوب")
    if not body.branch.strip(): raise HTTPException(400, "الفرع مطلوب")
    # تطبيع الشهر: 2026-05-01 → 2026-05 (لا نعدّل body مباشرةً لتفادي مشكلة Pydantic v2)
    month_val = body.month.strip()[:7] if body.month and len(body.month) > 7 else body.month
    start_date_val = body.start_date.strip() if body.start_date and body.start_date.strip() else datetime.utcnow().strftime("%Y-%m-%d")
    try:
        with get_db() as conn:
            conn.execute("""INSERT INTO rental_equipment
                (branch,equipment_name,code,rental_source,work_days,daily_rate,monthly_rate,
                 fuel_liters,hours_start,hours_end,hours_diff,utilization_pct,
                 consumption_rate,project,notes,month,sector,start_date)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (body.branch, body.equipment_name, body.code, body.rental_source,
                 body.work_days, body.daily_rate, body.monthly_rate, body.fuel_liters,
                 body.hours_start, body.hours_end, body.hours_diff, body.utilization_pct,
                 body.consumption_rate, body.project, body.notes, month_val, body.sector,
                 start_date_val))
            rec_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        return {"id": rec_id, "message": "تم الإضافة"}
    except Exception as e:
        log.error(f"create_rental error: {e}")
        raise HTTPException(500, f"خطأ في الحفظ: {e}")


@app.put("/rental-equipment/{rec_id}")
async def update_rental(rec_id: int, body: RentalCreate, cu: dict = Depends(require_admin)):
    month_val = body.month.strip()[:7] if body.month and len(body.month) > 7 else body.month
    try:
        with get_db() as conn:
            if not conn.execute("SELECT id FROM rental_equipment WHERE id=?", (rec_id,)).fetchone():
                raise HTTPException(404, "السجل غير موجود")
            conn.execute("""UPDATE rental_equipment SET
                branch=?,equipment_name=?,code=?,rental_source=?,work_days=?,daily_rate=?,
                monthly_rate=?,fuel_liters=?,hours_start=?,hours_end=?,hours_diff=?,
                utilization_pct=?,consumption_rate=?,project=?,notes=?,month=?,sector=?,start_date=?
                WHERE id=?""",
                (body.branch, body.equipment_name, body.code, body.rental_source,
                 body.work_days, body.daily_rate, body.monthly_rate, body.fuel_liters,
                 body.hours_start, body.hours_end, body.hours_diff, body.utilization_pct,
                 body.consumption_rate, body.project, body.notes, month_val, body.sector,
                 body.start_date.strip() if body.start_date and body.start_date.strip() else datetime.utcnow().strftime("%Y-%m-%d"),
                 rec_id))
        return {"message": "تم التعديل"}
    except HTTPException:
        raise
    except Exception as e:
        log.error(f"update_rental error: {e}")
        raise HTTPException(500, f"خطأ في التعديل: {e}")


@app.delete("/rental-equipment/{rec_id}")
async def delete_rental(rec_id: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        conn.execute("DELETE FROM rental_equipment WHERE id=?", (rec_id,))
    return {"message": "تم الحذف"}


@app.delete("/rental-equipment/month/{month}")
async def delete_rental_month(month: str, cu: dict = Depends(require_superuser)):
    """حذف كل بيانات شهر كامل — سوبر يوزر فقط"""
    month = month.strip()[:7]
    with get_db() as conn:
        result = conn.execute("DELETE FROM rental_equipment WHERE month=?", (month,))
        deleted = result.rowcount
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "delete_rental_month", f"حذف {deleted} سجل للشهر {month}")
    return {"message": f"تم حذف {deleted} سجل للشهر {month}", "deleted": deleted}


@app.post("/rental-equipment/import/preview")
async def rental_import_preview(file: UploadFile = File(...), cu: dict = Depends(require_admin)):
    content = await file.read()
    if len(content) > 5*1024*1024: raise HTTPException(400, "الملف كبير جداً")
    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows: raise HTTPException(400, "الملف فارغ")
    rows = [_validate_rental_row(r, i+1) for i, r in enumerate(raw_rows)]
    valid = sum(1 for r in rows if r["valid"])
    return {"import_type":"rental","rows":rows,
            "summary":{"total":len(rows),"valid":valid,"invalid":len(rows)-valid}}


@app.post("/rental-equipment/import/confirm")
async def rental_import_confirm(
    file: UploadFile = File(...),
    skip_errors: bool = False,
    replace_month: bool = False,
    cu: dict = Depends(require_admin),
):
    content = await file.read()
    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows: raise HTTPException(400, "الملف فارغ")
    rows = [_validate_rental_row(r, i+1) for i, r in enumerate(raw_rows)]
    invalid = [r for r in rows if not r["valid"]]
    if invalid and not skip_errors:
        raise HTTPException(422, {"message":f"{len(invalid)} صف بأخطاء","invalid_rows":invalid})

    inserted, skipped = 0, []
    with get_db() as conn:
        # لو replace_month: احذف بيانات نفس الشهر والفرع أولاً
        if replace_month:
            months = {r["month"] for r in rows if r["valid"] and r["month"]}
            for m in months:
                conn.execute("DELETE FROM rental_equipment WHERE month=?", (m,))

        for row in [r for r in rows if r["valid"]]:
            try:
                conn.execute("""INSERT INTO rental_equipment
                    (branch,equipment_name,code,rental_source,work_days,daily_rate,monthly_rate,
                     fuel_liters,hours_start,hours_end,hours_diff,utilization_pct,
                     consumption_rate,project,notes,month,sector)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (row["branch"],row["equipment_name"],row["code"],row["rental_source"],
                     row["work_days"],row["daily_rate"],row["monthly_rate"],row["fuel_liters"],
                     row["hours_start"],row["hours_end"],row["hours_diff"],row["utilization_pct"],
                     row["consumption_rate"],row["project"],row["notes"],row["month"],row["sector"]))
                inserted += 1
            except Exception as e:
                row["errors"]=[str(e)]; skipped.append(row)

    write_audit_log(cu["user_id"],cu["username"],cu["role"],
                    "bulk_import_rental",f"استيراد {inserted} معدة مستأجرة")
    return {"inserted":inserted,"skipped":len(skipped),"skipped_items":skipped}


@app.get("/rental-equipment/import/template")
async def rental_import_template(cu: dict = Depends(require_admin)):
    from fastapi.responses import Response
    cols = ["branch","equipment_name","code","rental_source","work_days","daily_rate",
            "monthly_rate","fuel_liters","hours_start","hours_end","hours_diff",
            "utilization_pct","consumption_rate","project","notes","month","sector"]
    samples = [
        ["حلوان","مان لفت 926","","وحدة الرفع","24","","84000","0","0","0","عاطل","","","مونوريل اكتوبر","","2026-03","قطاع القاهرة الكبرى"],
        ["القاهرة","حفار على كاتينة","12/11/171","ادارة تحريك التربة","6","3630","21780","400","9937","9961","24","14.3","","","","2026-03","قطاع القاهرة الكبرى"],
        ["صيانة القصور","اسكيد لودر CAT","12/5/340","إدارة معدات تحريك التربة","24","","48000","80","7347","7366","19","11.3","4.2","المركز الطبي","","2026-03","قطاع القاهرة الكبرى"],
    ]
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(cols)
    for s in samples: writer.writerow(s)
    return Response(content=buf.getvalue().encode("utf-8-sig"), media_type="text/csv",
                    headers={"Content-Disposition":"attachment; filename=rental_template.csv"})



# ══════════════════════════════════════════════════════════════════
# EQUIPMENT — معدات (جدول منفصل - بدون لوحة، الكود هو المعرف)
# ══════════════════════════════════════════════════════════════════

EQUIPMENT_TYPES = [
    "معدات تحريك تربة",
    "معدات رفع",
    "معدات ثابتة",
    "معدات نقل ثقيل",
    "معدات نقل",
    "وسائل انتقال",
    "معدات إنتاجية",
    "معدات مساعدة",
    "معدات وآلات ورش",
    "أخرى",
]

def _validate_equipment_row(row: dict, idx: int, admin_branch: Optional[str]) -> dict:
    """التحقق من صف معدة واحد من ملف CSV"""
    errors = []
    code = str(row.get("car_code") or row.get("الكود") or "").strip()
    name = str(row.get("equipment_name") or row.get("اسم المعدة") or row.get("نوع المعدة") or "").strip()
    if not code:
        errors.append("الكود مطلوب")
    if not name:
        errors.append("اسم المعدة مطلوب")
    branch = admin_branch or str(row.get("branch") or row.get("الفرع") or "").strip()
    eq_type = str(row.get("equipment_type") or row.get("نوع التصنيف") or "معدات تحريك تربة").strip()
    if eq_type not in EQUIPMENT_TYPES:
        eq_type = "معدات تحريك تربة"
    status = str(row.get("status") or row.get("الحالة") or "active").strip()
    if status not in ("active","inactive","maintenance"):
        status = "active"
    return {
        "row_index":      idx,
        "valid":          len(errors) == 0,
        "errors":         errors,
        "car_code":       code,
        "equipment_name": name,
        "brand":          str(row.get("brand") or row.get("الماركة") or "").strip(),
        "model":          str(row.get("model") or row.get("الموديل") or "").strip(),
        "year":           str(row.get("year") or row.get("سنة الصنع") or row.get("Y.O.M") or "").strip(),
        "chassis":        str(row.get("chassis") or row.get("رقم الشاسيه") or "").strip(),
        "engine_number":  str(row.get("engine_number") or row.get("رقم المحرك") or "").strip(),
        "serial_no":      str(row.get("serial_no") or row.get("الرقم التسلسلي") or row.get("SERIAL NO.") or "").strip(),
        "capacity":       str(row.get("capacity") or row.get("الطاقة") or row.get("CAPACITY") or "").strip(),
        "equipment_type": eq_type,
        "branch":         branch,
        "sector":         str(row.get("sector") or row.get("القطاع") or "").strip(),
        "project":        str(row.get("project") or row.get("المشروع") or row.get("ARABIC LOCATION") or "").strip(),
        "license_expiry": _normalize_date(row.get("license_expiry") or row.get("انتهاء الرخصة") or ""),
        "status":         status,
        "notes":          str(row.get("notes") or row.get("ملاحظات") or "").strip(),
    }


@app.get("/equipment")
async def list_equipment(
    branch:    str = Query(None),
    eq_type:   str = Query(None),
    status:    str = Query(None),
    search:    str = Query(None),
    cu: dict = Depends(require_admin_or_reporter),
):
    """قائمة المعدات"""
    eff_branch = _branch_filter(cu) or branch
    with get_db() as conn:
        q = "SELECT * FROM equipment WHERE 1=1"
        params = []
        if eff_branch:
            q += " AND branch=?"; params.append(eff_branch)
        if eq_type:
            q += " AND equipment_type=?"; params.append(eq_type)
        if status:
            q += " AND status=?"; params.append(status)
        if search:
            q += " AND (equipment_name LIKE ? OR car_code LIKE ? OR brand LIKE ?)"
            s = f"%{search}%"; params += [s, s, s]
        q += " ORDER BY branch, equipment_name"
        rows = [dict(r) for r in conn.execute(q, params).fetchall()]

    summary = {
        "total":       len(rows),
        "active":      sum(1 for r in rows if r["status"] == "active"),
        "inactive":    sum(1 for r in rows if r["status"] == "inactive"),
        "maintenance": sum(1 for r in rows if r["status"] == "maintenance"),
        "by_type": {}
    }
    for t in EQUIPMENT_TYPES:
        summary["by_type"][t] = sum(1 for r in rows if r["equipment_type"] == t)

    return {"equipment": rows, "summary": summary}


@app.get("/equipment/{eq_id}")
async def get_equipment(eq_id: int, cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        row = conn.execute("SELECT * FROM equipment WHERE id=?", (eq_id,)).fetchone()
    if not row:
        raise HTTPException(404, "المعدة غير موجودة")
    return dict(row)


class EquipmentCreate(BaseModel):
    car_code:       str
    equipment_name: str
    brand:          str = ""
    model:          str = ""
    year:           str = ""
    chassis:        str = ""
    engine_number:  str = ""
    equipment_type: str = "معدات تحريك تربة"
    branch:         str = ""
    sector:         str = ""
    project:        str = ""
    license_expiry: str = ""
    status:         str = "active"
    notes:          str = ""


@app.post("/equipment")
async def create_equipment(body: EquipmentCreate, cu: dict = Depends(require_admin)):
    if not body.car_code.strip():
        raise HTTPException(400, "الكود مطلوب")
    if not body.equipment_name.strip():
        raise HTTPException(400, "اسم المعدة مطلوب")
    try:
        with get_db() as conn:
            conn.execute("""
                INSERT INTO equipment
                (car_code,equipment_name,brand,model,year,chassis,engine_number,
                 equipment_type,branch,sector,project,license_expiry,status,notes)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (body.car_code.strip(), body.equipment_name.strip(),
                  body.brand, body.model, body.year, body.chassis, body.engine_number,
                  body.equipment_type, body.branch, body.sector, body.project,
                  body.license_expiry, body.status, body.notes))
            eq_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        return {"id": eq_id, "message": "تم إضافة المعدة بنجاح"}
    except Exception as e:
        if "UNIQUE" in str(e):
            raise HTTPException(409, f"الكود '{body.car_code}' موجود بالفعل")
        raise HTTPException(500, str(e))


@app.put("/equipment/{eq_id}")
async def update_equipment(eq_id: int, body: EquipmentCreate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        row = conn.execute("SELECT id FROM equipment WHERE id=?", (eq_id,)).fetchone()
        if not row:
            raise HTTPException(404, "المعدة غير موجودة")
        conn.execute("""
            UPDATE equipment SET
                car_code=?, equipment_name=?, brand=?, model=?, year=?, chassis=?,
                engine_number=?, equipment_type=?, branch=?, sector=?, project=?,
                license_expiry=?, status=?, notes=?
            WHERE id=?
        """, (body.car_code.strip(), body.equipment_name.strip(),
              body.brand, body.model, body.year, body.chassis, body.engine_number,
              body.equipment_type, body.branch, body.sector, body.project,
              body.license_expiry, body.status, body.notes, eq_id))
    return {"message": "تم التعديل بنجاح"}


@app.delete("/equipment/{eq_id}")
async def delete_equipment(eq_id: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        conn.execute("DELETE FROM equipment WHERE id=?", (eq_id,))
    return {"message": "تم الحذف"}


# ── Equipment Import ──
@app.post("/equipment/import/preview")
async def equipment_import_preview(
    file: UploadFile = File(...),
    cu: dict = Depends(require_admin),
):
    content = await file.read()
    if len(content) > 5 * 1024 * 1024:
        raise HTTPException(400, "حجم الملف يتجاوز 5 MB")
    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows:
        raise HTTPException(400, "الملف فارغ")
    admin_branch = _branch_filter(cu)
    rows = [_validate_equipment_row(r, i+1, admin_branch) for i, r in enumerate(raw_rows)]
    # تحقق من تكرار الكود
    with get_db() as conn:
        existing_codes = {r[0] for r in conn.execute("SELECT car_code FROM equipment").fetchall()}
    for row in rows:
        if row["valid"] and row["car_code"] in existing_codes:
            row["duplicate"] = True
        else:
            row["duplicate"] = False
    valid = sum(1 for r in rows if r["valid"])
    dups  = sum(1 for r in rows if r.get("duplicate"))
    return {
        "import_type": "equipment",
        "rows": rows,
        "summary": {"total": len(rows), "valid": valid,
                    "invalid": len(rows)-valid, "duplicates": dups},
    }


@app.post("/equipment/import/confirm")
async def equipment_import_confirm(
    file: UploadFile = File(...),
    skip_errors: bool = False,
    update_existing: bool = False,
    cu: dict = Depends(require_admin),
):
    content = await file.read()
    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows:
        raise HTTPException(400, "الملف فارغ")
    admin_branch = _branch_filter(cu)
    rows = [_validate_equipment_row(r, i+1, admin_branch) for i, r in enumerate(raw_rows)]
    invalid = [r for r in rows if not r["valid"]]
    if invalid and not skip_errors:
        raise HTTPException(422, {"message": f"{len(invalid)} صف بأخطاء", "invalid_rows": invalid})

    inserted, updated, skipped = 0, 0, []
    with get_db() as conn:
        for row in [r for r in rows if r["valid"]]:
            try:
                if update_existing:
                    conn.execute("""
                        INSERT INTO equipment
                        (car_code,equipment_name,brand,model,year,chassis,engine_number,
                         serial_no,capacity,equipment_type,branch,sector,project,
                         license_expiry,status,notes)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                        ON CONFLICT(car_code) DO UPDATE SET
                            equipment_name=excluded.equipment_name,
                            brand=excluded.brand, model=excluded.model,
                            year=excluded.year, chassis=excluded.chassis,
                            engine_number=excluded.engine_number,
                            serial_no=excluded.serial_no,
                            capacity=excluded.capacity,
                            equipment_type=excluded.equipment_type,
                            branch=excluded.branch, sector=excluded.sector,
                            project=excluded.project, status=excluded.status,
                            notes=excluded.notes
                    """, (row["car_code"], row["equipment_name"], row["brand"],
                          row["model"], row["year"], row["chassis"], row["engine_number"],
                          row["serial_no"], row["capacity"],
                          row["equipment_type"], row["branch"], row["sector"],
                          row["project"], row["license_expiry"], row["status"], row["notes"]))
                    updated += 1
                else:
                    conn.execute("""
                        INSERT OR IGNORE INTO equipment
                        (car_code,equipment_name,brand,model,year,chassis,engine_number,
                         serial_no,capacity,equipment_type,branch,sector,project,
                         license_expiry,status,notes)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (row["car_code"], row["equipment_name"], row["brand"],
                          row["model"], row["year"], row["chassis"], row["engine_number"],
                          row["serial_no"], row["capacity"],
                          row["equipment_type"], row["branch"], row["sector"],
                          row["project"], row["license_expiry"], row["status"], row["notes"]))
                    inserted += 1
            except Exception as e:
                row["errors"] = [str(e)]; skipped.append(row)

    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "bulk_import_equipment", f"استيراد {inserted+updated} معدة")
    return {
        "inserted": inserted, "updated": updated,
        "skipped": len(skipped), "skipped_items": skipped,
    }


@app.get("/equipment/import/template")
async def equipment_import_template(cu: dict = Depends(require_admin)):
    """تحميل قالب CSV للمعدات"""
    from fastapi.responses import Response
    cols = ["car_code","equipment_name","brand","model","year","chassis",
            "engine_number","equipment_type","branch","sector","project","status","notes"]
    samples = [
        ["12/5/347","اسكيد لودر","CATERPILLAR","246C","2015","CH123456","",
         "معدات تحريك تربة","إدارة صيانة القصور","قطاع القاهرة الكبرى","","active",""],
        ["12/5/926","مان لفت","MANITOU","180 ATJ","2018","","",
         "معدات رفع","حلوان","قطاع القاهرة الكبرى","","active",""],
        ["27/2/2333","مولد كهرباء 50 ك","FG WILSON","P50","2016","","",
         "معدات ثابتة","مدينة نصر","قطاع القاهرة الكبرى","","active",""],
    ]
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(cols)
    for s in samples:
        writer.writerow(s)
    return Response(
        content=buf.getvalue().encode("utf-8-sig"),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=equipment_template.csv"}
    )

# ══════════════════════════════════════════════════════════════════
# EQUIPMENT OPERATORS — مشغلو المعدات (نظام ورديات)
# ══════════════════════════════════════════════════════════════════

def _ensure_operator_tables(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS equipment_operators (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        name          TEXT NOT NULL,
        phone         TEXT DEFAULT '',
        national_id   TEXT DEFAULT '',
        birth_date    TEXT DEFAULT '',
        branch        TEXT DEFAULT '',
        sector        TEXT DEFAULT '',
        fixed_number  TEXT DEFAULT '',
        license_type  TEXT DEFAULT '',
        license_expiry TEXT DEFAULT '',
        status        TEXT DEFAULT 'active',
        user_id       INTEGER UNIQUE,
        notes         TEXT DEFAULT '',
        created_at    TEXT DEFAULT (datetime('now')),
        FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
    )""")
    conn.execute("""CREATE TABLE IF NOT EXISTS operator_shifts (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        operator_id     INTEGER NOT NULL,
        equipment_id    TEXT DEFAULT '',
        equipment_name  TEXT DEFAULT '',
        shift_type      TEXT DEFAULT 'day' CHECK(shift_type IN ('day','night','extended')),
        start_time      TEXT,
        end_time        TEXT,
        start_hours     REAL DEFAULT 0,
        end_hours       REAL DEFAULT 0,
        fuel_liters     REAL DEFAULT 0,
        project         TEXT DEFAULT '',
        branch          TEXT DEFAULT '',
        notes           TEXT DEFAULT '',
        status          TEXT DEFAULT 'active' CHECK(status IN ('active','ended')),
        created_at      TEXT DEFAULT (datetime('now')),
        FOREIGN KEY (operator_id) REFERENCES equipment_operators(id) ON DELETE CASCADE
    )""")
    try: conn.execute("CREATE INDEX IF NOT EXISTS idx_op_shifts_op ON operator_shifts(operator_id)")
    except: pass
    try: conn.execute("CREATE INDEX IF NOT EXISTS idx_op_shifts_status ON operator_shifts(status)")
    except: pass
    # Migration: أضف user_id لو مش موجود
    try: conn.execute("ALTER TABLE equipment_operators ADD COLUMN user_id INTEGER DEFAULT NULL")
    except: pass
    # Migration: أضف حقول المعدة الحالية
    try: conn.execute("ALTER TABLE equipment_operators ADD COLUMN current_equipment_id TEXT DEFAULT ''")
    except: pass
    try: conn.execute("ALTER TABLE equipment_operators ADD COLUMN current_equipment_name TEXT DEFAULT ''")
    except: pass


class OperatorCreate(BaseModel):
    name:          str
    phone:         str = ""
    national_id:   str = ""
    birth_date:    str = ""
    branch:        str = ""
    sector:        str = ""
    fixed_number:  str = ""
    license_type:  str = ""
    license_expiry:str = ""
    status:        str = "active"
    notes:         str = ""
    username:      str = ""   # اختياري — لو موجود يُنشئ حساب login
    password:      str = ""   # اختياري


class ShiftStart(BaseModel):
    operator_id:    int
    equipment_id:   str = ""
    equipment_name: str = ""
    shift_type:     str = "day"
    start_hours:    float = 0
    fuel_liters:    float = 0
    project:        str = ""
    branch:         str = ""
    notes:          str = ""


class ShiftEnd(BaseModel):
    end_hours:  float = 0
    fuel_liters:float = 0
    notes:      str = ""


@app.on_event("startup")
async def _migrate_operators():
    with get_db() as conn:
        _ensure_operator_tables(conn)


@app.get("/operators/me")
async def get_my_operator_profile(cu: dict = Depends(get_user)):
    """المشغل يجيب بياناته الخاصة بدون صلاحيات أدمن"""
    if cu["role"] != "operator":
        raise HTTPException(403, "هذا المسار للمشغلين فقط")
    with get_db() as conn:
        _ensure_operator_tables(conn)
        row = conn.execute(
            "SELECT * FROM equipment_operators WHERE user_id=?", (cu["user_id"],)
        ).fetchone()
    if not row:
        raise HTTPException(404, "لم يتم ربط حسابك بسجل مشغل — تواصل مع الإدارة")
    return dict(row)


@app.get("/operators")
async def list_operators(
    branch: str = Query(None),
    status: str = Query(None),
    search: str = Query(None),
    cu: dict = Depends(require_admin_or_reporter),
):
    with get_db() as conn:
        _ensure_operator_tables(conn)
        q = "SELECT * FROM equipment_operators WHERE 1=1"
        params = []
        eff = _branch_filter(cu) or branch
        if eff: q += " AND branch=?"; params.append(eff)
        if status: q += " AND status=?"; params.append(status)
        if search:
            s = f"%{search}%"
            q += " AND (name LIKE ? OR phone LIKE ? OR fixed_number LIKE ?)"
            params += [s, s, s]
        q += " ORDER BY name"
        rows = [dict(r) for r in conn.execute(q, params).fetchall()]
    return {"total": len(rows), "operators": rows}


@app.post("/operators")
async def create_operator(body: OperatorCreate, cu: dict = Depends(require_admin)):
    if not body.name.strip(): raise HTTPException(400, "الاسم مطلوب")
    with get_db() as conn:
        _ensure_operator_tables(conn)
        # إنشاء حساب login لو username وpassword موجودين
        user_id = None
        if body.username.strip() and body.password.strip():
            try:
                conn.execute(
                    "INSERT INTO users(username,password,role) VALUES(?,?,?)",
                    (body.username.strip(), _hash(body.password.strip()), "operator")
                )
                user_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            except Exception as e:
                raise HTTPException(400, f"اسم المستخدم موجود مسبقاً: {body.username}")

        conn.execute("""INSERT INTO equipment_operators
            (name,phone,national_id,birth_date,branch,sector,fixed_number,
             license_type,license_expiry,status,notes,user_id)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
            (body.name.strip(), body.phone, body.national_id, body.birth_date,
             body.branch, body.sector, body.fixed_number, body.license_type,
             body.license_expiry, body.status, body.notes, user_id))
        oid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    write_audit_log(cu["user_id"], cu["username"], cu["role"], "create_operator", f"مشغل: {body.name}")
    return {"id": oid, "message": "تم الإضافة", "user_id": user_id}


@app.put("/operators/{oid}")
async def update_operator(oid: int, body: OperatorCreate, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        _ensure_operator_tables(conn)
        if not conn.execute("SELECT id FROM equipment_operators WHERE id=?", (oid,)).fetchone():
            raise HTTPException(404, "المشغل غير موجود")
        conn.execute("""UPDATE equipment_operators SET
            name=?,phone=?,national_id=?,birth_date=?,branch=?,sector=?,fixed_number=?,
            license_type=?,license_expiry=?,status=?,notes=? WHERE id=?""",
            (body.name.strip(), body.phone, body.national_id, body.birth_date,
             body.branch, body.sector, body.fixed_number, body.license_type,
             body.license_expiry, body.status, body.notes, oid))
    return {"message": "تم التعديل"}



@app.put("/operators/{oid}/link-user")
async def link_operator_user(oid: int, body: dict, cu: dict = Depends(require_admin)):
    """
    ربط مشغل بـ user_id موجود أو إنشاء حساب جديد وربطه.
    body: { user_id: int }  — ربط بـ user موجود
    أو: { username: str, password: str }  — إنشاء حساب جديد
    """
    with get_db() as conn:
        _ensure_operator_tables(conn)
        c = conn.cursor()
        c.execute("SELECT id, name FROM equipment_operators WHERE id=?", (oid,))
        op = c.fetchone()
        if not op:
            raise HTTPException(404, "المشغل غير موجود")

        user_id = body.get("user_id")

        if user_id:
            # ربط بـ user موجود مباشرة
            c.execute("SELECT id, role FROM users WHERE id=?", (user_id,))
            user = c.fetchone()
            if not user:
                raise HTTPException(404, "المستخدم غير موجود")
            # تحديث الـ role ليكون operator
            c.execute("UPDATE users SET role='operator' WHERE id=?", (user_id,))
            c.execute("UPDATE equipment_operators SET user_id=? WHERE id=?", (user_id, oid))
        else:
            # إنشاء حساب جديد
            username = (body.get("username") or "").strip()
            password = (body.get("password") or "").strip()
            if not username or not password:
                raise HTTPException(400, "user_id أو (username + password) مطلوب")
            c.execute("SELECT id FROM users WHERE username=?", (username,))
            if c.fetchone():
                raise HTTPException(400, f"اسم المستخدم '{username}' موجود مسبقاً")
            c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                      (username, _hash(password), "operator"))
            user_id = c.lastrowid
            c.execute("UPDATE equipment_operators SET user_id=? WHERE id=?", (user_id, oid))

    log_event("operator_linked", operator_id=oid, user_id=user_id, admin=cu["username"])
    return {"ok": True, "operator_id": oid, "user_id": user_id}


@app.put("/operators/{oid}/assign-equipment")
async def assign_equipment_to_operator(oid: int, body: dict, cu: dict = Depends(require_admin)):
    """
    تعيين معدة حالية للمشغل (بدون بدء وردية).
    body: { equipment_id: str, equipment_name: str }
    """
    with get_db() as conn:
        _ensure_operator_tables(conn)
        op = conn.execute("SELECT id, name FROM equipment_operators WHERE id=?", (oid,)).fetchone()
        if not op:
            raise HTTPException(404, "المشغل غير موجود")
        eq_id   = (body.get("equipment_id") or "").strip()
        eq_name = (body.get("equipment_name") or "").strip()
        if not eq_name:
            raise HTTPException(400, "اسم المعدة مطلوب")
        conn.execute(
            "UPDATE equipment_operators SET current_equipment_id=?, current_equipment_name=? WHERE id=?",
            (eq_id, eq_name, oid)
        )
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "assign_equipment", f"تعيين {eq_name} ({eq_id}) للمشغل {op['name']}")
    return {"ok": True, "equipment_id": eq_id, "equipment_name": eq_name}

@app.delete("/operators/{oid}")
async def delete_operator(oid: int, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        conn.execute("DELETE FROM equipment_operators WHERE id=?", (oid,))
    return {"message": "تم الحذف"}


# ── Shifts ──
@app.get("/operators/{oid}/shifts")
async def get_operator_shifts(
    oid: int,
    limit: int = Query(20),
    cu: dict = Depends(require_admin_or_reporter),
):
    with get_db() as conn:
        _ensure_operator_tables(conn)
        rows = [dict(r) for r in conn.execute(
            "SELECT * FROM operator_shifts WHERE operator_id=? ORDER BY start_time DESC LIMIT ?",
            (oid, limit)).fetchall()]
    return {"shifts": rows}


@app.get("/operators/{oid}/active-shift")
async def get_active_shift(oid: int, cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        _ensure_operator_tables(conn)
        row = conn.execute(
            "SELECT * FROM operator_shifts WHERE operator_id=? AND status='active' ORDER BY start_time DESC LIMIT 1",
            (oid,)).fetchone()
    return dict(row) if row else None


@app.post("/operators/shifts/start")
async def start_shift(body: ShiftStart, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        _ensure_operator_tables(conn)
        # تحقق من عدم وجود وردية نشطة
        active = conn.execute(
            "SELECT id FROM operator_shifts WHERE operator_id=? AND status='active'",
            (body.operator_id,)).fetchone()
        if active:
            raise HTTPException(400, "يوجد وردية جارية بالفعل — أنهِ الوردية الحالية أولاً")
        conn.execute("""INSERT INTO operator_shifts
            (operator_id,equipment_id,equipment_name,shift_type,start_time,
             start_hours,fuel_liters,project,branch,notes,status)
            VALUES(?,?,?,?,?,?,?,?,?,?,'active')""",
            (body.operator_id, body.equipment_id, body.equipment_name, body.shift_type,
             datetime.utcnow().isoformat(), body.start_hours, body.fuel_liters,
             body.project, body.branch, body.notes))
        sid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    return {"id": sid, "message": "بدأت الوردية"}


@app.post("/operators/shifts/{sid}/end")
async def end_shift(sid: int, body: ShiftEnd, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        _ensure_operator_tables(conn)
        shift = conn.execute("SELECT * FROM operator_shifts WHERE id=?", (sid,)).fetchone()
        if not shift: raise HTTPException(404, "الوردية غير موجودة")
        if shift["status"] == "ended": raise HTTPException(400, "الوردية منتهية بالفعل")
        conn.execute("""UPDATE operator_shifts SET
            end_time=?, end_hours=?, fuel_liters=?, notes=?, status='ended'
            WHERE id=?""",
            (datetime.utcnow().isoformat(), body.end_hours, body.fuel_liters,
             body.notes or shift["notes"], sid))
    return {"message": "انتهت الوردية"}


@app.get("/shifts/active")
async def all_active_shifts(cu: dict = Depends(require_admin_or_reporter)):
    """كل الورديات النشطة الآن"""
    with get_db() as conn:
        _ensure_operator_tables(conn)
        rows = conn.execute("""
            SELECT s.*, o.name as operator_name, o.phone as operator_phone
            FROM operator_shifts s
            JOIN equipment_operators o ON o.id = s.operator_id
            WHERE s.status='active'
            ORDER BY s.start_time DESC
        """).fetchall()
    return {"active_shifts": [dict(r) for r in rows]}


# ── CSV تصدير / استيراد المعدات المستأجرة بفورمات الشيت ──
@app.get("/rental-equipment/export/excel-format")
async def export_rental_excel_format(
    month: str = Query(None),
    branch: str = Query(None),
    cu: dict = Depends(require_admin_or_reporter),
):
    """تصدير بفورمات الشيت الأصلي"""
    from fastapi.responses import Response
    eff_branch = _branch_filter(cu) or branch
    if month: month = month.strip()[:7]
    with get_db() as conn:
        q = "SELECT * FROM rental_equipment WHERE 1=1"
        params = []
        if eff_branch: q += " AND branch=?"; params.append(eff_branch)
        if month: q += " AND (month=? OR month LIKE ?)"; params.extend([month, month+"-%"])
        q += " ORDER BY branch, equipment_name"
        rows = [dict(r) for r in conn.execute(q, params).fetchall()]

    buf = io.StringIO()
    w = csv.writer(buf)
    # هيدر بالعربي زي الشيت
    w.writerow(["الفرع","نوع المعدة","كود المعدة","جهة الايجار","ايام العمل",
                "القيمة الايجارية الشهرية","استهلاك السولار باللتر",
                "عداد الساعات اول الشهر","عداد الساعات اخر الشهر","الفرق",
                "نسبة الاستخدام","معدل الاستهلاك","المشروع","ملاحظات",
                "الشهر","القطاع"])
    for r in rows:
        w.writerow([
            r.get("branch",""), r.get("equipment_name",""), r.get("code",""),
            r.get("rental_source",""), r.get("work_days",""), r.get("monthly_rate",""),
            r.get("fuel_liters",""), r.get("hours_start",""), r.get("hours_end",""),
            r.get("hours_diff",""), r.get("utilization_pct",""), r.get("consumption_rate",""),
            r.get("project",""), r.get("notes",""), r.get("month",""), r.get("sector",""),
        ])
    fname = f"rental_{month or 'all'}_{eff_branch or 'all'}.csv"
    return Response(
        content=buf.getvalue().encode("utf-8-sig"),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={fname}"}
    )


@app.get("/rental-equipment/template/excel-format")
async def rental_template_excel_format(cu: dict = Depends(require_admin_or_reporter)):
    """قالب CSV بنفس فورمات الشيت"""
    from fastapi.responses import Response
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["الفرع","نوع المعدة","كود المعدة","جهة الايجار","ايام العمل",
                "القيمة الايجارية الشهرية","استهلاك السولار باللتر",
                "عداد الساعات اول الشهر","عداد الساعات اخر الشهر","الفرق",
                "نسبة الاستخدام","معدل الاستهلاك","المشروع","ملاحظات","الشهر","القطاع"])
    # أمثلة من الشيتات الحقيقية
    samples = [
        ["حلوان","مان لفت 926","","وحده الرفع","24","84000","0","0","0","عاطل","","","مونوريل اكتوبر","","2026-03","قطاع القاهرة الكبرى"],
        ["حلوان","ونش شوكة متعدد12","","","24","48000","50","5795","5804","9","5.4","5.6","المحكمة","","2026-03","قطاع القاهرة الكبرى"],
        ["صيانة القصور","اسكيدلودر CAT","12/5/340","إدارة معدات تحريك التربة","24","48000","80","7347","7366","19","11.3","4.2","المركز الطبي","","2026-03","قطاع القاهرة الكبرى"],
        ["المنشات المتميزة","حفار","240/11/12","","23","50600","200","5835","5885","50","31.1","4","الازبكية","","2026-03","قطاع القاهرة الكبرى"],
        ["مدينة نصر","ونش على كاوتش 60 طن","","احمد محمد احمد عثمان","22","88000","400","70","117","47","30.5","","مونوريل اكتوبر","مؤجر من خارج","2026-03","قطاع القاهرة الكبرى"],
    ]
    for s in samples: w.writerow(s)
    return Response(
        content=buf.getvalue().encode("utf-8-sig"),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=rental_template_arabic.csv"}
    )


# ── استيراد بالفورمات العربي ──
def _parse_arabic_rental_row(row: dict, idx: int) -> dict:
    def g(*keys):
        for k in keys:
            v = row.get(k, "")
            if v and str(v).strip() not in ("", "nan", "None"):
                return str(v).strip()
        return ""
    errors = []
    name = g("نوع المعدة","equipment_name")
    branch = g("الفرع","branch")
    if not name: errors.append("نوع المعدة مطلوب")
    if not branch: errors.append("الفرع مطلوب")
    month_raw = g("الشهر","month")
    month_val = month_raw[:7] if len(month_raw) >= 7 else month_raw
    return {
        "row_index": idx, "valid": len(errors)==0, "errors": errors,
        "branch":          branch,
        "equipment_name":  name,
        "code":            g("كود المعدة","code"),
        "rental_source":   g("جهة الايجار","rental_source"),
        "work_days":       g("ايام العمل","work_days"),
        "monthly_rate":    g("القيمة الايجارية الشهرية","monthly_rate"),
        "fuel_liters":     g("استهلاك السولار باللتر","fuel_liters"),
        "hours_start":     g("عداد الساعات اول الشهر","hours_start"),
        "hours_end":       g("عداد الساعات اخر الشهر","hours_end"),
        "hours_diff":      g("الفرق","hours_diff"),
        "utilization_pct": g("نسبة الاستخدام","utilization_pct"),
        "consumption_rate":g("معدل الاستهلاك","consumption_rate"),
        "project":         g("المشروع","project"),
        "notes":           g("ملاحظات","notes"),
        "month":           month_val,
        "sector":          g("القطاع","sector"),
        "daily_rate":      g("القيمة اليومية","daily_rate"),
    }


@app.post("/rental-equipment/import/arabic")
async def import_rental_arabic(
    file: UploadFile = File(...),
    replace_month: bool = False,
    cu: dict = Depends(require_admin),
):
    """استيراد بفورمات الشيت العربي"""
    content = await file.read()
    if len(content) > 5*1024*1024: raise HTTPException(400, "الملف كبير جداً")
    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows: raise HTTPException(400, "الملف فارغ")
    rows = [_parse_arabic_rental_row(r, i+1) for i, r in enumerate(raw_rows)]
    valid = [r for r in rows if r["valid"]]
    invalid = [r for r in rows if not r["valid"]]
    inserted = 0
    with get_db() as conn:
        if replace_month:
            months = {r["month"] for r in valid if r["month"]}
            for m in months:
                conn.execute("DELETE FROM rental_equipment WHERE month=?", (m,))
        for row in valid:
            try:
                conn.execute("""INSERT INTO rental_equipment
                    (branch,equipment_name,code,rental_source,work_days,daily_rate,monthly_rate,
                     fuel_liters,hours_start,hours_end,hours_diff,utilization_pct,
                     consumption_rate,project,notes,month,sector)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (row["branch"],row["equipment_name"],row["code"],row["rental_source"],
                     row["work_days"],row["daily_rate"],row["monthly_rate"],row["fuel_liters"],
                     row["hours_start"],row["hours_end"],row["hours_diff"],row["utilization_pct"],
                     row["consumption_rate"],row["project"],row["notes"],row["month"],row["sector"]))
                inserted += 1
            except Exception as e:
                row["errors"]=[str(e)]; invalid.append(row)
    write_audit_log(cu["user_id"],cu["username"],cu["role"],
                    "import_rental_arabic", f"استيراد {inserted} معدة")
    return {"inserted": inserted, "skipped": len(invalid), "invalid_rows": invalid}

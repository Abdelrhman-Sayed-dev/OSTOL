# ══════════════════════════════════════════════════════
# FLEET MANAGEMENT SYSTEM — main.py
# ══════════════════════════════════════════════════════
# فهرس الأقسام (بترتيب ظهورها الفعلي في الملف) — كل رقم هنا يطابق
# رقم نفس القسم كـ تعليق داخل الكود، ابحث عنه (Ctrl+F) للوصول السريع.
#
#   1. ENVIRONMENT — REQUIRED (no fallback secrets)
#   2. STRUCTURED LOGGING
#   3. DATABASE — SQLite with WAL mode + proper indexing
#   4. AUTH — Access + Refresh Token system
#   5. RATE LIMITER
#   6. SECURITY HEADERS MIDDLEWARE
#   7. PYDANTIC MODELS
#   8. APP SETUP
#   9. SUPER ADMIN — Role Management Endpoints
#   10. IMPORT: جدول الصيانة الدورية (maintenance_schedule)
#   11. GLOBAL EXCEPTION HANDLER
#   12. HELPERS
#   13. AUTH ENDPOINTS
#   14. USER AVATAR — رفع وجلب صورة البروفايل للمستخدم
#   15. AUDIO UPLOAD — File system, NOT base64 in DB
#   16. DRIVERS
#   17. CARS
#   18. PERMISSIONS
#   19. DRIVER PHOTOS — رفع وعرض صور السائقين
#   20. LIVE LOCATION — real-time driver tracking
#   21. TRIPS — with pagination
#   22. USERS
#   23. WORKSHOPS
#   24. OPERATIONAL CARD PAGE 2 — زيوت وكاوتش وبطاريات
#   25. GARAGE RECORDS
#   26. EMERGENCY — Audio stored as file, not base64
#   27. SETTINGS
#   28. ADMIN — RESET DATA (protected with password)
#   29. SUPERUSER — Admin Management + Audit Logs
#   30. DRIVER REQUESTS
#   31. MAINTENANCE SCHEDULE (جدول الصيانة الدورية)
#   32. BRANCHES LIST — لقائمة الفروع المتاحة
#   33. OPERATIONAL CARD REPORT
#   34. OPERATIONAL REPORT — PER DRIVER BREAKDOWN
#   35. ANOMALOUS TRIPS DIAGNOSTIC
#   36. BULK IMPORT — CSV / Excel
#   37. PERMISSIONS BULK IMPORT — CSV / Excel
#   38. LAST ODOMETER — آخر عداد لمركبة (للسائق عند بدء الرحلة)
#   39. FRAUD DETECTION — كشف التلاعب
#   40. VOICE NOTES — رسائل صوتية من السوبر
#   41. MAINTENANCE SCHEDULE — CRUD + ALERTS
#   42. FUEL EFFICIENCY REPORT (تقرير كفاءة الوقود)
#   43. MONTHLY REPORT — تقرير شهري شامل
#   44. DRIVER KPI — مؤشرات أداء السائقين
#   45. RENTAL EQUIPMENT — المعدات المستأجرة
#   46. RENTAL EQUIPMENT — معدات مستأجرة (من الداتابيز)
#   47. EQUIPMENT — معدات (جدول منفصل - بدون لوحة، الكود هو المعرف)
#   48. EQUIPMENT OPERATORS — مشغلو المعدات (نظام ورديات)
#   49. السركي — تقرير حضور السائقين من الرحلات
#   50. 📷 OCR Odometer — EasyOCR (مجاني 100%، pip فقط، بدون apt)
#   51. VIRTUAL INVENTORY MODULE — إدارة المخزون
#   52. FORECAST ENGINE — Planning & Forecast Module
#   53. مراجعة العمليات — Workshop Approval Workflow (تفويل + صيانة معاً)
#   54. ENTRY POINT
# ══════════════════════════════════════════════════════

import asyncio
import base64
import csv
import hashlib
import io
import json
import math
import os
import uuid
import logging
from collections import defaultdict
from datetime import datetime, timedelta
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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
            role TEXT NOT NULL CHECK(role IN('superuser','super_admin','admin','driver','reporter','supervisor_workshop','supervisor_field','operator','workshop_admin')),
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
            odometer_photo TEXT DEFAULT '',
            FOREIGN KEY(driver_id) REFERENCES drivers(id),
            FOREIGN KEY(vehicle_id) REFERENCES cars(id)
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_driver ON workshop_records(driver_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_type ON workshop_records(type)")
        # ملحوظة: فهرس odometer_photo بيتعمل بعد ما الميجريشن الدفاعي تحت يضيف العمود فعلاً (راجع additions dict)

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
                role TEXT NOT NULL CHECK(role IN('superuser','super_admin','admin','driver','reporter','supervisor_workshop','supervisor_field','operator','workshop_admin')),
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

    # ── operator_equipment_permissions: migration آمنة ──
    try:
        c.execute("""CREATE TABLE IF NOT EXISTS operator_equipment_permissions (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            operator_id  INTEGER NOT NULL,
            equipment_id TEXT    NOT NULL,
            equipment_name TEXT  DEFAULT '',
            granted_by   TEXT   DEFAULT '',
            granted_at   TEXT   DEFAULT (datetime('now')),
            UNIQUE(operator_id, equipment_id),
            FOREIGN KEY (operator_id) REFERENCES equipment_operators(id) ON DELETE CASCADE
        )""")
        c.execute("CREATE INDEX IF NOT EXISTS idx_oep_operator ON operator_equipment_permissions(operator_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_oep_equip ON operator_equipment_permissions(equipment_id)")
    except Exception as _e:
        log.warning(f"operator_equipment_permissions migration: {_e}")
    for _col, _def in [("equipment_name","TEXT DEFAULT ''"),("granted_by","TEXT DEFAULT ''"),("granted_at","TEXT DEFAULT ''")]:
        try: c.execute(f"ALTER TABLE operator_equipment_permissions ADD COLUMN {_col} {_def}")
        except Exception: pass

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
        driver_id           INTEGER,
        operator_id         INTEGER,
        is_operator         INTEGER DEFAULT 0,
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
        created_at          TEXT NOT NULL
    )""")
    # ── Migration: rebuild driver_requests — إزالة FK + NOT NULL + إضافة operator columns ──
    try:
        c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='driver_requests'")
        _dr_row = c.fetchone()
        _dr_sql = _dr_row["sql"] if _dr_row else ""
        _needs_rebuild = (
            "NOT NULL" in _dr_sql
            or "FOREIGN KEY" in _dr_sql
            or "CHECK" in _dr_sql
            or "operator_id" not in _dr_sql
        )
        if _needs_rebuild:
            log.info("🔄 Rebuilding driver_requests — removing FK/NOT NULL, adding operator support...")
            c.execute("PRAGMA foreign_keys=OFF")
            # جيب كل الأعمدة الموجودة
            _existing = [r[1] for r in c.execute("PRAGMA table_info(driver_requests)").fetchall()]
            c.execute("""CREATE TABLE driver_requests_new (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                driver_id           INTEGER,
                operator_id         INTEGER,
                is_operator         INTEGER DEFAULT 0,
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
                created_at          TEXT NOT NULL
            )""")
            # انقل البيانات — فقط الأعمدة المشتركة
            # ⚠️ operator_id و is_operator يجب أن يكونا هنا دائماً لمنع data loss
            _target_cols = ["id","driver_id","operator_id","is_operator",
                            "type","notes","new_odometer","trip_id",
                            "status","priority","admin_notes","admin_message","handled_by",
                            "handled_at","forwarded_to_super","super_decision","super_notes",
                            "super_handled_by","super_handled_at","created_at"]
            _copy_cols = [col for col in _target_cols if col in _existing]
            _col_str = ",".join(_copy_cols)
            _coalesce = []
            for col in _copy_cols:
                if col in ("priority",):
                    _coalesce.append(f"COALESCE({col},'medium')")
                elif col in ("admin_notes","admin_message","handled_by","handled_at",
                             "super_decision","super_notes","super_handled_by","super_handled_at"):
                    _coalesce.append(f"COALESCE({col},'')")
                elif col in ("forwarded_to_super",):
                    _coalesce.append(f"COALESCE({col},0)")
                else:
                    _coalesce.append(col)
            _coalesce_str = ",".join(_coalesce)
            c.execute(f"INSERT OR IGNORE INTO driver_requests_new ({_col_str}) SELECT {_coalesce_str} FROM driver_requests")
            c.execute("DROP TABLE driver_requests")
            c.execute("ALTER TABLE driver_requests_new RENAME TO driver_requests")
            c.execute("CREATE INDEX IF NOT EXISTS idx_req_driver   ON driver_requests(driver_id)")
            c.execute("CREATE INDEX IF NOT EXISTS idx_req_operator ON driver_requests(operator_id)")
            c.execute("CREATE INDEX IF NOT EXISTS idx_req_status   ON driver_requests(status)")
            c.execute("PRAGMA foreign_keys=ON")
            # ── Verification: تحقق من سلامة البيانات بعد الـ rebuild ──
            _v = c.execute("""SELECT
                COUNT(*) as total,
                COUNT(CASE WHEN is_operator=1 THEN 1 END) as op_count,
                COUNT(CASE WHEN is_operator=1 AND operator_id IS NOT NULL THEN 1 END) as op_with_id
            FROM driver_requests""").fetchone()
            log.info(f"✅ driver_requests rebuilt: total={_v['total']}, operator_requests={_v['op_count']}, with_operator_id={_v['op_with_id']}")
            if _v["op_count"] > 0 and _v["op_with_id"] < _v["op_count"]:
                log.error(f"⚠️  DATA INTEGRITY WARNING: {_v['op_count'] - _v['op_with_id']} operator requests lost their operator_id after rebuild!")
    except Exception as _e:
        log.warning(f"driver_requests rebuild migration: {_e}")

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

    # Migration: أضف operator_id + is_operator لـ driver_requests
    for _col, _def in [("operator_id","INTEGER"), ("is_operator","INTEGER DEFAULT 0")]:
        try: c.execute(f"ALTER TABLE driver_requests ADD COLUMN {_col} {_def}")
        except Exception: pass

    additions = {
        "users":              [("refresh_token","TEXT"),("refresh_exp","TEXT"),("last_login","TEXT"),
                               ("branch","TEXT DEFAULT ''")],
        "emergency_reports":  [("audio_url","TEXT DEFAULT ''"),("is_handled","INTEGER DEFAULT 0"),
                               ("action_taken","TEXT DEFAULT ''"),("handled_by","TEXT DEFAULT ''"),
                               ("action_time","TEXT DEFAULT ''"),("location","TEXT DEFAULT ''"),
                               ("driver_message","TEXT DEFAULT ''"),
                               ("operator_id","INTEGER"),
                               ("is_operator","INTEGER DEFAULT 0")] ,
        "trips":              [("garage_location","TEXT DEFAULT ''")],
        "workshop_records":   [("location","TEXT DEFAULT ''"),("operation_type","TEXT DEFAULT ''"),
                               ("vehicle_id","INTEGER"),("odometer_reading","REAL"),
                               ("description","TEXT DEFAULT ''"),("tire_action","TEXT DEFAULT ''"),
                               ("odometer_photo","TEXT DEFAULT ''"),
                               ("doc_number","TEXT DEFAULT ''"),
                               ("engine_hours","REAL"),
                               ("supply_source","TEXT DEFAULT ''"),
                               ("item_name","TEXT DEFAULT ''"),
                               ("item_spec","TEXT DEFAULT ''"),
                               ("receiver_name","TEXT DEFAULT ''"),
                               ("operator_id","INTEGER"),
                               ("is_operator","INTEGER DEFAULT 0"),
                               ("fuel_source","TEXT DEFAULT ''"),
                               ("station_name","TEXT DEFAULT ''"),
                               ("pump_number","TEXT DEFAULT ''"),
                               ("invoice_photo","TEXT DEFAULT ''"),
                               ("parts_source","TEXT DEFAULT ''"),
                               ("inventory_product_id","INTEGER"),
                               ("direct_purchase_supplier","TEXT DEFAULT ''"),
                               ("direct_purchase_cost","REAL"),
                               ("approval_status","TEXT DEFAULT 'pending'")],
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

    # ── الآن العمود odometer_photo موجود بالتأكيد (سواء من CREATE TABLE أو من الميجريشن فوق) — أنشئ الفهرس بأمان ──
    try:
        c.execute("CREATE INDEX IF NOT EXISTS idx_ws_photo ON workshop_records(odometer_photo)")
    except Exception:
        pass

    # ══════════════════════════════════════════════════════
    # VIRTUAL INVENTORY MODULE — إدارة المخزون
    # ══════════════════════════════════════════════════════
    c.execute("""CREATE TABLE IF NOT EXISTS inventory_categories(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        notes TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now'))
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS inventory_suppliers(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT DEFAULT '',
        email TEXT DEFAULT '',
        address TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        company_name TEXT DEFAULT '',
        manager_name TEXT DEFAULT '',
        business_sector TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now'))
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS inventory_products(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        sku TEXT DEFAULT '',
        category_id INTEGER,
        supplier_id INTEGER,
        brand TEXT DEFAULT '',
        purchase_price REAL DEFAULT 0,
        sale_price REAL DEFAULT 0,
        quantity REAL DEFAULT 0,
        min_quantity REAL DEFAULT 0,
        unit TEXT DEFAULT 'قطعة',
        notes TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(category_id) REFERENCES inventory_categories(id) ON DELETE SET NULL,
        FOREIGN KEY(supplier_id) REFERENCES inventory_suppliers(id) ON DELETE SET NULL
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS inventory_movements(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER NOT NULL,
        movement_type TEXT NOT NULL CHECK(movement_type IN ('in','out','return','adjust')),
        quantity REAL NOT NULL,
        ref_type TEXT DEFAULT '',
        ref_id INTEGER,
        username TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(product_id) REFERENCES inventory_products(id) ON DELETE CASCADE
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS inventory_purchase_orders(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_number TEXT DEFAULT '',
        supplier_id INTEGER,
        order_date TEXT DEFAULT '',
        status TEXT DEFAULT 'draft' CHECK(status IN ('draft','open','received','cancelled')),
        total_value REAL DEFAULT 0,
        items_json TEXT DEFAULT '[]',
        notes TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(supplier_id) REFERENCES inventory_suppliers(id) ON DELETE SET NULL
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS workshop_repair_quotes(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        quote_number TEXT DEFAULT '',
        car_id INTEGER,
        equipment_id TEXT DEFAULT '',
        driver_id INTEGER DEFAULT NULL,
        quote_date TEXT DEFAULT '',
        status TEXT DEFAULT 'draft' CHECK(status IN ('draft','approved','done','cancelled')),
        items_json TEXT DEFAULT '[]',
        total_value REAL DEFAULT 0,
        notes TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE SET NULL,
        FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE SET NULL
    )""")
    # Migration: handle existing DB where driver_id/car_id were NOT NULL، وأضف equipment_id
    try:
        col_info = c.execute("PRAGMA table_info(workshop_repair_quotes)").fetchall()
        driver_col = next((col for col in col_info if col[1] == 'driver_id'), None)
        car_col = next((col for col in col_info if col[1] == 'car_id'), None)
        has_equip = any(col[1] == 'equipment_id' for col in col_info)
        if (driver_col and driver_col[3] == 1) or (car_col and car_col[3] == 1) or not has_equip:
            # Recreate table بدون NOT NULL على driver_id/car_id، ومع عمود equipment_id
            c.execute("ALTER TABLE workshop_repair_quotes RENAME TO workshop_repair_quotes_old")
            c.execute("""CREATE TABLE workshop_repair_quotes(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                quote_number TEXT DEFAULT '',
                car_id INTEGER,
                equipment_id TEXT DEFAULT '',
                driver_id INTEGER DEFAULT NULL,
                quote_date TEXT DEFAULT '',
                status TEXT DEFAULT 'draft' CHECK(status IN ('draft','approved','done','cancelled')),
                items_json TEXT DEFAULT '[]',
                total_value REAL DEFAULT 0,
                notes TEXT DEFAULT '',
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE SET NULL,
                FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE SET NULL
            )""")
            old_cols = {col[1] for col in c.execute("PRAGMA table_info(workshop_repair_quotes_old)").fetchall()}
            equip_select = "equipment_id" if "equipment_id" in old_cols else "''"
            c.execute(f"""INSERT INTO workshop_repair_quotes
                SELECT id,quote_number,
                       CASE WHEN car_id=0 THEN NULL ELSE car_id END,
                       {equip_select},
                       CASE WHEN driver_id=0 THEN NULL
                            WHEN driver_id NOT IN (SELECT id FROM drivers) THEN NULL
                            ELSE driver_id END,
                       quote_date,status,items_json,total_value,notes,created_at
                FROM workshop_repair_quotes_old""")
            c.execute("DROP TABLE workshop_repair_quotes_old")
        else:
            # Already nullable, just clean up zeros
            c.execute("UPDATE workshop_repair_quotes SET driver_id=NULL WHERE driver_id=0")
    except Exception as _e:
        pass

    # ── جدول محاضر الفحص (Inspection Reports) — لدور ادمن الورش ──
    c.execute("""CREATE TABLE IF NOT EXISTS workshop_inspection_reports(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_number TEXT DEFAULT '',
        car_id INTEGER,
        equipment_id TEXT DEFAULT '',
        driver_id INTEGER DEFAULT NULL,
        quote_id INTEGER DEFAULT NULL,
        inspection_date TEXT DEFAULT '',
        inspector_name TEXT DEFAULT '',
        odometer_reading REAL,
        result TEXT DEFAULT 'pending' CHECK(result IN ('pass','fail','needs_repair','pending')),
        findings TEXT DEFAULT '',
        items_json TEXT DEFAULT '[]',
        notes TEXT DEFAULT '',
        created_by TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE SET NULL,
        FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE SET NULL,
        FOREIGN KEY(quote_id) REFERENCES workshop_repair_quotes(id) ON DELETE SET NULL
    )""")
    # migration: إضافة quote_id / equipment_id لو الجدول كان موجود من قبل بدونهم، وخلي car_id يقبل NULL
    try:
        insp_cols_info = c.execute("PRAGMA table_info(workshop_inspection_reports)").fetchall()
        insp_cols = [r[1] for r in insp_cols_info]
        if "quote_id" not in insp_cols:
            c.execute("ALTER TABLE workshop_inspection_reports ADD COLUMN quote_id INTEGER DEFAULT NULL")
        if "equipment_id" not in insp_cols:
            c.execute("ALTER TABLE workshop_inspection_reports ADD COLUMN equipment_id TEXT DEFAULT ''")
        insp_cols_info = c.execute("PRAGMA table_info(workshop_inspection_reports)").fetchall()
        car_col = next((col for col in insp_cols_info if col[1] == 'car_id'), None)
        if car_col and car_col[3] == 1:
            c.execute("ALTER TABLE workshop_inspection_reports RENAME TO workshop_inspection_reports_old")
            c.execute("""CREATE TABLE workshop_inspection_reports(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                report_number TEXT DEFAULT '',
                car_id INTEGER,
                equipment_id TEXT DEFAULT '',
                driver_id INTEGER DEFAULT NULL,
                quote_id INTEGER DEFAULT NULL,
                inspection_date TEXT DEFAULT '',
                inspector_name TEXT DEFAULT '',
                odometer_reading REAL,
                result TEXT DEFAULT 'pending' CHECK(result IN ('pass','fail','needs_repair','pending')),
                findings TEXT DEFAULT '',
                items_json TEXT DEFAULT '[]',
                notes TEXT DEFAULT '',
                created_by TEXT DEFAULT '',
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE SET NULL,
                FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE SET NULL,
                FOREIGN KEY(quote_id) REFERENCES workshop_repair_quotes(id) ON DELETE SET NULL
            )""")
            c.execute("""INSERT INTO workshop_inspection_reports
                (id,report_number,car_id,equipment_id,driver_id,quote_id,inspection_date,inspector_name,
                 odometer_reading,result,findings,items_json,notes,created_by,created_at)
                SELECT id,report_number,
                       CASE WHEN car_id=0 THEN NULL ELSE car_id END,
                       COALESCE(equipment_id,''),driver_id,quote_id,inspection_date,inspector_name,
                       odometer_reading,result,findings,items_json,notes,created_by,created_at
                FROM workshop_inspection_reports_old""")
            c.execute("DROP TABLE workshop_inspection_reports_old")
    except Exception:
        pass

    # ── جدول طلبات التشغيل الخارجي (الورشة) ──
    c.execute("""CREATE TABLE IF NOT EXISTS workshop_external_ops(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        request_number TEXT DEFAULT '',
        quote_id INTEGER DEFAULT NULL,
        car_id INTEGER,
        equipment_id TEXT DEFAULT '',
        driver_id INTEGER DEFAULT NULL,
        workshop_name TEXT DEFAULT '',
        reason TEXT DEFAULT '',
        estimated_cost REAL DEFAULT 0,
        workmanship_cost REAL DEFAULT 0,
        request_date TEXT DEFAULT '',
        status TEXT DEFAULT 'pending' CHECK(status IN ('pending','approved','rejected','done')),
        notes TEXT DEFAULT '',
        created_by TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE SET NULL,
        FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE SET NULL,
        FOREIGN KEY(quote_id) REFERENCES workshop_repair_quotes(id) ON DELETE SET NULL
    )""")
    # migration: أضف equipment_id وخلي car_id يقبل NULL
    try:
        eo_cols_info = c.execute("PRAGMA table_info(workshop_external_ops)").fetchall()
        eo_cols = [r[1] for r in eo_cols_info]
        if "equipment_id" not in eo_cols:
            c.execute("ALTER TABLE workshop_external_ops ADD COLUMN equipment_id TEXT DEFAULT ''")
        eo_cols_info = c.execute("PRAGMA table_info(workshop_external_ops)").fetchall()
        car_col = next((col for col in eo_cols_info if col[1] == 'car_id'), None)
        if car_col and car_col[3] == 1:
            c.execute("ALTER TABLE workshop_external_ops RENAME TO workshop_external_ops_old")
            c.execute("""CREATE TABLE workshop_external_ops(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                request_number TEXT DEFAULT '',
                quote_id INTEGER DEFAULT NULL,
                car_id INTEGER,
                equipment_id TEXT DEFAULT '',
                driver_id INTEGER DEFAULT NULL,
                workshop_name TEXT DEFAULT '',
                reason TEXT DEFAULT '',
                estimated_cost REAL DEFAULT 0,
                workmanship_cost REAL DEFAULT 0,
                request_date TEXT DEFAULT '',
                status TEXT DEFAULT 'pending' CHECK(status IN ('pending','approved','rejected','done')),
                notes TEXT DEFAULT '',
                created_by TEXT DEFAULT '',
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE SET NULL,
                FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE SET NULL,
                FOREIGN KEY(quote_id) REFERENCES workshop_repair_quotes(id) ON DELETE SET NULL
            )""")
            c.execute("""INSERT INTO workshop_external_ops
                (id,request_number,quote_id,car_id,equipment_id,driver_id,workshop_name,reason,
                 estimated_cost,workmanship_cost,request_date,status,notes,created_by,created_at)
                SELECT id,request_number,quote_id,
                       CASE WHEN car_id=0 THEN NULL ELSE car_id END,
                       COALESCE(equipment_id,''),driver_id,workshop_name,reason,
                       estimated_cost,workmanship_cost,request_date,status,notes,created_by,created_at
                FROM workshop_external_ops_old""")
            c.execute("DROP TABLE workshop_external_ops_old")
    except Exception:
        pass

    # ── جدول إذن الصرف (الورشة) ──
    c.execute("""CREATE TABLE IF NOT EXISTS workshop_vouchers(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        voucher_number TEXT DEFAULT '',
        quote_id INTEGER DEFAULT NULL,
        car_id INTEGER DEFAULT NULL,
        equipment_id TEXT DEFAULT '',
        driver_id INTEGER DEFAULT NULL,
        recipient_name TEXT DEFAULT '',
        purpose TEXT DEFAULT '',
        items_json TEXT DEFAULT '[]',
        total_amount REAL DEFAULT 0,
        issue_date TEXT DEFAULT '',
        issued_by TEXT DEFAULT '',
        status TEXT DEFAULT 'draft' CHECK(status IN ('draft','issued','cancelled')),
        notes TEXT DEFAULT '',
        created_by TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE SET NULL,
        FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE SET NULL,
        FOREIGN KEY(quote_id) REFERENCES workshop_repair_quotes(id) ON DELETE SET NULL
    )""")
    # migration: equipment_id لو الجدول كان موجود من قبل بدونه، وخلي car_id يقبل NULL
    try:
        vch_cols = [r[1] for r in c.execute("PRAGMA table_info(workshop_vouchers)").fetchall()]
        if "equipment_id" not in vch_cols:
            c.execute("ALTER TABLE workshop_vouchers ADD COLUMN equipment_id TEXT DEFAULT ''")
        # لو car_id لسه NOT NULL من قبل (جدول قديم) — نعيد إنشاء الجدول بدون القيد عشان نسمح بالصرف على معدة بدل مركبة
        vch_col_info = c.execute("PRAGMA table_info(workshop_vouchers)").fetchall()
        car_col = next((col for col in vch_col_info if col[1] == 'car_id'), None)
        if car_col and car_col[3] == 1:  # notnull=1
            c.execute("ALTER TABLE workshop_vouchers RENAME TO workshop_vouchers_old")
            c.execute("""CREATE TABLE workshop_vouchers(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                voucher_number TEXT DEFAULT '',
                quote_id INTEGER DEFAULT NULL,
                car_id INTEGER DEFAULT NULL,
                equipment_id TEXT DEFAULT '',
                driver_id INTEGER DEFAULT NULL,
                recipient_name TEXT DEFAULT '',
                purpose TEXT DEFAULT '',
                items_json TEXT DEFAULT '[]',
                total_amount REAL DEFAULT 0,
                issue_date TEXT DEFAULT '',
                issued_by TEXT DEFAULT '',
                status TEXT DEFAULT 'draft' CHECK(status IN ('draft','issued','cancelled')),
                notes TEXT DEFAULT '',
                created_by TEXT DEFAULT '',
                created_at TEXT DEFAULT (datetime('now')),
                FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE SET NULL,
                FOREIGN KEY(driver_id) REFERENCES drivers(id) ON DELETE SET NULL,
                FOREIGN KEY(quote_id) REFERENCES workshop_repair_quotes(id) ON DELETE SET NULL
            )""")
            c.execute("""INSERT INTO workshop_vouchers
                (id,voucher_number,quote_id,car_id,equipment_id,driver_id,recipient_name,purpose,
                 items_json,total_amount,issue_date,issued_by,status,notes,created_by,created_at)
                SELECT id,voucher_number,quote_id,
                       CASE WHEN car_id=0 THEN NULL ELSE car_id END,
                       COALESCE(equipment_id,''),driver_id,recipient_name,purpose,
                       items_json,total_amount,issue_date,issued_by,status,notes,created_by,created_at
                FROM workshop_vouchers_old""")
            c.execute("DROP TABLE workshop_vouchers_old")
    except Exception:
        pass

    # ── منصات الإمداد (مخازن متحركة — تانك سولار على عربة، ممكن يحمل وقود أو قطع غيار) ──
    c.execute("""CREATE TABLE IF NOT EXISTS supply_platforms(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT NOT NULL UNIQUE,
        plate TEXT DEFAULT '',
        driver_name TEXT DEFAULT '',
        manager_name TEXT DEFAULT '',
        license_number TEXT DEFAULT '',
        status TEXT DEFAULT 'active' CHECK(status IN ('active','inactive')),
        notes TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now'))
    )""")
    # رصيد كل منصة من كل منتج (سولار/بنزين/قطع غيار...) — مخزون فرعي متحرك
    c.execute("""CREATE TABLE IF NOT EXISTS supply_platform_stock(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        platform_id INTEGER NOT NULL,
        product_id INTEGER NOT NULL,
        quantity REAL DEFAULT 0,
        updated_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(platform_id) REFERENCES supply_platforms(id) ON DELETE CASCADE,
        FOREIGN KEY(product_id) REFERENCES inventory_products(id) ON DELETE CASCADE,
        UNIQUE(platform_id, product_id)
    )""")
    # حركات كل منصة: تعبئة (fill) من المخزن الرئيسي، أو سحب (draw) عند التفويل/الصرف
    c.execute("""CREATE TABLE IF NOT EXISTS supply_platform_movements(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        platform_id INTEGER NOT NULL,
        product_id INTEGER NOT NULL,
        movement_type TEXT NOT NULL CHECK(movement_type IN ('fill','draw','adjust')),
        quantity REAL NOT NULL,
        ref_type TEXT DEFAULT '',
        ref_id INTEGER,
        username TEXT DEFAULT '',
        notes TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now')),
        FOREIGN KEY(platform_id) REFERENCES supply_platforms(id) ON DELETE CASCADE,
        FOREIGN KEY(product_id) REFERENCES inventory_products(id) ON DELETE CASCADE
    )""")

    for idx_sql in [
        "CREATE INDEX IF NOT EXISTS idx_inv_products_cat ON inventory_products(category_id)",
        "CREATE INDEX IF NOT EXISTS idx_inv_products_sup ON inventory_products(supplier_id)",
        "CREATE INDEX IF NOT EXISTS idx_inv_movements_prod ON inventory_movements(product_id)",
        "CREATE INDEX IF NOT EXISTS idx_inv_po_supplier ON inventory_purchase_orders(supplier_id)",
        "CREATE INDEX IF NOT EXISTS idx_rq_car ON workshop_repair_quotes(car_id)",
        "CREATE INDEX IF NOT EXISTS idx_rq_driver ON workshop_repair_quotes(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_ws_approval ON workshop_records(approval_status)",
        "CREATE INDEX IF NOT EXISTS idx_insp_car ON workshop_inspection_reports(car_id)",
        "CREATE INDEX IF NOT EXISTS idx_insp_driver ON workshop_inspection_reports(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_insp_quote ON workshop_inspection_reports(quote_id)",
        "CREATE INDEX IF NOT EXISTS idx_extop_car ON workshop_external_ops(car_id)",
        "CREATE INDEX IF NOT EXISTS idx_extop_quote ON workshop_external_ops(quote_id)",
        "CREATE INDEX IF NOT EXISTS idx_voucher_car ON workshop_vouchers(car_id)",
        "CREATE INDEX IF NOT EXISTS idx_voucher_quote ON workshop_vouchers(quote_id)",
        "CREATE INDEX IF NOT EXISTS idx_platform_stock_platform ON supply_platform_stock(platform_id)",
        "CREATE INDEX IF NOT EXISTS idx_platform_stock_product ON supply_platform_stock(product_id)",
        "CREATE INDEX IF NOT EXISTS idx_platform_mv_platform ON supply_platform_movements(platform_id)",
        "CREATE INDEX IF NOT EXISTS idx_platform_mv_product ON supply_platform_movements(product_id)",
    ]:
        try: c.execute(idx_sql)
        except Exception: pass
    # ── Migration: أعمدة جديدة للمنتجات والموردين (لقواعد البيانات المنشورة مسبقاً) ──
    try: c.execute("ALTER TABLE inventory_products ADD COLUMN brand TEXT DEFAULT ''")
    except Exception: pass
    for _sup_col, _sup_def in [("company_name", "TEXT DEFAULT ''"), ("manager_name", "TEXT DEFAULT ''"), ("business_sector", "TEXT DEFAULT ''")]:
        try: c.execute(f"ALTER TABLE inventory_suppliers ADD COLUMN {_sup_col} {_sup_def}")
        except Exception: pass
    # تصنيفات أساسية افتراضية (مرة واحدة فقط لو الجدول فاضي)
    try:
        if c.execute("SELECT COUNT(*) FROM inventory_categories").fetchone()[0] == 0:
            for cat_name in ["الوقود","الإطارات","البطاريات","العمرات","الفرامل","الزيوت والفلاتر","الكهرباء","التبريد","العفشة","أخرى"]:
                c.execute("INSERT INTO inventory_categories(name) VALUES(?)", (cat_name,))
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
                    role TEXT NOT NULL CHECK(role IN('superuser','super_admin','admin','driver','reporter','supervisor_workshop','supervisor_field','operator','workshop_admin')),
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

    # ── Migration: إضافة workshop_admin role (ادمن ورش) في users CHECK constraint ──
    try:
        c.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='users'")
        row = c.fetchone()
        if row:
            existing_sql = row['sql'] or ''
            if "'workshop_admin'" not in existing_sql:
                log.info("🔄 Adding workshop_admin role to users table...")
                c.execute("PRAGMA foreign_keys=OFF")
                c.execute("""CREATE TABLE IF NOT EXISTS _users_wa_fix(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT NOT NULL,
                    password TEXT NOT NULL,
                    role TEXT NOT NULL CHECK(role IN('superuser','super_admin','admin','driver','reporter','supervisor_workshop','supervisor_field','operator','workshop_admin')),
                    branch TEXT DEFAULT '',
                    created_at TEXT DEFAULT(datetime('now')),
                    last_login TEXT,
                    refresh_token TEXT,
                    refresh_exp TEXT,
                    avatar_url TEXT DEFAULT ''
                )""")
                c.execute("INSERT OR IGNORE INTO _users_wa_fix SELECT id,username,password,role,COALESCE(branch,''),created_at,last_login,refresh_token,refresh_exp,COALESCE(avatar_url,'') FROM users")
                c.execute("DROP TABLE users")
                c.execute("ALTER TABLE _users_wa_fix RENAME TO users")
                c.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")
                c.execute("PRAGMA foreign_keys=ON")
                log.info("✅ workshop_admin role added to users table")
    except Exception as _e:
        log.warning(f"workshop_admin role migration: {_e}")

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

    # ── migration: managed_by (ربط الأدمن بالسوبر يوزر المسؤول عنه) ──
    try:
        c.execute("ALTER TABLE users ADD COLUMN managed_by INTEGER DEFAULT NULL")
        log.info("✅ Migrated users: added managed_by column")
    except Exception:
        pass  # العمود موجود بالفعل
    try:
        c.execute("CREATE INDEX IF NOT EXISTS idx_users_managed_by ON users(managed_by)")
    except Exception:
        pass

    # ── migration: account lockout (قفل الحساب بعد محاولات دخول فاشلة متتالية) ──
    try:
        c.execute("ALTER TABLE users ADD COLUMN failed_attempts INTEGER DEFAULT 0")
        log.info("✅ Migrated users: added failed_attempts column")
    except Exception:
        pass  # العمود موجود بالفعل
    try:
        c.execute("ALTER TABLE users ADD COLUMN locked_until TEXT DEFAULT NULL")
        log.info("✅ Migrated users: added locked_until column")
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
    # migration: إزالة CHECK constraint من voice_notes نهائياً
    try:
        vn_sql = c.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name='voice_notes'"
        ).fetchone()
        if vn_sql and 'CHECK' in (vn_sql[0] or ''):
            c.executescript("""
                CREATE TABLE IF NOT EXISTS _vn_tmp(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sender_id INTEGER NOT NULL,
                    sender_role TEXT NOT NULL DEFAULT 'superuser',
                    target_group TEXT DEFAULT '',
                    target_user_id INTEGER DEFAULT NULL,
                    audio_data TEXT NOT NULL,
                    duration_sec REAL DEFAULT 0,
                    created_at TEXT NOT NULL,
                    expires_at TEXT NOT NULL,
                    play_count INTEGER DEFAULT 0,
                    max_plays INTEGER DEFAULT 2,
                    is_deleted INTEGER DEFAULT 0
                );
                INSERT OR IGNORE INTO _vn_tmp
                    SELECT id,sender_id,
                           COALESCE(sender_role,'superuser'),
                           target_group,target_user_id,audio_data,
                           duration_sec,created_at,expires_at,
                           play_count,max_plays,is_deleted
                    FROM voice_notes;
                DROP TABLE voice_notes;
                ALTER TABLE _vn_tmp RENAME TO voice_notes;
            """)
            log.info('✅ voice_notes CHECK constraint removed')
    except Exception as e:
        log.warning(f'voice_notes constraint migration: {e}')

    # تصحيح الرسائل الشخصية القديمة → target_group = 'personal'
    try:
        c.execute("""
            UPDATE voice_notes SET target_group = 'personal'
            WHERE target_user_id IS NOT NULL AND target_user_id != 0
              AND target_group IN ('admins', 'drivers')
        """)
    except Exception: pass
    try: c.execute("CREATE INDEX IF NOT EXISTS idx_vnotes_target ON voice_notes(target_user_id)")
    except Exception: pass

# ══════════════════════════════════════════════════════
# 4. AUTH — Access + Refresh Token system
# ══════════════════════════════════════════════════════

def _hash(pw: str) -> str:
    return bcrypt.hashpw(pw.encode(), bcrypt.gensalt(rounds=12)).decode()

# alias — يُستخدم في عدة endpoints (تغيير كلمة المرور، إنشاء/تعديل سوبر أدمن)
hash_password = _hash

def _verify(plain: str, hashed: str) -> bool:
    try:
        return bcrypt.checkpw(plain.encode(), hashed.encode())
    except Exception:
        return False

def _hash_refresh_token(token: str) -> str:
    """
    نخزن الـ hash فقط لـ refresh token في قاعدة البيانات (مش القيمة الخام).
    لو حد قدر يوصل لنسخة من قاعدة البيانات (باكاب، تسريب، ريبو...) مش
    هيقدر يستخدمها كـ refresh token صالح لانتحال شخصية أي مستخدم.
    """
    return hashlib.sha256(token.encode()).hexdigest()

# ── Account Lockout — قفل الحساب بعد محاولات دخول فاشلة متتالية ──
# بعد MAX_FAILED_LOGIN_ATTEMPTS محاولة فاشلة على نفس الحساب، يتقفل الحساب
# تلقائياً لمدة LOCKOUT_MINUTES دقيقة، حتى لو المحاولات جاية من IPs مختلفة.
# ده إضافي فوق الـ rate limiting الحالي (اللي بيراقب IP)، مش بديل عنه.
MAX_FAILED_LOGIN_ATTEMPTS = 5
LOCKOUT_MINUTES = 15

def validate_password(pw: str):
    """حد أدنى من متطلبات قوة كلمة المرور — يطبّق على أي كلمة مرور جديدة أو معدّلة."""
    if not pw:
        raise HTTPException(400, "كلمة المرور مطلوبة")
    if len(pw) < 8:
        raise HTTPException(400, "كلمة المرور يجب ألا تقل عن 8 أحرف")
    if pw.strip().lower() in (
        "12345678", "123456789", "password", "qwerty123", "11111111", "00000000"
    ):
        raise HTTPException(400, "كلمة المرور ضعيفة جداً وسهلة التخمين — اختر كلمة مرور أقوى")

def validate_driver_password(pw: str):
    """
    كلمة مرور/PIN السائقين أخف من حسابات الإدارة (عادة أرقام قصيرة يسهل إدخالها
    من الموبايل)، لكن برضه بحد أدنى يمنع قيم فارغة أو رقم واحد سهل التخمين.
    """
    if not pw:
        raise HTTPException(400, "كلمة المرور مطلوبة")
    if len(pw) < 4:
        raise HTTPException(400, "كلمة المرور يجب ألا تقل عن 4 أحرف/أرقام")

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

ADMIN_LIKE_ROLES = ("superuser","admin","reporter","supervisor_workshop","supervisor_field","workshop_admin")

def require_admin_or_reporter(cu: dict = Depends(get_user)):
    """Allows superuser, admin, reporter and supervisors (read-only) roles."""
    if cu["role"] not in ADMIN_LIKE_ROLES:
        raise HTTPException(status.HTTP_403_FORBIDDEN, "صلاحيات غير كافية")
    return cu

# دور "ادمن ورش" — إدارة كاملة (إضافة/تعديل/حذف) لكل ما يخص الورشة:
# سجلات الورشة والصيانة، جداول ومقايسات الإصلاح، محاضر الفحص، وبقية بنود الورشة.
WORKSHOP_ADMIN_ROLES = ("superuser", "admin", "workshop_admin")

def require_workshop_admin(cu: dict = Depends(get_user)):
    """Allows superuser, admin, and workshop_admin (دور ادمن الورش الكامل)."""
    if cu["role"] not in WORKSHOP_ADMIN_ROLES:
        raise HTTPException(status.HTTP_403_FORBIDDEN, "صلاحيات إدارة الورشة مطلوبة")
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
    "filter",        # فلتر (legacy)
    "filter_oil",       # فلتر زيت
    "filter_solar",     # فلتر سولار
    "filter_petrol",    # فلتر بنزين
    "filter_air",       # فلتر هواء
    "filter_separator", # فلتر فاصل
    "filter_ac",        # فلتر تكيف
    "filter_dryer",     # فلتر مجفف
    "filter_breather",  # فلتر منفس
    "filter_hydraulic", # فلتر هيدروليك
    "tire", "battery", "belt", "other",
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
        allowed = ("admin","superuser","driver","reporter","supervisor_workshop","supervisor_field","operator","workshop_admin")
        if v not in allowed:
            raise ValueError(f"دور غير صالح. المتاح: {allowed}")
        return v

class WorkshopCreate(BaseModel):
    driver_id: int; type: str; quantity: Optional[float] = None
    price: float = 0; notes: Optional[str] = ""
    operation_type: Optional[str] = ""; vehicle_id: Optional[int] = None
    odometer_reading: Optional[float] = None; description: Optional[str] = ""
    tire_action: Optional[str] = ""; location: Optional[str] = ""
    odometer_photo: Optional[str] = ""   # base64 من الفرونت — صورة العداد عند التفويل/الصيانة
    # حقول البطاقة التشغيلية
    doc_number:    Optional[str]   = ""    # رقم مستند الصرف
    engine_hours:  Optional[float] = None  # ساعات المحرك
    supply_source: Optional[str]   = ""    # جهة الصرف
    item_name:     Optional[str]   = ""    # اسم الصنف (كاوتش/بطارية)
    item_spec:     Optional[str]   = ""    # المقاس
    receiver_name: Optional[str]   = ""    # اسم المتسلم
    # ── Virtual Inventory: مصدر الوقود (للتفويل) ──
    fuel_source:   Optional[str]   = ""    # CompanyStation | ExternalStation | SupplyPlatform
    station_name:  Optional[str]   = ""    # اسم المحطة الخارجية
    pump_number:   Optional[str]   = ""    # رقم المضخة (محطة الشركة)
    invoice_photo: Optional[str]   = ""    # base64 صورة الفاتورة
    # ── Virtual Inventory: مصدر القطعة (للصيانة) ──
    parts_source:  Optional[str]   = ""    # Inventory | DirectPurchase | SupplyPlatform
    inventory_product_id: Optional[int] = None   # المنتج المصروف من المخزن أو من منصة الإمداد
    direct_purchase_supplier: Optional[str] = ""  # المورد/المحل (شراء مباشر)
    direct_purchase_cost: Optional[float] = None  # تكلفة الشراء المباشر
    # ── منصات الإمداد: المخزن المتحرك اللي تم السحب منه (وقود أو قطعة) ──
    platform_id: Optional[int] = None


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

# Startup
# ── حسابات البذر (seed accounts) ────────────────────────────────────────
# بدل ما تكون أسماء المستخدمين وكلمات المرور مكتوبة صراحة في الكود (خطر أمني
# لو الكود اتشاف من حد تاني أو اترفع مكان عام)، بقت بتتقرأ من متغيرات البيئة.
#
# الصيغة المطلوبة لكل متغير: "username:password,username2:password2"
# مثال: SEED_ADMINS="myadmin:MyStr0ngP@ssw0rd,secondadmin:AnotherStr0ngPass"
#
# لو المتغير مش موجود في البيئة (None) → مفيش أي تعديل على حسابات الدور ده
# خالص (مش هيتحذف ولا هيتضاف حد) — عشان الحسابات الموجودة فعلاً في قاعدة
# البيانات تفضل شغالة عادي زي ما هي من غير ما تتأثر بغياب المتغير.
def _parse_seed_accounts(env_var: str):
    raw = os.environ.get(env_var, "")
    if not raw.strip():
        return None  # المتغير غير مُعرَّف — تخطي مزامنة هذا الدور بالكامل
    accounts = []
    for pair in raw.split(","):
        pair = pair.strip()
        if not pair or ":" not in pair:
            continue
        uname, pw = pair.split(":", 1)
        uname, pw = uname.strip(), pw.strip()
        if uname and pw:
            accounts.append((uname, pw))
    return accounts

ADMINS     = _parse_seed_accounts("SEED_ADMINS")
REPORTERS  = _parse_seed_accounts("SEED_REPORTERS")
SUPERUSERS = _parse_seed_accounts("SEED_SUPERUSERS")

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


app = FastAPI(
    title="Fleet Management API",
    version="2.0.0",
    docs_url="/docs" if ENVIRONMENT != "production" else None,
    redoc_url=None,
)

@app.on_event("startup")
async def startup():
    migrate_db()
    with get_db() as conn:
        c = conn.cursor()
        if ADMINS is not None:
            _sync_accounts(c, ADMINS, "admin")
        else:
            log.info("ℹ️ SEED_ADMINS غير مُعرَّف — تم تخطي مزامنة حسابات الأدمن (الحسابات الحالية لم تتأثر)")
        if REPORTERS is not None:
            _sync_accounts(c, REPORTERS, "reporter")
        else:
            log.info("ℹ️ SEED_REPORTERS غير مُعرَّف — تم تخطي مزامنة حسابات المراقبين (الحسابات الحالية لم تتأثر)")
        if SUPERUSERS is not None:
            _sync_accounts(c, SUPERUSERS, "superuser")
        else:
            log.info("ℹ️ SEED_SUPERUSERS غير مُعرَّف — تم تخطي مزامنة حسابات السوبر يوزر (الحسابات الحالية لم تتأثر)")
    log.info("🚀 Fleet Management API started")

@app.get("/superuser/free-admins")
async def list_free_admins(cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        rows = conn.execute(
            "SELECT id,username,branch FROM users WHERE role='admin' AND (managed_by IS NULL OR managed_by=0) ORDER BY username"
        ).fetchall()
        return [dict(r) for r in rows]

# ══════════════════════════════════════════════════════
# 9. SUPER ADMIN — Role Management Endpoints
# ══════════════════════════════════════════════════════

class SuperAdminCreate(BaseModel):
    username: str
    password: str
    branch:   Optional[str] = ""
    managed_admin_ids: Optional[List[int]] = []   # الأدمنز التابعين له

class SuperAdminUpdate(BaseModel):
    username: Optional[str] = None
    password: Optional[str] = None
    branch:   Optional[str] = None
    managed_admin_ids: Optional[List[int]] = None


@app.get("/superuser/super-admins")
async def list_super_admins(cu: dict = Depends(require_superuser)):
    """قائمة كل السوبر أدمنز مع الأدمنز التابعين لهم وإحصائياتهم"""
    with get_db() as conn:
        c = conn.cursor()
        super_admins = [dict(r) for r in c.execute(
            "SELECT id,username,branch,created_at,last_login FROM users WHERE role='super_admin' ORDER BY id"
        ).fetchall()]
        for sa in super_admins:
            # الأدمنز التابعين
            managed = [dict(r) for r in c.execute(
                "SELECT id,username,branch,last_login FROM users WHERE managed_by=? AND role='admin'",
                (sa["id"],)
            ).fetchall()]
            sa["managed_admins"] = managed
            sa["managed_count"]  = len(managed)
            # إحصائيات: عدد السائقين والمركبات في نطاق الأدمنز التابعين
            admin_branches = list({m["branch"] for m in managed if m.get("branch")})
            if admin_branches:
                ph = ",".join("?" * len(admin_branches))
                sa["total_drivers"] = c.execute(
                    f"SELECT COUNT(*) FROM drivers WHERE branch IN ({ph})", admin_branches
                ).fetchone()[0]
                sa["total_cars"] = c.execute(
                    f"SELECT COUNT(*) FROM cars WHERE branch IN ({ph})", admin_branches
                ).fetchone()[0]
                sa["active_trips"] = c.execute(
                    "SELECT COUNT(*) FROM trips WHERE end_time IS NULL AND car_id IN "
                    f"(SELECT id FROM cars WHERE branch IN ({ph}))", admin_branches
                ).fetchone()[0]
            else:
                sa["total_drivers"] = 0
                sa["total_cars"]    = 0
                sa["active_trips"]  = 0
        return super_admins


@app.post("/superuser/super-admins", status_code=201)
async def create_super_admin(body: SuperAdminCreate, cu: dict = Depends(require_superuser)):
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        if c.execute("SELECT id FROM users WHERE username=?", (body.username,)).fetchone():
            raise HTTPException(409, "اسم المستخدم موجود بالفعل")
        hashed = hash_password(body.password)
        cur = c.execute(
            "INSERT INTO users(username,password,role,branch,created_at) VALUES(?,?,?,?,?)",
            (body.username, hashed, "super_admin", body.branch or "", now)
        )
        sa_id = cur.lastrowid
        # تعيين الأدمنز التابعين
        for admin_id in (body.managed_admin_ids or []):
            c.execute("UPDATE users SET managed_by=? WHERE id=? AND role='admin'", (sa_id, admin_id))
        log_event("super_admin_created", username=body.username, by=cu["username"])
        return {"id": sa_id, "username": body.username, "role": "super_admin"}


@app.put("/superuser/super-admins/{uid}")
async def update_super_admin(uid: int, body: SuperAdminUpdate, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        c = conn.cursor()
        if not c.execute("SELECT id FROM users WHERE id=? AND role='super_admin'", (uid,)).fetchone():
            raise HTTPException(404, "السوبر أدمن غير موجود")
        if body.username:
            ex = c.execute("SELECT id FROM users WHERE username=? AND id!=?", (body.username, uid)).fetchone()
            if ex: raise HTTPException(409, "اسم المستخدم مستخدم")
            c.execute("UPDATE users SET username=? WHERE id=?", (body.username, uid))
        if body.password:
            c.execute("UPDATE users SET password=? WHERE id=?", (hash_password(body.password), uid))
        if body.branch is not None:
            c.execute("UPDATE users SET branch=? WHERE id=?", (body.branch, uid))
        if body.managed_admin_ids is not None:
            # فك الربط القديم
            c.execute("UPDATE users SET managed_by=NULL WHERE managed_by=? AND role='admin'", (uid,))
            # الربط الجديد
            for admin_id in body.managed_admin_ids:
                c.execute("UPDATE users SET managed_by=? WHERE id=? AND role='admin'", (uid, admin_id))
        log_event("super_admin_updated", uid=uid, by=cu["username"])
        return {"ok": True}


@app.delete("/superuser/super-admins/{uid}")
async def delete_super_admin(uid: int, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        c = conn.cursor()
        if not c.execute("SELECT id FROM users WHERE id=? AND role='super_admin'", (uid,)).fetchone():
            raise HTTPException(404, "السوبر أدمن غير موجود")
        # فك ربط الأدمنز التابعين
        c.execute("UPDATE users SET managed_by=NULL WHERE managed_by=? AND role='admin'", (uid,))
        c.execute("DELETE FROM users WHERE id=?", (uid,))
        log_event("super_admin_deleted", uid=uid, by=cu["username"])
        return {"ok": True}


@app.get("/superuser/super-admins/{uid}/stats")
async def get_super_admin_stats(uid: int, cu: dict = Depends(require_superuser)):
    """إحصائيات تفصيلية لسوبر أدمن محدد"""
    with get_db() as conn:
        c = conn.cursor()
        sa = c.execute("SELECT * FROM users WHERE id=? AND role='super_admin'", (uid,)).fetchone()
        if not sa: raise HTTPException(404, "غير موجود")
        managed = [dict(r) for r in c.execute(
            "SELECT id,username,branch,last_login FROM users WHERE managed_by=? AND role='admin'", (uid,)
        ).fetchall()]
        branches = list({m["branch"] for m in managed if m.get("branch")})
        stats = {"managed_admins": managed, "branches": branches}
        if branches:
            ph = ",".join("?" * len(branches))
            stats["drivers"]      = c.execute(f"SELECT COUNT(*) FROM drivers WHERE branch IN ({ph})", branches).fetchone()[0]
            stats["cars"]         = c.execute(f"SELECT COUNT(*) FROM cars WHERE branch IN ({ph})", branches).fetchone()[0]
            stats["active_trips"] = c.execute("SELECT COUNT(*) FROM trips WHERE end_time IS NULL AND car_id IN (SELECT id FROM cars WHERE branch IN ("+ph+"))", branches).fetchone()[0]
            stats["total_trips_30d"] = c.execute("SELECT COUNT(*) FROM trips WHERE start_time >= date('now','-30 days') AND car_id IN (SELECT id FROM cars WHERE branch IN ("+ph+"))", branches).fetchone()[0]
            stats["workshop_30d"] = c.execute("SELECT COUNT(*) FROM workshop_records WHERE created_at >= date('now','-30 days') AND vehicle_id IN (SELECT id FROM cars WHERE branch IN ("+ph+"))", branches).fetchone()[0]
        return stats


# ══════════════════════════════════════════════════════
# 10. IMPORT: جدول الصيانة الدورية (maintenance_schedule)
# منطق UPSERT: لو المركبة+النوع موجودين بالفعل → تحديث
#              لو مش موجودين → إضافة سجل جديد
# ══════════════════════════════════════════════════════

def _validate_maintenance_row(row: dict, idx: int, cars_by_plate: dict, existing_map: dict) -> dict:
    """
    يتحقق من صحة صف صيانة دورية واحد.
    cars_by_plate: {plate_normalized: car_id}
    existing_map:  {(car_id, maintenance_type): schedule_id}  — للتحديد upsert أم insert
    """
    errors = []
    plate = (row.get("plate") or "").strip()
    maint_type = (row.get("maintenance_type") or "").strip()
    interval_km_raw   = row.get("interval_km")
    interval_days_raw = row.get("interval_days")
    last_km_raw   = row.get("last_done_km")
    last_date_raw = row.get("last_done_date")

    car_id = None
    if not plate:
        errors.append("رقم اللوحة مطلوب")
    else:
        car_id = cars_by_plate.get(plate.strip().lower())
        if not car_id:
            errors.append(f"المركبة '{plate}' غير موجودة في النظام")

    if not maint_type:
        errors.append("نوع الصيانة مطلوب")
    elif maint_type not in MAINTENANCE_TYPES:
        errors.append(f"نوع صيانة غير معروف: '{maint_type}'")

    def _to_float(v):
        try:
            if v is None or str(v).strip() in ("", "—", "None"):
                return None
            return float(str(v).replace(",", "").strip())
        except Exception:
            return None

    def _to_int(v):
        f = _to_float(v)
        return int(f) if f is not None else None

    interval_km   = _to_int(interval_km_raw)
    interval_days = _to_int(interval_days_raw)
    last_km       = _to_float(last_km_raw)
    last_date     = _normalize_date(str(last_date_raw)) if last_date_raw else ""

    if not interval_km and not interval_days:
        errors.append("يجب تحديد 'كل كام كم' أو 'كل كام يوم' على الأقل")

    is_update = False
    existing_id = None
    if car_id and maint_type and maint_type in MAINTENANCE_TYPES:
        existing_id = existing_map.get((car_id, maint_type))
        is_update = existing_id is not None

    return {
        "row_index":      idx,
        "plate":          plate,
        "maintenance_type": maint_type,
        "interval_km":    interval_km,
        "interval_days":  interval_days,
        "last_done_km":   last_km,
        "last_done_date": last_date,
        "car_id":         car_id,
        "existing_id":    existing_id,
        "is_update":      is_update,
        "valid":          len(errors) == 0,
        "errors":         errors,
    }


def _build_maintenance_maps(c):
    """يبني خرائط المركبات (باللوحة) وجداول الصيانة الموجودة فعلاً."""
    cars_by_plate = {}
    for r in c.execute("SELECT id, plate FROM cars").fetchall():
        if r["plate"]:
            cars_by_plate[r["plate"].strip().lower()] = r["id"]

    existing_map = {}
    for r in c.execute("SELECT id, car_id, maintenance_type FROM maintenance_schedule").fetchall():
        existing_map[(r["car_id"], r["maintenance_type"])] = r["id"]

    return cars_by_plate, existing_map


@app.post("/import/maintenance/preview")
async def import_maintenance_preview(
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
        cars_by_plate, existing_map = _build_maintenance_maps(conn.cursor())

    rows = [_validate_maintenance_row(r, i + 1, cars_by_plate, existing_map) for i, r in enumerate(raw_rows)]

    valid   = sum(1 for r in rows if r["valid"])
    updates = sum(1 for r in rows if r["valid"] and r["is_update"])
    inserts = sum(1 for r in rows if r["valid"] and not r["is_update"])

    return {
        "import_type": "maintenance",
        "rows": rows,
        "summary": {
            "total": len(rows), "valid": valid, "invalid": len(rows) - valid,
            "will_update": updates, "will_insert": inserts,
        },
    }


@app.post("/import/maintenance/confirm")
async def import_maintenance_confirm(
    file: UploadFile = File(...),
    skip_errors: bool = False,
    cu: dict = Depends(require_admin)
):
    content = await file.read()
    raw_rows = _parse_upload_file(content, file.filename or "file.csv")
    if not raw_rows:
        raise HTTPException(400, "الملف فارغ")

    with get_db() as conn:
        cars_by_plate, existing_map = _build_maintenance_maps(conn.cursor())

    rows = [_validate_maintenance_row(r, i + 1, cars_by_plate, existing_map) for i, r in enumerate(raw_rows)]

    invalid = [r for r in rows if not r["valid"]]
    if invalid and not skip_errors:
        raise HTTPException(422, {"message": f"يوجد {len(invalid)} صف بأخطاء", "invalid_rows": invalid})

    now = datetime.utcnow().isoformat() + "Z"
    inserted_items, updated_items, skipped_rows = [], [], []

    with get_db() as conn:
        c = conn.cursor()
        for row in [r for r in rows if r["valid"]]:
            try:
                if row["is_update"]:
                    # ── UPSERT: تحديث السجل الموجود بدل تجاهله ──
                    c.execute("""UPDATE maintenance_schedule
                                 SET interval_km=COALESCE(?, interval_km),
                                     interval_days=COALESCE(?, interval_days),
                                     last_done_km=COALESCE(?, last_done_km),
                                     last_done_date=CASE WHEN ?!='' THEN ? ELSE last_done_date END,
                                     updated_at=?
                                 WHERE id=?""",
                              (row["interval_km"], row["interval_days"], row["last_done_km"],
                               row["last_done_date"], row["last_done_date"], now, row["existing_id"]))
                    updated_items.append({
                        "row_index": row["row_index"], "plate": row["plate"],
                        "maintenance_type": row["maintenance_type"],
                    })
                else:
                    c.execute("""INSERT INTO maintenance_schedule
                                 (car_id,maintenance_type,interval_km,interval_days,
                                  last_done_km,last_done_date,alert_km_before,alert_days_before,
                                  notes,created_at,updated_at)
                                 VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                              (row["car_id"], row["maintenance_type"], row["interval_km"], row["interval_days"],
                               row["last_done_km"], row["last_done_date"] or None,
                               500, 7, "", now, now))
                    inserted_items.append({
                        "row_index": row["row_index"], "plate": row["plate"],
                        "maintenance_type": row["maintenance_type"],
                    })
            except Exception as e:
                row["errors"] = [str(e)]
                skipped_rows.append(row)

    write_audit_log(cu["user_id"], cu["username"], cu["role"], "bulk_import_maintenance",
                    f"استيراد صيانة دورية: {len(inserted_items)} إضافة، {len(updated_items)} تحديث")
    log_event("maintenance_bulk_import", inserted=len(inserted_items), updated=len(updated_items),
              skipped=len(skipped_rows), admin=cu["username"])

    return {
        "import_type": "maintenance",
        "inserted": len(inserted_items),
        "updated":  len(updated_items),
        "skipped":  len(skipped_rows),
        "inserted_items": inserted_items,
        "updated_items":  updated_items,
        "skipped_items":  skipped_rows,
    }


@app.get("/import/template/maintenance")
async def import_maintenance_template(cu: dict = Depends(require_admin)):
    cols   = ["رقم اللوحة", "نوع الصيانة", "كل كام كم", "كل كام يوم", "آخر قراءة", "تاريخ آخر صيانة"]
    sample = ["أ-ب-1234", "تغيير زيت", "5000", "90", "120000", "2026-05-01"]
    output = io.StringIO()
    csv.writer(output).writerows([cols, sample])
    from fastapi.responses import Response
    return Response(
        content=output.getvalue().encode("utf-8-sig"),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=maintenance_template.csv"},
    )


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
# 11. GLOBAL EXCEPTION HANDLER
# ══════════════════════════════════════════════════════

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    log.error(f"Unhandled error | path={request.url.path} | {type(exc).__name__}: {exc}")
    return JSONResponse(
        status_code=500,
        content={"detail": "خطأ داخلي في الخادم — يرجى المحاولة لاحقاً"}
    )

# ══════════════════════════════════════════════════════
# 12. HELPERS
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
# 13. AUTH ENDPOINTS
# ══════════════════════════════════════════════════════

@app.post("/login", response_model=LoginResp)
@limiter.limit("10/minute")   # strict: prevent brute-force
async def login(request: Request, data: LoginReq):
    with get_db() as conn:
        c = conn.cursor()
        # username ممكن يتكرر عند السائقين — نجيب كل الحسابات بنفس الاسم
        # ونبحث عن الأول اللي كلمة مروره تطابق
        c.execute("SELECT id,username,password,role,branch,avatar_url,failed_attempts,locked_until "
                  "FROM users WHERE username=?", (data.username,))
        candidates = c.fetchall()

        # ── فحص القفل: لو أي حساب بنفس اليوزرنيم مقفول حالياً، نرفض فوراً ──
        now = datetime.utcnow()
        for cand in candidates:
            if cand["locked_until"]:
                locked_until_dt = datetime.fromisoformat(cand["locked_until"])
                if locked_until_dt > now:
                    remaining_min = max(1, int((locked_until_dt - now).total_seconds() // 60) + 1)
                    log_event("login_blocked_locked", username=data.username, ip=request.client.host)
                    raise HTTPException(
                        423,
                        f"الحساب مقفول مؤقتاً بسبب محاولات دخول فاشلة متكررة — "
                        f"حاول مرة أخرى بعد {remaining_min} دقيقة"
                    )

        u = None
        for candidate in candidates:
            if _verify(data.password, candidate["password"]):
                u = candidate
                break

        if not u:
            # ── تسجيل محاولة فاشلة وزيادة العداد على كل الحسابات المطابقة للاسم ──
            for cand in candidates:
                new_count = (cand["failed_attempts"] or 0) + 1
                if new_count >= MAX_FAILED_LOGIN_ATTEMPTS:
                    lock_until = (now + timedelta(minutes=LOCKOUT_MINUTES)).isoformat()
                    c.execute("UPDATE users SET failed_attempts=?, locked_until=? WHERE id=?",
                              (new_count, lock_until, cand["id"]))
                    log_event("account_locked", user_id=cand["id"], username=cand["username"],
                              ip=request.client.host)
                else:
                    c.execute("UPDATE users SET failed_attempts=? WHERE id=?", (new_count, cand["id"]))
            log_event("login_failed", username=data.username, ip=request.client.host)
            raise HTTPException(401, "اسم المستخدم أو كلمة المرور غير صحيحة")

        # ── تسجيل دخول ناجح: تصفير عداد المحاولات الفاشلة وفك القفل ──
        if (u["failed_attempts"] or 0) > 0 or u["locked_until"]:
            c.execute("UPDATE users SET failed_attempts=0, locked_until=NULL WHERE id=?", (u["id"],))

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

        # Persist refresh token (one per user) — نخزن الـ hash فقط
        c.execute("UPDATE users SET refresh_token=?,refresh_exp=?,last_login=? WHERE id=?",
                  (_hash_refresh_token(refresh_token), refresh_exp, datetime.utcnow().isoformat() + "Z", u["id"]))

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
                  (_hash_refresh_token(body.refresh_token),))
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
                  (_hash_refresh_token(new_refresh), new_exp, u["id"]))
        return {"access_token": new_access, "refresh_token": new_refresh, "token_type": "bearer"}

@app.post("/logout")
async def logout(cu: dict = Depends(get_user)):
    with get_db() as conn:
        conn.cursor().execute("UPDATE users SET refresh_token=NULL,refresh_exp=NULL WHERE id=?",
                              (cu["user_id"],))
    return {"ok": True}

# ══════════════════════════════════════════════════════
# 14. USER AVATAR — رفع وجلب صورة البروفايل للمستخدم
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

    fname = f"avatar_user_{cu['user_id']}_{uuid.uuid4().hex}.{ext}"
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
# 15. AUDIO UPLOAD — File system, NOT base64 in DB
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
# 16. DRIVERS
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
                         ORDER BY d.id DESC""", (branch,))
        else:
            c.execute("""SELECT d.*,u.username FROM drivers d
                         LEFT JOIN users u ON d.user_id=u.id
                         ORDER BY d.id DESC""")
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
        validate_driver_password(new_password)
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
# 17. CARS
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
            c.execute(base_q + " WHERE branch=? ORDER BY id DESC", (branch,))
        else:
            c.execute(base_q + " ORDER BY id DESC")
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

@app.get("/cars/{cid}/drivers")
async def car_drivers(cid: int, cu: dict = Depends(get_user)):
    """
    كل السائقين المصرح لهم بقيادة مركبة معينة.
    مفيد بشكل خاص لمركبات (سيارات متعددة السائقين) — طوارئ / إطفاء / إسعاف —
    التي يمكن أن يُصرَّح لأكثر من سائق بقيادتها في نفس الوقت.
    """
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT d.id,d.name,d.phone,d.status,p.id AS permission_id
                     FROM drivers d
                     INNER JOIN driver_car_permissions p ON d.id=p.driver_id
                     WHERE p.car_id=?
                     ORDER BY d.name""", (cid,))
        return [dict(r) for r in c.fetchall()]

# ══════════════════════════════════════════════════════
# 18. PERMISSIONS
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
async def reassign_permission(body: dict, cu: dict = Depends(require_admin)):
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
        if cu["role"] == "admin" and cu.get("branch"):
            if new_driver["branch"] != cu["branch"]:
                raise HTTPException(403, "السائق ليس من فرعك — يمكنك فقط إدارة سائقي فرعك")

        # جيب المركبة
        c.execute("SELECT id, plate, model, car_name, branch FROM cars WHERE plate=?", (plate,))
        car = c.fetchone()
        if not car:
            raise HTTPException(404, f"لا توجد مركبة بلوحة '{plate}'")
        if cu["role"] == "admin" and cu.get("branch"):
            if car["branch"] and car["branch"] != cu["branch"]:
                raise HTTPException(403, "المركبة ليست من فرعك — يمكنك فقط إدارة مركبات فرعك")

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
# 19. DRIVER PHOTOS — رفع وعرض صور السائقين
# ══════════════════════════════════════════════════════

@app.post("/driver/photo/upload")
async def upload_driver_photo(body: dict, cu: dict = Depends(get_user)):
    """السائق يرفع صورة (base64) — تُحفظ وتُعرض عند السوبر يوزر."""
    if cu["role"] not in ("driver", "operator"):
        raise HTTPException(403, "للسائقين والمشغلين فقط")

    driver_id = cu.get("driver_id")
    if not driver_id and cu["role"] == "operator":
        driver_id = cu.get("operator_id")
    if not driver_id and cu["role"] == "driver":
        with get_db() as _c:
            row = _c.execute("SELECT id FROM drivers WHERE user_id=?", (cu["user_id"],)).fetchone()
            driver_id = row["id"] if row else None
    if not driver_id:
        raise HTTPException(400, "لا يوجد حساب مرتبط")

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
    fname = f"photo_{driver_id}_{uuid.uuid4().hex}.{ext}"
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
# 20. LIVE LOCATION — real-time driver tracking
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
# 21. TRIPS — with pagination
# ══════════════════════════════════════════════════════

@app.get("/trips")
async def get_trips(cu: dict = Depends(get_user), branch: Optional[str] = None):
    with get_db() as conn:
        c = conn.cursor()
        branch = _effective_branch(cu, branch)
        if cu["role"] in ("admin", "reporter", "superuser"):
            # عرض كل الرحلات بدون أي حد أقصى (مطلوب من الإدارة لرؤية كل البيانات في صفحة الرحلات وخريطة الأسطول)
            if branch:
                c.execute("""SELECT t.* FROM trips t
                             JOIN drivers d ON t.driver_id=d.id
                             WHERE d.branch=?
                             ORDER BY t.id DESC""", (branch,))
            else:
                c.execute("SELECT * FROM trips ORDER BY id DESC")
        else:
            did = cu.get("driver_id")
            if not did: return []
            c.execute("SELECT * FROM trips WHERE driver_id=? ORDER BY id DESC", (did,))
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
        c.execute("UPDATE trips SET end_time=?,end_odometer=?,end_location=?,notes=? WHERE id=?",
                  (now, end_odo, loc, notes, active["id"]))
        c.execute("SELECT * FROM trips WHERE id=?", (active["id"],))
        return enrich(dict(c.fetchone()))

# ══════════════════════════════════════════════════════
# 22. USERS
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
        c.execute("SELECT id,username,role,branch,last_login FROM users ORDER BY id DESC")
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
# 23. WORKSHOPS
# ══════════════════════════════════════════════════════

@app.post("/workshops")
async def create_workshop(rec: WorkshopCreate, cu: dict = Depends(get_user)):
    if rec.type not in WORKSHOP_TYPES:
        raise HTTPException(400, "نوع غير صالح")
    # المشغل: لا يحتاج driver_id صالح في جدول drivers — يُسمح له دائماً
    if cu["role"] not in ("operator", "admin", "superuser"):
        if rec.driver_id != cu.get("driver_id"):
            raise HTTPException(403, "غير مصرح")
    # ── احسب السعر الإجمالي بشكل موحّد لكل الأدوار ──
    # الأنواع اللي ليها كمية: الوقود بكل أنواعه + الزيت + الكوتش
    QUANTITY_TYPES = {t for t in WORKSHOP_TYPES if t.startswith("fuel_") or t.startswith("oil") or t in ("tire",)}
    HAS_QUANTITY   = rec.type in QUANTITY_TYPES
    IS_FUEL_TYPE   = rec.type.startswith("fuel_")
    IS_PARTS_TYPE  = (not IS_FUEL_TYPE) and rec.type != "other"

    # ── حد أقصى لكمية الوقود للسيارات: 200 لتر في التفويلة الواحدة ──
    if IS_FUEL_TYPE and rec.quantity and rec.quantity > 200:
        raise HTTPException(400, "الحد الأقصى لكمية الوقود للسيارات هو 200 لتر")

    # ── صورة العداد (تفويل/صيانة): إلزامية لأنواع الوقود + الصيانة، اختيارية لباقي الأنواع ──
    PHOTO_REQUIRED_TYPES = {t for t in WORKSHOP_TYPES if t.startswith("fuel_") or t.startswith("oil")
                             or t.startswith("filter") or t in ("tire","battery","belt")}
    if rec.type in PHOTO_REQUIRED_TYPES and not (rec.odometer_photo or "").strip():
        raise HTTPException(400, "صورة العداد مطلوبة لتسجيل هذه العملية")

    # ── Virtual Inventory: مصدر الوقود (إلزامي لكل أنواع التفويل) ──
    if IS_FUEL_TYPE:
        if rec.fuel_source not in ("CompanyStation", "ExternalStation", "SupplyPlatform"):
            raise HTTPException(400, "اختر مصدر الوقود: محطة الشركة أو محطة خارجية أو منصة إمداد")
        if rec.fuel_source == "SupplyPlatform":
            if not rec.platform_id:
                raise HTTPException(400, "اختر منصة الإمداد")
            if not rec.inventory_product_id:
                raise HTTPException(400, "اختر نوع الوقود المسحوب من منصة الإمداد")

    # ── Virtual Inventory: مصدر القطعة (إلزامي لكل أنواع الصيانة، إلا "أخرى") ──
    if IS_PARTS_TYPE:
        if rec.parts_source not in ("Inventory", "DirectPurchase", "SupplyPlatform"):
            raise HTTPException(400, "اختر مصدر القطعة: من المخزن أو شراء مباشر من الخارج أو منصة إمداد")
        if rec.parts_source == "Inventory" and not rec.inventory_product_id:
            raise HTTPException(400, "اختر المنتج من المخزون")
        if rec.parts_source == "DirectPurchase" and not (rec.direct_purchase_supplier or "").strip():
            raise HTTPException(400, "اسم المورد أو المحل مطلوب للشراء المباشر")
        if rec.parts_source == "SupplyPlatform":
            if not rec.platform_id:
                raise HTTPException(400, "اختر منصة الإمداد")
            if not rec.inventory_product_id:
                raise HTTPException(400, "اختر القطعة المسحوبة من منصة الإمداد")

    def _save_b64_image(b64_data, prefix):
        if not b64_data:
            return ""
        import re as _re_ws
        try:
            _m = _re_ws.match(r"data:image/(\w+);base64,(.+)", b64_data, _re_ws.DOTALL)
            _ext = _m.group(1) if _m else "jpg"
            _b64data = _m.group(2) if _m else b64_data
            _raw_bytes = base64.b64decode(_b64data)
            if len(_raw_bytes) >= 2 and _raw_bytes[0] == 0xFF and _raw_bytes[1] == 0xD8: _ext = "jpg"
            elif len(_raw_bytes) >= 4 and _raw_bytes[0] == 0x89 and _raw_bytes[1:4] == b"PNG": _ext = "png"
            elif len(_raw_bytes) >= 3 and _raw_bytes[:3] == b"GIF": _ext = "gif"
            elif len(_raw_bytes) >= 12 and _raw_bytes[:4] == b"RIFF" and _raw_bytes[8:12] == b"WEBP": _ext = "webp"
            _fname = f"{prefix}_{rec.driver_id or 0}_{uuid.uuid4().hex}.{_ext}"
            (UPLOAD_DIR / _fname).write_bytes(_raw_bytes)
            return f"/uploads/{_fname}"
        except Exception as _photo_err:
            log.error(f"[WORKSHOP] photo save failed ({prefix}): {_photo_err}")
            raise HTTPException(400, "صورة غير صالحة")

    odometer_photo_url = _save_b64_image(rec.odometer_photo, "odo")
    invoice_photo_url  = _save_b64_image(rec.invoice_photo, "invoice")

    with get_db() as conn:
        c = conn.cursor()
        # جيب الفرع لتحديد السعر — المشغل من equipment_operators بالـ token operator_id
        if cu["role"] == "operator":
            c.execute("SELECT branch FROM equipment_operators WHERE id=?", (cu.get("operator_id"),))
        else:
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

    # ── شراء مباشر (تفويل خارجي أو قطعة من الخارج): التكلفة من المستخدم نفسه، مش من إعدادات الأسعار ──
    if IS_FUEL_TYPE and rec.fuel_source == "ExternalStation" and rec.price and rec.price > 0:
        unit_price_cfg = rec.price
    if IS_PARTS_TYPE and rec.parts_source == "DirectPurchase" and rec.direct_purchase_cost:
        unit_price_cfg = rec.direct_purchase_cost

    if cu["role"] in ("admin", "superuser"):
        # الأدمن يرسل price من الإعدادات (سعر/وحدة) — نضرب في الكمية
        unit_price = rec.price if rec.price and rec.price > 0 else unit_price_cfg
    else:
        # السائق: دايماً سعر الإعدادات (أو التكلفة المُدخلة في حالة الشراء المباشر)
        unit_price = unit_price_cfg

    qty = rec.quantity if rec.quantity and rec.quantity > 0 else 1.0
    # الإجمالي = سعر الوحدة × الكمية (للأنواع ذات الكمية) أو سعر الوحدة مباشرة
    final_price = unit_price * qty if HAS_QUANTITY else unit_price
    if IS_PARTS_TYPE and rec.parts_source == "DirectPurchase" and rec.direct_purchase_cost:
        final_price = rec.direct_purchase_cost
    if final_price < 0:
        raise HTTPException(400, "السعر لا يمكن أن يكون سالباً")
    now = datetime.utcnow().isoformat() + "Z"
    with get_db() as conn:
        c = conn.cursor()
        # ضمان وجود الأعمدة الجديدة في الجدول (ميجريشن دفاعي)
        _ws_extra_cols = [
            ("doc_number", "TEXT DEFAULT ''"),
            ("engine_hours", "REAL"),
            ("supply_source", "TEXT DEFAULT ''"),
            ("item_name", "TEXT DEFAULT ''"),
            ("item_spec", "TEXT DEFAULT ''"),
            ("receiver_name", "TEXT DEFAULT ''"),
            ("operator_id", "INTEGER"),
            ("is_operator", "INTEGER DEFAULT 0"),
            ("odometer_photo", "TEXT DEFAULT ''"),
            ("fuel_source", "TEXT DEFAULT ''"),
            ("station_name", "TEXT DEFAULT ''"),
            ("pump_number", "TEXT DEFAULT ''"),
            ("invoice_photo", "TEXT DEFAULT ''"),
            ("parts_source", "TEXT DEFAULT ''"),
            ("inventory_product_id", "INTEGER"),
            ("direct_purchase_supplier", "TEXT DEFAULT ''"),
            ("direct_purchase_cost", "REAL"),
            ("approval_status", "TEXT DEFAULT 'pending'"),
            ("reviewed_by", "TEXT DEFAULT ''"),
            ("reviewed_at", "TEXT DEFAULT ''"),
            ("price_edited_by", "TEXT DEFAULT ''"),
            ("price_edited_at", "TEXT DEFAULT ''"),
            ("quantity_edited_by", "TEXT DEFAULT ''"),
            ("quantity_edited_at", "TEXT DEFAULT ''"),
            ("platform_id", "INTEGER"),
        ]
        _ws_existing = {r["name"] for r in c.execute("PRAGMA table_info(workshop_records)").fetchall()}
        for _wc, _wd in _ws_extra_cols:
            if _wc not in _ws_existing:
                try: c.execute(f"ALTER TABLE workshop_records ADD COLUMN {_wc} {_wd}")
                except Exception: pass

        # ── Virtual Inventory: لو المصدر "من المخزن" تأكد من توافر الكمية ──
        if IS_PARTS_TYPE and rec.parts_source == "Inventory":
            prow = c.execute("SELECT quantity, name FROM inventory_products WHERE id=?", (rec.inventory_product_id,)).fetchone()
            if not prow:
                raise HTTPException(404, "المنتج غير موجود في المخزون")
            if (prow["quantity"] or 0) < qty:
                raise HTTPException(400, f"الكمية المتاحة من «{prow['name']}» غير كافية (متاح {prow['quantity']})")

        # ── منصات الإمداد: لو المصدر "منصة إمداد" (وقود أو قطعة) تأكد من توافر الكمية في تانك المنصة ──
        if (IS_FUEL_TYPE and rec.fuel_source == "SupplyPlatform") or (IS_PARTS_TYPE and rec.parts_source == "SupplyPlatform"):
            plat_row = c.execute("SELECT code, status FROM supply_platforms WHERE id=?", (rec.platform_id,)).fetchone()
            if not plat_row:
                raise HTTPException(404, "منصة الإمداد غير موجودة")
            if plat_row["status"] != "active":
                raise HTTPException(400, "منصة الإمداد غير نشطة")
            pstock = c.execute("""SELECT s.quantity, p.name, p.unit FROM supply_platform_stock s
                                   JOIN inventory_products p ON p.id = s.product_id
                                   WHERE s.platform_id=? AND s.product_id=?""",
                                (rec.platform_id, rec.inventory_product_id)).fetchone()
            if not pstock or (pstock["quantity"] or 0) < qty:
                avail = pstock["quantity"] if pstock else 0
                pname = pstock["name"] if pstock else "الصنف"
                raise HTTPException(400, f"الكمية المتاحة من «{pname}» في منصة الإمداد «{plat_row['code']}» غير كافية (متاح {avail})")

        # المشغل: driver_id يبقى NULL، نحفظ operator_id بشكل منفصل
        try:
            if cu["role"] == "operator":
                op_id_ws = cu.get("operator_id") or 0
                log.info(f"[WORKSHOP] operator insert: operator_id={op_id_ws} type={rec.type}")
                c.execute("""INSERT INTO workshop_records
                             (driver_id,operator_id,is_operator,type,quantity,price,notes,created_at,
                              operation_type,vehicle_id,odometer_reading,description,tire_action,location,
                              doc_number,engine_hours,supply_source,item_name,item_spec,receiver_name,odometer_photo,
                              fuel_source,station_name,pump_number,invoice_photo,
                              parts_source,inventory_product_id,direct_purchase_supplier,direct_purchase_cost,approval_status,
                              platform_id)
                             VALUES(NULL,?,1,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'pending',?)""",
                          (op_id_ws, rec.type, rec.quantity, final_price, rec.notes or "", now,
                           rec.operation_type or "", rec.vehicle_id, rec.odometer_reading,
                           rec.description or "", rec.tire_action or "", rec.location or "",
                           rec.doc_number or "", rec.engine_hours, rec.supply_source or "",
                           rec.item_name or "", rec.item_spec or "", rec.receiver_name or "",
                           odometer_photo_url,
                           rec.fuel_source or "", rec.station_name or "", rec.pump_number or "", invoice_photo_url,
                           rec.parts_source or "", rec.inventory_product_id, rec.direct_purchase_supplier or "", rec.direct_purchase_cost,
                           rec.platform_id))
            else:
                c.execute("""INSERT INTO workshop_records
                             (driver_id,is_operator,type,quantity,price,notes,created_at,
                              operation_type,vehicle_id,odometer_reading,description,tire_action,location,
                              doc_number,engine_hours,supply_source,item_name,item_spec,receiver_name,odometer_photo,
                              fuel_source,station_name,pump_number,invoice_photo,
                              parts_source,inventory_product_id,direct_purchase_supplier,direct_purchase_cost,approval_status,
                              platform_id)
                             VALUES(?,0,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'pending',?)""",
                      (rec.driver_id, rec.type, rec.quantity, final_price, rec.notes or "", now,
                       rec.operation_type or "", rec.vehicle_id, rec.odometer_reading,
                       rec.description or "", rec.tire_action or "", rec.location or "",
                       rec.doc_number or "", rec.engine_hours, rec.supply_source or "",
                       rec.item_name or "", rec.item_spec or "", rec.receiver_name or "",
                       odometer_photo_url,
                       rec.fuel_source or "", rec.station_name or "", rec.pump_number or "", invoice_photo_url,
                       rec.parts_source or "", rec.inventory_product_id, rec.direct_purchase_supplier or "", rec.direct_purchase_cost,
                       rec.platform_id))
            rid = c.lastrowid
        except Exception as _ws_err:
            log.error(f"[WORKSHOP] INSERT failed: {_ws_err}")
            raise HTTPException(500, f"خطأ في حفظ سجل الورشة: {str(_ws_err)}")

        # ── Virtual Inventory: تنفيذ حركة المخزون حسب مصدر القطعة ──
        if IS_PARTS_TYPE and rec.parts_source == "Inventory":
            c.execute("UPDATE inventory_products SET quantity = quantity - ?, updated_at=? WHERE id=?",
                      (qty, now, rec.inventory_product_id))
            c.execute("""INSERT INTO inventory_movements(product_id,movement_type,quantity,ref_type,ref_id,username,notes,created_at)
                         VALUES(?,'out',?,'workshop',?,?,?,?)""",
                      (rec.inventory_product_id, qty, rid, cu.get("username",""), rec.notes or "", now))

        # ── منصات الإمداد: سحب الكمية من رصيد المنصة (وقود أو قطعة) وتسجيل حركة "سحب" ──
        if (IS_FUEL_TYPE and rec.fuel_source == "SupplyPlatform") or (IS_PARTS_TYPE and rec.parts_source == "SupplyPlatform"):
            c.execute("""UPDATE supply_platform_stock SET quantity = quantity - ?, updated_at=?
                         WHERE platform_id=? AND product_id=?""",
                      (qty, now, rec.platform_id, rec.inventory_product_id))
            c.execute("""INSERT INTO supply_platform_movements(platform_id,product_id,movement_type,quantity,ref_type,ref_id,username,notes,created_at)
                         VALUES(?,?,'draw',?,'workshop',?,?,?,?)""",
                      (rec.platform_id, rec.inventory_product_id, qty, rid, cu.get("username",""), rec.notes or "", now))

        vp = None
        if rec.vehicle_id:
            c.execute("SELECT plate FROM cars WHERE id=?", (rec.vehicle_id,))
            cv = c.fetchone()
            vp = cv["plate"] if cv else None
        # ── تحديث جدول الصيانة الدورية تلقائياً ──
        _override_km = rec.odometer_reading
        if rec.vehicle_id and not _override_km:
            # اجيب آخر عداد من الرحلات لو السائق ما حطش قراءة
            _km_row = c.execute("SELECT MAX(end_odometer) FROM trips WHERE car_id=? AND end_odometer IS NOT NULL", (rec.vehicle_id,)).fetchone()
            if _km_row and _km_row[0]: _override_km = _km_row[0]
        if rec.vehicle_id and _override_km:
            MAINT_MAP = {
                # ── زيوت ──
                "oil":        ["تغيير زيت", "فلتر زيت"],
                "oil_motor":  ["تغيير زيت"],
                "oil_gear":   ["زيت تروس"],
                "oil_brake":  ["زيت فرامل"],
                "oil_hydro":  ["زيت هيدروليك"],
                "oil_grease": ["شحم", "تشحيم"],
                # ── فلاتر ──
                "filter":         ["فلتر هواء", "فلتر وقود", "فلتر مكيف"],
                "filter_oil":     ["فلتر زيت"],
                "filter_solar":   ["فلتر وقود", "فلتر سولار"],
                "filter_petrol":  ["فلتر بنزين", "فلتر وقود"],
                "filter_air":     ["فلتر هواء"],
                "filter_separator":["فلتر فاصل"],
                "filter_ac":      ["فلتر مكيف", "فلتر تكيف"],
                "filter_dryer":   ["فلتر مجفف"],
                "filter_breather":["فلتر منفس"],
                "filter_hydraulic":["فلتر هيدروليك"],
                # ── قطع وإطارات ──
                "tire":    ["إطارات"],
                "battery": ["بطارية"],
                "belt":    ["سير توزيع", "سيور"],
                # ── عمرات وصيانة كبيرة ──
                "timing_belt":           ["سير كاتينة", "سير توزيع"],
                "engine_overhaul":       ["عمرة محرك", "عمرة كاملة"],
                "engine_half_overhaul":  ["نصف عمرة محرك", "عمرة نصفية"],
                "head_gasket":           ["جوان وش سلندر", "سلندر"],
                "fuel_pump_overhaul":    ["عمرة طرمبة وقود", "طرمبة وقود", "وحدات الحقن"],
            }
            related_types = MAINT_MAP.get(rec.type, [])
            if related_types:
                placeholders_m = ",".join("?" * len(related_types))
                c.execute(f"""UPDATE maintenance_schedule
                              SET last_done_km=?, last_done_date=?, updated_at=?
                              WHERE car_id=? AND maintenance_type IN ({placeholders_m})""",
                          [_override_km, now[:10], now, rec.vehicle_id] + related_types)
                updated_m = conn.execute("SELECT changes()").fetchone()[0]
                if updated_m:
                    log_event("maintenance_auto_updated",
                              car_id=rec.vehicle_id, types=related_types, km=_override_km)

        eff_driver_id = None if cu["role"] == "operator" else rec.driver_id
        return {"id":rid,"driver_id":eff_driver_id,"operator_id":cu.get("operator_id") if cu["role"]=="operator" else None,
                "type":rec.type,"quantity":rec.quantity,
                "price":final_price,"notes":rec.notes or "","created_at":now,
                "operation_type":rec.operation_type or "","vehicle_id":rec.vehicle_id,
                "odometer_reading":rec.odometer_reading,"description":rec.description or "",
                "tire_action":rec.tire_action or "","vehicle_plate":vp,"location":rec.location or "",
                "odometer_photo":odometer_photo_url,
                "fuel_source":rec.fuel_source or "","station_name":rec.station_name or "",
                "invoice_photo":invoice_photo_url,"parts_source":rec.parts_source or "",
                "approval_status":"pending"}


@app.get("/workshops")
async def get_workshops(cu: dict = Depends(get_user), branch: Optional[str] = None, month: Optional[str] = None):
    q = """SELECT w.*,
                  COALESCE(d.name, op.name) as driver_name,
                  c.plate as vehicle_plate,
                  sp.code as platform_code,
                  ip.name as platform_product_name
           FROM workshop_records w
           LEFT JOIN drivers d ON w.driver_id=d.id AND (w.is_operator IS NULL OR w.is_operator=0)
           LEFT JOIN equipment_operators op ON w.operator_id=op.id
           LEFT JOIN cars c ON w.vehicle_id=c.id
           LEFT JOIN supply_platforms sp ON w.platform_id=sp.id
           LEFT JOIN inventory_products ip ON w.inventory_product_id=ip.id"""
    with get_db() as conn:
        c = conn.cursor()
        branch = _effective_branch(cu, branch)
        if cu["role"] in ("admin", "reporter", "superuser", "workshop_admin"):
            conditions = []
            params = []
            if branch:
                conditions.append("d.branch=?")
                params.append(branch)
            if month:
                conditions.append("w.created_at LIKE ?")
                params.append(month + "%")
            where = (" WHERE " + " AND ".join(conditions)) if conditions else ""
            c.execute(q + where + " ORDER BY w.id DESC", params)
        elif cu["role"] == "operator":
            op_id3 = cu.get("operator_id")
            if not op_id3: return []
            month_cond = " AND w.created_at LIKE ?" if month else ""
            month_params = [month + "%"] if month else []
            # المشغل: operator_id column أو driver_id=0 sentinel
            c.execute(q + " WHERE (w.operator_id=? OR (w.driver_id=0 AND w.is_operator=1 AND w.operator_id IS NULL AND w.created_at IS NOT NULL))" + month_cond + " ORDER BY w.id DESC",
                      [op_id3] + month_params)
        else:
            did = cu.get("driver_id")
            if not did: return []
            month_cond = " AND w.created_at LIKE ?" if month else ""
            month_params = [month + "%"] if month else []
            c.execute(q + " WHERE w.driver_id=?" + month_cond + " ORDER BY w.id DESC", [did] + month_params)
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

@app.get("/workshops/attachments")
async def get_workshop_attachments(cu: dict = Depends(get_user), branch: Optional[str] = None,
                                    month: Optional[str] = None, type: Optional[str] = None):
    """تتبع المرفقات — قائمة سجلات الورشة (تفويل/صيانة) المرفق بها صورة عداد، للمراجعة من لوحة الإدارة فقط."""
    if cu["role"] not in ("admin", "superuser", "reporter", "workshop_admin"):
        raise HTTPException(403, "غير مصرح")
    q = """SELECT w.id, w.type, w.quantity, w.price, w.notes, w.created_at,
                  w.operation_type, w.vehicle_id, w.odometer_reading, w.description,
                  w.tire_action, w.location, w.odometer_photo, w.invoice_photo,
                  COALESCE(d.name, op.name) as driver_name,
                  c.plate as vehicle_plate, c.model as vehicle_model,
                  d.branch as branch
           FROM workshop_records w
           LEFT JOIN drivers d ON w.driver_id=d.id AND (w.is_operator IS NULL OR w.is_operator=0)
           LEFT JOIN equipment_operators op ON w.operator_id=op.id
           LEFT JOIN cars c ON w.vehicle_id=c.id
           WHERE w.odometer_photo IS NOT NULL AND w.odometer_photo != ''"""
    params = []
    branch = _effective_branch(cu, branch)
    if branch:
        q += " AND d.branch=?"
        params.append(branch)
    if month:
        q += " AND w.created_at LIKE ?"
        params.append(month + "%")
    if type:
        q += " AND w.type=?"
        params.append(type)
    q += " ORDER BY w.id DESC"
    with get_db() as conn:
        c = conn.cursor()
        c.execute(q, params)
        return [dict(r) for r in c.fetchall()]

# ══════════════════════════════════════════════════════
# 24. OPERATIONAL CARD PAGE 2 — زيوت وكاوتش وبطاريات
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

    OIL_TYPES     = {"oil", "oil_motor", "oil_gear", "oil_brake", "oil_hydro", "oil_grease",
                     "filter", "filter_oil", "filter_solar", "filter_petrol", "filter_air",
                     "filter_separator", "filter_ac", "filter_dryer", "filter_breather", "filter_hydraulic"}
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
        "تغيير زيت":   "motor",
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
# 25. GARAGE RECORDS
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
            c.execute(q + " ORDER BY g.id DESC")
        else:
            did = cu.get("driver_id")
            if not did: return []
            c.execute(q + " WHERE g.driver_id=? ORDER BY g.id DESC", (did,))
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
# 26. EMERGENCY — Audio stored as file, not base64
# ══════════════════════════════════════════════════════

@app.post("/emergency/report")
@limiter.limit("5/minute")  # very strict on emergency endpoint
async def create_emergency(request: Request, rep: EmergencyCreate, cu: dict = Depends(get_user)):
    if rep.type not in ("emergency", "accident"):
        raise HTTPException(400, "نوع البلاغ غير صالح")
    # المشغل: لا يحتاج driver_id — يُسمح له دائماً بإرسال بلاغ
    if cu["role"] not in ("operator", "admin", "superuser"):
        if rep.driver_id != cu.get("driver_id"):
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
        # المشغل: نحفظ بيانات البلاغ بدون driver_id (FK على drivers) — نستخدم operator_id
        if cu["role"] == "operator":
            op_id_emg = cu.get("operator_id") or 0
            log.info(f"[EMERGENCY] operator insert: operator_id={op_id_emg} type={rep.type}")
            c.execute("""INSERT INTO emergency_reports
                         (driver_id,car_id,type,audio_url,notes,created_at,is_read,location,operator_id,is_operator)
                         VALUES(NULL,NULL,?,?,?,?,0,?,?,1)""",
                      (rep.type, audio_url, rep.notes or "", now, rep.location or "", op_id_emg))
        else:
            c.execute("""INSERT INTO emergency_reports
                         (driver_id,car_id,type,audio_url,notes,created_at,is_read,location,is_operator)
                         VALUES(?,?,?,?,?,?,0,?,0)""",
                      (rep.driver_id, rep.car_id, rep.type, audio_url, rep.notes or "", now, rep.location or ""))
        rid = c.lastrowid
        log_event("emergency_reported", report_id=rid, driver_id=rep.driver_id, type=rep.type)
        return {"id": rid, "created_at": now, "type": rep.type}

@app.get("/emergency/reports")
async def get_emergency_reports(cu: dict = Depends(get_user), branch: Optional[str] = None):
    q = """SELECT e.*,
                  COALESCE(d.name, op.name) as driver_name,
                  COALESCE(d.name, op.name) as sender_name,
                  COALESCE(d.phone, op.phone) as driver_phone,
                  c.plate as car_plate,
                  op.name as operator_name
           FROM emergency_reports e
           LEFT JOIN drivers d ON e.driver_id=d.id AND e.is_operator=0
           LEFT JOIN equipment_operators op ON e.operator_id=op.id
           LEFT JOIN cars c ON e.car_id=c.id"""
    with get_db() as conn:
        c = conn.cursor()
        branch = _effective_branch(cu, branch)
        if cu["role"] in ("admin", "reporter", "superuser"):
            if branch:
                c.execute(q + " WHERE d.branch=? ORDER BY e.id DESC", (branch,))
            else:
                c.execute(q + " ORDER BY e.id DESC")
        elif cu["role"] == "operator":
            # المشغل يشوف بلاغاته باستخدام operator_id (المخزّن كـ driver_id في emergency_reports)
            op_id = cu.get("operator_id")
            if not op_id: return []
            c.execute(q + " WHERE e.driver_id=? ORDER BY e.id DESC", (op_id,))
        else:
            did = cu.get("driver_id")
            if not did: return []
            c.execute(q + " WHERE e.driver_id=? ORDER BY e.id DESC", (did,))
        rows = []
        for r in c.fetchall():
            d = dict(r)
            d["is_read"]    = bool(d.get("is_read", 0))
            d["is_handled"] = bool(d.get("is_handled", 0))
            d.setdefault("driver_name", None); d.setdefault("car_plate", None)
            d.setdefault("driver_phone", "")
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
# 27. SETTINGS
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
            # رجّع بس الأسعار العامة (مش الفروع)
            # الأسعار المقبولة هي كل ws_type موجود في WORKSHOP_TYPES
            VALID_PRICE_KEYS = {f"price_{t}" for t in WORKSHOP_TYPES}
            return {
                k.replace("price_", ""): v
                for k, v in all_rows.items()
                if k in VALID_PRICE_KEYS
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
# 28. ADMIN — RESET DATA (protected with password)
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
# 29. SUPERUSER — Admin Management + Audit Logs
# ══════════════════════════════════════════════════════

class AdminCreate(BaseModel):
    username: str
    password: str
    branch: Optional[str] = ""   # الفرع التابع له الأدمن/المراقب

class AdminUpdate(BaseModel):
    username: str
    password: Optional[str] = None   # اختياري — لو فارغ مش هيتغير
    branch: Optional[str] = None     # اختياري — لو None مش هيتغير


@app.put("/superuser/change-password")
async def superuser_change_password(body: dict, cu: dict = Depends(require_superuser)):
    """السوبر يوزر يغير كلمة مروره هو أو أي حساب آخر (بصلاحيات كاملة)"""
    current_password = body.get("current_password", "").strip()
    new_password     = body.get("new_password", "").strip()
    target_user_id   = body.get("user_id")   # اختياري — لو مش موجود يغير كلمة مرور السوبر يوزر نفسه

    if not new_password or len(new_password) < 6:
        raise HTTPException(400, "كلمة المرور الجديدة يجب أن تكون 6 أحرف على الأقل")

    with get_db() as conn:
        c = conn.cursor()

        if target_user_id:
            # تغيير كلمة مرور مستخدم آخر (بدون الحاجة للكلمة الحالية)
            target = c.execute("SELECT id, username, role FROM users WHERE id=?", (target_user_id,)).fetchone()
            if not target:
                raise HTTPException(404, "المستخدم غير موجود")
            target = dict(target)
            c.execute("UPDATE users SET password=?, refresh_token=NULL, refresh_exp=NULL WHERE id=?",
                      (hash_password(new_password), target_user_id))
            log_event("password_changed_by_superuser",
                      target_user=target["username"], target_role=target["role"],
                      by=cu["username"])
            return {"ok": True, "message": f"تم تغيير كلمة مرور {target['username']} بنجاح"}
        else:
            # تغيير كلمة مرور السوبر يوزر نفسه — الكلمة الحالية مطلوبة
            if not current_password:
                raise HTTPException(400, "كلمة المرور الحالية مطلوبة")
            user = c.execute("SELECT id, password FROM users WHERE id=?", (cu["user_id"],)).fetchone()
            if not user:
                raise HTTPException(404, "المستخدم غير موجود")
            import bcrypt
            if not bcrypt.checkpw(current_password.encode(), user["password"].encode()):
                raise HTTPException(401, "كلمة المرور الحالية غير صحيحة")
            c.execute("UPDATE users SET password=?, refresh_token=NULL, refresh_exp=NULL WHERE id=?",
                      (hash_password(new_password), cu["user_id"]))
            log_event("password_changed_self", user=cu["username"])
            return {"ok": True, "message": "تم تغيير كلمة مرورك بنجاح — ستحتاج لتسجيل الدخول مجدداً"}


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


# ── Workshop Admins management (ادمن الورش) ─────────────────────────────────

@app.get("/superuser/locked-accounts")
async def list_locked_accounts(cu: dict = Depends(require_superuser)):
    """عرض كل الحسابات المقفولة حالياً بسبب محاولات دخول فاشلة متكررة."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT id, username, role, failed_attempts, locked_until
                     FROM users WHERE locked_until IS NOT NULL""")
        now = datetime.utcnow()
        rows = []
        for r in c.fetchall():
            row = dict(r)
            try:
                is_locked_now = datetime.fromisoformat(row["locked_until"]) > now
            except Exception:
                is_locked_now = False
            row["is_locked_now"] = is_locked_now
            rows.append(row)
        return rows

@app.post("/superuser/users/{uid}/unlock")
async def unlock_account(uid: int, request: Request, cu: dict = Depends(require_superuser)):
    """فك قفل حساب فوراً (بدل انتظار الـ 15 دقيقة) — لحالات الطوارئ."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, username FROM users WHERE id=?", (uid,))
        target = c.fetchone()
        if not target:
            raise HTTPException(404, "المستخدم غير موجود")
        c.execute("UPDATE users SET failed_attempts=0, locked_until=NULL WHERE id=?", (uid,))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "unlock_account", f"فك قفل الحساب: {target['username']} (id={uid})",
                        request.client.host if request.client else "")
        return {"message": f"تم فك قفل الحساب: {target['username']}"}


@app.get("/superuser/workshop-admins")
async def superuser_list_workshop_admins(cu: dict = Depends(require_superuser)):
    """عرض جميع حسابات ادمن الورش."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("""SELECT id, username, role, branch, created_at, last_login
                     FROM users WHERE role='workshop_admin' ORDER BY id""")
        return [dict(r) for r in c.fetchall()]


@app.post("/superuser/workshop-admins", status_code=201)
async def superuser_create_workshop_admin(
    request: Request,
    body: AdminCreate,
    cu: dict = Depends(require_superuser)
):
    """إضافة ادمن ورش جديد."""
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
                  (username, _hash(body.password), "workshop_admin", body.branch or ""))
        new_id = c.lastrowid
    log_event("workshop_admin_created_by_super", new_username=username, superuser=cu["username"])
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "create_workshop_admin", f"إنشاء ادمن ورش جديد: {username} (فرع: {body.branch or 'بدون'})",
                    request.client.host if request.client else "")
    return {"id": new_id, "username": username, "role": "workshop_admin", "branch": body.branch or ""}


@app.put("/superuser/workshop-admins/{uid}")
async def superuser_update_workshop_admin(
    request: Request,
    uid: int,
    body: AdminUpdate,
    cu: dict = Depends(require_superuser)
):
    """تعديل اسم المستخدم أو كلمة مرور ادمن ورش."""
    new_username = body.username.strip()
    if len(new_username) < 2:
        raise HTTPException(400, "اسم المستخدم قصير جداً")
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, username, role FROM users WHERE id=?", (uid,))
        target = c.fetchone()
        if not target:
            raise HTTPException(404, "المستخدم غير موجود")
        if target["role"] != "workshop_admin":
            raise HTTPException(403, "لا يمكن تعديل إلا حسابات ادمن الورش من هنا")
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
    log_event("workshop_admin_updated_by_super", uid=uid, new_username=new_username, superuser=cu["username"])
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "update_workshop_admin",
                    f"تعديل ادمن ورش #{uid} من '{old_username}' إلى '{new_username}'"
                    + (" (مع تغيير كلمة المرور)" if body.password else ""),
                    request.client.host if request.client else "")
    return {"message": "تم التحديث", "id": uid, "username": new_username}


@app.delete("/superuser/workshop-admins/{uid}")
async def superuser_delete_workshop_admin(
    request: Request,
    uid: int,
    cu: dict = Depends(require_superuser)
):
    """حذف حساب ادمن ورش."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, username, role FROM users WHERE id=?", (uid,))
        target = c.fetchone()
        if not target:
            raise HTTPException(404, "المستخدم غير موجود")
        if target["role"] != "workshop_admin":
            raise HTTPException(403, "لا يمكن حذف إلا حسابات ادمن الورش من هنا")
        deleted_username = target["username"]
        c.execute("DELETE FROM users WHERE id=?", (uid,))
    log_event("workshop_admin_deleted_by_super", uid=uid, username=deleted_username, superuser=cu["username"])
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "delete_workshop_admin", f"حذف ادمن ورش: {deleted_username} (id={uid})",
                    request.client.host if request.client else "")
    return {"message": f"تم حذف ادمن الورش: {deleted_username}"}


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
# 30. DRIVER REQUESTS
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
    """إنشاء طلب جديد — يدعم السائق والمشغل بنفس الـ endpoint"""
    import traceback as _tb_mod
    log.info(f"[REQUEST] role={cu.get('role')} op_id={cu.get('operator_id')} drv_id={cu.get('driver_id')} type={body.get('type')}")
    try:
        req_type        = (body.get("type") or "").strip()
        notes           = (body.get("notes") or "").strip()
        new_odometer    = body.get("new_odometer")
        trip_id         = body.get("trip_id")
        renewal_date    = (body.get("renewal_date") or "").strip()
        new_expiry_date = (body.get("new_expiry_date") or "").strip()
        license_type    = (body.get("license_type") or "").strip()

        if req_type not in REQUEST_TYPES:
            raise HTTPException(400, f"نوع الطلب غير صالح: {req_type}")

        is_op     = cu["role"] == "operator"
        driver_id = cu.get("driver_id") if not is_op else None
        op_id     = cu.get("operator_id") if is_op else None

        if not is_op and not driver_id:
            raise HTTPException(403, "فقط السائقون والمشغلون يمكنهم تقديم الطلبات")

        # تحذير لو المشغل غير مربوط بـ equipment_operators
        if is_op and not op_id:
            log.warning(f"[REQUEST] Operator user_id={cu['user_id']} has no equipment_operators record — request will be stored without operator_id")

        if req_type == "license_renew":
            if not renewal_date:    raise HTTPException(400, "تاريخ التجديد مطلوب")
            if not new_expiry_date: raise HTTPException(400, "تاريخ انتهاء الرخصة الجديد مطلوب")
            if not license_type:    raise HTTPException(400, "نوع الرخصة مطلوب")
            notes = (f"نوع الرخصة: {'رخصة القيادة' if license_type=='driver' else 'رخصة العربة'}"
                     f" | تاريخ التجديد: {renewal_date} | تاريخ الانتهاء الجديد: {new_expiry_date}")

        now = datetime.utcnow().isoformat() + "Z"
        odo = float(new_odometer) if new_odometer else None
        tid = int(trip_id)        if trip_id      else None

        with get_db() as conn:
            c = conn.cursor()

            # ── guard: لو operator columns ناقصة لسبب ما أضفها فقط (ALTER آمن) ──
            # هذا السيناريو لا يجب أن يحدث أبداً بعد migrate_db() في startup
            # لكن كـ safety net نضيف فقط ولا نعمل rebuild أبداً
            _cols = {r[1] for r in c.execute("PRAGMA table_info(driver_requests)").fetchall()}
            for _mc, _md in [("operator_id","INTEGER"), ("is_operator","INTEGER DEFAULT 0")]:
                if _mc not in _cols:
                    log.error(f"[REQUEST] UNEXPECTED: column '{_mc}' missing from driver_requests after startup migration!")
                    try: c.execute(f"ALTER TABLE driver_requests ADD COLUMN {_mc} {_md}")
                    except Exception: pass

            # ── Safeguard: تحقق من consistency قبل INSERT ──
            if is_op and not op_id:
                log.error(f"[REQUEST] SAFEGUARD: is_operator=1 but operator_id=None — user_id={cu['user_id']}")
            if not is_op and not driver_id:
                raise HTTPException(403, "driver_id مطلوب للسائق")

            log.info(f"[REQUEST] pre-insert: role={cu['role']} is_op={is_op} op_id={op_id} drv_id={driver_id} type={req_type}")

            # ── INSERT ──
            if is_op:
                c.execute(
                    """INSERT INTO driver_requests
                       (driver_id,operator_id,is_operator,type,notes,new_odometer,trip_id,status,created_at)
                       VALUES(NULL,?,1,?,?,?,?,'pending',?)""",
                    (op_id, req_type, notes, odo, tid, now)
                )
            else:
                c.execute(
                    """INSERT INTO driver_requests
                       (driver_id,is_operator,type,notes,new_odometer,trip_id,status,created_at)
                       VALUES(?,0,?,?,?,?,'pending',?)""",
                    (driver_id, req_type, notes, odo, tid, now)
                )
            rid = c.lastrowid
            log.info(f"[REQUEST] post-insert: id={rid} is_operator={1 if is_op else 0} operator_id={op_id} driver_id={driver_id}")

        log.info(f"[REQUEST] ✅ created id={rid} type={req_type} is_op={is_op} op_id={op_id}")
        log_event("request_created",
                  driver_id=op_id if is_op else driver_id,
                  request_id=rid, type=req_type)
        return {"id": rid, "status": "pending"}

    except HTTPException:
        raise
    except Exception as _err:
        _tb = _tb_mod.format_exc()
        log.error(f"[REQUEST ERROR] {_err}\n{_tb}")
        raise HTTPException(500, f"خطأ في إنشاء الطلب: {str(_err)}")

@app.get("/requests")
async def get_requests(
    cu: dict = Depends(get_user),
    branch:    Optional[str] = None,
    user_type: Optional[str] = None,   # 'driver' | 'operator' — للأدمن/السوبر
):
    with get_db() as conn:
        c = conn.cursor()
        if cu["role"] == "driver":
            driver_id = cu.get("driver_id")
            c.execute("""SELECT r.*,
                                d.name   as driver_name,
                                d.branch as driver_branch
                         FROM driver_requests r
                         LEFT JOIN drivers d ON d.id = r.driver_id
                         WHERE r.driver_id=?
                         ORDER BY r.id DESC""", (driver_id,))
        elif cu["role"] == "operator":
            op_id2 = cu.get("operator_id")
            if op_id2:
                # مشغل مرتبط بـ equipment_operators → جيب طلباته بالـ operator_id
                c.execute("""SELECT r.*,
                                    op.name   as operator_name,
                                    op.name   as driver_name,
                                    op.branch as driver_branch
                             FROM driver_requests r
                             LEFT JOIN equipment_operators op ON op.id = r.operator_id
                             WHERE r.is_operator=1
                               AND r.operator_id=?
                             ORDER BY r.id DESC""", (op_id2,))
            else:
                # مشغل غير مرتبط بـ equipment_operators بعد
                # نرجع فارغ لأننا مش عارفين هو مين
                return []
        else:
            eff_branch = _effective_branch(cu, branch)
            # user_type filter: driver=is_operator=0, operator=is_operator=1
            ut_cond = ""
            if user_type == "driver":   ut_cond = " AND (r.is_operator=0 OR r.is_operator IS NULL)"
            elif user_type == "operator": ut_cond = " AND r.is_operator=1"
            if eff_branch:
                c.execute(f"""SELECT r.*,
                                    COALESCE(d.name, op.name)   as driver_name,
                                    COALESCE(op.name, d.name)   as operator_name,
                                    COALESCE(d.branch,op.branch) as driver_branch
                             FROM driver_requests r
                             LEFT JOIN drivers d ON r.driver_id=d.id AND (r.is_operator IS NULL OR r.is_operator=0)
                             LEFT JOIN equipment_operators op ON r.operator_id=op.id
                             WHERE COALESCE(d.branch,op.branch)=? {ut_cond}
                             ORDER BY r.id DESC""", (eff_branch,))
            else:
                c.execute(f"""SELECT r.*,
                                    COALESCE(d.name, op.name)   as driver_name,
                                    COALESCE(op.name, d.name)   as operator_name,
                                    COALESCE(d.branch,op.branch) as driver_branch
                             FROM driver_requests r
                             LEFT JOIN drivers d ON r.driver_id=d.id AND (r.is_operator IS NULL OR r.is_operator=0)
                             LEFT JOIN equipment_operators op ON r.operator_id=op.id
                             WHERE 1=1 {ut_cond}
                             ORDER BY r.id DESC""")
        return [dict(row) for row in c.fetchall()]


@app.get("/requests/pending-count")
async def requests_pending_count(cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        c = conn.cursor()
        eff_branch = _branch_filter(cu)
        if eff_branch:
            c.execute("""SELECT COUNT(*) as cnt FROM driver_requests r
                         LEFT JOIN drivers d ON r.driver_id=d.id AND (r.is_operator IS NULL OR r.is_operator=0)
                         LEFT JOIN equipment_operators op ON r.operator_id=op.id
                         WHERE r.status='pending'
                           AND COALESCE(d.branch, op.branch)=?""", (eff_branch,))
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
            # تحديث admin_message عشان السائق يشوف الرسالة
            eff_msg = super_notes or (body.get("admin_message") or "").strip()
            c.execute("""UPDATE driver_requests
                         SET status=?, super_decision=?, super_notes=?,
                             super_handled_by=?, super_handled_at=?,
                             admin_message=CASE WHEN ? != '' THEN ? ELSE admin_message END
                         WHERE id=?""",
                      (status, status, super_notes, cu["username"], now,
                       eff_msg, eff_msg, rid))
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
# 31. MAINTENANCE SCHEDULE (جدول الصيانة الدورية)
# ══════════════════════════════════════════════════════

MAINTENANCE_TYPES = [
    # ── زيوت وفلاتر ──
    "تغيير زيت", "زيت تروس", "زيت فرامل", "زيت هيدروليك", "شحم",
    "فلتر زيت", "فلتر هواء", "فلتر وقود", "فلتر سولار", "فلتر بنزين",
    "فلتر فاصل", "فلتر مكيف", "فلتر تكيف", "فلتر مجفف", "فلتر منفس", "فلتر هيدروليك",
    # ── قطع ──
    "إطارات", "بطارية", "فرامل أمامية", "فرامل خلفية",
    # ── سيور ──
    "سير توزيع", "سيور", "سير كاتينة",
    # ── عمرات وصيانة كبيرة ──
    "عمرة محرك", "عمرة كاملة", "نصف عمرة محرك", "عمرة نصفية",
    "جوان وش سلندر", "سلندر",
    "عمرة طرمبة وقود", "طرمبة وقود", "وحدات الحقن",
    # ── أخرى ──
    "ماء تبريد", "سائل فرامل", "فحص دوري عام", "أخرى",
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
# 32. BRANCHES LIST — لقائمة الفروع المتاحة
# ══════════════════════════════════════════════════════

@app.get("/reports/branches")
async def list_branches(cu: dict = Depends(require_admin_or_reporter)):
    """قائمة الفروع الثابتة للنظام."""
    eff_branch = _branch_filter(cu)
    if eff_branch:
        return {"branches": [eff_branch]}
    return {"branches": BRANCHES}

# ══════════════════════════════════════════════════════
# 33. OPERATIONAL CARD REPORT
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
                COUNT(DISTINCT date(t.start_time))                AS work_days,
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
            "work_days":   int(tot.get("work_days") or 0),
            "min_odo":     tot["first_odo"],
            "max_odo":     tot["last_odo"],
            "report_type": report_type,
            "group_by":    group_by,
        }



# ══════════════════════════════════════════════════════
# 34. OPERATIONAL REPORT — PER DRIVER BREAKDOWN
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
# 35. ANOMALOUS TRIPS DIAGNOSTIC
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
# 36. BULK IMPORT — CSV / Excel
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
    # للصيانة الدورية
    "نوع الصيانة":        "maintenance_type",
    "maintenance_type":  "maintenance_type",
    "كل كام كم":          "interval_km",
    "الفترة كم":          "interval_km",
    "interval_km":       "interval_km",
    "كل كام يوم":         "interval_days",
    "الفترة يوم":         "interval_days",
    "interval_days":     "interval_days",
    "آخر قراءة":          "last_done_km",
    "آخر قراءة كم":        "last_done_km",
    "last_done_km":      "last_done_km",
    "تاريخ آخر صيانة":     "last_done_date",
    "آخر تاريخ":          "last_done_date",
    "last_done_date":    "last_done_date",
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
# 37. PERMISSIONS BULK IMPORT — CSV / Excel
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
# 38. LAST ODOMETER — آخر عداد لمركبة (للسائق عند بدء الرحلة)
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
# 39. FRAUD DETECTION — كشف التلاعب
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
# 40. VOICE NOTES — رسائل صوتية من السوبر
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
    if body.target_user_id is None and body.target_group not in ("admins", "drivers"):
        raise HTTPException(400, "حدد مجموعة مستلمين (admins/drivers) أو شخصاً محدداً")
    now     = datetime.utcnow()
    expires = (now + timedelta(hours=body.hours_ttl)).isoformat() + "Z"
    # لو رسالة شخصية → target_group فاضي لضمان عدم وصولها للمجموعة
    tg = '' if body.target_user_id is not None else (
        body.target_group if body.target_group in ('admins', 'drivers') else 'admins'
    )
    try:
        with get_db() as conn:
            c = conn.cursor()
            # تأكد إن الـ CHECK constraint اتشال (migration)
            try:
                c.execute("""INSERT INTO voice_notes
                    (sender_id, target_group, target_user_id, audio_data,
                     duration_sec, created_at, expires_at, max_plays)
                    VALUES(?,?,?,?,?,?,?,?)""",
                    (cu["user_id"], tg, body.target_user_id, body.audio_data,
                     body.duration_sec, now.isoformat()+"Z", expires, body.max_plays))
            except Exception:
                # CHECK constraint قديم → نستخدم executescript لإعادة بناء الجدول
                try:
                    conn.executescript("""
                        CREATE TABLE IF NOT EXISTS _vn_fix(
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            sender_id INTEGER NOT NULL,
                            sender_role TEXT NOT NULL DEFAULT 'superuser',
                            target_group TEXT DEFAULT '',
                            target_user_id INTEGER DEFAULT NULL,
                            audio_data TEXT NOT NULL,
                            duration_sec REAL DEFAULT 0,
                            created_at TEXT NOT NULL,
                            expires_at TEXT NOT NULL,
                            play_count INTEGER DEFAULT 0,
                            max_plays INTEGER DEFAULT 2,
                            is_deleted INTEGER DEFAULT 0
                        );
                        INSERT OR IGNORE INTO _vn_fix
                            SELECT id, sender_id,
                                   COALESCE(sender_role,'superuser'),
                                   target_group, target_user_id, audio_data,
                                   duration_sec, created_at, expires_at,
                                   play_count, max_plays, is_deleted
                            FROM voice_notes;
                        DROP TABLE voice_notes;
                        ALTER TABLE _vn_fix RENAME TO voice_notes;
                    """)
                except Exception:
                    pass
                c.execute("""INSERT INTO voice_notes
                    (sender_id, target_group, target_user_id, audio_data,
                     duration_sec, created_at, expires_at, max_plays)
                    VALUES(?,?,?,?,?,?,?,?)""",
                    (cu["user_id"], tg, body.target_user_id, body.audio_data,
                     body.duration_sec, now.isoformat()+"Z", expires, body.max_plays))
            note_id = c.lastrowid
        return {"id": note_id, "expires_at": expires}
    except Exception as e:
        log.error(f"create_voice_note error: {e}")
        raise HTTPException(500, f"خطأ في حفظ الرسالة: {str(e)}")

@app.get("/voice-notes")
async def list_voice_notes(cu: dict = Depends(get_user)):
    role = cu["role"]
    uid  = cu["user_id"]
    now  = datetime.utcnow().isoformat() + "Z"

    # ── صفحة إدارة الرسائل (superuser أو admin) ──
    if role in ("superuser", "admin"):
        try:
            with get_db() as conn:
                c = conn.cursor()
                c.execute("""
                    SELECT vn.id, vn.target_group, vn.target_user_id,
                           vn.duration_sec, vn.created_at,
                           vn.expires_at, vn.play_count, vn.max_plays, vn.is_deleted,
                           u.username  AS target_uname,
                           d.name      AS target_dname
                    FROM voice_notes vn
                    LEFT JOIN users   u ON u.id = vn.target_user_id
                    LEFT JOIN drivers d ON d.user_id = vn.target_user_id
                    WHERE vn.sender_id = ?
                    ORDER BY vn.created_at DESC
                """, (uid,))
                rows = []
                for r in c.fetchall():
                    row = dict(r)
                    tid = row.get("target_user_id")
                    if tid:
                        row["target_name"] = row.get("target_uname") or row.get("target_dname") or "مستخدم"
                    else:
                        row["target_name"] = None
                    row.pop("target_uname", None)
                    row.pop("target_dname", None)
                    rows.append(row)
                return rows
        except Exception as e:
            log.error(f"list_voice_notes error: {e}")
            raise HTTPException(500, f"خطأ في جلب الرسائل: {str(e)}")

    # ── استقبال الرسائل (driver, reporter …) ──
    if role == "driver":
        group = "drivers"
    elif role in ("reporter", "supervisor_workshop", "supervisor_field", "workshop_admin"):
        group = "admins"
    else:
        return []

    with get_db() as conn:
        c = conn.cursor()
        c.execute("""
            SELECT id, duration_sec, created_at, expires_at,
                   play_count, max_plays, audio_data
            FROM voice_notes
            WHERE (
                    -- رسائل المجموعة
                    (target_group = ? AND target_user_id IS NULL)
                    OR (target_group = ? AND target_user_id = 0)
                    -- رسائل شخصية بأي target_group
                    OR target_user_id = ?
                  )
              AND sender_id != ?
              AND is_deleted = 0
              AND expires_at > ?
              AND play_count < max_plays
            ORDER BY created_at DESC
        """, (group, group, uid, uid, now))
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

    if role in ("admin", "reporter", "supervisor_workshop", "supervisor_field", "workshop_admin"):
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
            WHERE (
                (target_group = ? AND (target_user_id IS NULL OR target_user_id = 0))
                OR (target_user_id = ? AND target_group NOT IN ('admins','drivers'))
                OR (target_user_id = ? AND target_group IN ('','personal'))
            )
              AND sender_id != ?
              AND is_deleted = 0
              AND expires_at > ?
              AND play_count < max_plays
            ORDER BY created_at DESC
        """, (group, uid, uid, uid, now))
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


# ══════════════════════════════════════════════════════
# 41. MAINTENANCE SCHEDULE — CRUD + ALERTS
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


@app.get("/driver/maintenance-alerts")
async def get_driver_maintenance_alerts(cu: dict = Depends(get_user)):
    """
    تنبيهات الصيانة الدورية للسائق:
    - يجيب العربية المخصصة للسائق
    - يرجع السجلات warning أو overdue فقط
    - يظل ظاهر لحد ما الصيانة تتعمل (last_done_km / last_done_date يتحدثوا)
    """
    driver_id = cu.get("driver_id")
    if not driver_id:
        return []

    with get_db() as conn:
        c = conn.cursor()

        # جيب العربية الأساسية للسائق
        c.execute("""
            SELECT c.id, c.plate, c.model
            FROM cars c
            INNER JOIN driver_car_permissions p ON c.id = p.car_id
            WHERE p.driver_id = ?
            ORDER BY p.id ASC
            LIMIT 1
        """, (driver_id,))
        car_row = c.fetchone()
        if not car_row:
            return []

        car = dict(car_row)
        car_id = car["id"]

        # آخر عداد مسجل للعربية من الرحلات
        c.execute("""
            SELECT t.end_odometer
            FROM trips t
            INNER JOIN (
                SELECT car_id, MAX(end_time) as latest
                FROM trips
                WHERE end_odometer IS NOT NULL AND end_time IS NOT NULL AND car_id = ?
                GROUP BY car_id
            ) lx ON t.car_id = lx.car_id AND t.end_time = lx.latest
            WHERE t.end_odometer IS NOT NULL AND t.car_id = ?
        """, (car_id, car_id))
        odo_row = c.fetchone()
        current_km = float(odo_row["end_odometer"]) if odo_row else None

        # جدول الصيانة للعربية دي
        c.execute("SELECT * FROM maintenance_schedule WHERE car_id = ?", (car_id,))
        rows = c.fetchall()

    cars_map = {car_id: {"plate": car["plate"], "model": car["model"], "branch": ""}}
    alerts = []
    for row in rows:
        d = dict(row)
        d["_current_km"] = current_km
        enriched = _maintenance_row(d, cars_map)
        if enriched["alert_status"] in ("warning", "overdue"):
            remaining = None
            if enriched.get("next_due_km") and current_km is not None:
                remaining = round(enriched["next_due_km"] - current_km, 1)
            alerts.append({
                "maintenance_type": enriched["maintenance_type"],
                "alert_status":     enriched["alert_status"],
                "next_due_km":      enriched.get("next_due_km"),
                "current_km":       enriched.get("current_km"),
                "remaining_km":     remaining,
                "next_due_date":    enriched.get("next_due_date"),
                "car_plate":        car["plate"],
                "car_model":        car["model"],
                "notes":            enriched.get("notes", ""),
            })
    return alerts


@app.post("/maintenance/schedule", status_code=201)
async def create_maintenance(rec: MaintenanceScheduleCreate, cu: dict = Depends(require_workshop_admin)):
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
async def update_maintenance(mid: int, rec: MaintenanceScheduleUpdate, cu: dict = Depends(require_workshop_admin)):
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
async def delete_maintenance(mid: int, cu: dict = Depends(require_workshop_admin)):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("DELETE FROM maintenance_schedule WHERE id=?", (mid,))
        if conn.execute("SELECT changes()").fetchone()[0] == 0:
            raise HTTPException(404, "السجل غير موجود")
        log_event("maintenance_deleted", id=mid)
        return {"ok": True}


# ══════════════════════════════════════════════════════
# 42. FUEL EFFICIENCY REPORT (تقرير كفاءة الوقود)
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
# 43. MONTHLY REPORT — تقرير شهري شامل
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
# 44. DRIVER KPI — مؤشرات أداء السائقين
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

            # ── وقود — سيارات فقط (استبعاد سولار المعدات) ──
            # نجيب الـ car_ids المرتبطة بالسائق من رحلاته في الفترة
            c.execute("""
                SELECT DISTINCT t.car_id
                FROM trips t
                JOIN cars c ON c.id = t.car_id
                WHERE t.driver_id=?
                  AND t.end_time IS NOT NULL
                  AND date(t.start_time) >= ? AND date(t.start_time) < ?
                  AND (c.equipment_type IS NULL OR c.equipment_type = '')
            """, (did, month_start, month_end))
            car_ids_no_equip = [r[0] for r in c.fetchall()]

            # لو مفيش سيارات عادية، معدل الاستهلاك مش قابل للحساب
            c.execute("""
                SELECT quantity, odometer_reading, created_at, price
                FROM workshop_records
                WHERE driver_id=? AND type LIKE 'fuel_%'
                  AND date(created_at) >= ? AND date(created_at) < ?
                  AND (vehicle_id IS NULL OR vehicle_id IN ({}))
                ORDER BY created_at ASC
            """.format(
                ",".join("?" * len(car_ids_no_equip)) if car_ids_no_equip else "SELECT -1"
            ), (did, month_start, month_end, *car_ids_no_equip) if car_ids_no_equip else (did, month_start, month_end))
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
                WHERE driver_id=? AND type IN (
                    'oil','filter','filter_oil','filter_solar','filter_petrol',
                    'filter_air','filter_separator','filter_ac','filter_dryer',
                    'filter_breather','filter_hydraulic'
                )
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


def _calc_operator_kpi_score(
    shifts: int, hours: float, fuel_liters: float,
    avg_hours_all_operators: float = 0.0, target_hours: float = 160.0,
) -> dict:
    """
    حساب نقاط KPI لمشغل معدة واحد:
      عدد الورديات       → 20%
      ساعات التشغيل      → 60%
      كفاءة استهلاك الوقود (لتر/ساعة) → 20%  (لو مفيش بيانات وقود يُعوَّض من الساعات)
    """
    # ── 1. عدد الورديات (20%) ──
    if shifts >= 20: shifts_score = 100.0
    elif shifts >= 15: shifts_score = 85.0
    elif shifts >= 10: shifts_score = 65.0
    elif shifts >= 5:  shifts_score = 40.0
    elif shifts >= 1:  shifts_score = 20.0
    else: shifts_score = 0.0

    # ── 2. ساعات التشغيل (60%) — مقارنة بمتوسط باقي المشغلين ──
    ref_hours = avg_hours_all_operators if avg_hours_all_operators > 0 else target_hours
    if ref_hours > 0:
        ratio = hours / ref_hours
        if ratio >= 1.2: hours_score = 100.0
        elif ratio >= 1.0: hours_score = 90.0
        elif ratio >= 0.8: hours_score = 75.0
        elif ratio >= 0.6: hours_score = 55.0
        elif ratio >= 0.4: hours_score = 35.0
        else: hours_score = 10.0
    else:
        hours_score = 0.0

    # ── 3. كفاءة الوقود (لتر/ساعة) (20%) ──
    if hours > 5 and fuel_liters > 0:
        consumption = fuel_liters / hours
        if consumption <= 8: fuel_score = 100.0
        elif consumption <= 12: fuel_score = 85.0
        elif consumption <= 16: fuel_score = 65.0
        elif consumption <= 20: fuel_score = 45.0
        else: fuel_score = 20.0
    else:
        fuel_score = None
        consumption = None

    W_SHIFTS, W_HOURS, W_FUEL = 0.20, 0.60, 0.20
    weighted = shifts_score * W_SHIFTS + hours_score * W_HOURS
    if fuel_score is not None:
        weighted += fuel_score * W_FUEL
    else:
        weighted += hours_score * W_FUEL   # عوّض الوقود بالساعات

    total = round(weighted, 1)
    if total >= 85: grade, grade_color = "ممتاز", "#10b981"
    elif total >= 70: grade, grade_color = "جيد", "#3b82f6"
    elif total >= 50: grade, grade_color = "مقبول", "#f97316"
    else: grade, grade_color = "يحتاج تحسين", "#ef4444"

    return {
        "total_score": total, "grade": grade, "grade_color": grade_color,
        "shifts_score": shifts_score, "hours_score": hours_score, "fuel_score": fuel_score,
        "consumption": round(consumption, 2) if consumption is not None else None,
    }


@app.get("/reports/equipment-operator-kpi")
async def equipment_operator_kpi_report(
    month:       str = Query(None, description="YYYY-MM"),
    operator_id: int = Query(None),
    branch:      str = Query(None),
    cu: dict = Depends(require_admin_or_reporter),
):
    """تقرير مؤشرات أداء مشغلي المعدات KPI — نفس فلسفة تقرير السائقين لكن بالساعات بدل الكيلومترات."""
    if not month:
        month = datetime.utcnow().strftime("%Y-%m")
    try:
        y, m = month.split("-")
        month_start = f"{y}-{m}-01"
        next_m, next_y = int(m) + 1, int(y)
        if next_m > 12:
            next_m, next_y = 1, next_y + 1
        month_end = f"{next_y}-{next_m:02d}-01"
    except Exception:
        raise HTTPException(400, "صيغة الشهر خاطئة — استخدم YYYY-MM")

    eff_branch = _branch_filter(cu) or branch

    with get_db() as conn:
        _ensure_operator_tables(conn)
        c = conn.cursor()
        q = "SELECT id, name, branch, status FROM equipment_operators WHERE status='active'"
        params: list = []
        if eff_branch:
            q += " AND branch=?"; params.append(eff_branch)
        if operator_id:
            q += " AND id=?"; params.append(operator_id)
        q += " ORDER BY name"
        operators = [dict(r) for r in c.execute(q, params).fetchall()]

        raw_data = []
        for op in operators:
            oid = op["id"]
            c.execute("""
                SELECT id, equipment_id, equipment_name, start_time, end_time,
                       start_hours, end_hours, fuel_liters, status, productivity
                FROM operator_shifts
                WHERE operator_id=? AND date(start_time) >= ? AND date(start_time) < ?
                ORDER BY start_time DESC
            """, (oid, month_start, month_end))
            shift_rows = [dict(r) for r in c.fetchall()]

            shift_count = len(shift_rows)
            total_hours = 0.0
            fuel_total  = 0.0
            equip_codes = set()
            productivity_notes = []
            for s in shift_rows:
                if s.get("status") == "ended" and s.get("end_hours") is not None and s.get("start_hours") is not None \
                   and s["end_hours"] > s["start_hours"]:
                    total_hours += (s["end_hours"] - s["start_hours"])
                fuel_total += float(s.get("fuel_liters") or 0)
                if s.get("equipment_id"):
                    equip_codes.add(s["equipment_id"])
                if s.get("productivity"):
                    productivity_notes.append(s["productivity"])

            # ── تكلفة الصيانة المرتبطة بالمعدات اللي شغّلها المشغل في نفس الفترة (من أذون الصرف) ──
            maint_cost = 0.0
            if equip_codes:
                placeholders = ",".join("?" * len(equip_codes))
                c.execute(f"""
                    SELECT COALESCE(SUM(total_amount),0) FROM workshop_vouchers
                    WHERE equipment_id IN ({placeholders})
                      AND date(created_at) >= ? AND date(created_at) < ?
                """, (*equip_codes, month_start, month_end))
                maint_cost = float(c.fetchone()[0] or 0)

            raw_data.append({
                "operator_id":         oid,
                "operator_name":       op["name"],
                "branch":              op["branch"] or "",
                "shifts":              shift_count,
                "hours":               round(total_hours, 2),
                "fuel_liters":         round(fuel_total, 1),
                "equipment_operated":  sorted(equip_codes),
                "productivity_notes":  productivity_notes,
                "maint_cost":          maint_cost,
                "first_shift":         shift_rows[-1]["start_time"] if shift_rows else None,
                "last_shift":          shift_rows[0]["start_time"] if shift_rows else None,
            })

        hours_values = [d["hours"] for d in raw_data if d["hours"] > 0]
        avg_hours = (sum(hours_values) / len(hours_values)) if hours_values else 0.0

        results = []
        for d in raw_data:
            scores = _calc_operator_kpi_score(
                shifts=d["shifts"], hours=d["hours"], fuel_liters=d["fuel_liters"],
                avg_hours_all_operators=avg_hours,
            )
            results.append({
                "operator_id":        d["operator_id"],
                "operator_name":      d["operator_name"],
                "branch":             d["branch"],
                "month":              month,
                "shifts":             d["shifts"],
                "hours":              d["hours"],
                "fuel_liters":        d["fuel_liters"],
                "equipment_operated": d["equipment_operated"],
                "productivity_notes": d["productivity_notes"],
                "maint_cost":         round(d["maint_cost"], 2),
                "first_shift":        d["first_shift"],
                "last_shift":         d["last_shift"],
                "avg_hours_peers":    round(avg_hours, 1),
                **scores,
            })

    results.sort(key=lambda x: x["total_score"], reverse=True)
    for i, r in enumerate(results, 1):
        r["rank"] = i

    return {
        "month":            month,
        "branch":           eff_branch or "كل الفروع",
        "generated_at":     datetime.utcnow().isoformat() + "Z",
        "total_operators":  len(results),
        "avg_hours":        round(avg_hours, 1),
        "operators":        results,
    }


# ══════════════════════════════════════════════════════
# 45. RENTAL EQUIPMENT — المعدات المستأجرة
# ══════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════
# 46. RENTAL EQUIPMENT — معدات مستأجرة (من الداتابيز)
# ══════════════════════════════════════════════════════

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
    if body.fuel_liters and body.fuel_liters.strip():
        try:
            if float(body.fuel_liters) > 400:
                raise HTTPException(400, "الحد الأقصى لكمية الوقود للمعدات هو 400 لتر")
        except ValueError:
            pass
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
    if body.fuel_liters and body.fuel_liters.strip():
        try:
            if float(body.fuel_liters) > 400:
                raise HTTPException(400, "الحد الأقصى لكمية الوقود للمعدات هو 400 لتر")
        except ValueError:
            pass
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



# ══════════════════════════════════════════════════════
# 47. EQUIPMENT — معدات (جدول منفصل - بدون لوحة، الكود هو المعرف)
# ══════════════════════════════════════════════════════

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


@app.get("/equipment/log")
async def get_equipment_log(
    branch:  str = Query(None),
    eq_type: str = Query(None),
    search:  str = Query(None),
    cu: dict = Depends(require_admin_or_reporter),
):
    """
    سجل المعدات — لكل معدة: عدد الورديات، عدد ساعات التشغيل الفعلية،
    أسماء المشغلين اللي شغّلوها، ومعدل الإنتاجية (متوسط ساعات العمل في الوردية).
    """
    with get_db() as conn:
        _ensure_operator_tables(conn)
        eff_branch = _branch_filter(cu) or branch
        q = "SELECT * FROM equipment WHERE 1=1"
        params: list = []
        if eff_branch:
            q += " AND branch=?"; params.append(eff_branch)
        if eq_type:
            q += " AND equipment_type=?"; params.append(eq_type)
        if search:
            q += " AND (equipment_name LIKE ? OR car_code LIKE ?)"
            s = f"%{search}%"; params += [s, s]
        q += " ORDER BY branch, equipment_name"
        equip_rows = [dict(r) for r in conn.execute(q, params).fetchall()]

        shift_rows = conn.execute("""
            SELECT s.equipment_id, s.status, s.start_hours, s.end_hours, o.name AS operator_name
            FROM operator_shifts s
            LEFT JOIN equipment_operators o ON o.id = s.operator_id
        """).fetchall()

    by_equip: dict = {}
    for r in shift_rows:
        eqid = r["equipment_id"] or ""
        if not eqid:
            continue
        agg = by_equip.setdefault(eqid, {"shift_count": 0, "total_hours": 0.0, "operators": set()})
        agg["shift_count"] += 1
        if r["operator_name"]:
            agg["operators"].add(r["operator_name"])
        if r["status"] == "ended" and r["end_hours"] is not None and r["start_hours"] is not None \
           and r["end_hours"] > r["start_hours"]:
            agg["total_hours"] += (r["end_hours"] - r["start_hours"])

    log_rows = []
    for e in equip_rows:
        agg = by_equip.get(e["car_code"], {"shift_count": 0, "total_hours": 0.0, "operators": set()})
        shift_count  = agg["shift_count"]
        total_hours  = round(agg["total_hours"], 2)
        productivity = round(total_hours / shift_count, 2) if shift_count else 0.0
        log_rows.append({
            "id":              e["id"],
            "car_code":        e["car_code"],
            "equipment_name":  e["equipment_name"],
            "equipment_type":  e.get("equipment_type", ""),
            "branch":          e.get("branch", ""),
            "status":          e.get("status", ""),
            "shift_count":     shift_count,
            "total_hours":     total_hours,
            "operators":       sorted(agg["operators"]),
            "productivity":    productivity,   # متوسط ساعات التشغيل لكل وردية
        })
    return {"equipment_log": log_rows}


@app.get("/equipment/unused")
async def get_unused_equipment(
    branch: str = Query(None),
    min_idle_days: float = Query(1.0),
    cu: dict = Depends(require_admin_or_reporter),
):
    """
    المعدات غير المستخدمة — المعدات اللي عدى عليها يوم (أو المدة المحددة) على الأقل
    وهي مش شغالة (مفيش وردية نشطة، وآخر وردية انتهت من مدة أطول من الحد المسموح،
    أو مفيش أي وردية لها أساساً).
    """
    with get_db() as conn:
        _ensure_operator_tables(conn)
        eff_branch = _branch_filter(cu) or branch
        q = "SELECT * FROM equipment WHERE status='active'"
        params: list = []
        if eff_branch:
            q += " AND branch=?"; params.append(eff_branch)
        rows = [dict(r) for r in conn.execute(q, params).fetchall()]

        last_shifts = {}
        for r in conn.execute("""
            SELECT equipment_id, status, end_time, start_time
            FROM operator_shifts
            ORDER BY start_time DESC
        """).fetchall():
            eqid = r["equipment_id"] or ""
            if eqid and eqid not in last_shifts:
                last_shifts[eqid] = dict(r)

    now = datetime.utcnow()
    unused = []
    for e in rows:
        last = last_shifts.get(e["car_code"])
        if last and last["status"] == "active":
            continue  # شغالة دلوقتي
        if last:
            ref_time_str = last.get("end_time") or last.get("start_time")
        else:
            ref_time_str = e.get("created_at")
        idle_days = None
        if ref_time_str:
            try:
                ref_time = datetime.fromisoformat(ref_time_str.replace("Z", ""))
                idle_days = round((now - ref_time).total_seconds() / 86400, 1)
            except Exception:
                idle_days = None
        # لو مفيش أي تاريخ مرجعي، اعتبرها غير مستخدمة (لم تُستخدم إطلاقاً)
        if idle_days is None or idle_days >= min_idle_days:
            unused.append({
                "id":             e["id"],
                "car_code":       e["car_code"],
                "equipment_name": e["equipment_name"],
                "equipment_type": e.get("equipment_type", ""),
                "branch":         e.get("branch", ""),
                "last_used_at":   ref_time_str or None,
                "idle_days":      idle_days,
            })
    unused.sort(key=lambda x: (x["idle_days"] is None, -(x["idle_days"] or 0)))
    return {"unused_equipment": unused, "total": len(unused)}


@app.get("/equipment/report")
async def get_equipment_report(car_code: str = Query(...), cu: dict = Depends(require_admin_or_reporter)):
    """تقرير تفصيلي وقابل للطباعة عن معدة واحدة — كل الورديات، المشغلين، الساعات، الإنتاجية
    ملحوظة: car_code بتيجي كـ query param مش path param، لأن أكواد بعض المعدات فيها '/' (زي 00/01/65)
    وده كان بيكسر الراوت لو اتبعت جوا الـ path."""
    with get_db() as conn:
        _ensure_operator_tables(conn)
        eq_row = conn.execute("SELECT * FROM equipment WHERE car_code=?", (car_code,)).fetchone()
        if not eq_row:
            raise HTTPException(404, "المعدة غير موجودة")
        equipment = dict(eq_row)

        shift_rows = conn.execute("""
            SELECT s.*, o.name AS operator_name, o.phone AS operator_phone, o.national_id AS operator_national_id
            FROM operator_shifts s
            LEFT JOIN equipment_operators o ON o.id = s.operator_id
            WHERE s.equipment_id=?
            ORDER BY s.start_time DESC
        """, (car_code,)).fetchall()

    shifts = []
    total_hours = 0.0
    operators_set = set()
    for r in shift_rows:
        d = dict(r)
        actual_hours = None
        if d.get("status") == "ended" and d.get("end_hours") is not None and d.get("start_hours") is not None \
           and d["end_hours"] > d["start_hours"]:
            actual_hours = round(d["end_hours"] - d["start_hours"], 2)
            total_hours += actual_hours
        if d.get("operator_name"):
            operators_set.add(d["operator_name"])
        shifts.append({
            "id":              d["id"],
            "operator_name":   d.get("operator_name") or "—",
            "operator_phone":  d.get("operator_phone") or "",
            "shift_type":      d.get("shift_type", ""),
            "start_time":      d.get("start_time", ""),
            "end_time":        d.get("end_time", ""),
            "start_hours":     d.get("start_hours", 0),
            "end_hours":       d.get("end_hours", 0),
            "actual_hours":    actual_hours,
            "fuel_liters":     d.get("fuel_liters", 0),
            "fuel_type":       d.get("fuel_type", ""),
            "ownership_type":  d.get("ownership_type", ""),
            "project":         d.get("project", ""),
            "branch":          d.get("branch", ""),
            "notes":           d.get("notes", ""),
            "status":          d.get("status", ""),
            "start_photo":     d.get("start_photo", ""),
            "productivity":    d.get("productivity", ""),
        })

    shift_count = len(shifts)
    productivity = round(total_hours / shift_count, 2) if shift_count else 0.0
    summary = {
        "shift_count":    shift_count,
        "total_hours":    round(total_hours, 2),
        "operators":      sorted(operators_set),
        "productivity":   productivity,
        "first_shift_at": shifts[-1]["start_time"] if shifts else None,
        "last_shift_at":  shifts[0]["start_time"] if shifts else None,
    }
    return {"equipment": equipment, "shifts": shifts, "summary": summary}


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
    ownership_type: str = "مملوكة"
    branch:         str = ""
    sector:         str = ""
    project:        str = ""
    license_expiry: str = ""
    status:         str = "active"
    notes:          str = ""
    rental_source:  str = ""
    monthly_rate:   float = 0.0
    rental_start:   str = ""
    rental_end:     str = ""


@app.post("/equipment")
async def create_equipment(body: EquipmentCreate, cu: dict = Depends(require_admin)):
    if not body.car_code.strip():
        raise HTTPException(400, "الكود مطلوب")
    if not body.equipment_name.strip():
        raise HTTPException(400, "اسم المعدة مطلوب")
    try:
        with get_db() as conn:
            # migration دفاعي — أضف الأعمدة الجديدة لو مش موجودة
            for _col, _def in [
                ("ownership_type", "TEXT DEFAULT 'مملوكة'"),
                ("rental_source",  "TEXT DEFAULT ''"),
                ("monthly_rate",   "REAL DEFAULT 0"),
                ("rental_start",   "TEXT DEFAULT ''"),
                ("rental_end",     "TEXT DEFAULT ''"),
            ]:
                try: conn.execute(f"ALTER TABLE equipment ADD COLUMN {_col} {_def}")
                except Exception: pass
            conn.execute("""
                INSERT INTO equipment
                (car_code,equipment_name,brand,model,year,chassis,engine_number,
                 equipment_type,ownership_type,branch,sector,project,license_expiry,status,notes,
                 rental_source,monthly_rate,rental_start,rental_end)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (body.car_code.strip(), body.equipment_name.strip(),
                  body.brand, body.model, body.year, body.chassis, body.engine_number,
                  body.equipment_type, body.ownership_type,
                  body.branch, body.sector, body.project,
                  body.license_expiry, body.status, body.notes,
                  body.rental_source, body.monthly_rate, body.rental_start, body.rental_end))
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
        # migration دفاعي
        for _col, _def in [
            ("ownership_type", "TEXT DEFAULT 'مملوكة'"),
            ("rental_source",  "TEXT DEFAULT ''"),
            ("monthly_rate",   "REAL DEFAULT 0"),
            ("rental_start",   "TEXT DEFAULT ''"),
            ("rental_end",     "TEXT DEFAULT ''"),
        ]:
            try: conn.execute(f"ALTER TABLE equipment ADD COLUMN {_col} {_def}")
            except Exception: pass
        conn.execute("""
            UPDATE equipment SET
                car_code=?, equipment_name=?, brand=?, model=?, year=?, chassis=?,
                engine_number=?, equipment_type=?, ownership_type=?,
                branch=?, sector=?, project=?, license_expiry=?, status=?, notes=?,
                rental_source=?, monthly_rate=?, rental_start=?, rental_end=?
            WHERE id=?
        """, (body.car_code.strip(), body.equipment_name.strip(),
              body.brand, body.model, body.year, body.chassis, body.engine_number,
              body.equipment_type, body.ownership_type,
              body.branch, body.sector, body.project,
              body.license_expiry, body.status, body.notes,
              body.rental_source, body.monthly_rate, body.rental_start, body.rental_end,
              eq_id))
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


# ══════════════════════════════════════════════════════
# 48. EQUIPMENT OPERATORS — مشغلو المعدات (نظام ورديات)
# ══════════════════════════════════════════════════════

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
        fuel_type       TEXT DEFAULT '',
        ownership_type  TEXT DEFAULT 'مملوكة',
        project         TEXT DEFAULT '',
        branch          TEXT DEFAULT '',
        notes           TEXT DEFAULT '',
        status          TEXT DEFAULT 'active' CHECK(status IN ('active','ended')),
        start_location  TEXT DEFAULT '',
        end_location    TEXT DEFAULT '',
        created_at      TEXT DEFAULT (datetime('now')),
        FOREIGN KEY (operator_id) REFERENCES equipment_operators(id) ON DELETE CASCADE
    )""")
    # Migration: أضف location columns لو مش موجودة
    try: conn.execute("ALTER TABLE operator_shifts ADD COLUMN start_location TEXT DEFAULT ''")
    except Exception: pass
    try: conn.execute("ALTER TABLE operator_shifts ADD COLUMN end_location TEXT DEFAULT ''")
    except Exception: pass
    # Migration: أضف ownership_type و fuel_type لو مش موجودين
    try: conn.execute("ALTER TABLE operator_shifts ADD COLUMN ownership_type TEXT DEFAULT 'مملوكة'")
    except Exception: pass
    try: conn.execute("ALTER TABLE operator_shifts ADD COLUMN fuel_type TEXT DEFAULT ''")
    except Exception: pass
    # Migration: أضف صورة بدء الوردية (إلزامية عند البدء)
    try: conn.execute("ALTER TABLE operator_shifts ADD COLUMN start_photo TEXT DEFAULT ''")
    except Exception: pass
    try: conn.execute("CREATE INDEX IF NOT EXISTS idx_op_shifts_photo ON operator_shifts(start_photo)")
    except Exception: pass
    # Migration: أضف عمود إنتاجية المشغل (اختياري، بيدخله المشغل نفسه عند إنهاء الوردية)
    try: conn.execute("ALTER TABLE operator_shifts ADD COLUMN productivity TEXT DEFAULT ''")
    except Exception: pass
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

    # ── جدول صلاحيات المشغلين على المعدات (operator_equipment_permissions) ──
    conn.execute("""CREATE TABLE IF NOT EXISTS operator_equipment_permissions (
        id             INTEGER PRIMARY KEY AUTOINCREMENT,
        operator_id    INTEGER NOT NULL,
        equipment_id   TEXT    NOT NULL,
        equipment_name TEXT    DEFAULT '',
        granted_by     TEXT    DEFAULT '',
        granted_at     TEXT    DEFAULT (datetime('now')),
        UNIQUE(operator_id, equipment_id),
        FOREIGN KEY (operator_id) REFERENCES equipment_operators(id) ON DELETE CASCADE
    )""")
    try:
        conn.execute("CREATE INDEX IF NOT EXISTS idx_oep_operator ON operator_equipment_permissions(operator_id)")
    except Exception: pass
    try:
        conn.execute("CREATE INDEX IF NOT EXISTS idx_oep_equip ON operator_equipment_permissions(equipment_id)")
    except Exception: pass
    # Migration: أضف أعمدة جديدة لو مش موجودة
    for _oep_col, _oep_def in [
        ("equipment_name", "TEXT DEFAULT ''"),
        ("granted_by",     "TEXT DEFAULT ''"),
        ("granted_at",     "TEXT DEFAULT ''"),
    ]:
        try: conn.execute(f"ALTER TABLE operator_equipment_permissions ADD COLUMN {_oep_col} {_oep_def}")
        except Exception: pass


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
    fuel_type:      str = ""
    ownership_type: str = "مملوكة"
    project:        str = ""
    branch:         str = ""
    notes:          str = ""
    start_location: str = ""
    start_photo:    str = ""   # إلزامي — صورة المعدة/العداد عند بدء الوردية (base64)


class ShiftEnd(BaseModel):
    end_hours:    float = 0
    fuel_liters:  float = 0
    fuel_type:    str = ""
    notes:        str = ""
    end_location: str = ""
    productivity: str = ""   # اختياري — إنتاجية المشغل في اليوم، بيدخلها هو نفسه عند إنهاء الوردية


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


# ─────────────────────────────────────────────────────────
# Operator: GET by ID
# ─────────────────────────────────────────────────────────
@app.get("/operators/{oid}")
async def get_operator_by_id(oid: int, cu: dict = Depends(require_admin_or_reporter)):
    """جلب بيانات مشغل محدد"""
    with get_db() as conn:
        _ensure_operator_tables(conn)
        row = conn.execute("SELECT * FROM equipment_operators WHERE id=?", (oid,)).fetchone()
    if not row:
        raise HTTPException(404, "المشغل غير موجود")
    return dict(row)


# ─────────────────────────────────────────────────────────
# Operator Equipment Permissions — CRUD كامل
# ─────────────────────────────────────────────────────────

@app.get("/operators/{oid}/permissions")
async def get_operator_permissions(oid: int, cu: dict = Depends(get_user)):
    """جلب قائمة المعدات المصرح بها للمشغل."""
    if cu["role"] == "operator":
        if cu.get("operator_id") != oid:
            raise HTTPException(403, "يمكنك الاطلاع على صلاحياتك فقط")
    elif cu["role"] not in ("admin", "superuser", "reporter", "supervisor_workshop", "supervisor_field", "workshop_admin"):
        raise HTTPException(403, "صلاحيات غير كافية")
    with get_db() as conn:
        _ensure_operator_tables(conn)
        perms = conn.execute(
            """SELECT oep.*, e.equipment_name as eq_name_ref, e.equipment_type, e.branch as eq_branch
               FROM operator_equipment_permissions oep
               LEFT JOIN equipment e ON e.car_code = oep.equipment_id
               WHERE oep.operator_id=?
               ORDER BY oep.equipment_name""",
            (oid,)
        ).fetchall()
    return {"operator_id": oid, "permissions": [dict(p) for p in perms]}


@app.post("/operators/{oid}/permissions")
async def add_operator_permission(oid: int, body: dict, cu: dict = Depends(require_admin)):
    """منح مشغل صلاحية تشغيل معدة. body: { equipment_id, equipment_name }"""
    with get_db() as conn:
        _ensure_operator_tables(conn)
        op = conn.execute("SELECT id, name FROM equipment_operators WHERE id=?", (oid,)).fetchone()
        if not op:
            raise HTTPException(404, "المشغل غير موجود")
        eq_id   = (body.get("equipment_id") or "").strip()
        eq_name = (body.get("equipment_name") or "").strip()
        if not eq_id and not eq_name:
            raise HTTPException(400, "equipment_id أو equipment_name مطلوب")
        # استكمال البيانات من جدول equipment لو ناقصة
        if eq_id and not eq_name:
            r = conn.execute("SELECT equipment_name FROM equipment WHERE car_code=? LIMIT 1", (eq_id,)).fetchone()
            if r: eq_name = r["equipment_name"]
        if not eq_id and eq_name:
            r = conn.execute("SELECT car_code FROM equipment WHERE equipment_name=? LIMIT 1", (eq_name,)).fetchone()
            if r: eq_id = r["car_code"]
        key = eq_id or eq_name
        try:
            conn.execute(
                """INSERT INTO operator_equipment_permissions
                   (operator_id, equipment_id, equipment_name, granted_by, granted_at)
                   VALUES (?, ?, ?, ?, ?)""",
                (oid, key, eq_name, cu["username"], datetime.utcnow().isoformat())
            )
            perm_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        except Exception as e:
            if "UNIQUE" in str(e):
                raise HTTPException(400, f"الصلاحية ممنوحة بالفعل: {eq_name or eq_id}")
            raise HTTPException(500, str(e))
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "grant_operator_permission", f"منح مشغل [{op['name']}] صلاحية [{eq_name or eq_id}]")
    return {"ok": True, "id": perm_id, "operator_id": oid, "equipment_id": key, "equipment_name": eq_name}


@app.delete("/operators/{oid}/permissions/{pid}")
async def remove_operator_permission(oid: int, pid: int, cu: dict = Depends(require_admin)):
    """سحب صلاحية معدة من مشغل."""
    with get_db() as conn:
        _ensure_operator_tables(conn)
        perm = conn.execute(
            "SELECT * FROM operator_equipment_permissions WHERE id=? AND operator_id=?", (pid, oid)
        ).fetchone()
        if not perm:
            raise HTTPException(404, "الصلاحية غير موجودة")
        conn.execute("DELETE FROM operator_equipment_permissions WHERE id=?", (pid,))
    write_audit_log(cu["user_id"], cu["username"], cu["role"],
                    "revoke_operator_permission", f"سحب صلاحية [{perm['equipment_name']}] من مشغل [{oid}]")
    return {"ok": True, "deleted": pid}


@app.get("/operators/{oid}/allowed-equipment")
async def get_allowed_equipment(oid: int, cu: dict = Depends(get_user)):
    """المعدات المتاحة للمشغل — إذا لم تكن هناك صلاحيات محددة يرجع كل المعدات النشطة."""
    if cu["role"] == "operator":
        if cu.get("operator_id") != oid:
            raise HTTPException(403, "يمكنك الاطلاع على صلاحياتك فقط")
    elif cu["role"] not in ("admin", "superuser", "reporter", "supervisor_workshop", "supervisor_field", "workshop_admin"):
        raise HTTPException(403, "صلاحيات غير كافية")
    with get_db() as conn:
        _ensure_operator_tables(conn)
        perms = conn.execute(
            "SELECT equipment_id, equipment_name FROM operator_equipment_permissions WHERE operator_id=?",
            (oid,)
        ).fetchall()
        if perms:
            eq_ids = [p["equipment_id"] for p in perms]
            ph = ",".join("?" * len(eq_ids))
            equipment = conn.execute(
                f"""SELECT car_code as equipment_id, equipment_name, equipment_type,
                           branch, sector, project, status
                    FROM equipment WHERE car_code IN ({ph}) AND status='active'
                    ORDER BY equipment_name""", eq_ids
            ).fetchall()
            eq_from_table = {e["equipment_id"] for e in equipment}
            extras = [{"equipment_id": p["equipment_id"], "equipment_name": p["equipment_name"],
                       "equipment_type": "", "branch": "", "sector": "", "project": "", "status": "active"}
                      for p in perms if p["equipment_id"] not in eq_from_table]
            result = [dict(e) for e in equipment] + extras
        else:
            equipment = conn.execute(
                """SELECT car_code as equipment_id, equipment_name, equipment_type,
                          branch, sector, project, status
                   FROM equipment WHERE status='active' ORDER BY equipment_name"""
            ).fetchall()
            result = [dict(e) for e in equipment]
    return {"operator_id": oid, "equipment": result, "restricted": len(perms) > 0}


@app.get("/operators/{oid}/stats")
async def get_operator_stats(oid: int, cu: dict = Depends(get_user)):
    """إحصائيات مشغل محدد."""
    if cu["role"] == "operator":
        if cu.get("operator_id") != oid:
            raise HTTPException(403, "يمكنك الاطلاع على بياناتك فقط")
    elif cu["role"] not in ("admin", "superuser", "reporter", "supervisor_workshop", "supervisor_field", "workshop_admin"):
        raise HTTPException(403, "صلاحيات غير كافية")
    with get_db() as conn:
        _ensure_operator_tables(conn)
        total_shifts  = conn.execute("SELECT COUNT(*) FROM operator_shifts WHERE operator_id=?", (oid,)).fetchone()[0]
        active_shifts = conn.execute("SELECT COUNT(*) FROM operator_shifts WHERE operator_id=? AND status='active'", (oid,)).fetchone()[0]
        ended_shifts  = conn.execute("SELECT COUNT(*) FROM operator_shifts WHERE operator_id=? AND status='ended'", (oid,)).fetchone()[0]
        hours_row     = conn.execute(
            "SELECT SUM(end_hours - start_hours) FROM operator_shifts WHERE operator_id=? AND status='ended' AND end_hours > start_hours", (oid,)
        ).fetchone()
        total_hours   = round(hours_row[0] or 0, 2)
        fuel_row      = conn.execute("SELECT SUM(fuel_liters) FROM operator_shifts WHERE operator_id=?", (oid,)).fetchone()
        total_fuel    = round(fuel_row[0] or 0, 2)
        last_shift    = conn.execute(
            "SELECT * FROM operator_shifts WHERE operator_id=? ORDER BY start_time DESC LIMIT 1", (oid,)
        ).fetchone()
        week_ago      = (datetime.utcnow() - timedelta(days=7)).isoformat()
        week_shifts   = conn.execute(
            "SELECT COUNT(*) FROM operator_shifts WHERE operator_id=? AND start_time >= ?", (oid, week_ago)
        ).fetchone()[0]
        allowed_count = conn.execute(
            "SELECT COUNT(*) FROM operator_equipment_permissions WHERE operator_id=?", (oid,)
        ).fetchone()[0]
    return {
        "operator_id": oid, "total_shifts": total_shifts, "active_shifts": active_shifts,
        "ended_shifts": ended_shifts, "total_hours": total_hours, "total_fuel": total_fuel,
        "week_shifts": week_shifts, "allowed_equipment_count": allowed_count,
        "last_shift": dict(last_shift) if last_shift else None,
    }


# ─────────────────────────────────────────────────────────
# Shifts: قائمة كاملة مع فلاتر + تصدير CSV
# ─────────────────────────────────────────────────────────

@app.get("/shifts")
async def list_all_shifts(
    operator_id: int = Query(None),
    equipment_id: str = Query(None),
    branch:      str = Query(None),
    project:     str = Query(None),
    status:      str = Query(None),
    date_from:   str = Query(None),
    date_to:     str = Query(None),
    search:      str = Query(None),
    limit:       int = Query(100),
    offset:      int = Query(0),
    cu: dict = Depends(require_admin_or_reporter),
):
    """كل الورديات مع فلاتر كاملة"""
    with get_db() as conn:
        _ensure_operator_tables(conn)
        q = """SELECT s.*, o.name as operator_name, o.phone as operator_phone,
                      o.fixed_number as operator_fixed, o.branch as operator_branch,
                      e.ownership_type as eq_ownership_type,
                      e.equipment_type as eq_type_label
               FROM operator_shifts s
               LEFT JOIN equipment_operators o ON o.id = s.operator_id
               LEFT JOIN equipment e ON e.car_code = s.equipment_id
               WHERE 1=1"""
        params: list = []
        eff_branch = _branch_filter(cu) or branch
        if eff_branch:
            q += " AND (s.branch=? OR o.branch=?)"; params.extend([eff_branch, eff_branch])
        if operator_id:
            q += " AND s.operator_id=?"; params.append(operator_id)
        if equipment_id:
            q += " AND s.equipment_id=?"; params.append(equipment_id)
        if project:
            q += " AND s.project LIKE ?"; params.append(f"%{project}%")
        if status:
            q += " AND s.status=?"; params.append(status)
        if date_from:
            q += " AND s.start_time >= ?"; params.append(date_from)
        if date_to:
            q += " AND s.start_time <= ?"; params.append(date_to + "T23:59:59")
        if search:
            s = f"%{search}%"
            q += " AND (o.name LIKE ? OR s.equipment_name LIKE ? OR s.project LIKE ?)"
            params.extend([s, s, s])
        total = conn.execute(f"SELECT COUNT(*) FROM ({q})", params).fetchone()[0]
        q += " ORDER BY s.start_time DESC LIMIT ? OFFSET ?"; params.extend([limit, offset])
        rows = [dict(r) for r in conn.execute(q, params).fetchall()]
        for r in rows:
            if r.get("start_hours") is not None and r.get("end_hours") is not None and r["end_hours"] > r["start_hours"]:
                r["working_hours"] = round(r["end_hours"] - r["start_hours"], 2)
            else:
                r["working_hours"] = None
            r["ownership_type"] = r.get("eq_ownership_type") or r.get("ownership_type") or "مملوكة"
            r["equipment_type_label"] = r.get("eq_type_label") or ""
    return {"total": total, "shifts": rows}


@app.get("/shifts/export")
async def export_shifts_csv(
    operator_id: int = Query(None),
    branch:      str = Query(None),
    date_from:   str = Query(None),
    date_to:     str = Query(None),
    cu: dict = Depends(require_admin_or_reporter),
):
    """تصدير الورديات CSV"""
    from fastapi.responses import Response
    with get_db() as conn:
        _ensure_operator_tables(conn)
        q = """SELECT s.*, o.name as operator_name, o.fixed_number as operator_fixed
               FROM operator_shifts s
               LEFT JOIN equipment_operators o ON o.id = s.operator_id WHERE 1=1"""
        params: list = []
        eff_branch = _branch_filter(cu) or branch
        if eff_branch:
            q += " AND (s.branch=? OR o.branch=?)"; params.extend([eff_branch, eff_branch])
        if operator_id:
            q += " AND s.operator_id=?"; params.append(operator_id)
        if date_from:
            q += " AND s.start_time >= ?"; params.append(date_from)
        if date_to:
            q += " AND s.start_time <= ?"; params.append(date_to + "T23:59:59")
        q += " ORDER BY s.start_time DESC"
        rows = [dict(r) for r in conn.execute(q, params).fetchall()]
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["#","المشغل","الرقم الثابت","المعدة","كود المعدة",
                "نوع الوردية","وقت البداية","وقت النهاية",
                "ساعات البداية","ساعات النهاية","مدة التشغيل (ساعة)",
                "نوع الوقود","الوقود (لتر)","الملكية","المشروع","الفرع",
                "موقع البداية","موقع النهاية","الحالة","ملاحظات"])
    type_labels = {"day":"نهارية","night":"ليلية","extended":"ممتدة"}
    fuel_labels = {"fuel_solar":"سولار","fuel_92":"بنزين 92","fuel_95":"بنزين 95","fuel_80":"بنزين 80","fuel_cng":"غاز طبيعي"}
    for i, r in enumerate(rows, 1):
        wh = None
        if r.get("start_hours") is not None and r.get("end_hours") is not None and r["end_hours"] > r["start_hours"]:
            wh = round(r["end_hours"] - r["start_hours"], 2)
        w.writerow([i, r.get("operator_name",""), r.get("operator_fixed",""),
                    r.get("equipment_name",""), r.get("equipment_id",""),
                    type_labels.get(r.get("shift_type",""), r.get("shift_type","")),
                    (r.get("start_time") or "")[:16].replace("T"," "),
                    (r.get("end_time") or "")[:16].replace("T"," "),
                    r.get("start_hours",""), r.get("end_hours",""), wh or "",
                    fuel_labels.get(r.get("fuel_type",""), r.get("fuel_type","")),
                    r.get("fuel_liters",""),
                    r.get("ownership_type","مملوكة"),
                    r.get("project",""), r.get("branch",""),
                    r.get("start_location",""), r.get("end_location",""),
                    "نشطة" if r.get("status")=="active" else "منتهية", r.get("notes","")])
    fname = f"shifts_{date_from or 'all'}_{date_to or 'all'}.csv"
    return Response(content=buf.getvalue().encode("utf-8-sig"),
                    media_type="text/csv",
                    headers={"Content-Disposition": f"attachment; filename={fname}"})


# ── Shifts ──
@app.get("/operators/{oid}/shifts")
async def get_operator_shifts(
    oid: int,
    limit: int = Query(20),
    cu: dict = Depends(get_user),
):
    # المشغل يقدر يجيب ورديات نفسه فقط
    if cu["role"] == "operator":
        if cu.get("operator_id") != oid:
            raise HTTPException(403, "يمكنك الاطلاع على ورديات حسابك فقط")
    elif cu["role"] not in ("admin", "superuser", "reporter", "supervisor_workshop", "supervisor_field", "workshop_admin"):
        raise HTTPException(403, "صلاحيات غير كافية")
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
async def start_shift(body: ShiftStart, cu: dict = Depends(get_user)):
    # المشغل يبدأ وردية لنفسه فقط
    if cu["role"] == "operator":
        if cu.get("operator_id") != body.operator_id:
            raise HTTPException(403, "يمكنك بدء وردية لحسابك فقط")
    elif cu["role"] not in ("admin", "superuser"):
        raise HTTPException(403, "صلاحيات غير كافية")
    with get_db() as conn:
        _ensure_operator_tables(conn)
        # تحقق من عدم وجود وردية نشطة
        active = conn.execute(
            "SELECT id FROM operator_shifts WHERE operator_id=? AND status='active'",
            (body.operator_id,)).fetchone()
        if active:
            raise HTTPException(400, "يوجد وردية جارية بالفعل — أنهِ الوردية الحالية أولاً")
        if body.fuel_liters and body.fuel_liters > 400:
            raise HTTPException(400, "الحد الأقصى لكمية الوقود للمعدات هو 400 لتر")
        # ── صورة بدء الوردية: إلزامية دائماً ──
        if not (body.start_photo or "").strip():
            raise HTTPException(400, "📷 صورة بدء الوردية مطلوبة — ارفع صورة قبل البدء")

        def _save_shift_photo(b64_data, prefix):
            if not b64_data:
                return ""
            import re as _re_sh
            try:
                _m = _re_sh.match(r"data:image/(\w+);base64,(.+)", b64_data, _re_sh.DOTALL)
                _ext = _m.group(1) if _m else "jpg"
                _b64data = _m.group(2) if _m else b64_data
                _raw_bytes = base64.b64decode(_b64data)
                if len(_raw_bytes) >= 2 and _raw_bytes[0] == 0xFF and _raw_bytes[1] == 0xD8: _ext = "jpg"
                elif len(_raw_bytes) >= 4 and _raw_bytes[0] == 0x89 and _raw_bytes[1:4] == b"PNG": _ext = "png"
                elif len(_raw_bytes) >= 3 and _raw_bytes[:3] == b"GIF": _ext = "gif"
                elif len(_raw_bytes) >= 12 and _raw_bytes[:4] == b"RIFF" and _raw_bytes[8:12] == b"WEBP": _ext = "webp"
                _fname = f"{prefix}_{body.operator_id or 0}_{uuid.uuid4().hex}.{_ext}"
                (UPLOAD_DIR / _fname).write_bytes(_raw_bytes)
                return f"/uploads/{_fname}"
            except Exception as _photo_err:
                log.error(f"[SHIFT] photo save failed ({prefix}): {_photo_err}")
                raise HTTPException(400, "صورة غير صالحة")

        start_photo_url = _save_shift_photo(body.start_photo, "shift_start")
        if not start_photo_url:
            raise HTTPException(400, "📷 صورة بدء الوردية مطلوبة — ارفع صورة قبل البدء")

        # اسحب نوع الملكية الحقيقي من جدول المعدات (مش من الـ body) — الملكية خاصية المعدة نفسها
        eq_row = conn.execute("SELECT ownership_type FROM equipment WHERE car_code=?", (body.equipment_id,)).fetchone()
        real_ownership = (eq_row["ownership_type"] if eq_row and eq_row["ownership_type"] else None) or body.ownership_type
        conn.execute("""INSERT INTO operator_shifts
            (operator_id,equipment_id,equipment_name,shift_type,start_time,
             start_hours,fuel_liters,fuel_type,ownership_type,project,branch,notes,start_location,start_photo,status)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,'active')""",
            (body.operator_id, body.equipment_id, body.equipment_name, body.shift_type,
             datetime.utcnow().isoformat(), body.start_hours, body.fuel_liters,
             body.fuel_type, real_ownership,
             body.project, body.branch, body.notes, body.start_location or "", start_photo_url))
        sid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    return {"id": sid, "message": "بدأت الوردية", "start_photo": start_photo_url}


@app.post("/operators/shifts/{sid}/end")
async def end_shift(sid: int, body: ShiftEnd, cu: dict = Depends(get_user)):
    with get_db() as conn:
        _ensure_operator_tables(conn)
        shift = conn.execute("SELECT * FROM operator_shifts WHERE id=?", (sid,)).fetchone()
        if not shift: raise HTTPException(404, "الوردية غير موجودة")
        if shift["status"] == "ended": raise HTTPException(400, "الوردية منتهية بالفعل")
        if body.fuel_liters and body.fuel_liters > 400:
            raise HTTPException(400, "الحد الأقصى لكمية الوقود للمعدات هو 400 لتر")
        # المشغل يقدر ينهي وردية نفسه فقط
        if cu["role"] == "operator":
            if cu.get("operator_id") != shift["operator_id"]:
                raise HTTPException(403, "يمكنك إنهاء ورديات حسابك فقط")
        elif cu["role"] not in ("admin", "superuser"):
            raise HTTPException(403, "صلاحيات غير كافية")
        conn.execute("""UPDATE operator_shifts SET
            end_time=?, end_hours=?, fuel_liters=?, fuel_type=?, notes=?, end_location=?, productivity=?, status='ended'
            WHERE id=?""",
            (datetime.utcnow().isoformat(), body.end_hours, body.fuel_liters,
             body.fuel_type or shift["fuel_type"] or "",
             body.notes or shift["notes"], body.end_location or "",
             (body.productivity or "").strip(), sid))
    return {"message": "انتهت الوردية"}


@app.put("/shifts/{sid}/hours")
async def edit_shift_hours(sid: int, body: dict, cu: dict = Depends(require_admin)):
    """تعديل ساعات عداد بداية/نهاية الوردية — أدمن / سوبر يوزر فقط."""
    start_hours = body.get("start_hours")
    end_hours   = body.get("end_hours")
    if start_hours is None:
        raise HTTPException(400, "ساعات البداية مطلوبة")
    start_hours = float(start_hours)
    if start_hours < 0:
        raise HTTPException(400, "ساعات البداية يجب أن تكون موجبة")
    if end_hours is not None:
        end_hours = float(end_hours)
        if end_hours > 0 and end_hours <= start_hours:
            raise HTTPException(400, "ساعات النهاية يجب أن تكون أكبر من البداية")
    with get_db() as conn:
        _ensure_operator_tables(conn)
        row = conn.execute("SELECT id, start_hours, end_hours, status FROM operator_shifts WHERE id=?", (sid,)).fetchone()
        if not row:
            raise HTTPException(404, "الوردية غير موجودة")
        eff_end = end_hours if end_hours is not None else row["end_hours"]
        conn.execute(
            "UPDATE operator_shifts SET start_hours=?, end_hours=? WHERE id=?",
            (start_hours, eff_end, sid)
        )
        log_event("shift_hours_edited", admin=cu["username"], shift_id=sid,
                  start_hours=start_hours, end_hours=eff_end)
    return {"ok": True, "shift_id": sid, "start_hours": start_hours, "end_hours": eff_end}


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

# ══════════════════════════════════════════════════════
# 49. السركي — تقرير حضور السائقين من الرحلات
# ══════════════════════════════════════════════════════

@app.get("/reports/sarky")
async def sarky_report(
    driver_id: int  = Query(None),
    date_from: str  = Query(None, description="YYYY-MM-DD"),
    date_to:   str  = Query(None, description="YYYY-MM-DD"),
    branch:    str  = Query(None),
    cu: dict = Depends(require_admin_or_reporter),
):
    """
    السركي: تقرير حضور السائقين المُحتسب من الرحلات.
    يوم الحضور = أي يوم فيه رحلة واحدة على الأقل منتهية مع garage_location مسجّل.
    """
    eff_branch = _branch_filter(cu) or branch

    # defaults: الشهر الحالي
    if not date_from:
        date_from = datetime.utcnow().strftime("%Y-%m-01")
    if not date_to:
        date_to = datetime.utcnow().strftime("%Y-%m-%d")

    with get_db() as conn:
        c = conn.cursor()

        # ── جيب السائقين ──
        q_drv = "SELECT d.id, d.name, d.branch, d.fixed_number FROM drivers d WHERE 1=1"
        p_drv: list = []
        if eff_branch:
            q_drv += " AND d.branch=?"
            p_drv.append(eff_branch)
        if driver_id:
            q_drv += " AND d.id=?"
            p_drv.append(driver_id)
        q_drv += " ORDER BY d.name"
        c.execute(q_drv, p_drv)
        drivers = [dict(r) for r in c.fetchall()]

        results = []
        for drv in drivers:
            did = drv["id"]

            # رحلات منتهية في الفترة — مع أول وقت بداية وآخر وقت نهاية لكل يوم
            c.execute("""
                SELECT
                    date(start_time)                            AS trip_date,
                    COUNT(*)                                    AS trip_count,
                    COALESCE(SUM(
                        CASE WHEN end_odometer > start_odometer
                             THEN end_odometer - start_odometer
                             ELSE 0 END
                    ), 0)                                       AS km_total,
                    MIN(start_time)                             AS first_trip_time,
                    MAX(end_time)                               AS last_trip_time
                FROM trips
                WHERE driver_id = ?
                  AND end_time   IS NOT NULL
                  AND date(start_time) BETWEEN ? AND ?
                GROUP BY date(start_time)
                ORDER BY date(start_time)
            """, (did, date_from, date_to))
            days_raw = [dict(r) for r in c.fetchall()]

            days = []
            for day in days_raw:
                ft = day.get("first_trip_time")
                lt = day.get("last_trip_time")
                first_fmt = ""
                last_fmt  = ""
                tz_offset = timedelta(hours=3)   # UTC+3 القاهرة
                if ft:
                    try:
                        dt_ft = datetime.fromisoformat(ft.replace("Z","")) + tz_offset
                        first_fmt = dt_ft.strftime("%H:%M")
                    except Exception:
                        first_fmt = ""
                if lt:
                    try:
                        dt_lt = datetime.fromisoformat(lt.replace("Z","")) + tz_offset
                        last_fmt = dt_lt.strftime("%H:%M")
                    except Exception:
                        pass
                days.append({**day, "first_trip_time": first_fmt, "garage_time": last_fmt})

            total_days  = len(days)
            total_trips = sum(d["trip_count"] for d in days)
            total_km    = round(sum(d["km_total"] for d in days), 2)

            # ── جيب بيانات KPI الخام (بدون حساب score دلوقتي) ──
            c.execute("""
                SELECT COUNT(*) as cnt,
                       COALESCE(SUM(end_odometer - start_odometer),0) as km
                FROM trips
                WHERE driver_id=? AND end_time IS NOT NULL
                  AND end_odometer IS NOT NULL
                  AND date(start_time) BETWEEN ? AND ?
            """, (did, date_from, date_to))
            kpi_tr = dict(c.fetchone())

            c.execute("""
                SELECT COALESCE(SUM(quantity),0) as liters
                FROM workshop_records
                WHERE driver_id=? AND type LIKE 'fuel_%'
                  AND date(created_at) BETWEEN ? AND ?
            """, (did, date_from, date_to))
            kpi_fuel = float(c.fetchone()[0] or 0)

            c.execute("""
                SELECT COALESCE(SUM(quantity),0) as liters
                FROM workshop_records
                WHERE driver_id=? AND type IN (
                    'oil','filter','filter_oil','filter_solar','filter_petrol',
                    'filter_air','filter_separator','filter_ac','filter_dryer',
                    'filter_breather','filter_hydraulic'
                ) AND date(created_at) BETWEEN ? AND ?
            """, (did, date_from, date_to))
            kpi_oil = float(c.fetchone()[0] or 0)

            c.execute("""
                SELECT COUNT(*) FROM emergency_reports
                WHERE driver_id=? AND date(created_at) BETWEEN ? AND ?
            """, (did, date_from, date_to))
            kpi_emg = c.fetchone()[0] or 0

            results.append({
                "driver_id":    did,
                "driver_name":  drv["name"],
                "branch":       drv["branch"] or "",
                "fixed_number": drv["fixed_number"] or "",
                "total_days":   total_days,
                "total_trips":  total_trips,
                "total_km":     total_km,
                "days_detail":  days,
            })

        # رتّب: الأكتر حضوراً أول
        results.sort(key=lambda x: (-x["total_days"], -x["total_km"]))
        return {
            "date_from": date_from,
            "date_to":   date_to,
            "branch":    eff_branch or "الكل",
            "drivers":   results,
        }

# ══════════════════════════════════════════════════════
# 50. 📷 OCR Odometer — EasyOCR (مجاني 100%، pip فقط، بدون apt)
#  السائق يرفع صورة العداد → الباك-إند يقرأها بـ EasyOCR
# ══════════════════════════════════════════════════════
@app.post("/ocr/odometer")
async def ocr_odometer(
    file: UploadFile = File(...),
    cu: dict = Depends(get_user)
):
    """قراءة رقم العداد من صورة عبر EasyOCR (مجاني ومحلي)."""
    import re as _re
    try:
        import easyocr
        import numpy as np
        from PIL import Image, ImageEnhance, ImageFilter
        import io
    except ImportError:
        raise HTTPException(500, "مكتبة easyocr غير مثبتة — أضف easyocr لـ requirements.txt")

    raw_bytes = await file.read()
    if len(raw_bytes) == 0:
        raise HTTPException(400, "الصورة فارغة")
    if len(raw_bytes) > 10 * 1024 * 1024:
        raise HTTPException(400, "حجم الصورة كبير جداً (الحد الأقصى 10 ميجا)")

    try:
        # تحميل الصورة وتحسينها
        img = Image.open(io.BytesIO(raw_bytes)).convert("RGB")

        # تكبير لو صغيرة
        w, h = img.size
        if min(w, h) < 600:
            scale = 600 // min(w, h)
            img = img.resize((w * scale, h * scale), Image.LANCZOS)

        # زيادة contrast
        img = ImageEnhance.Contrast(img).enhance(2.0)
        img = img.filter(ImageFilter.SHARPEN)

        img_array = np.array(img)

        # EasyOCR — digit_only mode
        reader = easyocr.Reader(['en'], gpu=False, verbose=False)
        results = reader.readtext(img_array, allowlist='0123456789', detail=0)

        raw_text = " ".join(results)
        log.info(f"[OCR] EasyOCR raw: {raw_text!r}  size={len(raw_bytes)}")

        # استخراج أكبر رقم في نطاق عداد السيارة
        numbers = _re.findall(r'\d+', raw_text.replace(",", "").replace(".", ""))
        candidates = [int(n) for n in numbers if 1000 <= int(n) <= 999999]
        if candidates:
            reading = str(max(candidates))
        else:
            all_nums = [int(n) for n in numbers if n]
            reading = str(max(all_nums)) if all_nums else ""

        log.info(f"[OCR] reading={reading!r} candidates={candidates}")
        return {"ok": True, "reading": reading, "raw": raw_text}

    except HTTPException:
        raise
    except Exception as e:
        log.error(f"[OCR] EasyOCR error: {e}")
        raise HTTPException(502, "خطأ في قراءة الصورة — يرجى المحاولة مرة أخرى أو إدخال الرقم يدوياً")

# ══════════════════════════════════════════════════════
# 51. VIRTUAL INVENTORY MODULE — إدارة المخزون
# ══════════════════════════════════════════════════════

class InventoryCategoryCreate(BaseModel):
    name: str; notes: Optional[str] = ""

class InventorySupplierCreate(BaseModel):
    name: str; phone: Optional[str] = ""; email: Optional[str] = ""
    address: Optional[str] = ""; notes: Optional[str] = ""
    company_name: Optional[str] = ""; manager_name: Optional[str] = ""; business_sector: Optional[str] = ""

class InventoryProductCreate(BaseModel):
    name: str; sku: Optional[str] = ""
    category_id: Optional[int] = None; supplier_id: Optional[int] = None
    brand: Optional[str] = ""
    purchase_price: float = 0; sale_price: float = 0
    quantity: float = 0; min_quantity: float = 0
    unit: Optional[str] = "قطعة"; notes: Optional[str] = ""

class InventoryMovementCreate(BaseModel):
    product_id: int; movement_type: str  # in | out | return | adjust
    quantity: float; notes: Optional[str] = ""

class SupplyPlatformCreate(BaseModel):
    code: str; plate: Optional[str] = ""
    driver_name: Optional[str] = ""; manager_name: Optional[str] = ""
    license_number: Optional[str] = ""; status: Optional[str] = "active"
    notes: Optional[str] = ""

class SupplyPlatformFill(BaseModel):
    product_id: int; quantity: float; notes: Optional[str] = ""

class SupplyPlatformAdjust(BaseModel):
    product_id: int; quantity: float; notes: Optional[str] = ""  # تعيين الكمية مباشرة

class RepairQuoteCreate(BaseModel):
    quote_number: Optional[str] = ""
    car_id: Optional[int] = None
    equipment_id: Optional[str] = ""   # كود المعدة (car_code) — بديل عن car_id لو المقايسة على معدة مش مركبة
    driver_id: Optional[int] = None
    quote_date: Optional[str] = ""; status: Optional[str] = "draft"
    items_json: Optional[str] = "[]"; total_value: Optional[float] = 0
    notes: Optional[str] = ""

def _validate_quote_target(conn, body: "RepairQuoteCreate"):
    """لازم يتحدد إما مركبة (car_id) أو معدة (equipment_id) — واحد منهم إجباري."""
    car_id = body.car_id if (body.car_id and body.car_id > 0) else None
    equipment_id = (body.equipment_id or "").strip()
    if not car_id and not equipment_id:
        raise HTTPException(400, "لازم تدخل رقم المعدة أو نمرة العربية")
    if car_id and not conn.execute("SELECT id FROM cars WHERE id=?", (car_id,)).fetchone():
        raise HTTPException(404, "المركبة غير موجودة")
    if equipment_id and not conn.execute("SELECT id FROM equipment WHERE car_code=?", (equipment_id,)).fetchone():
        raise HTTPException(404, "رقم المعدة غير موجود")
    return car_id, equipment_id

class WorkshopApproval(BaseModel):
    action: str  # approve | reject
    note: Optional[str] = ""


def _inv_product_row(r: dict, cat_map: dict, sup_map: dict) -> dict:
    d = dict(r)
    qty = d.get("quantity") or 0
    minq = d.get("min_quantity") or 0
    d["category_name"] = cat_map.get(d.get("category_id"), "")
    d["supplier_name"] = sup_map.get(d.get("supplier_id"), "")
    if qty <= 0:
        d["status"] = "غير متوفر"
    elif minq > 0 and qty <= minq:
        d["status"] = "منخفض"
    else:
        d["status"] = "متوفر"
    return d


# ── الفئات ──
@app.get("/inventory/categories")
async def list_inventory_categories(cu: dict = Depends(get_user)):
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM inventory_categories ORDER BY name").fetchall()
        return [dict(r) for r in rows]

@app.post("/inventory/categories")
async def create_inventory_category(body: InventoryCategoryCreate, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        try:
            cur = conn.execute("INSERT INTO inventory_categories(name,notes) VALUES(?,?)", (body.name.strip(), body.notes or ""))
            return {"id": cur.lastrowid, "name": body.name, "notes": body.notes or ""}
        except sqlite3.IntegrityError:
            raise HTTPException(400, "الفئة موجودة مسبقاً")

@app.put("/inventory/categories/{cid}")
async def update_inventory_category(cid: int, body: InventoryCategoryCreate, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        conn.execute("UPDATE inventory_categories SET name=?, notes=? WHERE id=?", (body.name.strip(), body.notes or "", cid))
        return {"ok": True}

@app.delete("/inventory/categories/{cid}")
async def delete_inventory_category(cid: int, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        in_use = conn.execute("SELECT COUNT(*) c FROM inventory_products WHERE category_id=?", (cid,)).fetchone()["c"]
        if in_use:
            raise HTTPException(400, f"لا يمكن حذف الفئة — مرتبطة بـ {in_use} منتج")
        conn.execute("DELETE FROM inventory_categories WHERE id=?", (cid,))
        return {"ok": True}


# ── الموردون ──
@app.get("/inventory/suppliers")
async def list_inventory_suppliers(cu: dict = Depends(get_user)):
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM inventory_suppliers ORDER BY name").fetchall()
        return [dict(r) for r in rows]

@app.post("/inventory/suppliers")
async def create_inventory_supplier(body: InventorySupplierCreate, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        cur = conn.execute("""INSERT INTO inventory_suppliers(name,phone,email,address,notes,company_name,manager_name,business_sector)
                              VALUES(?,?,?,?,?,?,?,?)""",
                            (body.name.strip(), body.phone or "", body.email or "", body.address or "", body.notes or "",
                             body.company_name or "", body.manager_name or "", body.business_sector or ""))
        return {"id": cur.lastrowid}

@app.put("/inventory/suppliers/{sid}")
async def update_inventory_supplier(sid: int, body: InventorySupplierCreate, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        conn.execute("""UPDATE inventory_suppliers SET name=?,phone=?,email=?,address=?,notes=?,
                       company_name=?,manager_name=?,business_sector=? WHERE id=?""",
                      (body.name.strip(), body.phone or "", body.email or "", body.address or "", body.notes or "",
                       body.company_name or "", body.manager_name or "", body.business_sector or "", sid))
        return {"ok": True}

@app.delete("/inventory/suppliers/{sid}")
async def delete_inventory_supplier(sid: int, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        conn.execute("UPDATE inventory_products SET supplier_id=NULL WHERE supplier_id=?", (sid,))
        conn.execute("DELETE FROM inventory_suppliers WHERE id=?", (sid,))
        return {"ok": True}


# ── المنتجات ──
@app.get("/inventory/products")
async def list_inventory_products(cu: dict = Depends(get_user), category_id: Optional[int] = None,
                                   status: Optional[str] = None, search: Optional[str] = None):
    with get_db() as conn:
        cats = {r["id"]: r["name"] for r in conn.execute("SELECT id,name FROM inventory_categories").fetchall()}
        sups = {r["id"]: r["name"] for r in conn.execute("SELECT id,name FROM inventory_suppliers").fetchall()}
        q = "SELECT * FROM inventory_products WHERE 1=1"
        params = []
        if category_id:
            q += " AND category_id=?"; params.append(category_id)
        if search:
            q += " AND (name LIKE ? OR sku LIKE ?)"; params += [f"%{search}%", f"%{search}%"]
        q += " ORDER BY id DESC"
        rows = [_inv_product_row(dict(r), cats, sups) for r in conn.execute(q, params).fetchall()]
        if status:
            rows = [r for r in rows if r["status"] == status]
        return rows

@app.post("/inventory/products")
async def create_inventory_product(body: InventoryProductCreate, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        now = datetime.utcnow().isoformat() + "Z"
        cur = conn.execute("""INSERT INTO inventory_products
                              (name,sku,category_id,supplier_id,brand,purchase_price,sale_price,quantity,min_quantity,unit,notes,created_at,updated_at)
                              VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (body.name.strip(), body.sku or "", body.category_id, body.supplier_id, body.brand or "",
                             body.purchase_price, body.sale_price, body.quantity, body.min_quantity,
                             body.unit or "قطعة", body.notes or "", now, now))
        pid = cur.lastrowid
        if body.quantity and body.quantity > 0:
            conn.execute("""INSERT INTO inventory_movements(product_id,movement_type,quantity,ref_type,username,notes,created_at)
                           VALUES(?,'in',?,'initial',?,?,?)""", (pid, body.quantity, cu.get("username",""), "رصيد افتتاحي", now))
        return {"id": pid}

@app.put("/inventory/products/{pid}")
async def update_inventory_product(pid: int, body: InventoryProductCreate, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        now = datetime.utcnow().isoformat() + "Z"
        conn.execute("""UPDATE inventory_products SET name=?,sku=?,category_id=?,supplier_id=?,brand=?,
                       purchase_price=?,sale_price=?,min_quantity=?,unit=?,notes=?,updated_at=? WHERE id=?""",
                      (body.name.strip(), body.sku or "", body.category_id, body.supplier_id, body.brand or "",
                       body.purchase_price, body.sale_price, body.min_quantity, body.unit or "قطعة",
                       body.notes or "", now, pid))
        return {"ok": True}

@app.delete("/inventory/products/{pid}")
async def delete_inventory_product(pid: int, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        conn.execute("DELETE FROM inventory_products WHERE id=?", (pid,))
        conn.execute("DELETE FROM inventory_movements WHERE product_id=?", (pid,))
        return {"ok": True}


# ── حركات المخزون ──
@app.get("/inventory/movements")
async def list_inventory_movements(cu: dict = Depends(require_superuser), product_id: Optional[int] = None, limit: int = 500):
    with get_db() as conn:
        q = """SELECT m.*, p.name as product_name, p.unit as product_unit
               FROM inventory_movements m LEFT JOIN inventory_products p ON p.id=m.product_id WHERE 1=1"""
        params = []
        if product_id:
            q += " AND m.product_id=?"; params.append(product_id)
        q += " ORDER BY m.id DESC LIMIT ?"; params.append(limit)
        return [dict(r) for r in conn.execute(q, params).fetchall()]

@app.post("/inventory/movements")
async def create_inventory_movement(body: InventoryMovementCreate, cu: dict = Depends(require_superuser)):
    if body.movement_type not in ("in", "out", "return", "adjust"):
        raise HTTPException(400, "نوع حركة غير صالح")
    with get_db() as conn:
        prow = conn.execute("SELECT quantity FROM inventory_products WHERE id=?", (body.product_id,)).fetchone()
        if not prow:
            raise HTTPException(404, "المنتج غير موجود")
        now = datetime.utcnow().isoformat() + "Z"
        if body.movement_type in ("in", "return"):
            conn.execute("UPDATE inventory_products SET quantity=quantity+?, updated_at=? WHERE id=?", (body.quantity, now, body.product_id))
        elif body.movement_type == "out":
            if (prow["quantity"] or 0) < body.quantity:
                raise HTTPException(400, "الكمية المتاحة غير كافية")
            conn.execute("UPDATE inventory_products SET quantity=quantity-?, updated_at=? WHERE id=?", (body.quantity, now, body.product_id))
        else:  # adjust → تعيين الكمية مباشرة
            conn.execute("UPDATE inventory_products SET quantity=?, updated_at=? WHERE id=?", (body.quantity, now, body.product_id))
        cur = conn.execute("""INSERT INTO inventory_movements(product_id,movement_type,quantity,ref_type,username,notes,created_at)
                              VALUES(?,?,?,'manual',?,?,?)""",
                            (body.product_id, body.movement_type, body.quantity, cu.get("username",""), body.notes or "", now))
        return {"id": cur.lastrowid}


# ══════════════════════════════════════════════════════
# منصات الإمداد — مخازن متحركة (تانك سولار/عربة تحمل وقود أو قطع غيار)
# ══════════════════════════════════════════════════════
@app.get("/supply-platforms")
async def list_supply_platforms(cu: dict = Depends(get_user)):
    with get_db() as conn:
        rows = [dict(r) for r in conn.execute("SELECT * FROM supply_platforms ORDER BY code").fetchall()]
        for r in rows:
            stock = conn.execute("""SELECT p.id product_id, p.name, p.unit, s.quantity
                                     FROM supply_platform_stock s JOIN inventory_products p ON p.id=s.product_id
                                     WHERE s.platform_id=? ORDER BY p.name""", (r["id"],)).fetchall()
            r["stock"] = [dict(x) for x in stock]
        return rows

@app.post("/supply-platforms")
async def create_supply_platform(body: SupplyPlatformCreate, cu: dict = Depends(require_superuser)):
    code = body.code.strip()
    if not code:
        raise HTTPException(400, "كود المنصة مطلوب")
    with get_db() as conn:
        if conn.execute("SELECT id FROM supply_platforms WHERE code=?", (code,)).fetchone():
            raise HTTPException(400, "الكود مستخدم بالفعل")
        now = datetime.utcnow().isoformat() + "Z"
        cur = conn.execute("""INSERT INTO supply_platforms(code,plate,driver_name,manager_name,license_number,status,notes,created_at,updated_at)
                               VALUES(?,?,?,?,?,?,?,?,?)""",
                            (code, body.plate or "", body.driver_name or "", body.manager_name or "",
                             body.license_number or "", body.status or "active", body.notes or "", now, now))
        return {"id": cur.lastrowid}

@app.put("/supply-platforms/{pid}")
async def update_supply_platform(pid: int, body: SupplyPlatformCreate, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        if not conn.execute("SELECT id FROM supply_platforms WHERE id=?", (pid,)).fetchone():
            raise HTTPException(404, "المنصة غير موجودة")
        dup = conn.execute("SELECT id FROM supply_platforms WHERE code=? AND id!=?", (body.code.strip(), pid)).fetchone()
        if dup:
            raise HTTPException(400, "الكود مستخدم بالفعل")
        now = datetime.utcnow().isoformat() + "Z"
        conn.execute("""UPDATE supply_platforms SET code=?,plate=?,driver_name=?,manager_name=?,
                        license_number=?,status=?,notes=?,updated_at=? WHERE id=?""",
                     (body.code.strip(), body.plate or "", body.driver_name or "", body.manager_name or "",
                      body.license_number or "", body.status or "active", body.notes or "", now, pid))
        return {"ok": True}

@app.delete("/supply-platforms/{pid}")
async def delete_supply_platform(pid: int, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        conn.execute("DELETE FROM supply_platform_stock WHERE platform_id=?", (pid,))
        conn.execute("DELETE FROM supply_platform_movements WHERE platform_id=?", (pid,))
        conn.execute("DELETE FROM supply_platforms WHERE id=?", (pid,))
        return {"ok": True}

@app.get("/supply-platforms/{pid}/movements")
async def list_supply_platform_movements(pid: int, cu: dict = Depends(require_admin_or_reporter), limit: int = 500):
    with get_db() as conn:
        rows = conn.execute("""SELECT m.*, p.name product_name, p.unit product_unit
                                FROM supply_platform_movements m JOIN inventory_products p ON p.id=m.product_id
                                WHERE m.platform_id=? ORDER BY m.id DESC LIMIT ?""", (pid, limit)).fetchall()
        return [dict(r) for r in rows]

@app.post("/supply-platforms/{pid}/fill")
async def fill_supply_platform(pid: int, body: SupplyPlatformFill, cu: dict = Depends(require_admin)):
    """تعبئة تانك المنصة — يسحب الكمية من المخزن الرئيسي ويضيفها لرصيد المنصة."""
    if body.quantity <= 0:
        raise HTTPException(400, "الكمية يجب أن تكون أكبر من صفر")
    with get_db() as conn:
        plat = conn.execute("SELECT id FROM supply_platforms WHERE id=?", (pid,)).fetchone()
        if not plat:
            raise HTTPException(404, "المنصة غير موجودة")
        prow = conn.execute("SELECT quantity, name FROM inventory_products WHERE id=?", (body.product_id,)).fetchone()
        if not prow:
            raise HTTPException(404, "المنتج غير موجود في المخزون الرئيسي")
        if (prow["quantity"] or 0) < body.quantity:
            raise HTTPException(400, f"الكمية المتاحة من «{prow['name']}» بالمخزن الرئيسي غير كافية (متاح {prow['quantity']})")
        now = datetime.utcnow().isoformat() + "Z"
        # سحب من المخزن الرئيسي
        conn.execute("UPDATE inventory_products SET quantity=quantity-?, updated_at=? WHERE id=?",
                     (body.quantity, now, body.product_id))
        conn.execute("""INSERT INTO inventory_movements(product_id,movement_type,quantity,ref_type,ref_id,username,notes,created_at)
                        VALUES(?,'out',?,'platform_fill',?,?,?,?)""",
                     (body.product_id, body.quantity, pid, cu.get("username",""), body.notes or "", now))
        # إضافة لرصيد المنصة
        existing = conn.execute("SELECT id, quantity FROM supply_platform_stock WHERE platform_id=? AND product_id=?",
                                (pid, body.product_id)).fetchone()
        if existing:
            conn.execute("UPDATE supply_platform_stock SET quantity=quantity+?, updated_at=? WHERE id=?",
                         (body.quantity, now, existing["id"]))
        else:
            conn.execute("INSERT INTO supply_platform_stock(platform_id,product_id,quantity,updated_at) VALUES(?,?,?,?)",
                         (pid, body.product_id, body.quantity, now))
        conn.execute("""INSERT INTO supply_platform_movements(platform_id,product_id,movement_type,quantity,ref_type,username,notes,created_at)
                        VALUES(?,?,'fill',?,'manual',?,?,?)""",
                     (pid, body.product_id, body.quantity, cu.get("username",""), body.notes or "", now))
        return {"ok": True}

@app.post("/supply-platforms/{pid}/adjust")
async def adjust_supply_platform_stock(pid: int, body: SupplyPlatformAdjust, cu: dict = Depends(require_superuser)):
    """تعديل يدوي لرصيد صنف داخل منصة (جرد)."""
    with get_db() as conn:
        if not conn.execute("SELECT id FROM supply_platforms WHERE id=?", (pid,)).fetchone():
            raise HTTPException(404, "المنصة غير موجودة")
        if not conn.execute("SELECT id FROM inventory_products WHERE id=?", (body.product_id,)).fetchone():
            raise HTTPException(404, "المنتج غير موجود")
        now = datetime.utcnow().isoformat() + "Z"
        existing = conn.execute("SELECT id FROM supply_platform_stock WHERE platform_id=? AND product_id=?",
                                (pid, body.product_id)).fetchone()
        if existing:
            conn.execute("UPDATE supply_platform_stock SET quantity=?, updated_at=? WHERE id=?",
                         (body.quantity, now, existing["id"]))
        else:
            conn.execute("INSERT INTO supply_platform_stock(platform_id,product_id,quantity,updated_at) VALUES(?,?,?,?)",
                         (pid, body.product_id, body.quantity, now))
        conn.execute("""INSERT INTO supply_platform_movements(platform_id,product_id,movement_type,quantity,ref_type,username,notes,created_at)
                        VALUES(?,?,'adjust',?,'manual',?,?,?)""",
                     (pid, body.product_id, body.quantity, cu.get("username",""), body.notes or "", now))
        return {"ok": True}


# ── مقايسات الإصلاح (الورشة) ──
@app.get("/workshop/repair-quotes")
async def list_repair_quotes(cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        cars = {r["id"]: r["plate"] for r in conn.execute("SELECT id,plate FROM cars").fetchall()}
        drvs = {r["id"]: r["name"] for r in conn.execute("SELECT id,name FROM drivers").fetchall()}
        equip = {r["car_code"]: r["equipment_name"] for r in conn.execute("SELECT car_code,equipment_name FROM equipment").fetchall()}
        rows = [dict(r) for r in conn.execute("SELECT * FROM workshop_repair_quotes ORDER BY id DESC").fetchall()]
        for r in rows:
            r["car_plate"] = cars.get(r.get("car_id"), "")
            r["driver_name"] = drvs.get(r.get("driver_id"), "")
            r["equipment_name"] = equip.get(r.get("equipment_id"), "") if r.get("equipment_id") else ""
        return rows

@app.post("/workshop/repair-quotes")
async def create_repair_quote(body: RepairQuoteCreate, cu: dict = Depends(require_workshop_admin)):
    with get_db() as conn:
        car_id, equipment_id = _validate_quote_target(conn, body)
        now = datetime.utcnow().isoformat() + "Z"
        quote_number = body.quote_number or f"RQ-{int(datetime.utcnow().timestamp())}"
        driver_id_val = body.driver_id if (body.driver_id and body.driver_id > 0) else None
        cur = conn.execute("""INSERT INTO workshop_repair_quotes
                              (quote_number,car_id,equipment_id,driver_id,quote_date,status,items_json,total_value,notes,created_at)
                              VALUES(?,?,?,?,?,?,?,?,?,?)""",
                            (quote_number, car_id, equipment_id, driver_id_val, body.quote_date or now[:10], body.status or "draft",
                             body.items_json or "[]", body.total_value or 0, body.notes or "", now))
        new_qid = cur.lastrowid
        # ── إنشاء محضر فحص تلقائياً بنفس رقم ووصف المقايسة ──
        conn.execute("""INSERT INTO workshop_inspection_reports
                        (report_number,car_id,equipment_id,driver_id,quote_id,inspection_date,result,notes,created_by,created_at)
                        VALUES(?,?,?,?,?,?,?,?,?,?)""",
                      (quote_number, car_id, equipment_id, driver_id_val, new_qid, body.quote_date or now[:10],
                       "pending", f"تم إنشاؤه تلقائياً من مقايسة الإصلاح رقم {quote_number}",
                       cu.get("username", ""), now))
        return {"id": new_qid, "quote_number": quote_number}

@app.put("/workshop/repair-quotes/{qid}")
async def update_repair_quote(qid: int, body: RepairQuoteCreate, cu: dict = Depends(require_workshop_admin)):
    if body.status not in ("draft", "approved", "done", "cancelled"):
        raise HTTPException(400, "حالة غير صالحة")
    with get_db() as conn:
        car_id, equipment_id = _validate_quote_target(conn, body)
        driver_id_val = body.driver_id if (body.driver_id and body.driver_id > 0) else None
        conn.execute("""UPDATE workshop_repair_quotes SET quote_number=?,car_id=?,equipment_id=?,driver_id=?,quote_date=?,
                       status=?,items_json=?,total_value=?,notes=? WHERE id=?""",
                      (body.quote_number or "", car_id, equipment_id, driver_id_val, body.quote_date or "", body.status,
                       body.items_json or "[]", body.total_value or 0, body.notes or "", qid))
        # ── مزامنة محضر الفحص المرتبط تلقائياً (نفس الرقم والمركبة/المعدة والسائق والتاريخ) ──
        conn.execute("""UPDATE workshop_inspection_reports
                        SET report_number=?, car_id=?, equipment_id=?, driver_id=?, inspection_date=?
                        WHERE quote_id=?""",
                      (body.quote_number or "", car_id, equipment_id, driver_id_val, body.quote_date or "", qid))
        return {"ok": True}

@app.delete("/workshop/repair-quotes/{qid}")
async def delete_repair_quote(qid: int, cu: dict = Depends(require_workshop_admin)):
    with get_db() as conn:
        conn.execute("DELETE FROM workshop_inspection_reports WHERE quote_id=?", (qid,))
        conn.execute("DELETE FROM workshop_repair_quotes WHERE id=?", (qid,))
        return {"ok": True}


# ── محاضر الفحص (الورشة) ──
class InspectionReportCreate(BaseModel):
    report_number: Optional[str] = ""
    car_id: Optional[int] = None
    equipment_id: Optional[str] = ""   # كود المعدة (car_code) — بديل عن car_id
    driver_id: Optional[int] = None
    quote_id: Optional[int] = None
    inspection_date: Optional[str] = ""
    inspector_name: Optional[str] = ""
    odometer_reading: Optional[float] = None
    result: Optional[str] = "pending"
    findings: Optional[str] = ""
    items_json: Optional[str] = "[]"
    notes: Optional[str] = ""

INSPECTION_RESULTS = ("pass", "fail", "needs_repair", "pending")


def _validate_inspection_decisions(items_json: str):
    """أي بند اتحدد 'مرفوض' لازم يكون له سبب رفض مكتوب."""
    try:
        decisions = json.loads(items_json or "{}")
    except Exception:
        return
    if not isinstance(decisions, dict):
        return
    for key, val in decisions.items():
        if isinstance(val, dict) and val.get("status") == "rejected":
            if not (val.get("reason") or "").strip():
                raise HTTPException(400, f"لازم تدخل سبب الرفض للبند المرفوض ({key})")
        elif val == "rejected":
            # صيغة قديمة (نص فقط) بدون سبب — نرفضها لضمان وجود السبب دايماً
            raise HTTPException(400, f"لازم تدخل سبب الرفض للبند المرفوض ({key})")


def _validate_inspection_target(conn, body: "InspectionReportCreate"):
    """لازم يتحدد إما مركبة (car_id) أو معدة (equipment_id) — واحد منهم إجباري."""
    car_id = body.car_id if (body.car_id and body.car_id > 0) else None
    equipment_id = (body.equipment_id or "").strip()
    if not car_id and not equipment_id:
        raise HTTPException(400, "لازم تدخل رقم المعدة أو نمرة العربية")
    if car_id and not conn.execute("SELECT id FROM cars WHERE id=?", (car_id,)).fetchone():
        raise HTTPException(404, "المركبة غير موجودة")
    if equipment_id and not conn.execute("SELECT id FROM equipment WHERE car_code=?", (equipment_id,)).fetchone():
        raise HTTPException(404, "رقم المعدة غير موجود")
    return car_id, equipment_id

@app.get("/workshop/inspection-reports")
async def list_inspection_reports(cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        cars = {r["id"]: r["plate"] for r in conn.execute("SELECT id,plate FROM cars").fetchall()}
        drvs = {r["id"]: r["name"] for r in conn.execute("SELECT id,name FROM drivers").fetchall()}
        equip = {r["car_code"]: r["equipment_name"] for r in conn.execute("SELECT car_code,equipment_name FROM equipment").fetchall()}
        quotes = {r["id"]: dict(r) for r in conn.execute(
            "SELECT id,quote_number,items_json,total_value FROM workshop_repair_quotes").fetchall()}
        rows = [dict(r) for r in conn.execute("SELECT * FROM workshop_inspection_reports ORDER BY id DESC").fetchall()]
        for r in rows:
            r["car_plate"] = cars.get(r.get("car_id"), "")
            r["driver_name"] = drvs.get(r.get("driver_id"), "")
            r["equipment_name"] = equip.get(r.get("equipment_id"), "") if r.get("equipment_id") else ""
            q = quotes.get(r.get("quote_id"))
            if q:
                r["quote_number"] = q.get("quote_number", "")
                r["quote_items_json"] = q.get("items_json", "[]")
                r["quote_total_value"] = q.get("total_value", 0)
        return rows

@app.post("/workshop/inspection-reports", status_code=201)
async def create_inspection_report(body: InspectionReportCreate, cu: dict = Depends(require_workshop_admin)):
    if body.result not in INSPECTION_RESULTS:
        raise HTTPException(400, "نتيجة فحص غير صالحة")
    _validate_inspection_decisions(body.items_json)
    with get_db() as conn:
        car_id, equipment_id = _validate_inspection_target(conn, body)
        now = datetime.utcnow().isoformat() + "Z"
        report_number = body.report_number or f"IR-{int(datetime.utcnow().timestamp())}"
        driver_id_val = body.driver_id if (body.driver_id and body.driver_id > 0) else None
        cur = conn.execute("""INSERT INTO workshop_inspection_reports
                              (report_number,car_id,equipment_id,driver_id,quote_id,inspection_date,inspector_name,
                               odometer_reading,result,findings,items_json,notes,created_by,created_at)
                              VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (report_number, car_id, equipment_id, driver_id_val, body.quote_id, body.inspection_date or now[:10],
                             body.inspector_name or "", body.odometer_reading, body.result or "pending",
                             body.findings or "", body.items_json or "[]", body.notes or "",
                             cu.get("username", ""), now))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "create_inspection_report", f"إنشاء محضر فحص جديد: {report_number}")
        return {"id": cur.lastrowid, "report_number": report_number}

@app.put("/workshop/inspection-reports/{rid}")
async def update_inspection_report(rid: int, body: InspectionReportCreate, cu: dict = Depends(require_workshop_admin)):
    if body.result not in INSPECTION_RESULTS:
        raise HTTPException(400, "نتيجة فحص غير صالحة")
    _validate_inspection_decisions(body.items_json)
    with get_db() as conn:
        if not conn.execute("SELECT id FROM workshop_inspection_reports WHERE id=?", (rid,)).fetchone():
            raise HTTPException(404, "المحضر غير موجود")
        car_id, equipment_id = _validate_inspection_target(conn, body)
        driver_id_val = body.driver_id if (body.driver_id and body.driver_id > 0) else None
        conn.execute("""UPDATE workshop_inspection_reports SET report_number=?,car_id=?,equipment_id=?,driver_id=?,quote_id=COALESCE(?,quote_id),
                       inspection_date=?,inspector_name=?,odometer_reading=?,result=?,findings=?,
                       items_json=?,notes=? WHERE id=?""",
                      (body.report_number or "", car_id, equipment_id, driver_id_val, body.quote_id, body.inspection_date or "",
                       body.inspector_name or "", body.odometer_reading, body.result,
                       body.findings or "", body.items_json or "[]", body.notes or "", rid))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "update_inspection_report", f"تعديل محضر فحص #{rid}")
        return {"ok": True}

@app.delete("/workshop/inspection-reports/{rid}")
async def delete_inspection_report(rid: int, cu: dict = Depends(require_workshop_admin)):
    with get_db() as conn:
        conn.execute("DELETE FROM workshop_inspection_reports WHERE id=?", (rid,))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "delete_inspection_report", f"حذف محضر فحص #{rid}")
        return {"ok": True}


# ── طلبات التشغيل الخارجي (الورشة) ──
class ExternalOpCreate(BaseModel):
    request_number: Optional[str] = ""
    quote_id: Optional[int] = None
    car_id: Optional[int] = None
    equipment_id: Optional[str] = ""   # كود المعدة (car_code) — بديل عن car_id
    driver_id: Optional[int] = None
    workshop_name: Optional[str] = ""
    reason: Optional[str] = ""
    estimated_cost: Optional[float] = 0
    workmanship_cost: Optional[float] = 0
    request_date: Optional[str] = ""
    status: Optional[str] = "pending"
    notes: Optional[str] = ""

EXTERNAL_OP_STATUSES = ("pending", "approved", "rejected", "done")


def _validate_extop_target(conn, body: "ExternalOpCreate"):
    """لازم يتحدد إما مركبة (car_id) أو معدة (equipment_id) — واحد منهم إجباري."""
    car_id = body.car_id if (body.car_id and body.car_id > 0) else None
    equipment_id = (body.equipment_id or "").strip()
    if not car_id and not equipment_id:
        raise HTTPException(400, "لازم تدخل رقم المعدة أو نمرة العربية")
    if car_id and not conn.execute("SELECT id FROM cars WHERE id=?", (car_id,)).fetchone():
        raise HTTPException(404, "المركبة غير موجودة")
    if equipment_id and not conn.execute("SELECT id FROM equipment WHERE car_code=?", (equipment_id,)).fetchone():
        raise HTTPException(404, "رقم المعدة غير موجود")
    return car_id, equipment_id

@app.get("/workshop/external-ops")
async def list_external_ops(cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        cars = {r["id"]: r["plate"] for r in conn.execute("SELECT id,plate FROM cars").fetchall()}
        drvs = {r["id"]: r["name"] for r in conn.execute("SELECT id,name FROM drivers").fetchall()}
        equip = {r["car_code"]: r["equipment_name"] for r in conn.execute("SELECT car_code,equipment_name FROM equipment").fetchall()}
        quotes = {r["id"]: r["quote_number"] for r in conn.execute("SELECT id,quote_number FROM workshop_repair_quotes").fetchall()}
        rows = [dict(r) for r in conn.execute("SELECT * FROM workshop_external_ops ORDER BY id DESC").fetchall()]
        for r in rows:
            r["car_plate"] = cars.get(r.get("car_id"), "")
            r["driver_name"] = drvs.get(r.get("driver_id"), "")
            r["equipment_name"] = equip.get(r.get("equipment_id"), "") if r.get("equipment_id") else ""
            r["quote_number"] = quotes.get(r.get("quote_id"), "")
        return rows

@app.post("/workshop/external-ops", status_code=201)
async def create_external_op(body: ExternalOpCreate, cu: dict = Depends(require_workshop_admin)):
    if body.status not in EXTERNAL_OP_STATUSES:
        raise HTTPException(400, "حالة غير صالحة")
    with get_db() as conn:
        car_id, equipment_id = _validate_extop_target(conn, body)
        now = datetime.utcnow().isoformat() + "Z"
        request_number = body.request_number or f"EXT-{int(datetime.utcnow().timestamp())}"
        driver_id_val = body.driver_id if (body.driver_id and body.driver_id > 0) else None
        cur = conn.execute("""INSERT INTO workshop_external_ops
                              (request_number,quote_id,car_id,equipment_id,driver_id,workshop_name,reason,
                               estimated_cost,workmanship_cost,request_date,status,notes,created_by,created_at)
                              VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (request_number, body.quote_id, car_id, equipment_id, driver_id_val, body.workshop_name or "",
                             body.reason or "", body.estimated_cost or 0, body.workmanship_cost or 0,
                             body.request_date or now[:10], body.status or "pending", body.notes or "",
                             cu.get("username", ""), now))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "create_external_op", f"إنشاء طلب تشغيل خارجي: {request_number}")
        return {"id": cur.lastrowid, "request_number": request_number}

@app.put("/workshop/external-ops/{eid}")
async def update_external_op(eid: int, body: ExternalOpCreate, cu: dict = Depends(require_workshop_admin)):
    if body.status not in EXTERNAL_OP_STATUSES:
        raise HTTPException(400, "حالة غير صالحة")
    with get_db() as conn:
        if not conn.execute("SELECT id FROM workshop_external_ops WHERE id=?", (eid,)).fetchone():
            raise HTTPException(404, "الطلب غير موجود")
        car_id, equipment_id = _validate_extop_target(conn, body)
        driver_id_val = body.driver_id if (body.driver_id and body.driver_id > 0) else None
        conn.execute("""UPDATE workshop_external_ops SET request_number=?,quote_id=COALESCE(?,quote_id),
                       car_id=?,equipment_id=?,driver_id=?,workshop_name=?,reason=?,estimated_cost=?,workmanship_cost=?,
                       request_date=?,status=?,notes=? WHERE id=?""",
                      (body.request_number or "", body.quote_id, car_id, equipment_id, driver_id_val, body.workshop_name or "",
                       body.reason or "", body.estimated_cost or 0, body.workmanship_cost or 0,
                       body.request_date or "", body.status, body.notes or "", eid))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "update_external_op", f"تعديل طلب تشغيل خارجي #{eid}")
        return {"ok": True}

@app.delete("/workshop/external-ops/{eid}")
async def delete_external_op(eid: int, cu: dict = Depends(require_workshop_admin)):
    with get_db() as conn:
        conn.execute("DELETE FROM workshop_external_ops WHERE id=?", (eid,))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "delete_external_op", f"حذف طلب تشغيل خارجي #{eid}")
        return {"ok": True}


# ── إذن الصرف (الورشة) ──
class VoucherCreate(BaseModel):
    voucher_number: Optional[str] = ""
    quote_id: Optional[int] = None
    car_id: Optional[int] = None
    equipment_id: Optional[str] = ""   # كود المعدة (car_code) — بديل عن car_id لو الصنف معدة مش مركبة
    driver_id: Optional[int] = None
    recipient_name: Optional[str] = ""
    purpose: Optional[str] = ""
    items_json: Optional[str] = "[]"
    total_amount: Optional[float] = 0
    issue_date: Optional[str] = ""
    issued_by: Optional[str] = ""
    status: Optional[str] = "draft"
    notes: Optional[str] = ""

VOUCHER_STATUSES = ("draft", "issued", "cancelled")


def _validate_voucher_target(conn, body: "VoucherCreate"):
    """لازم يتحدد إما مركبة (car_id) أو معدة (equipment_id) — واحد منهم إجباري."""
    car_id = body.car_id if (body.car_id and body.car_id > 0) else None
    equipment_id = (body.equipment_id or "").strip()
    if not car_id and not equipment_id:
        raise HTTPException(400, "لازم تدخل رقم المعدة أو نمرة العربية")
    if car_id and not conn.execute("SELECT id FROM cars WHERE id=?", (car_id,)).fetchone():
        raise HTTPException(404, "المركبة غير موجودة")
    if equipment_id and not conn.execute("SELECT id FROM equipment WHERE car_code=?", (equipment_id,)).fetchone():
        raise HTTPException(404, "رقم المعدة غير موجود")
    return car_id, equipment_id
    # لازم تدخل كمية الصرف لكل بند — يتم التحقق منها في items_json من الفرونت إند

@app.get("/workshop/repair-quotes/external-items")
async def list_external_quote_items(cu: dict = Depends(require_admin_or_reporter)):
    """كل البنود (من كل المقايسات) اللي نوع تشغيلها 'خارجي' — تُستخدم فى صفحة طلبات التشغيل الخارجي"""
    with get_db() as conn:
        cars = {r["id"]: r["plate"] for r in conn.execute("SELECT id,plate FROM cars").fetchall()}
        quotes = [dict(r) for r in conn.execute(
            "SELECT id,quote_number,car_id,items_json,quote_date FROM workshop_repair_quotes ORDER BY id DESC"
        ).fetchall()]
    result = []
    for q in quotes:
        try:
            parsed = json.loads(q.get("items_json") or "[]")
        except Exception:
            continue
        sections = parsed.get("sections", []) if isinstance(parsed, dict) else (parsed if isinstance(parsed, list) else [])
        for si, sec in enumerate(sections):
            for ii, it in enumerate(sec.get("items", []) if isinstance(sec, dict) else []):
                if isinstance(it, dict) and it.get("item_type") == "external":
                    result.append({
                        "quote_id":       q["id"],
                        "quote_number":   q.get("quote_number", ""),
                        "quote_date":     q.get("quote_date", ""),
                        "car_plate":      cars.get(q.get("car_id"), ""),
                        "section_title":  sec.get("title", "") if isinstance(sec, dict) else "",
                        "item_name":      it.get("name", ""),
                        "category":       it.get("category", ""),
                        "qty":            it.get("qty", 0),
                        "unit":           it.get("unit", ""),
                        "price":          it.get("price", 0),
                        "mfg":            it.get("mfg", 0),
                        "parts":          it.get("parts", 0),
                    })
    return {"external_items": result}


@app.get("/workshop/vouchers")
async def list_vouchers(cu: dict = Depends(require_admin_or_reporter)):
    with get_db() as conn:
        cars = {r["id"]: r["plate"] for r in conn.execute("SELECT id,plate FROM cars").fetchall()}
        drvs = {r["id"]: r["name"] for r in conn.execute("SELECT id,name FROM drivers").fetchall()}
        equip = {r["car_code"]: r["equipment_name"] for r in conn.execute("SELECT car_code,equipment_name FROM equipment").fetchall()}
        quotes = {r["id"]: r["quote_number"] for r in conn.execute("SELECT id,quote_number FROM workshop_repair_quotes").fetchall()}
        rows = [dict(r) for r in conn.execute("SELECT * FROM workshop_vouchers ORDER BY id DESC").fetchall()]
        for r in rows:
            r["car_plate"] = cars.get(r.get("car_id"), "")
            r["driver_name"] = drvs.get(r.get("driver_id"), "")
            r["equipment_name"] = equip.get(r.get("equipment_id"), "") if r.get("equipment_id") else ""
            r["quote_number"] = quotes.get(r.get("quote_id"), "")
        return rows

@app.post("/workshop/vouchers", status_code=201)
async def create_voucher(body: VoucherCreate, cu: dict = Depends(require_workshop_admin)):
    if body.status not in VOUCHER_STATUSES:
        raise HTTPException(400, "حالة غير صالحة")
    with get_db() as conn:
        car_id, equipment_id = _validate_voucher_target(conn, body)
        now = datetime.utcnow().isoformat() + "Z"
        voucher_number = body.voucher_number or f"VCH-{int(datetime.utcnow().timestamp())}"
        driver_id_val = body.driver_id if (body.driver_id and body.driver_id > 0) else None
        cur = conn.execute("""INSERT INTO workshop_vouchers
                              (voucher_number,quote_id,car_id,equipment_id,driver_id,recipient_name,purpose,
                               items_json,total_amount,issue_date,issued_by,status,notes,created_by,created_at)
                              VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (voucher_number, body.quote_id, car_id, equipment_id, driver_id_val, body.recipient_name or "",
                             body.purpose or "", body.items_json or "[]", body.total_amount or 0,
                             body.issue_date or now[:10], body.issued_by or cu.get("username", ""),
                             body.status or "draft", body.notes or "", cu.get("username", ""), now))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "create_voucher", f"إنشاء إذن صرف: {voucher_number}")
        return {"id": cur.lastrowid, "voucher_number": voucher_number}

@app.put("/workshop/vouchers/{vid}")
async def update_voucher(vid: int, body: VoucherCreate, cu: dict = Depends(require_workshop_admin)):
    if body.status not in VOUCHER_STATUSES:
        raise HTTPException(400, "حالة غير صالحة")
    with get_db() as conn:
        if not conn.execute("SELECT id FROM workshop_vouchers WHERE id=?", (vid,)).fetchone():
            raise HTTPException(404, "الإذن غير موجود")
        car_id, equipment_id = _validate_voucher_target(conn, body)
        driver_id_val = body.driver_id if (body.driver_id and body.driver_id > 0) else None
        conn.execute("""UPDATE workshop_vouchers SET voucher_number=?,quote_id=COALESCE(?,quote_id),
                       car_id=?,equipment_id=?,driver_id=?,recipient_name=?,purpose=?,items_json=?,total_amount=?,
                       issue_date=?,issued_by=?,status=?,notes=? WHERE id=?""",
                      (body.voucher_number or "", body.quote_id, car_id, equipment_id, driver_id_val, body.recipient_name or "",
                       body.purpose or "", body.items_json or "[]", body.total_amount or 0, body.issue_date or "",
                       body.issued_by or "", body.status, body.notes or "", vid))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "update_voucher", f"تعديل إذن صرف #{vid}")
        return {"ok": True}

@app.delete("/workshop/vouchers/{vid}")
async def delete_voucher(vid: int, cu: dict = Depends(require_workshop_admin)):
    with get_db() as conn:
        conn.execute("DELETE FROM workshop_vouchers WHERE id=?", (vid,))
        write_audit_log(cu["user_id"], cu["username"], cu["role"],
                        "delete_voucher", f"حذف إذن صرف #{vid}")
        return {"ok": True}

# ── تصحيح تلقائي عند الـ startup لأي سجل ورش لم يتم تطبيق override عليه ──
FULL_MAINT_MAP_GLOBAL = {
    "oil":        ["تغيير زيت", "فلتر زيت"],
    "oil_motor":  ["تغيير زيت"],
    "oil_gear":   ["زيت تروس"],
    "oil_brake":  ["زيت فرامل"],
    "oil_hydro":  ["زيت هيدروليك"],
    "oil_grease": ["شحم", "تشحيم"],
    "filter":         ["فلتر هواء", "فلتر وقود", "فلتر مكيف"],
    "filter_oil":     ["فلتر زيت"],
    "filter_solar":   ["فلتر وقود", "فلتر سولار"],
    "filter_petrol":  ["فلتر بنزين", "فلتر وقود"],
    "filter_air":     ["فلتر هواء"],
    "filter_separator":["فلتر فاصل"],
    "filter_ac":      ["فلتر مكيف", "فلتر تكيف"],
    "filter_dryer":   ["فلتر مجفف"],
    "filter_breather":["فلتر منفس"],
    "filter_hydraulic":["فلتر هيدروليك"],
    "tire":    ["إطارات"],
    "battery": ["بطارية"],
    "belt":    ["سير توزيع", "سيور"],
    "timing_belt":           ["سير كاتينة", "سير توزيع"],
    "engine_overhaul":       ["عمرة محرك", "عمرة كاملة"],
    "engine_half_overhaul":  ["نصف عمرة محرك", "عمرة نصفية"],
    "head_gasket":           ["جوان وش سلندر", "سلندر"],
    "fuel_pump_overhaul":    ["عمرة طرمبة وقود", "طرمبة وقود", "وحدات الحقن"],
}

def _apply_override_for_record(c, rec_dict, now_str):
    """يطبق override على سجل ورش واحد — يُستخدم من الـ startup والـ endpoint"""
    ws_type = rec_dict.get("type", "")
    car_id  = rec_dict.get("vehicle_id")
    km      = rec_dict.get("odometer_reading")
    if not car_id:
        return 0
    # لو مفيش عداد، نجيب آخر عداد من الرحلات
    if not km:
        km_row = c.execute(
            "SELECT MAX(end_odometer) FROM trips WHERE car_id=? AND end_odometer IS NOT NULL",
            (car_id,)
        ).fetchone()
        if km_row and km_row[0]:
            km = km_row[0]
    if not km:
        return 0
    related = FULL_MAINT_MAP_GLOBAL.get(ws_type, [])
    if not related:
        return 0
    placeholders = ",".join("?" * len(related))
    c.execute(
        f"""UPDATE maintenance_schedule
              SET last_done_km=?, last_done_date=?, updated_at=?
              WHERE car_id=? AND maintenance_type IN ({placeholders})
              AND (last_done_km IS NULL OR last_done_km < ?)""",
        [km, now_str[:10], now_str, car_id] + related + [km]
    )
    return c.execute("SELECT changes()").fetchone()[0]

@app.on_event("startup")
async def ensure_forecast_tables():
    """Migration: تأكد من وجود جداول الـ Forecast حتى لو الـ DB قديمة"""
    try:
        with get_db() as conn:
            c = conn.cursor()
            c.execute("""CREATE TABLE IF NOT EXISTS forecast_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                car_id INTEGER NOT NULL,
                forecast_month TEXT NOT NULL,
                generated_at TEXT NOT NULL,
                km_predicted REAL DEFAULT 0,
                km_trend REAL DEFAULT 0,
                km_confidence REAL DEFAULT 0,
                fuel_liters REAL DEFAULT 0,
                fuel_cost REAL DEFAULT 0,
                fuel_l_per_100 REAL DEFAULT 0,
                maint_cost REAL DEFAULT 0,
                maint_visits REAL DEFAULT 0,
                maint_cost_per_km REAL DEFAULT 0,
                budget_total REAL DEFAULT 0,
                confidence REAL DEFAULT 0,
                data_points INTEGER DEFAULT 0,
                algorithm TEXT DEFAULT 'WMA_v1',
                FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE CASCADE
            )""")
            try: c.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_fc_car_month ON forecast_results(car_id, forecast_month)")
            except Exception: pass
            c.execute("""CREATE TABLE IF NOT EXISTS forecast_parts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                car_id INTEGER NOT NULL,
                forecast_month TEXT NOT NULL,
                part_type TEXT NOT NULL,
                part_label TEXT DEFAULT '',
                prob REAL DEFAULT 0,
                est_cost REAL DEFAULT 0,
                due_km REAL,
                due_date TEXT,
                avg_life_km REAL,
                avg_life_days INTEGER,
                change_count INTEGER DEFAULT 0,
                FOREIGN KEY(car_id) REFERENCES cars(id) ON DELETE CASCADE
            )""")
            c.execute("""CREATE TABLE IF NOT EXISTS forecast_settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                note TEXT DEFAULT ''
            )""")
            for _k, _v, _n in [
                ('fuel_price_per_liter','20','سعر اللتر'),
                ('wma_weights','1,2,3,4','أوزان WMA'),
                ('confidence_data_w','0.4',''),
                ('confidence_reg_w','0.3',''),
                ('confidence_stab_w','0.3',''),
                ('forecast_horizon_months','3',''),
            ]:
                c.execute("INSERT OR IGNORE INTO forecast_settings VALUES(?,?,?)", (_k,_v,_n))
            log.info("[STARTUP] Forecast tables ensured")
    except Exception as e:
        log.error(f"[STARTUP] ensure_forecast_tables failed: {e}")

@app.on_event("startup")
async def fix_missing_overrides():
    """عند الـ startup: يمسح كل سجلات الورش اللي عندها مركبة وعداد ويطبق override لو ما اتطبقش"""
    await asyncio.sleep(3)  # نستنى شوية حتى الـ DB يكون جاهز
    try:
        now = datetime.utcnow().isoformat() + "Z"
        with get_db() as conn:
            c = conn.cursor()
            rows = c.execute(
                """SELECT id, type, vehicle_id, odometer_reading FROM workshop_records
                   WHERE vehicle_id IS NOT NULL
                   AND type IN ({})
                   ORDER BY id ASC""".format(
                    ",".join(f"'{t}'" for t in FULL_MAINT_MAP_GLOBAL.keys())
                )
            ).fetchall()
            total = 0
            for row in rows:
                total += _apply_override_for_record(c, dict(row), now)
            if total:
                log.info(f"[STARTUP] Applied maintenance override for {total} schedule rows from {len(rows)} workshop records")
    except Exception as e:
        log.error(f"[STARTUP] fix_missing_overrides failed: {e}")

@app.post("/workshop/records/{rid}/apply-maintenance-override")
async def apply_maintenance_override(rid: int, cu: dict = Depends(require_admin)):
    """يطبق override يدوي لسجل ورش على الصيانات الدورية"""
    FULL_MAINT_MAP = {
        "oil":        ["تغيير زيت", "فلتر زيت"],
        "oil_motor":  ["تغيير زيت"],
        "oil_gear":   ["زيت تروس"],
        "oil_brake":  ["زيت فرامل"],
        "oil_hydro":  ["زيت هيدروليك"],
        "oil_grease": ["شحم", "تشحيم"],
        "filter":         ["فلتر هواء", "فلتر وقود", "فلتر مكيف"],
        "filter_oil":     ["فلتر زيت"],
        "filter_solar":   ["فلتر وقود", "فلتر سولار"],
        "filter_petrol":  ["فلتر بنزين", "فلتر وقود"],
        "filter_air":     ["فلتر هواء"],
        "filter_separator":["فلتر فاصل"],
        "filter_ac":      ["فلتر مكيف", "فلتر تكيف"],
        "filter_dryer":   ["فلتر مجفف"],
        "filter_breather":["فلتر منفس"],
        "filter_hydraulic":["فلتر هيدروليك"],
        "tire":    ["إطارات"],
        "battery": ["بطارية"],
        "belt":    ["سير توزيع", "سيور"],
        "timing_belt":           ["سير كاتينة", "سير توزيع"],
        "engine_overhaul":       ["عمرة محرك", "عمرة كاملة"],
        "engine_half_overhaul":  ["نصف عمرة محرك", "عمرة نصفية"],
        "head_gasket":           ["جوان وش سلندر", "سلندر"],
        "fuel_pump_overhaul":    ["عمرة طرمبة وقود", "طرمبة وقود", "وحدات الحقن"],
    }
    with get_db() as conn:
        c = conn.cursor()
        row = c.execute("SELECT * FROM workshop_records WHERE id=?", (rid,)).fetchone()
        if not row:
            raise HTTPException(404, "سجل غير موجود")
        rec = dict(row)
        ws_type = rec.get("type","")
        car_id  = rec.get("vehicle_id")
        km      = rec.get("odometer_reading")
        now     = datetime.utcnow().isoformat() + "Z"

        if not car_id:
            raise HTTPException(400, "السجل لا يحتوي على مركبة — لا يمكن التطبيق")

        # لو مفيش عداد، نجيب آخر عداد من الرحلات
        if not km:
            km_row = c.execute("SELECT MAX(end_odometer) FROM trips WHERE car_id=? AND end_odometer IS NOT NULL", (car_id,)).fetchone()
            if km_row and km_row[0]: km = km_row[0]

        if not km:
            raise HTTPException(400, "لا توجد قراءة عداد — أضف قراءة العداد للسجل أولاً")

        related = FULL_MAINT_MAP.get(ws_type, [])
        if not related:
            raise HTTPException(400, f"نوع الورش '{ws_type}' لا يرتبط بأي صيانة دورية")

        placeholders = ",".join("?" * len(related))
        c.execute(f"""UPDATE maintenance_schedule
                      SET last_done_km=?, last_done_date=?, updated_at=?
                      WHERE car_id=? AND maintenance_type IN ({placeholders})""",
                  [km, now[:10], now, car_id] + related)
        updated = conn.execute("SELECT changes()").fetchone()[0]
        log_event("maintenance_manual_override", car_id=car_id,
                  ws_record_id=rid, types=related, km=km, by=cu["username"])
        return {"ok": True, "updated_rows": updated, "km": km, "types": related}





# ══════════════════════════════════════════════════════
# 52. FORECAST ENGINE — Planning & Forecast Module
# Algorithm: WMA + Linear Trend + Statistical Rules
# Designed for extensibility (Prophet/XGBoost ready)
# ══════════════════════════════════════════════════════

# ── Part type mapping: workshop type → label + avg life ──
PART_LIFECYCLE = {
    "oil":               {"label": "تغيير زيت",        "avg_km": 5000,  "avg_days": 90},
    "filter_oil":        {"label": "فلتر زيت",          "avg_km": 5000,  "avg_days": 90},
    "filter":            {"label": "فلتر هواء",          "avg_km": 10000, "avg_days": 180},
    "filter_air":        {"label": "فلتر هواء",          "avg_km": 10000, "avg_days": 180},
    "filter_solar":      {"label": "فلتر سولار",         "avg_km": 10000, "avg_days": 180},
    "tire":              {"label": "إطارات",             "avg_km": 40000, "avg_days": 730},
    "battery":           {"label": "بطارية",             "avg_km": 60000, "avg_days": 730},
    "belt":              {"label": "سيور",               "avg_km": 30000, "avg_days": 365},
    "timing_belt":       {"label": "سير كاتينة",         "avg_km": 60000, "avg_days": 730},
    "engine_overhaul":   {"label": "عمرة محرك",          "avg_km": 150000,"avg_days": 1825},
    "engine_half_overhaul":{"label":"نصف عمرة محرك",    "avg_km": 80000, "avg_days": 1095},
    "head_gasket":       {"label": "جوان وش سلندر",      "avg_km": 100000,"avg_days": 1825},
    "fuel_pump_overhaul":{"label": "عمرة طرمبة وقود",    "avg_km": 80000, "avg_days": 1825},
    "filter_ac":         {"label": "فلتر مكيف",          "avg_km": 15000, "avg_days": 365},
    "oil_gear":          {"label": "زيت تروس",           "avg_km": 30000, "avg_days": 365},
    "oil_brake":         {"label": "زيت فرامل",          "avg_km": 20000, "avg_days": 365},
    "oil_hydro":         {"label": "زيت هيدروليك",       "avg_km": 20000, "avg_days": 365},
}

FUEL_TYPES = {"fuel_solar", "fuel_92", "fuel_95", "fuel_80", "fuel_cng"}


def _wma(values: List[float], weights: List[float] = None) -> float:
    """Weighted Moving Average — الأحدث وزنه أعلى"""
    if not values:
        return 0.0
    n = len(values)
    if weights is None:
        weights = list(range(1, n + 1))
    w = weights[-n:] if len(weights) >= n else weights + [weights[-1]] * (n - len(weights))
    total_w = sum(w)
    if total_w == 0:
        return sum(values) / n
    return sum(v * wt for v, wt in zip(values, w)) / total_w


def _linear_trend(values: List[float]) -> Tuple[float, float]:
    """Simple linear regression → (slope, intercept)"""
    n = len(values)
    if n < 2:
        return 0.0, values[0] if values else 0.0
    x_mean = (n - 1) / 2
    y_mean = sum(values) / n
    num = sum((i - x_mean) * (v - y_mean) for i, v in enumerate(values))
    den = sum((i - x_mean) ** 2 for i in range(n))
    slope = num / den if den else 0.0
    intercept = y_mean - slope * x_mean
    return slope, intercept


def _coefficient_of_variation(values: List[float]) -> float:
    """CV = std/mean — يقيس الاستقرار (أقل = أفضل)"""
    if len(values) < 2:
        return 1.0
    mean = sum(values) / len(values)
    if mean == 0:
        return 1.0
    variance = sum((v - mean) ** 2 for v in values) / len(values)
    return math.sqrt(variance) / mean


def _confidence_score(
    data_points: int,
    cv_km: float,
    cv_cost: float,
    has_fuel: bool,
    has_maint: bool,
    completeness: float,
) -> float:
    """
    Confidence Score (0-1) مبني على:
    - كمية البيانات (40%)
    - انتظام الكيلومترات (30%)
    - اكتمال البيانات (30%)
    """
    # كمية البيانات: 3 رحلات = 30%، 10 = 70%، 30+ = 100%
    data_score = min(1.0, data_points / 30)

    # انتظام: CV أقل من 0.2 = ممتاز، فوق 1.0 = ضعيف
    regularity = max(0.0, 1.0 - cv_km)
    cost_reg   = max(0.0, 1.0 - cv_cost * 0.5) if has_maint else 0.5

    # اكتمال: هل عنده وقود وصيانة مسجلة؟
    completeness_score = completeness

    confidence = (
        0.40 * data_score +
        0.30 * regularity +
        0.20 * completeness_score +
        0.10 * ((1 if has_fuel else 0) + (1 if has_maint else 0)) / 2
    )
    return round(min(1.0, max(0.05, confidence)), 3)


def _months_between(d1: str, d2: str) -> int:
    """عدد الأشهر بين تاريخين YYYY-MM-DD"""
    try:
        a = datetime.strptime(d1[:10], "%Y-%m-%d")
        b = datetime.strptime(d2[:10], "%Y-%m-%d")
        return abs((b.year - a.year) * 12 + b.month - a.month)
    except Exception:
        return 0


def compute_forecast_for_car(
    car_id: int,
    car_plate: str,
    trips: List[Dict],
    workshop_records: List[Dict],
    maintenance_schedule: List[Dict],
    settings: Dict[str, str],
    target_month: str,  # YYYY-MM
) -> Dict[str, Any]:
    """
    الـ Forecast Engine الأساسي — قابل للاستبدال بـ Prophet/XGBoost
    يرجع dict كامل بكل التنبؤات
    """
    now = datetime.utcnow()
    target_dt = datetime.strptime(target_month, "%Y-%m")
    fuel_price = float(settings.get("fuel_price_per_liter", "20"))
    wma_w = [float(x) for x in settings.get("wma_weights", "1,2,3,4").split(",")]

    # ──────────────────────────────────────
    # A. تحليل الكيلومترات الشهرية
    # ──────────────────────────────────────
    # نجمّع الكيلومترات لكل شهر
    monthly_km: Dict[str, float] = defaultdict(float)
    for t in trips:
        if not t.get("start_odometer") or not t.get("end_odometer"):
            continue
        km = float(t["end_odometer"]) - float(t["start_odometer"])
        if km <= 0 or km > 5000:  # فلتر القيم الشاذة
            continue
        ts = (t.get("end_time") or t.get("start_time") or "")[:7]  # YYYY-MM
        if ts:
            monthly_km[ts] += km

    sorted_months = sorted(monthly_km.keys())
    km_series = [monthly_km[m] for m in sorted_months]
    data_points = len([t for t in trips if t.get("end_odometer")])

    # WMA للتنبؤ
    km_wma = _wma(km_series, wma_w) if km_series else 0.0

    # Linear Trend
    slope, intercept = _linear_trend(km_series)
    n_ahead = len(km_series)
    km_trend_predicted = intercept + slope * n_ahead

    # الكيلومترات المتوقعة = WMA (70%) + Trend (30%)
    km_predicted = km_wma * 0.7 + max(0, km_trend_predicted) * 0.3 if km_series else 0.0

    # CV للانتظام
    cv_km = _coefficient_of_variation(km_series)

    # ──────────────────────────────────────
    # B. تحليل الوقود
    # ──────────────────────────────────────
    fuel_records = [r for r in workshop_records if r.get("type") in FUEL_TYPES]
    total_fuel_liters = sum(float(r.get("quantity") or 0) for r in fuel_records)
    total_fuel_cost   = sum(float(r.get("price") or 0) for r in fuel_records)
    has_fuel = len(fuel_records) > 0

    # كيلومترات إجمالية من الرحلات
    total_km_all = sum(km_series) if km_series else 1

    # L/100km
    l_per_100 = (total_fuel_liters / total_km_all * 100) if total_km_all > 0 and total_fuel_liters > 0 else 0.0

    # تكلفة الوقود لكل كم
    fuel_cost_per_km = total_fuel_cost / total_km_all if total_km_all > 0 and total_fuel_cost > 0 else 0.0

    # التنبؤ
    pred_fuel_liters = (km_predicted * l_per_100 / 100) if l_per_100 > 0 else 0.0
    pred_fuel_cost   = (fuel_cost_per_km * km_predicted) if fuel_cost_per_km > 0 else (pred_fuel_liters * fuel_price)

    # ──────────────────────────────────────
    # C. تحليل الصيانة
    # ──────────────────────────────────────
    maint_records = [r for r in workshop_records if r.get("type") not in FUEL_TYPES]
    total_maint_cost = sum(float(r.get("price") or 0) for r in maint_records)
    has_maint = len(maint_records) > 0
    cv_cost = _coefficient_of_variation([float(r.get("price") or 0) for r in maint_records]) if maint_records else 1.0

    # تكلفة الصيانة لكل كم
    maint_cost_per_km = total_maint_cost / total_km_all if total_km_all > 0 and total_maint_cost > 0 else 0.0

    # تكلفة الصيانة لكل يوم
    if workshop_records:
        dates = [r["created_at"][:10] for r in maint_records if r.get("created_at")]
        if len(dates) >= 2:
            d_min = min(dates)
            d_max = max(dates)
            span_days = max(1, (datetime.strptime(d_max, "%Y-%m-%d") - datetime.strptime(d_min, "%Y-%m-%d")).days)
            maint_cost_per_day = total_maint_cost / span_days
        else:
            maint_cost_per_day = total_maint_cost / 30
    else:
        maint_cost_per_day = 0.0

    # معدل دخول الورشة الشهري
    monthly_visits: Dict[str, int] = defaultdict(int)
    for r in maint_records:
        ts = (r.get("created_at") or "")[:7]
        if ts:
            monthly_visits[ts] += 1
    avg_visits = sum(monthly_visits.values()) / len(monthly_visits) if monthly_visits else 0.0

    # التنبؤ بتكلفة الصيانة = weighted(per_km + per_day)
    pred_maint_from_km  = maint_cost_per_km * km_predicted
    pred_maint_from_day = maint_cost_per_day * 30
    # لو عنده بيانات كيلومتر كويسة → نعتمد per_km أكتر
    if km_predicted > 0 and maint_cost_per_km > 0:
        pred_maint_cost = pred_maint_from_km * 0.6 + pred_maint_from_day * 0.4
    else:
        pred_maint_cost = pred_maint_from_day

    pred_visits = _wma(list(monthly_visits.values()), wma_w) if monthly_visits else 0.0

    # ──────────────────────────────────────
    # D. تنبؤ قطع الغيار
    # ──────────────────────────────────────
    parts_forecast = []
    # متوسط الحركة اليومية (الأساس لحساب موعد الاستحقاق)
    daily_km_avg = km_predicted / 30.0 if km_predicted > 0 else 0.0
    # العداد الحالي للمركبة
    current_km = max([float(t.get("end_odometer") or 0) for t in trips if t.get("end_odometer")] or [0])

    # دمج oil_motor مع oil قبل المعالجة
    merged_ws = []
    for r in workshop_records:
        rec = dict(r)
        if rec.get("type") == "oil_motor":
            rec["type"] = "oil"
        merged_ws.append(rec)

    for part_type, life in PART_LIFECYCLE.items():
        part_recs = [r for r in merged_ws if r.get("type") == part_type]
        if not part_recs:
            continue

        change_count = len(part_recs)
        # آخر تغيير — نأخذ الأحدث بالكيلومتر ثم بالتاريخ
        last_rec = max(part_recs, key=lambda r: (float(r.get("odometer_reading") or 0), r.get("created_at") or ""))
        last_date_str = (last_rec.get("created_at") or "")[:10]
        last_km = float(last_rec.get("odometer_reading") or 0)

        # ── حساب التكلفة مع استبعاد الشواذ (Outlier Removal) ──
        all_costs = [float(r.get("price") or 0) for r in part_recs if float(r.get("price") or 0) > 0]
        if len(all_costs) >= 3:
            # نستخدم IQR لاستبعاد الشواذ
            sorted_c = sorted(all_costs)
            q1 = sorted_c[len(sorted_c)//4]
            q3 = sorted_c[3*len(sorted_c)//4]
            iqr = q3 - q1
            clean_costs = [c for c in all_costs if q1 - 1.5*iqr <= c <= q3 + 1.5*iqr]
            avg_cost = sum(clean_costs) / len(clean_costs) if clean_costs else sum(all_costs) / len(all_costs)
        elif len(all_costs) == 2:
            # لو اتنين فقط — خد الأقل (أكثر واقعية للتقدير)
            avg_cost = min(all_costs)
        elif len(all_costs) == 1:
            avg_cost = all_costs[0]
        else:
            avg_cost = 0.0

        # ── العمر الفعلي من البيانات الحقيقية ──
        actual_life_km = life["avg_km"]   # افتراضي
        km_readings = sorted([
            float(r.get("odometer_reading") or 0)
            for r in part_recs if r.get("odometer_reading") and float(r.get("odometer_reading") or 0) > 0
        ])
        if len(km_readings) >= 2 and life["avg_km"] > 0:
            diffs = [km_readings[i+1] - km_readings[i] for i in range(len(km_readings)-1)]
            # فلتر: الفارق لازم يكون بين 10% و 300% من العمر الافتراضي
            min_valid = life["avg_km"] * 0.1
            max_valid = life["avg_km"] * 3.0
            valid_diffs = [d for d in diffs if min_valid <= d <= max_valid]
            if valid_diffs:
                # متوسط مرجح — الأحدث وزنه أعلى
                weights = list(range(1, len(valid_diffs)+1))
                actual_life_km = sum(d*w for d,w in zip(valid_diffs, weights)) / sum(weights)
        elif life["avg_km"] == 0:
            actual_life_km = 0  # نعتمد على الأيام فقط

        # ── Safety guard: لو actual_life_km أصغر من 10% من العمر الافتراضي = بيانات غير موثوقة ──
        if life["avg_km"] > 0 and actual_life_km < life["avg_km"] * 0.1:
            actual_life_km = life["avg_km"]  # ارجع للعمر الافتراضي

        # ── حساب العداد المتوقع للاستحقاق القادم ──
        if last_km > 0 and actual_life_km > 0:
            if current_km > last_km:
                km_since_last = current_km - last_km
                if km_since_last >= actual_life_km:
                    # فات الموعد — الاستحقاق = العداد الحالي + الفرق
                    cycles_since = int(km_since_last / actual_life_km)
                    next_due_km = last_km + (cycles_since + 1) * actual_life_km
                else:
                    # لسه في العمر المتوقع
                    next_due_km = last_km + actual_life_km
            else:
                next_due_km = last_km + actual_life_km
        else:
            next_due_km = None  # نعتمد على الأيام فقط

        # ── الكيلومترات المتبقية حتى الاستحقاق ──
        km_remaining = (next_due_km - current_km) if next_due_km and current_km > 0 else None

        # ── الأيام المتبقية بناءً على متوسط الحركة اليومية ──
        days_until = None
        due_date   = None
        prob       = 0.0

        if km_remaining is not None and daily_km_avg > 0:
            days_until = km_remaining / daily_km_avg
            due_dt     = now + timedelta(days=days_until)
            due_date   = due_dt.strftime("%Y-%m-%d")

            # ── الاحتمالية بناءً على متى يقع الاستحقاق ──
            if km_remaining <= 0:
                prob = 0.97   # فات الموعد — عاجل جداً
            elif days_until <= 7:
                prob = 0.95   # خلال أسبوع
            elif days_until <= 30:
                prob = 0.85   # خلال الشهر
            elif days_until <= 45:
                prob = 0.50   # الشهر القادم
            elif days_until <= 60:
                prob = 0.25   # بعدين بشوية
            else:
                prob = 0.05   # بعيد

        elif last_date_str:
            # fallback: نعتمد على الأيام لو مفيش عداد
            try:
                days_since = (now.date() - datetime.strptime(last_date_str, "%Y-%m-%d").date()).days
            except Exception:
                days_since = 9999
            days_left_by_date = life["avg_days"] - days_since
            if days_left_by_date <= 0:
                prob = 0.85
                due_date = now.strftime("%Y-%m-%d")
            elif days_left_by_date <= 30:
                prob = 0.60
                due_date = (now + timedelta(days=days_left_by_date)).strftime("%Y-%m-%d")
            elif days_left_by_date <= 60:
                prob = 0.25
                due_date = (now + timedelta(days=days_left_by_date)).strftime("%Y-%m-%d")
            else:
                prob = 0.05

        # نضيف فقط القطع اللي احتمالها > 5%
        if prob > 0.05:
            parts_forecast.append({
                "part_type":    part_type,
                "part_label":   life["label"],
                "prob":         round(prob, 2),
                "est_cost":     round(avg_cost, 2),
                "due_km":       round(next_due_km, 0) if next_due_km else None,
                "km_remaining": round(km_remaining, 0) if km_remaining is not None else None,
                "days_until":   round(days_until, 1) if days_until is not None else None,
                "due_date":     due_date,
                "avg_life_km":  round(actual_life_km, 0),
                "avg_life_days":life["avg_days"],
                "change_count": change_count,
                "last_km":      round(last_km, 0),
                "last_date":    last_date_str,
                "daily_km_avg": round(daily_km_avg, 1),
            })

    # ترتيب: الأقرب أولاً (بالأيام)، ثم الأعلى احتمالاً
    parts_forecast.sort(key=lambda x: (
        x.get("days_until") if x.get("days_until") is not None else 9999,
        -x["prob"]
    ))

    # ──────────────────────────────────────
    # E. الميزانية
    # ──────────────────────────────────────
    parts_cost = sum(p["est_cost"] * p["prob"] for p in parts_forecast)
    budget_total = pred_fuel_cost + pred_maint_cost + parts_cost

    # ──────────────────────────────────────
    # F. نسبة الثقة
    # ──────────────────────────────────────
    completeness = min(1.0, (
        (0.4 if has_fuel else 0) +
        (0.4 if has_maint else 0) +
        (0.2 if len(km_series) >= 3 else len(km_series) * 0.07)
    ))
    confidence = _confidence_score(
        data_points=data_points,
        cv_km=cv_km,
        cv_cost=cv_cost,
        has_fuel=has_fuel,
        has_maint=has_maint,
        completeness=completeness,
    )

    return {
        "car_id":          car_id,
        "car_plate":       car_plate,
        "forecast_month":  target_month,
        "generated_at":    now.isoformat() + "Z",
        "algorithm":       "WMA_v1",
        "data_points":     data_points,
        # كيلومترات
        "km_predicted":    round(km_predicted, 1),
        "km_trend":        round(slope, 2),
        "km_confidence":   round(min(1.0, max(0.0, 1 - cv_km)), 2),
        "km_series":       {m: round(v, 1) for m, v in zip(sorted_months, km_series)},
        # وقود
        "fuel_liters":     round(pred_fuel_liters, 1),
        "fuel_cost":       round(pred_fuel_cost, 2),
        "fuel_l_per_100":  round(l_per_100, 2),
        "fuel_cost_per_km":round(fuel_cost_per_km, 4),
        # صيانة
        "maint_cost":      round(pred_maint_cost, 2),
        "maint_visits":    round(pred_visits, 1),
        "maint_cost_per_km": round(maint_cost_per_km, 4),
        # قطع الغيار
        "parts":           parts_forecast,
        # ميزانية
        "fuel_budget":     round(pred_fuel_cost, 2),
        "maint_budget":    round(pred_maint_cost + parts_cost, 2),
        "budget_total":    round(budget_total, 2),
        # ثقة
        "confidence":      confidence,
        "confidence_pct":  round(confidence * 100, 1),
    }


# ── POST /forecast/run ──
@app.post("/forecast/run")
async def run_forecast(month: Optional[str] = None, cu: dict = Depends(require_admin)):
    target_month = month or datetime.utcnow().strftime("%Y-%m")
    results = []
    with get_db() as conn:
        c = conn.cursor()
        settings = dict((r[0], r[1]) for r in c.execute("SELECT key,value FROM forecast_settings").fetchall())
        cars = [dict(r) for r in c.execute("SELECT * FROM cars WHERE status != 'inactive'").fetchall()]
        now = datetime.utcnow().isoformat() + "Z"
        for car in cars:
            cid = car["id"]
            trips   = [dict(r) for r in c.execute("SELECT * FROM trips WHERE car_id=? AND start_time >= date('now','-120 days')", (cid,)).fetchall()]
            ws_recs = [dict(r) for r in c.execute("SELECT * FROM workshop_records WHERE vehicle_id=? AND created_at >= date('now','-120 days')", (cid,)).fetchall()]
            maint   = [dict(r) for r in c.execute("SELECT * FROM maintenance_schedule WHERE car_id=?", (cid,)).fetchall()]
            fc = compute_forecast_for_car(
                car_id=cid, car_plate=car.get("plate",""),
                trips=trips, workshop_records=ws_recs,
                maintenance_schedule=maint, settings=settings, target_month=target_month,
            )
            c.execute(
                "INSERT OR REPLACE INTO forecast_results (car_id,forecast_month,generated_at,km_predicted,km_trend,km_confidence,fuel_liters,fuel_cost,fuel_l_per_100,maint_cost,maint_visits,maint_cost_per_km,budget_total,confidence,data_points,algorithm) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (cid, target_month, now, fc["km_predicted"], fc["km_trend"], fc["km_confidence"],
                 fc["fuel_liters"], fc["fuel_cost"], fc["fuel_l_per_100"],
                 fc["maint_cost"], fc["maint_visits"], fc["maint_cost_per_km"],
                 fc["budget_total"], fc["confidence"], fc["data_points"], fc["algorithm"])
            )
            c.execute("DELETE FROM forecast_parts WHERE car_id=? AND forecast_month=?", (cid, target_month))
            for p in fc.get("parts", []):
                c.execute(
                    "INSERT INTO forecast_parts (car_id,forecast_month,part_type,part_label,prob,est_cost,due_km,due_date,avg_life_km,avg_life_days,change_count) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                    (cid, target_month, p["part_type"], p["part_label"], p["prob"], p["est_cost"],
                     p.get("due_km"), p.get("due_date"), p.get("avg_life_km"), p.get("avg_life_days"), p["change_count"])
                )
            fc.update({"branch": car.get("branch",""), "project": car.get("project",""),
                       "model": car.get("model",""), "equipment_type": car.get("equipment_type",""),
                       "sector": car.get("sector","")})
            results.append(fc)
    return {"ok": True, "month": target_month, "cars": len(results), "results": results}


# ── GET /forecast/results ──
@app.get("/forecast/results")
async def get_forecast_results(month: Optional[str] = None, branch: Optional[str] = None,
                                cu: dict = Depends(require_admin_or_reporter)):
    target_month = month or datetime.utcnow().strftime("%Y-%m")
    with get_db() as conn:
        c = conn.cursor()
        query = "SELECT fr.*,c.plate,c.model,c.branch,c.project,c.equipment_type,c.sector,c.status as car_status,(SELECT MAX(end_odometer) FROM trips WHERE car_id=fr.car_id AND end_odometer IS NOT NULL) as current_km FROM forecast_results fr JOIN cars c ON c.id=fr.car_id WHERE fr.forecast_month=?"
        params = [target_month]
        if branch:
            query += " AND c.branch=?"
            params.append(branch)
        query += " ORDER BY fr.budget_total DESC"
        rows = [dict(r) for r in c.execute(query, params).fetchall()]
        parts_map = defaultdict(list)
        if rows:
            car_ids = [r["car_id"] for r in rows]
            ph = ",".join("?" * len(car_ids))
            for p in c.execute(f"SELECT * FROM forecast_parts WHERE car_id IN ({ph}) AND forecast_month=? ORDER BY prob DESC", car_ids + [target_month]).fetchall():
                d = dict(p)
                parts_map[d["car_id"]].append(d)
        for r in rows:
            r["parts"] = parts_map.get(r["car_id"], [])
            r["confidence_pct"] = round((r.get("confidence") or 0) * 100, 1)
        total_budget  = sum(r.get("budget_total") or 0 for r in rows)
        total_fuel    = sum(r.get("fuel_liters") or 0 for r in rows)
        workshop_cars = sum(1 for r in rows if (r.get("maint_visits") or 0) > 0.5)
        top_cost_car  = max(rows, key=lambda r: r.get("budget_total") or 0) if rows else None
        part_counts = defaultdict(lambda: {"count": 0, "label": ""})
        for r in rows:
            for p in r["parts"]:
                if (p.get("prob") or 0) >= 0.5:
                    pt = p["part_type"]
                    part_counts[pt]["count"] += 1
                    part_counts[pt]["label"] = p.get("part_label", pt)
        top_parts = sorted(part_counts.items(), key=lambda x: -x[1]["count"])[:5]
        return {
            "month": target_month, "cars": rows,
            "summary": {
                "total_budget": round(total_budget, 2),
                "total_fuel_liters": round(total_fuel, 1),
                "workshop_cars": workshop_cars,
                "total_cars": len(rows),
                "top_cost_car": top_cost_car,
                "top_parts": [{"type": k, **v} for k, v in top_parts],
                "generated": rows[0].get("generated_at") if rows else None,
            }
        }


# ── GET /forecast/car/{car_id} ──
@app.get("/forecast/car/{car_id}")
async def get_car_forecast(car_id: int, month: Optional[str] = None,
                            cu: dict = Depends(require_admin_or_reporter)):
    target_month = month or datetime.utcnow().strftime("%Y-%m")
    with get_db() as conn:
        c = conn.cursor()
        row = c.execute("SELECT * FROM forecast_results WHERE car_id=? AND forecast_month=?", (car_id, target_month)).fetchone()
        if not row:
            raise HTTPException(404, "لم يتم حساب التنبؤ — شغّل /forecast/run أولاً")
        result = dict(row)
        result["parts"] = [dict(p) for p in c.execute("SELECT * FROM forecast_parts WHERE car_id=? AND forecast_month=? ORDER BY prob DESC", (car_id, target_month)).fetchall()]
        result["confidence_pct"] = round((result.get("confidence") or 0) * 100, 1)
        result["maintenance_plan"] = [dict(r) for r in c.execute("SELECT * FROM maintenance_schedule WHERE car_id=?", (car_id,)).fetchall()]
        return result


# ── GET/PUT /forecast/settings ──
@app.get("/forecast/settings")
async def get_forecast_settings(cu: dict = Depends(require_admin)):
    with get_db() as conn:
        return dict((r[0], r[1]) for r in conn.execute("SELECT key,value FROM forecast_settings").fetchall())

@app.put("/forecast/settings")
async def update_forecast_settings(body: dict, cu: dict = Depends(require_admin)):
    with get_db() as conn:
        for k, v in body.items():
            conn.execute("INSERT OR REPLACE INTO forecast_settings(key,value,note) VALUES(?,?,'')", (k, str(v)))
    return {"ok": True}


# ── التنبيهات (منتجات منخفضة/منتهية) ──
@app.get("/inventory/alerts")
async def inventory_alerts(cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        cats = {r["id"]: r["name"] for r in conn.execute("SELECT id,name FROM inventory_categories").fetchall()}
        sups = {r["id"]: r["name"] for r in conn.execute("SELECT id,name FROM inventory_suppliers").fetchall()}
        rows = conn.execute("""SELECT * FROM inventory_products
                               WHERE min_quantity > 0 AND quantity <= min_quantity
                               ORDER BY quantity ASC""").fetchall()
        return [_inv_product_row(dict(r), cats, sups) for r in rows]


# ── Dashboard ──
@app.get("/inventory/dashboard")
async def inventory_dashboard(cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        total_items = conn.execute("SELECT COUNT(*) c FROM inventory_products").fetchone()["c"]
        stock_value = conn.execute("SELECT COALESCE(SUM(quantity*purchase_price),0) v FROM inventory_products").fetchone()["v"]
        low_stock = conn.execute("SELECT COUNT(*) c FROM inventory_products WHERE min_quantity>0 AND quantity<=min_quantity").fetchone()["c"]
        top_used = conn.execute("""SELECT p.id, p.name, SUM(m.quantity) as total_out
                                   FROM inventory_movements m JOIN inventory_products p ON p.id=m.product_id
                                   WHERE m.movement_type='out'
                                   GROUP BY m.product_id ORDER BY total_out DESC LIMIT 5""").fetchall()
        return {
            "total_items": total_items,
            "stock_value": stock_value,
            "low_stock_count": low_stock,
            "top_used_products": [dict(r) for r in top_used],
        }


# ══════════════════════════════════════════════════════
# 53. مراجعة العمليات — Workshop Approval Workflow (تفويل + صيانة معاً)
# ══════════════════════════════════════════════════════

@app.get("/workshops/pending-review")
async def list_pending_review(cu: dict = Depends(require_admin_or_reporter), status: Optional[str] = None,
                               kind: Optional[str] = None, branch: Optional[str] = None):
    """قائمة عمليات التفويل/الصيانة لشاشة مراجعة العمليات الموحدة."""
    with get_db() as conn:
        q = """SELECT w.*, COALESCE(d.name, op.name) as driver_name,
                      c.plate as vehicle_plate, c.model as vehicle_model
               FROM workshop_records w
               LEFT JOIN drivers d ON w.driver_id=d.id AND (w.is_operator IS NULL OR w.is_operator=0)
               LEFT JOIN equipment_operators op ON w.operator_id=op.id
               LEFT JOIN cars c ON w.vehicle_id=c.id
               WHERE w.type != 'other'"""
        params = []
        eff_branch = _branch_filter(cu) or branch
        if eff_branch:
            q += " AND (d.branch=? OR op.branch=?)"; params += [eff_branch, eff_branch]
        if status:
            q += " AND w.approval_status=?"; params.append(status)
        if kind == "fuel":
            q += " AND w.type LIKE 'fuel_%'"
        elif kind == "maintenance":
            q += " AND w.type NOT LIKE 'fuel_%'"
        q += " ORDER BY w.id DESC LIMIT 1000"
        rows = [dict(r) for r in conn.execute(q, params).fetchall()]
        for r in rows:
            r.setdefault("approval_status", "pending")
        return rows

@app.get("/workshops/fuel-consumption-analysis")
async def fuel_consumption_analysis(cu: dict = Depends(require_admin_or_reporter)):
    """
    يحسب معدل استهلاك الوقود لكل تفويلة بناءً على المعادلة:
    معدل الاستهلاك = (عداد التفويلة الحالية - عداد التفويلة السابقة) / كمية الوقود بالتفويلة الحالية
    (لتر / كم، ثم نحوله لـ لتر/100كم)
    يكتشف الشذوذ بناءً على الانحراف عن متوسط آخر 5 تفويلات لنفس المركبة.
    """
    with get_db() as conn:
        c = conn.cursor()
        q = """SELECT w.id, w.vehicle_id, w.driver_id, w.odometer_reading, w.quantity,
                      w.price, w.created_at, w.type,
                      c.plate as vehicle_plate, c.model as vehicle_model,
                      COALESCE(d.name, op.name) as driver_name
               FROM workshop_records w
               LEFT JOIN cars c ON w.vehicle_id = c.id
               LEFT JOIN drivers d ON w.driver_id = d.id AND (w.is_operator IS NULL OR w.is_operator = 0)
               LEFT JOIN equipment_operators op ON w.operator_id = op.id
               WHERE w.type LIKE 'fuel_%'
                 AND w.vehicle_id IS NOT NULL
                 AND w.odometer_reading IS NOT NULL
                 AND w.quantity IS NOT NULL
                 AND w.quantity > 0
               ORDER BY w.vehicle_id, w.odometer_reading ASC"""
        rows = [dict(r) for r in c.execute(q).fetchall()]

        # تجميع حسب المركبة
        by_car = {}
        for r in rows:
            by_car.setdefault(r["vehicle_id"], []).append(r)

        result = {}
        for car_id, recs in by_car.items():
            recs_sorted = sorted(recs, key=lambda x: x["odometer_reading"])
            enriched = []
            consumptions = []  # لتر/100كم لكل سجلات الحساب الصحيحة

            for i, rec in enumerate(recs_sorted):
                entry = dict(rec)
                if i == 0:
                    entry["prev_odometer"]   = None
                    entry["km_driven"]       = None
                    entry["consumption_l100"]= None
                    entry["is_anomaly"]      = False
                    entry["anomaly_reason"]  = None
                    enriched.append(entry)
                    continue

                prev = recs_sorted[i-1]
                km_driven = rec["odometer_reading"] - prev["odometer_reading"]
                qty = rec["quantity"]

                entry["prev_odometer"] = prev["odometer_reading"]
                entry["km_driven"]     = round(km_driven, 1)

                if km_driven <= 0:
                    # عداد رجع للخلف أو نفس القراءة — بيانات غير منطقية
                    entry["consumption_l100"] = None
                    entry["is_anomaly"]     = True
                    entry["anomaly_reason"] = "عداد غير منطقي (لم يتقدم منذ آخر تفويلة)"
                elif qty <= 0:
                    entry["consumption_l100"] = None
                    entry["is_anomaly"]     = True
                    entry["anomaly_reason"] = "كمية وقود غير صحيحة"
                else:
                    # معدل الاستهلاك حسب المعادلة المطلوبة: km_driven / qty = كم/لتر
                    km_per_liter = km_driven / qty
                    consumption_l100 = (qty / km_driven) * 100  # لتر/100كم — الصيغة القياسية
                    entry["km_per_liter"]      = round(km_per_liter, 2)
                    entry["consumption_l100"]  = round(consumption_l100, 2)

                    # ── كشف الشذوذ: مقارنة بمتوسط آخر 5 تفويلات سابقة صحيحة لنفس المركبة ──
                    recent_valid = [c_ for c_ in consumptions[-5:] if c_ is not None]
                    if len(recent_valid) >= 2:
                        avg_recent = sum(recent_valid) / len(recent_valid)
                        std_dev = (sum((x-avg_recent)**2 for x in recent_valid) / len(recent_valid)) ** 0.5
                        # عتبة الشذوذ: انحراف أكبر من 40% عن المتوسط أو أكبر من 2.5 انحراف معياري
                        deviation_pct = abs(consumption_l100 - avg_recent) / avg_recent if avg_recent > 0 else 0
                        threshold_std = max(std_dev * 2.5, avg_recent * 0.4)

                        if abs(consumption_l100 - avg_recent) > threshold_std:
                            entry["is_anomaly"] = True
                            if consumption_l100 > avg_recent:
                                entry["anomaly_reason"] = (
                                    f"استهلاك أعلى من المعتاد بنسبة {round(deviation_pct*100)}٪ "
                                    f"(المعتاد ≈ {round(avg_recent,1)} لتر/100كم)"
                                )
                            else:
                                entry["anomaly_reason"] = (
                                    f"استهلاك أقل من المعتاد بنسبة {round(deviation_pct*100)}٪ "
                                    f"— قد يشير لتفويلة جزئية أو خطأ تسجيل "
                                    f"(المعتاد ≈ {round(avg_recent,1)} لتر/100كم)"
                                )
                        else:
                            entry["is_anomaly"] = False
                            entry["anomaly_reason"] = None
                    else:
                        # مفيش بيانات كافية للمقارنة بعد
                        entry["is_anomaly"] = False
                        entry["anomaly_reason"] = None

                    entry["baseline_avg_l100"] = round(sum(recent_valid)/len(recent_valid), 2) if recent_valid else None
                    consumptions.append(consumption_l100)

                enriched.append(entry)

            result[str(car_id)] = enriched

        return result


@app.get("/workshops/fuel-consumption-analysis/{record_id}")
async def fuel_consumption_for_record(record_id: int, cu: dict = Depends(require_admin_or_reporter)):
    """يرجع تحليل الاستهلاك لتفويلة واحدة محددة (يُستخدم في شاشة مراجعة العمليات عند فتح تفاصيل التفويلة)."""
    with get_db() as conn:
        c = conn.cursor()
        rec = c.execute("SELECT * FROM workshop_records WHERE id=?", (record_id,)).fetchone()
        if not rec:
            raise HTTPException(404, "السجل غير موجود")
        rec = dict(rec)
        if not (rec.get("type") or "").startswith("fuel_") or not rec.get("vehicle_id"):
            return {"applicable": False}

        # نجيب آخر تفويلة سابقة بعداد أقل لنفس المركبة
        prev = c.execute(
            """SELECT * FROM workshop_records
               WHERE vehicle_id=? AND type LIKE 'fuel_%' AND id != ?
                 AND odometer_reading IS NOT NULL AND odometer_reading < ?
               ORDER BY odometer_reading DESC LIMIT 1""",
            (rec["vehicle_id"], record_id, rec.get("odometer_reading") or 999999999)
        ).fetchone()

        if not prev or not rec.get("odometer_reading") or not rec.get("quantity"):
            return {"applicable": True, "has_previous": False}

        prev = dict(prev)
        km_driven = rec["odometer_reading"] - prev["odometer_reading"]
        qty = rec["quantity"]

        if km_driven <= 0 or qty <= 0:
            return {
                "applicable": True, "has_previous": True,
                "is_anomaly": True,
                "anomaly_reason": "بيانات غير منطقية (عداد أو كمية غير صحيحة)",
                "km_driven": km_driven, "prev_odometer": prev["odometer_reading"],
            }

        consumption_l100 = (qty / km_driven) * 100

        # متوسط آخر 5 تفويلات قبل دي
        recent = c.execute(
            """SELECT odometer_reading, quantity FROM workshop_records
               WHERE vehicle_id=? AND type LIKE 'fuel_%' AND id != ?
                 AND odometer_reading IS NOT NULL AND odometer_reading <= ?
               ORDER BY odometer_reading DESC LIMIT 6""",
            (rec["vehicle_id"], record_id, prev["odometer_reading"])
        ).fetchall()
        recent = [dict(r) for r in recent]
        recent_consumptions = []
        for i in range(len(recent)-1):
            km_d = recent[i]["odometer_reading"] - recent[i+1]["odometer_reading"]
            q_ = recent[i]["quantity"]
            if km_d > 0 and q_ and q_ > 0:
                recent_consumptions.append((q_ / km_d) * 100)

        is_anomaly = False
        anomaly_reason = None
        baseline = None
        if len(recent_consumptions) >= 2:
            avg_recent = sum(recent_consumptions) / len(recent_consumptions)
            std_dev = (sum((x-avg_recent)**2 for x in recent_consumptions) / len(recent_consumptions)) ** 0.5
            threshold = max(std_dev * 2.5, avg_recent * 0.4)
            baseline = round(avg_recent, 2)
            if abs(consumption_l100 - avg_recent) > threshold:
                is_anomaly = True
                deviation_pct = round(abs(consumption_l100 - avg_recent) / avg_recent * 100) if avg_recent > 0 else 0
                if consumption_l100 > avg_recent:
                    anomaly_reason = f"استهلاك أعلى من المعتاد بنسبة {deviation_pct}٪"
                else:
                    anomaly_reason = f"استهلاك أقل من المعتاد بنسبة {deviation_pct}٪ — تحقق من صحة البيانات"

        return {
            "applicable": True, "has_previous": True,
            "prev_odometer": prev["odometer_reading"],
            "current_odometer": rec["odometer_reading"],
            "km_driven": round(km_driven, 1),
            "fuel_quantity": qty,
            "consumption_l100": round(consumption_l100, 2),
            "km_per_liter": round(km_driven / qty, 2),
            "baseline_avg_l100": baseline,
            "is_anomaly": is_anomaly,
            "anomaly_reason": anomaly_reason,
        }


@app.post("/workshops/{wid}/review")
async def review_workshop_record(wid: int, body: WorkshopApproval, cu: dict = Depends(require_workshop_admin)):
    if body.action not in ("approve", "reject"):
        raise HTTPException(400, "إجراء غير صالح")
    new_status = "approved" if body.action == "approve" else "rejected"
    reviewer_name = cu.get("username") or cu.get("name") or ""
    with get_db() as conn:
        row = conn.execute("SELECT id FROM workshop_records WHERE id=?", (wid,)).fetchone()
        if not row:
            raise HTTPException(404, "السجل غير موجود")
        conn.execute(
            "UPDATE workshop_records SET approval_status=?, reviewed_by=?, reviewed_at=? WHERE id=?",
            (new_status, reviewer_name, datetime.utcnow().isoformat(), wid)
        )
        log_event("workshop_review", record_id=wid, action=body.action, note=body.note or "")
        return {"ok": True, "approval_status": new_status, "reviewed_by": reviewer_name}


class WorkshopPriceUpdate(BaseModel):
    price: float


@app.put("/workshops/{wid}/price")
async def update_workshop_price(wid: int, body: WorkshopPriceUpdate, cu: dict = Depends(require_workshop_admin)):
    """تصحيح سعر/تكلفة سجل تفويل أو صيانة بعد المراجعة، لو مش مطابق للفاتورة."""
    if body.price is None or body.price < 0:
        raise HTTPException(400, "قيمة غير صالحة")
    editor_name = cu.get("username") or cu.get("name") or ""
    with get_db() as conn:
        row = conn.execute("SELECT id FROM workshop_records WHERE id=?", (wid,)).fetchone()
        if not row:
            raise HTTPException(404, "السجل غير موجود")
        conn.execute(
            "UPDATE workshop_records SET price=?, price_edited_by=?, price_edited_at=? WHERE id=?",
            (body.price, editor_name, datetime.utcnow().isoformat(), wid)
        )
        log_event("workshop_price_edit", record_id=wid, note=f"new_price={body.price}")
        return {"ok": True, "price": body.price, "price_edited_by": editor_name}


class WorkshopQuantityUpdate(BaseModel):
    quantity: float


@app.put("/workshops/{wid}/quantity")
async def update_workshop_quantity(wid: int, body: WorkshopQuantityUpdate, cu: dict = Depends(require_workshop_admin)):
    """تصحيح الكمية (لترات الوقود/عدد الإطارات...إلخ) لسجل بعد المراجعة، لو مش مطابقة للفاتورة أو صورة العداد."""
    if body.quantity is None or body.quantity < 0:
        raise HTTPException(400, "قيمة غير صالحة")
    editor_name = cu.get("username") or cu.get("name") or ""
    with get_db() as conn:
        # تأكد من وجود الأعمدة (لو السيرفر اتعمل له deploy جديد ولسه محدش أنشأ سجل صيانة يفعّل الـ migration)
        _existing_cols = {r["name"] for r in conn.execute("PRAGMA table_info(workshop_records)").fetchall()}
        for _col, _def in (("quantity_edited_by", "TEXT DEFAULT ''"), ("quantity_edited_at", "TEXT DEFAULT ''")):
            if _col not in _existing_cols:
                try: conn.execute(f"ALTER TABLE workshop_records ADD COLUMN {_col} {_def}")
                except Exception: pass
        row = conn.execute("SELECT id FROM workshop_records WHERE id=?", (wid,)).fetchone()
        if not row:
            raise HTTPException(404, "السجل غير موجود")
        conn.execute(
            "UPDATE workshop_records SET quantity=?, quantity_edited_by=?, quantity_edited_at=? WHERE id=?",
            (body.quantity, editor_name, datetime.utcnow().isoformat(), wid)
        )
        log_event("workshop_quantity_edit", record_id=wid, note=f"new_quantity={body.quantity}")
        return {"ok": True, "quantity": body.quantity, "quantity_edited_by": editor_name}


# ══════════════════════════════════════════════════════
# 53.5 SUPERUSER NOTES — ملاحظات مشتركة بين السوبر يوزرز
# ══════════════════════════════════════════════════════

def _ensure_superuser_notes_table(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS superuser_notes (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        author_username TEXT NOT NULL,
        content         TEXT NOT NULL,
        created_at      TEXT DEFAULT (datetime('now')),
        updated_at      TEXT DEFAULT ''
    )""")
    try: conn.execute("CREATE INDEX IF NOT EXISTS idx_su_notes_created ON superuser_notes(created_at)")
    except Exception: pass


@app.on_event("startup")
async def _migrate_superuser_notes():
    with get_db() as conn:
        _ensure_superuser_notes_table(conn)


class SuperNoteCreate(BaseModel):
    content: str


@app.get("/superuser-notes")
async def list_superuser_notes(cu: dict = Depends(require_superuser)):
    """كل الملاحظات المشتركة بين السوبر يوزرز — ترتيب من الأحدث للأقدم"""
    with get_db() as conn:
        _ensure_superuser_notes_table(conn)
        rows = [dict(r) for r in conn.execute(
            "SELECT * FROM superuser_notes ORDER BY id DESC"
        ).fetchall()]
    return {"notes": rows}


@app.post("/superuser-notes")
async def create_superuser_note(body: SuperNoteCreate, cu: dict = Depends(require_superuser)):
    content = (body.content or "").strip()
    if not content:
        raise HTTPException(400, "اكتب ملاحظة قبل الإرسال")
    if len(content) > 4000:
        raise HTTPException(400, "الملاحظة طويلة جداً")
    with get_db() as conn:
        _ensure_superuser_notes_table(conn)
        conn.execute(
            "INSERT INTO superuser_notes(author_username, content, created_at) VALUES(?,?,?)",
            (cu["username"], content, datetime.utcnow().isoformat())
        )
        nid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    return {"id": nid, "message": "تم إضافة الملاحظة"}


@app.put("/superuser-notes/{note_id}")
async def update_superuser_note(note_id: int, body: SuperNoteCreate, cu: dict = Depends(require_superuser)):
    content = (body.content or "").strip()
    if not content:
        raise HTTPException(400, "اكتب ملاحظة قبل الحفظ")
    with get_db() as conn:
        _ensure_superuser_notes_table(conn)
        row = conn.execute("SELECT * FROM superuser_notes WHERE id=?", (note_id,)).fetchone()
        if not row:
            raise HTTPException(404, "الملاحظة غير موجودة")
        if row["author_username"] != cu["username"]:
            raise HTTPException(403, "تقدر تعدّل ملاحظاتك أنت فقط")
        conn.execute(
            "UPDATE superuser_notes SET content=?, updated_at=? WHERE id=?",
            (content, datetime.utcnow().isoformat(), note_id)
        )
    return {"message": "تم تعديل الملاحظة"}


@app.delete("/superuser-notes/{note_id}")
async def delete_superuser_note(note_id: int, cu: dict = Depends(require_superuser)):
    with get_db() as conn:
        _ensure_superuser_notes_table(conn)
        row = conn.execute("SELECT * FROM superuser_notes WHERE id=?", (note_id,)).fetchone()
        if not row:
            raise HTTPException(404, "الملاحظة غير موجودة")
        if row["author_username"] != cu["username"]:
            raise HTTPException(403, "تقدر تحذف ملاحظاتك أنت فقط")
        conn.execute("DELETE FROM superuser_notes WHERE id=?", (note_id,))
    return {"message": "تم حذف الملاحظة"}


# ══════════════════════════════════════════════════════
# 54. ENTRY POINT
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
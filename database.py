# database.py — Fleet Management System
# يُستخدم لإنشاء قاعدة البيانات لأول مرة فقط
# التهجير (migration) يتم تلقائياً عبر main.py عند الإقلاع

import os
import sqlite3

try:
    import bcrypt
except ImportError:
    print("يرجى تثبيت bcrypt: pip install bcrypt")
    exit(1)


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt(rounds=12)).decode("utf-8")


def create_database():
    db_path = os.environ.get("DATABASE_PATH", "trip_tracker.db")
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA journal_mode=WAL")
    cursor = conn.cursor()

    # ── جدول المستخدمين ────────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT    UNIQUE NOT NULL,
            password TEXT    NOT NULL,
            role     TEXT    NOT NULL CHECK(role IN ('admin', 'driver'))
        )
    """)

    # ── جدول السائقين ──────────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS drivers (
            id                     INTEGER PRIMARY KEY AUTOINCREMENT,
            name                   TEXT NOT NULL,
            phone                  TEXT DEFAULT '',
            status                 TEXT DEFAULT 'active',
            user_id                INTEGER UNIQUE,
            national_id            TEXT DEFAULT '',
            birth_date             TEXT DEFAULT '',
            driver_license_expiry  TEXT DEFAULT '',
            vehicle_license_expiry TEXT DEFAULT '',
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
        )
    """)

    # ── جدول المركبات ──────────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cars (
            id     INTEGER PRIMARY KEY AUTOINCREMENT,
            plate  TEXT NOT NULL UNIQUE,
            model  TEXT NOT NULL,
            status TEXT DEFAULT 'available'
        )
    """)

    # ── جدول صلاحيات السائقين ──────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS driver_car_permissions (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER NOT NULL,
            car_id    INTEGER NOT NULL,
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE CASCADE,
            FOREIGN KEY (car_id)    REFERENCES cars(id)    ON DELETE CASCADE,
            UNIQUE(driver_id, car_id)
        )
    """)

    # ── جدول الرحلات ───────────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS trips (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id       INTEGER NOT NULL,
            car_id          INTEGER NOT NULL,
            start_time      TEXT,
            end_time        TEXT,
            start_odometer  REAL,
            end_odometer    REAL,
            start_location  TEXT DEFAULT '',
            end_location    TEXT DEFAULT '',
            garage_location TEXT DEFAULT '',
            notes           TEXT DEFAULT '',
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE CASCADE,
            FOREIGN KEY (car_id)    REFERENCES cars(id)    ON DELETE CASCADE
        )
    """)

    # ── جدول سجلات الورشة ──────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workshop_records (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id        INTEGER,
            type             TEXT NOT NULL,
            quantity         REAL,
            price            REAL NOT NULL DEFAULT 0,
            notes            TEXT DEFAULT '',
            created_at       TEXT NOT NULL,
            operation_type   TEXT DEFAULT '',
            vehicle_id       INTEGER,
            odometer_reading REAL,
            description      TEXT DEFAULT '',
            tire_action      TEXT DEFAULT '',
            location         TEXT DEFAULT '',
            FOREIGN KEY (driver_id)  REFERENCES drivers(id),
            FOREIGN KEY (vehicle_id) REFERENCES cars(id)
        )
    """)

    # ── جدول بلاغات الطوارئ ────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS emergency_reports (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id    INTEGER NOT NULL,
            car_id       INTEGER,
            type         TEXT NOT NULL CHECK(type IN ('emergency', 'accident')),
            audio_data   TEXT NOT NULL,
            notes        TEXT DEFAULT '',
            created_at   TEXT NOT NULL,
            is_read      INTEGER DEFAULT 0,
            location     TEXT DEFAULT '',
            is_handled   INTEGER DEFAULT 0,
            action_taken TEXT DEFAULT '',
            handled_by   TEXT DEFAULT '',
            action_time  TEXT DEFAULT '',
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE CASCADE,
            FOREIGN KEY (car_id)    REFERENCES cars(id)    ON DELETE SET NULL
        )
    """)

    # ── جدول سجلات الجراج ──────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS garage_records (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id   INTEGER NOT NULL,
            car_id      INTEGER,
            location    TEXT NOT NULL,
            recorded_at TEXT NOT NULL,
            notes       TEXT DEFAULT ''
        )
    """)

    # ── جدول الإعدادات ─────────────────────────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS app_settings (
            key        TEXT PRIMARY KEY,
            value      TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)

    # ── Indexes للأداء ─────────────────────────────────────────
    indexes = [
        "CREATE INDEX IF NOT EXISTS idx_trips_driver      ON trips(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_trips_car         ON trips(car_id)",
        "CREATE INDEX IF NOT EXISTS idx_emg_driver        ON emergency_reports(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_emg_unread        ON emergency_reports(is_read)",
        "CREATE INDEX IF NOT EXISTS idx_ws_driver         ON workshop_records(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_garage_driver     ON garage_records(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_dcp_driver        ON driver_car_permissions(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_dcp_car           ON driver_car_permissions(car_id)",
        "CREATE INDEX IF NOT EXISTS idx_users_username    ON users(username)",
        "CREATE INDEX IF NOT EXISTS idx_drivers_user      ON drivers(user_id)",
    ]
    for idx in indexes:
        cursor.execute(idx)

    # ── Seed أسعار افتراضية ────────────────────────────────────
    from datetime import datetime
    for ws_type in ['fuel', 'oil', 'filter', 'tire', 'battery', 'belt', 'other']:
        cursor.execute(
            "INSERT OR IGNORE INTO app_settings(key,value,updated_at) VALUES(?,?,?)",
            (f"price_{ws_type}", "0", datetime.now().isoformat())
        )

    # ── حذف كل الأدمنز القديمين وإنشاء الأدمنز الجدد ──────────
    ADMINS = [
        ("Eng mohamed mansour", "mo@mansour241"),
        ("Eng mohamed sayed",   "mo@sayed11214123"),
        ("Eng abdelrhman sayed","abdo@11214123"),
        ("admin", "ad11228898"),
    ]
    # Delete old admins (keep drivers)
    cursor.execute("SELECT id FROM users WHERE role='admin'")
    old_admin_ids = [r[0] for r in cursor.fetchall()]
    for aid in old_admin_ids:
        cursor.execute("DELETE FROM users WHERE id=?", (aid,))

    # Create new admins
    for uname, pw in ADMINS:
        cursor.execute("SELECT id FROM users WHERE username=?", (uname,))
        if not cursor.fetchone():
            hashed_pw = hash_password(pw)
            cursor.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                           (uname, hashed_pw, "admin"))
            print(f"✅ Admin: {uname}")

    conn.commit()
    conn.close()
    print(f"✅ تم إنشاء قاعدة البيانات: {db_path}")


if __name__ == "__main__":
    create_database()

# database.py — Fleet Management System
# يُستخدم لإنشاء قاعدة البيانات لأول مرة فقط
# التهجير (migration) يتم تلقائياً عبر main.py عند الإقلاع

import os
import sqlite3
from datetime import datetime

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

    # ── جدول المستخدمين ──
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            username   TEXT    NOT NULL,
            password   TEXT    NOT NULL,
            role       TEXT    NOT NULL CHECK(role IN ('superuser', 'admin', 'driver', 'reporter', 'supervisor_workshop', 'supervisor_field', 'operator', 'workshop_admin')),
            branch     TEXT    DEFAULT '',
            created_at TEXT    DEFAULT(datetime('now')),
            last_login TEXT,
            refresh_token TEXT,
            refresh_exp   TEXT,
            avatar_url    TEXT    DEFAULT '',
            managed_by    INTEGER DEFAULT NULL,
            failed_attempts INTEGER DEFAULT 0,
            locked_until    TEXT    DEFAULT NULL
        )
    """)

    # ── جدول السائقين ──
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
            branch                 TEXT DEFAULT '',
            fixed_number           TEXT DEFAULT '',
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
        )
    """)

    # ── جدول المركبات ──
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cars (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            plate              TEXT NOT NULL UNIQUE,
            model              TEXT NOT NULL,
            status             TEXT DEFAULT 'available',
            car_name           TEXT DEFAULT '',
            car_code           TEXT DEFAULT '',
            chassis            TEXT DEFAULT '',
            engine_number      TEXT DEFAULT '',
            year               TEXT DEFAULT '',
            project            TEXT DEFAULT '',
            branch             TEXT DEFAULT '',
            car_license_expiry TEXT DEFAULT '',
            equipment_type     TEXT DEFAULT '',
            sector             TEXT DEFAULT ''
        )
    """)

    # ── جدول صلاحيات السائقين ──
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

    # ── جدول الرحلات ──
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

    # ── جدول سجلات الورشة ──
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
            odometer_photo   TEXT DEFAULT '',
            FOREIGN KEY (driver_id)  REFERENCES drivers(id),
            FOREIGN KEY (vehicle_id) REFERENCES cars(id)
        )
    """)

    # ── جدول بلاغات الطوارئ ──
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS emergency_reports (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id    INTEGER NOT NULL,
            car_id       INTEGER,
            type         TEXT NOT NULL CHECK(type IN ('emergency', 'accident')),
            audio_url    TEXT DEFAULT '',
            notes        TEXT DEFAULT '',
            created_at   TEXT NOT NULL,
            is_read      INTEGER DEFAULT 0,
            location     TEXT DEFAULT '',
            is_handled   INTEGER DEFAULT 0,
            action_taken TEXT DEFAULT '',
            handled_by   TEXT DEFAULT '',
            action_time  TEXT DEFAULT '',
            driver_message TEXT DEFAULT '',
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE CASCADE,
            FOREIGN KEY (car_id)    REFERENCES cars(id)    ON DELETE SET NULL
        )
    """)

    # ── جدول سجلات الجراج ──
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

    # ── جدول طلبات السائقين ──
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS driver_requests (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id       INTEGER NOT NULL,
            type            TEXT NOT NULL CHECK(type IN ('odometer_change','permission','leave','other')),
            notes           TEXT DEFAULT '',
            new_odometer    REAL,
            trip_id         INTEGER,
            status          TEXT DEFAULT 'pending' CHECK(status IN ('pending','approved','rejected')),
            admin_notes     TEXT DEFAULT '',
            admin_message   TEXT DEFAULT '',
            handled_by      TEXT DEFAULT '',
            handled_at      TEXT DEFAULT '',
            created_at      TEXT NOT NULL,
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE CASCADE
        )
    """)

    # ── جدول الرسائل الصوتية ──
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS voice_notes (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            sender_id      INTEGER NOT NULL,
            sender_role    TEXT    NOT NULL DEFAULT 'superuser',
            target_group   TEXT    DEFAULT '',
            target_user_id INTEGER DEFAULT NULL,
            audio_data     TEXT    NOT NULL,
            duration_sec   REAL    DEFAULT 0,
            created_at     TEXT    NOT NULL,
            expires_at     TEXT    NOT NULL,
            play_count     INTEGER DEFAULT 0,
            max_plays      INTEGER DEFAULT 2,
            is_deleted     INTEGER DEFAULT 0
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_logs (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id     INTEGER,
            username    TEXT NOT NULL,
            role        TEXT NOT NULL,
            action      TEXT NOT NULL,
            details     TEXT DEFAULT '',
            ip_address  TEXT DEFAULT '',
            created_at  TEXT NOT NULL
        )
    """)

    # ── جدول محاضر الفحص (الورشة) ──
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workshop_inspection_reports (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            report_number    TEXT DEFAULT '',
            car_id           INTEGER NOT NULL,
            driver_id        INTEGER DEFAULT NULL,
            quote_id         INTEGER DEFAULT NULL,
            inspection_date  TEXT DEFAULT '',
            inspector_name   TEXT DEFAULT '',
            odometer_reading REAL,
            result           TEXT DEFAULT 'pending' CHECK(result IN ('pass','fail','needs_repair','pending')),
            findings         TEXT DEFAULT '',
            items_json       TEXT DEFAULT '[]',
            notes            TEXT DEFAULT '',
            created_by       TEXT DEFAULT '',
            created_at       TEXT DEFAULT (datetime('now')),
            FOREIGN KEY (car_id)    REFERENCES cars(id)    ON DELETE SET NULL,
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE SET NULL,
            FOREIGN KEY (quote_id)  REFERENCES workshop_repair_quotes(id) ON DELETE SET NULL
        )
    """)

    # ── جدول طلبات التشغيل الخارجي ──
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workshop_external_ops (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            request_number    TEXT DEFAULT '',
            quote_id          INTEGER DEFAULT NULL,
            car_id            INTEGER NOT NULL,
            driver_id         INTEGER DEFAULT NULL,
            workshop_name     TEXT DEFAULT '',
            reason            TEXT DEFAULT '',
            estimated_cost    REAL DEFAULT 0,
            workmanship_cost  REAL DEFAULT 0,
            request_date      TEXT DEFAULT '',
            status            TEXT DEFAULT 'pending' CHECK(status IN ('pending','approved','rejected','done')),
            notes             TEXT DEFAULT '',
            created_by        TEXT DEFAULT '',
            created_at        TEXT DEFAULT (datetime('now')),
            FOREIGN KEY (car_id)    REFERENCES cars(id)    ON DELETE SET NULL,
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE SET NULL,
            FOREIGN KEY (quote_id)  REFERENCES workshop_repair_quotes(id) ON DELETE SET NULL
        )
    """)

    # ── جدول إذن الصرف ──
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS workshop_vouchers (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            voucher_number   TEXT DEFAULT '',
            quote_id         INTEGER DEFAULT NULL,
            car_id           INTEGER NOT NULL,
            driver_id        INTEGER DEFAULT NULL,
            recipient_name   TEXT DEFAULT '',
            purpose          TEXT DEFAULT '',
            items_json       TEXT DEFAULT '[]',
            total_amount     REAL DEFAULT 0,
            issue_date       TEXT DEFAULT '',
            issued_by        TEXT DEFAULT '',
            status           TEXT DEFAULT 'draft' CHECK(status IN ('draft','issued','cancelled')),
            notes            TEXT DEFAULT '',
            created_by       TEXT DEFAULT '',
            created_at       TEXT DEFAULT (datetime('now')),
            FOREIGN KEY (car_id)    REFERENCES cars(id)    ON DELETE SET NULL,
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE SET NULL,
            FOREIGN KEY (quote_id)  REFERENCES workshop_repair_quotes(id) ON DELETE SET NULL
        )
    """)

    # ── جدول الإعدادات ──
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS app_settings (
            key        TEXT PRIMARY KEY,
            value      TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)

    # ── Indexes ──
    indexes = [
        "CREATE INDEX IF NOT EXISTS idx_trips_driver   ON trips(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_trips_car      ON trips(car_id)",
        "CREATE INDEX IF NOT EXISTS idx_emg_driver     ON emergency_reports(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_emg_unread     ON emergency_reports(is_read)",
        "CREATE INDEX IF NOT EXISTS idx_ws_driver      ON workshop_records(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_ws_photo       ON workshop_records(odometer_photo)",
        "CREATE INDEX IF NOT EXISTS idx_garage_driver  ON garage_records(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_dcp_driver     ON driver_car_permissions(driver_id)",
        "CREATE INDEX IF NOT EXISTS idx_dcp_car        ON driver_car_permissions(car_id)",
        "CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)",
        "CREATE INDEX IF NOT EXISTS idx_users_managed_by ON users(managed_by)",
        "CREATE INDEX IF NOT EXISTS idx_drivers_user   ON drivers(user_id)",
        "CREATE INDEX IF NOT EXISTS idx_audit_user     ON audit_logs(user_id)",
        "CREATE INDEX IF NOT EXISTS idx_audit_created  ON audit_logs(created_at)",
        "CREATE INDEX IF NOT EXISTS idx_vnotes_group   ON voice_notes(target_group)",
        "CREATE INDEX IF NOT EXISTS idx_vnotes_target  ON voice_notes(target_user_id)",
        "CREATE INDEX IF NOT EXISTS idx_vnotes_expires ON voice_notes(expires_at)",
        # الرقم الثابت يجب أن يكون فريداً (unique) — اللي ممكن يتكرر هو اسم المستخدم فقط
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_drivers_fixed_number ON drivers(fixed_number) WHERE fixed_number IS NOT NULL AND fixed_number != ''",
    ]
    for idx in indexes:
        cursor.execute(idx)

    # ── Seed أسعار ──
    for ws_type in ["fuel_solar","fuel_92","fuel_95","fuel_80","fuel_cng",
                    "oil","filter","tire","battery","belt","other"]:
        cursor.execute(
            "INSERT OR IGNORE INTO app_settings(key,value,updated_at) VALUES(?,?,?)",
            (f"price_{ws_type}", "0", datetime.now().isoformat())
        )

    # ── حسابات البذر (اختيارية) — تُقرأ من متغيرات البيئة بدل كتابتها في الكود ──
    # الصيغة: "username:password,username2:password2"
    # مثال:   SEED_ADMINS="admin:كلمة_مرور_قوية,Eng mohamed sayed:كلمة_مرور_تانية"
    # لو المتغير غير موجود، مفيش أي حساب هيتعمل لهذا الدور (تقدر تعمل الحسابات
    # يدوياً بعدين من واجهة السوبر يوزر).
    def _parse_seed_accounts(env_var):
        raw = os.environ.get(env_var, "")
        if not raw.strip():
            return []
        out = []
        for pair in raw.split(","):
            pair = pair.strip()
            if not pair or ":" not in pair:
                continue
            uname, pw = pair.split(":", 1)
            uname, pw = uname.strip(), pw.strip()
            if uname and pw:
                out.append((uname, pw))
        return out

    # ── حسابات الأدمن ──
    ADMINS = _parse_seed_accounts("SEED_ADMINS")
    if not ADMINS:
        print("ℹ️  SEED_ADMINS غير معرَّف — لم يتم إنشاء أي حساب أدمن تلقائياً")
    for uname, pw in ADMINS:
        cursor.execute("SELECT id FROM users WHERE username=?", (uname,))
        if not cursor.fetchone():
            cursor.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                           (uname, hash_password(pw), "admin"))
            print(f"✅ Admin: {uname}")

    # ── حسابات السوبر يوزر ──
    SUPERUSERS = _parse_seed_accounts("SEED_SUPERUSERS")
    if not SUPERUSERS:
        print("ℹ️  SEED_SUPERUSERS غير معرَّف — لم يتم إنشاء أي حساب سوبر يوزر تلقائياً")
    for uname, pw in SUPERUSERS:
        cursor.execute("SELECT id FROM users WHERE username=?", (uname,))
        if not cursor.fetchone():
            cursor.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                           (uname, hash_password(pw), "superuser"))
            print(f"✅ Superuser: {uname}")

    # ── حسابات الريبورتر (قراءة فقط) ──
    REPORTERS = _parse_seed_accounts("SEED_REPORTERS")
    if not REPORTERS:
        print("ℹ️  SEED_REPORTERS غير معرَّف — لم يتم إنشاء أي حساب مراقب تلقائياً")
    for uname, pw in REPORTERS:
        cursor.execute("SELECT id FROM users WHERE username=?", (uname,))
        if not cursor.fetchone():
            cursor.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                           (uname, hash_password(pw), "reporter"))
            print(f"✅ Reporter: {uname}")

    conn.commit()
    conn.close()
    print(f"✅ تم إنشاء قاعدة البيانات: {db_path}")


if __name__ == "__main__":
    create_database()
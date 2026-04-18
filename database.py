# database_schema.py
import sqlite3
import bcrypt


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')


def create_database():
    conn = sqlite3.connect('trip_tracker.db')
    cursor = conn.cursor()

    # جدول المستخدمين
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'driver'))
        )
    """)

    # جدول السائقين (مرتبط بـ users)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS drivers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            phone TEXT,
            status TEXT DEFAULT 'active',
            user_id INTEGER UNIQUE,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
        )
    """)

    # جدول المركبات
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cars (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plate TEXT NOT NULL UNIQUE,
            model TEXT NOT NULL,
            status TEXT DEFAULT 'available'
        )
    """)

    # جدول صلاحيات السائقين
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS driver_car_permissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER NOT NULL,
            car_id INTEGER NOT NULL,
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE CASCADE,
            FOREIGN KEY (car_id) REFERENCES cars(id) ON DELETE CASCADE,
            UNIQUE(driver_id, car_id)
        )
    """)

    # جدول الرحلات
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS trips (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER NOT NULL,
            car_id INTEGER NOT NULL,
            start_time TEXT,
            end_time TEXT,
            start_odometer REAL,
            end_odometer REAL,
            start_location TEXT DEFAULT '',
            end_location TEXT DEFAULT '',
            notes TEXT,
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE CASCADE,
            FOREIGN KEY (car_id) REFERENCES cars(id) ON DELETE CASCADE
        )
    """)

    for column in ["start_location", "end_location"]:
        try:
            cursor.execute(f"ALTER TABLE trips ADD COLUMN {column} TEXT DEFAULT ''")
        except sqlite3.OperationalError:
            pass

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS emergency_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER NOT NULL,
            car_id INTEGER NOT NULL,
            type TEXT NOT NULL CHECK(type IN ('emergency', 'accident')),
            audio_data TEXT NOT NULL,
            notes TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            is_read INTEGER DEFAULT 0,
            FOREIGN KEY (driver_id) REFERENCES drivers(id) ON DELETE CASCADE,
            FOREIGN KEY (car_id) REFERENCES cars(id) ON DELETE CASCADE
        )
    """)

    # إنشاء مستخدم admin
    cursor.execute("SELECT id FROM users WHERE username = 'admin'")
    if not cursor.fetchone():
        hashed_pw = hash_password("admin123")
        cursor.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
                       ("admin", hashed_pw, "admin"))
        print("✅ تم إنشاء حساب admin: admin / admin123")

    # إنشاء مستخدم admin2
    cursor.execute("SELECT id FROM users WHERE username = 'admin2'")
    if not cursor.fetchone():
        hashed_pw = hash_password("123")
        cursor.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
            ("admin2", hashed_pw, "admin")
        )
        print("✅ تم إنشاء حساب admin2: admin2 / 123")

    conn.commit()
    conn.close()
    print("✅ تم إنشاء قاعدة البيانات بنجاح!")


if __name__ == "__main__":
    try:
        import bcrypt
    except ImportError:
        print("يرجى تثبيت bcrypt: pip install bcrypt")
        exit(1)
    create_database()
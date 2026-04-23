from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
from datetime import datetime, timedelta
import sqlite3
import bcrypt
import jwt
from contextlib import contextmanager

DATABASE_PATH = "trip_tracker.db"


@contextmanager
def get_db():
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def migrate_db():
    with get_db() as conn:
        cursor = conn.cursor()
        # trips columns
        for col in ["start_location", "end_location", "garage_location"]:
            try:
                cursor.execute(f"ALTER TABLE trips ADD COLUMN {col} TEXT DEFAULT ''")
            except Exception:
                pass
        # cars new columns
        for col in [
            ("car_type", "TEXT DEFAULT ''"),
            ("code", "TEXT DEFAULT ''"),
            ("chassis_number", "TEXT DEFAULT ''"),
            ("brand", "TEXT DEFAULT ''"),
            ("manufacture_year", "TEXT DEFAULT ''"),
            ("vehicle_license_expiry", "TEXT DEFAULT ''"),
            ("project", "TEXT DEFAULT ''"),
            ("branch", "TEXT DEFAULT ''"),
            ("equipment_category", "TEXT DEFAULT ''"),
        ]:
            try:
                cursor.execute(f"ALTER TABLE cars ADD COLUMN {col[0]} {col[1]}")
            except Exception:
                pass
        # drivers new columns
        for col in [
            ("national_id", "TEXT DEFAULT ''"),
            ("birth_date", "TEXT DEFAULT ''"),
            ("driver_license_expiry", "TEXT DEFAULT ''"),
            ("vehicle_license_expiry", "TEXT DEFAULT ''"),
            ("staff_number", "TEXT DEFAULT ''"),
        ]:
            try:
                cursor.execute(f"ALTER TABLE drivers ADD COLUMN {col[0]} {col[1]}")
            except Exception:
                pass
        # workshop_records table
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


SECRET_KEY = "your-super-secret-key-change-this-in-production-2024"
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 480


def create_access_token(data: dict) -> str:
    to_encode = data.copy()
    expire = datetime.utcnow() + timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)


def verify_token(token: str) -> dict:
    try:
        return jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token has expired")
    except jwt.InvalidTokenError:
        raise HTTPException(status_code=401, detail="Invalid token")


class LoginRequest(BaseModel):
    username: str
    password: str


class UserResponse(BaseModel):
    id: int
    username: str
    role: str
    driver_id: Optional[int] = None


class LoginResponse(BaseModel):
    access_token: str
    token_type: str
    user: UserResponse


class DriverCreate(BaseModel):
    name: str
    phone: str
    status: str = "active"
    national_id: Optional[str] = ""
    birth_date: Optional[str] = ""
    driver_license_expiry: Optional[str] = ""
    staff_number: Optional[str] = ""


class DriverUpdate(BaseModel):
    name: str
    phone: str
    status: str
    national_id: Optional[str] = ""
    birth_date: Optional[str] = ""
    driver_license_expiry: Optional[str] = ""
    staff_number: Optional[str] = ""


class DriverResponse(BaseModel):
    id: int
    name: str
    phone: str
    status: str
    user_id: Optional[int] = None
    username: Optional[str] = None
    national_id: Optional[str] = None
    birth_date: Optional[str] = None
    driver_license_expiry: Optional[str] = None
    staff_number: Optional[str] = None


class CarCreate(BaseModel):
    plate: str
    model: str
    status: str = "available"
    car_type: Optional[str] = ""
    code: Optional[str] = ""
    chassis_number: Optional[str] = ""
    brand: Optional[str] = ""
    manufacture_year: Optional[str] = ""
    vehicle_license_expiry: Optional[str] = ""
    project: Optional[str] = ""
    branch: Optional[str] = ""
    equipment_category: Optional[str] = ""


class CarUpdate(BaseModel):
    plate: str
    model: str
    status: str
    car_type: Optional[str] = ""
    code: Optional[str] = ""
    chassis_number: Optional[str] = ""
    brand: Optional[str] = ""
    manufacture_year: Optional[str] = ""
    vehicle_license_expiry: Optional[str] = ""
    project: Optional[str] = ""
    branch: Optional[str] = ""
    equipment_category: Optional[str] = ""


class CarResponse(BaseModel):
    id: int
    plate: str
    model: str
    status: str
    car_type: Optional[str] = None
    code: Optional[str] = None
    chassis_number: Optional[str] = None
    brand: Optional[str] = None
    manufacture_year: Optional[str] = None
    vehicle_license_expiry: Optional[str] = None
    project: Optional[str] = None
    branch: Optional[str] = None
    equipment_category: Optional[str] = None


class PermissionCreate(BaseModel):
    driver_id: int
    car_id: int


class PermissionResponse(BaseModel):
    id: int
    driver_id: int
    car_id: int


class TripStart(BaseModel):
    driver_id: int
    car_id: int
    start_odometer: float
    start_location: Optional[str] = ""


class TripEnd(BaseModel):
    trip_id: int
    end_odometer: float
    end_location: Optional[str] = ""
    notes: Optional[str] = ""


class TripResponse(BaseModel):
    id: int
    driver_id: int
    car_id: int
    start_time: Optional[str]
    end_time: Optional[str]
    start_odometer: float
    end_odometer: Optional[float]
    start_location: Optional[str]
    end_location: Optional[str]
    notes: Optional[str]


class UserCreate(BaseModel):
    username: str
    password: str
    role: str


class WorkshopCreate(BaseModel):
    driver_id: int
    type: str
    quantity: float
    price: float
    notes: Optional[str] = ""


class WorkshopResponse(BaseModel):
    id: int
    driver_id: int
    type: str
    quantity: float
    price: float
    notes: Optional[str]
    created_at: str


app = FastAPI(title="Fleet Management System")


@app.on_event("startup")
async def startup_event():
    migrate_db()


@app.get("/")
async def serve_frontend():
    return FileResponse("index.html")


app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

security = HTTPBearer(auto_error=False)


async def get_current_user(credentials: Optional[HTTPAuthorizationCredentials] = Depends(security)) -> Dict[str, Any]:
    if not credentials:
        raise HTTPException(status_code=401, detail="Not authenticated")
    token = credentials.credentials
    payload = verify_token(token)
    if "user_id" not in payload or "role" not in payload:
        raise HTTPException(status_code=401, detail="Invalid token payload")
    result = {
        "user_id": payload.get("user_id"),
        "role": payload.get("role"),
        "username": payload.get("username")
    }
    if result["role"] == "driver":
        with get_db() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM drivers WHERE user_id = ?", (result["user_id"],))
            driver = cursor.fetchone()
            if driver:
                result["driver_id"] = driver["id"]
    return result


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')


def verify_password(plain_password: str, hashed_password: str) -> bool:
    return bcrypt.checkpw(plain_password.encode('utf-8'), hashed_password.encode('utf-8'))


@app.post("/login", response_model=LoginResponse)
async def login(login_data: LoginRequest):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, password, role FROM users WHERE username = ?", (login_data.username,))
        user = cursor.fetchone()
        if not user or not verify_password(login_data.password, user["password"]):
            raise HTTPException(status_code=401, detail="اسم المستخدم أو كلمة المرور غير صحيحة")
        driver_id = None
        if user["role"] == "driver":
            cursor.execute("SELECT id FROM drivers WHERE user_id = ?", (user["id"],))
            driver = cursor.fetchone()
            driver_id = driver["id"] if driver else None
        token = create_access_token({
            "user_id": user["id"],
            "role": user["role"],
            "username": user["username"]
        })
        return LoginResponse(
            access_token=token,
            token_type="bearer",
            user=UserResponse(id=user["id"], username=user["username"], role=user["role"], driver_id=driver_id)
        )


@app.get("/drivers", response_model=List[DriverResponse])
async def get_drivers(current_user: dict = Depends(get_current_user)):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT d.id, d.name, d.phone, d.status, d.user_id, u.username,
                   d.national_id, d.birth_date, d.driver_license_expiry, d.staff_number
            FROM drivers d
            LEFT JOIN users u ON d.user_id = u.id
            ORDER BY d.id DESC
        """)
        return [dict(d) for d in cursor.fetchall()]


@app.post("/drivers", response_model=DriverResponse)
async def create_driver(driver: DriverCreate, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        username_base = driver.name.replace(" ", "").lower()
        final_username = username_base
        counter = 1
        while True:
            cursor.execute("SELECT id FROM users WHERE username = ?", (final_username,))
            if not cursor.fetchone():
                break
            final_username = f"{username_base}{counter}"
            counter += 1
        default_password = f"pass{final_username}"
        hashed_pw = hash_password(default_password)
        cursor.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", (final_username, hashed_pw, "driver"))
        user_id = cursor.lastrowid
        cursor.execute("""
            INSERT INTO drivers (name, phone, status, user_id, national_id, birth_date, driver_license_expiry, staff_number)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (driver.name, driver.phone, driver.status, user_id,
              driver.national_id or "", driver.birth_date or "",
              driver.driver_license_expiry or "", driver.staff_number or ""))
        driver_id = cursor.lastrowid
        return {"id": driver_id, "name": driver.name, "phone": driver.phone, "status": driver.status,
                "user_id": user_id, "username": final_username,
                "national_id": driver.national_id, "birth_date": driver.birth_date,
                "driver_license_expiry": driver.driver_license_expiry, "staff_number": driver.staff_number}


@app.put("/drivers/{driver_id}", response_model=DriverResponse)
async def update_driver(driver_id: int, driver: DriverUpdate, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE drivers SET name=?, phone=?, status=?, national_id=?, birth_date=?,
            driver_license_expiry=?, staff_number=? WHERE id=?
        """, (driver.name, driver.phone, driver.status,
              driver.national_id or "", driver.birth_date or "",
              driver.driver_license_expiry or "", driver.staff_number or "", driver_id))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="السائق غير موجود")
        cursor.execute("SELECT user_id FROM drivers WHERE id = ?", (driver_id,))
        result = cursor.fetchone()
        return {"id": driver_id, "name": driver.name, "phone": driver.phone, "status": driver.status,
                "user_id": result["user_id"] if result else None,
                "national_id": driver.national_id, "birth_date": driver.birth_date,
                "driver_license_expiry": driver.driver_license_expiry, "staff_number": driver.staff_number}


@app.delete("/drivers/{driver_id}")
async def delete_driver(driver_id: int, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT user_id FROM drivers WHERE id = ?", (driver_id,))
        driver = cursor.fetchone()
        cursor.execute("DELETE FROM driver_car_permissions WHERE driver_id = ?", (driver_id,))
        cursor.execute("DELETE FROM trips WHERE driver_id = ?", (driver_id,))
        cursor.execute("DELETE FROM workshop_records WHERE driver_id = ?", (driver_id,))
        cursor.execute("DELETE FROM drivers WHERE id = ?", (driver_id,))
        if driver and driver["user_id"]:
            cursor.execute("DELETE FROM users WHERE id = ?", (driver["user_id"],))
        return {"message": "تم حذف السائق بنجاح"}


@app.get("/cars", response_model=List[CarResponse])
async def get_cars(current_user: dict = Depends(get_current_user)):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, plate, model, status, car_type, code, chassis_number, brand,
                   manufacture_year, vehicle_license_expiry, project, branch, equipment_category
            FROM cars ORDER BY id DESC
        """)
        return [dict(c) for c in cursor.fetchall()]


@app.post("/cars", response_model=CarResponse)
async def create_car(car: CarCreate, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO cars (plate, model, status, car_type, code, chassis_number, brand,
                              manufacture_year, vehicle_license_expiry, project, branch, equipment_category)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (car.plate, car.model, car.status, car.car_type or "", car.code or "",
              car.chassis_number or "", car.brand or "", car.manufacture_year or "",
              car.vehicle_license_expiry or "", car.project or "", car.branch or "",
              car.equipment_category or ""))
        new_id = cursor.lastrowid
        return {"id": new_id, "plate": car.plate, "model": car.model, "status": car.status,
                "car_type": car.car_type, "code": car.code, "chassis_number": car.chassis_number,
                "brand": car.brand, "manufacture_year": car.manufacture_year,
                "vehicle_license_expiry": car.vehicle_license_expiry, "project": car.project,
                "branch": car.branch, "equipment_category": car.equipment_category}


@app.put("/cars/{car_id}", response_model=CarResponse)
async def update_car(car_id: int, car: CarUpdate, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE cars SET plate=?, model=?, status=?, car_type=?, code=?, chassis_number=?,
            brand=?, manufacture_year=?, vehicle_license_expiry=?, project=?, branch=?, equipment_category=?
            WHERE id=?
        """, (car.plate, car.model, car.status, car.car_type or "", car.code or "",
              car.chassis_number or "", car.brand or "", car.manufacture_year or "",
              car.vehicle_license_expiry or "", car.project or "", car.branch or "",
              car.equipment_category or "", car_id))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="المركبة غير موجودة")
        return {"id": car_id, "plate": car.plate, "model": car.model, "status": car.status,
                "car_type": car.car_type, "code": car.code, "chassis_number": car.chassis_number,
                "brand": car.brand, "manufacture_year": car.manufacture_year,
                "vehicle_license_expiry": car.vehicle_license_expiry, "project": car.project,
                "branch": car.branch, "equipment_category": car.equipment_category}


@app.delete("/cars/{car_id}")
async def delete_car(car_id: int, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM driver_car_permissions WHERE car_id = ?", (car_id,))
        cursor.execute("DELETE FROM cars WHERE id = ?", (car_id,))
        return {"message": "تم حذف المركبة بنجاح"}


@app.get("/permissions", response_model=List[PermissionResponse])
async def get_permissions(current_user: dict = Depends(get_current_user)):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, driver_id, car_id FROM driver_car_permissions ORDER BY id DESC")
        return [dict(p) for p in cursor.fetchall()]


@app.post("/permissions", response_model=PermissionResponse)
async def create_permission(perm: PermissionCreate, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM driver_car_permissions WHERE driver_id = ? AND car_id = ?", (perm.driver_id, perm.car_id))
        if cursor.fetchone():
            raise HTTPException(status_code=400, detail="الصلاحية موجودة مسبقاً")
        cursor.execute("INSERT INTO driver_car_permissions (driver_id, car_id) VALUES (?, ?)", (perm.driver_id, perm.car_id))
        return {"id": cursor.lastrowid, "driver_id": perm.driver_id, "car_id": perm.car_id}


@app.delete("/permissions/{permission_id}")
async def delete_permission(permission_id: int, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM driver_car_permissions WHERE id = ?", (permission_id,))
        return {"message": "تم حذف الصلاحية"}


@app.get("/trips", response_model=List[TripResponse])
async def get_trips(current_user: dict = Depends(get_current_user)):
    with get_db() as conn:
        cursor = conn.cursor()
        if current_user["role"] == "admin":
            cursor.execute("SELECT * FROM trips ORDER BY id DESC")
        else:
            driver_id = current_user.get("driver_id")
            if not driver_id:
                return []
            cursor.execute("SELECT * FROM trips WHERE driver_id = ? ORDER BY id DESC", (driver_id,))
        rows = cursor.fetchall()
        result = []
        for t in rows:
            d = dict(t)
            d.setdefault("start_location", None)
            d.setdefault("end_location", None)
            result.append(d)
        return result


@app.post("/trips/start", response_model=TripResponse)
async def start_trip(trip: TripStart, current_user: dict = Depends(get_current_user)):
    if trip.start_odometer < 0:
        raise HTTPException(status_code=400, detail="عداد البداية يجب أن يكون رقماً موجباً")
    with get_db() as conn:
        cursor = conn.cursor()
        if current_user["role"] != "admin":
            driver_id_from_token = current_user.get("driver_id")
            if trip.driver_id != driver_id_from_token:
                raise HTTPException(status_code=403, detail="لا يمكنك بدء رحلة لسائق آخر")
            cursor.execute("SELECT id FROM driver_car_permissions WHERE driver_id = ? AND car_id = ?", (trip.driver_id, trip.car_id))
            if not cursor.fetchone():
                raise HTTPException(status_code=403, detail="غير مصرح لك بقيادة هذه المركبة")
        cursor.execute("SELECT id FROM trips WHERE driver_id = ? AND end_time IS NULL", (trip.driver_id,))
        if cursor.fetchone():
            raise HTTPException(status_code=400, detail="لديك رحلة نشطة بالفعل، أنهها أولاً")
        now = datetime.now().isoformat()
        cursor.execute(
            "INSERT INTO trips (driver_id, car_id, start_time, start_odometer, start_location) VALUES (?, ?, ?, ?, ?)",
            (trip.driver_id, trip.car_id, now, trip.start_odometer, trip.start_location or "")
        )
        trip_id = cursor.lastrowid
        return {"id": trip_id, "driver_id": trip.driver_id, "car_id": trip.car_id,
                "start_time": now, "end_time": None,
                "start_odometer": trip.start_odometer, "end_odometer": None,
                "start_location": trip.start_location or "", "end_location": None, "notes": ""}


@app.post("/trips/end", response_model=TripResponse)
async def end_trip(trip_end: TripEnd, current_user: dict = Depends(get_current_user)):
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT driver_id, start_odometer FROM trips WHERE id = ? AND end_time IS NULL", (trip_end.trip_id,))
        trip = cursor.fetchone()
        if not trip:
            raise HTTPException(status_code=404, detail="الرحلة النشطة غير موجودة")
        if current_user["role"] != "admin":
            if trip["driver_id"] != current_user.get("driver_id"):
                raise HTTPException(status_code=403, detail="لا يمكنك إنهاء رحلة سائق آخر")
        if trip_end.end_odometer <= trip["start_odometer"]:
            raise HTTPException(status_code=400, detail="عداد النهاية يجب أن يكون أكبر من عداد البداية")
        now = datetime.now().isoformat()
        cursor.execute(
            "UPDATE trips SET end_time = ?, end_odometer = ?, end_location = ?, notes = ? WHERE id = ?",
            (now, trip_end.end_odometer, trip_end.end_location or "", trip_end.notes, trip_end.trip_id)
        )
        cursor.execute("SELECT * FROM trips WHERE id = ?", (trip_end.trip_id,))
        row = dict(cursor.fetchone())
        row.setdefault("start_location", None)
        row.setdefault("end_location", None)
        return row


@app.post("/users")
async def create_user(user: UserCreate, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        hashed_pw = hash_password(user.password)
        try:
            cursor.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", (user.username, hashed_pw, user.role))
            user_id = cursor.lastrowid
            if user.role == "driver":
                cursor.execute("INSERT INTO drivers (name, phone, status, user_id) VALUES (?, ?, ?, ?)", (user.username, "", "active", user_id))
            return {"message": f"تم إنشاء المستخدم {user.username}", "user_id": user_id}
        except sqlite3.IntegrityError:
            raise HTTPException(status_code=400, detail="اسم المستخدم موجود مسبقاً")


@app.get("/users/list")
async def get_users_list(current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, role FROM users ORDER BY id DESC")
        return [dict(u) for u in cursor.fetchall()]


@app.delete("/users/{user_id}")
async def delete_user(user_id: int, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin":
        raise HTTPException(status_code=403, detail="Admin access required")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, username FROM users WHERE id = ?", (user_id,))
        user = cursor.fetchone()
        if not user:
            raise HTTPException(status_code=404, detail="المستخدم غير موجود")
        if user["username"] == "admin":
            raise HTTPException(status_code=400, detail="لا يمكن حذف الحساب الرئيسي للمدير")
        cursor.execute("UPDATE drivers SET user_id = NULL WHERE user_id = ?", (user_id,))
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        return {"message": "تم حذف المستخدم بنجاح"}


@app.get("/drivers/{driver_id}/allowed-cars")
async def get_allowed_cars_for_driver(driver_id: int, current_user: dict = Depends(get_current_user)):
    if current_user["role"] != "admin" and current_user.get("driver_id") != driver_id:
        raise HTTPException(status_code=403, detail="Not authorized")
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT c.id, c.plate, c.model, c.status
            FROM cars c
            INNER JOIN driver_car_permissions p ON c.id = p.car_id
            WHERE p.driver_id = ?
        """, (driver_id,))
        return [dict(c) for c in cursor.fetchall()]


@app.post("/workshops", response_model=WorkshopResponse)
async def create_workshop_record(record: WorkshopCreate, current_user: dict = Depends(get_current_user)):
    if record.quantity <= 0:
        raise HTTPException(status_code=400, detail="الكمية يجب أن تكون أكبر من صفر")
    if record.price < 0:
        raise HTTPException(status_code=400, detail="السعر لا يمكن أن يكون سالباً")
    if record.type not in ["oil", "filter", "tire"]:
        raise HTTPException(status_code=400, detail="نوع غير صالح")
    if current_user["role"] != "admin":
        if record.driver_id != current_user.get("driver_id"):
            raise HTTPException(status_code=403, detail="لا يمكنك تسجيل ورشة لسائق آخر")
    now = datetime.now().isoformat()
    with get_db() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO workshop_records (driver_id, type, quantity, price, notes, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            (record.driver_id, record.type, record.quantity, record.price, record.notes or "", now)
        )
        rec_id = cursor.lastrowid
        return {"id": rec_id, "driver_id": record.driver_id, "type": record.type,
                "quantity": record.quantity, "price": record.price,
                "notes": record.notes or "", "created_at": now}


@app.get("/workshops")
async def get_workshop_records(current_user: dict = Depends(get_current_user)):
    with get_db() as conn:
        cursor = conn.cursor()
        if current_user["role"] == "admin":
            cursor.execute("""
                SELECT w.*, d.name as driver_name
                FROM workshop_records w
                LEFT JOIN drivers d ON w.driver_id = d.id
                ORDER BY w.id DESC
            """)
        else:
            driver_id = current_user.get("driver_id")
            if not driver_id:
                return []
            cursor.execute("""
                SELECT w.*, d.name as driver_name
                FROM workshop_records w
                LEFT JOIN drivers d ON w.driver_id = d.id
                WHERE w.driver_id = ?
                ORDER BY w.id DESC
            """, (driver_id,))
        return [dict(r) for r in cursor.fetchall()]


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=True)

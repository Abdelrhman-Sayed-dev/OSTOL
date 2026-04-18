from fastapi import FastAPI, Depends, HTTPException
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
from datetime import datetime, timedelta
import sqlite3, bcrypt, jwt, os
from contextlib import contextmanager

DATABASE_PATH = os.environ.get("DATABASE_PATH", "trip_tracker.db")

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
        cur = conn.cursor()
        for col in ["start_location TEXT","end_location TEXT","garage_location TEXT"]:
            try: cur.execute(f"ALTER TABLE trips ADD COLUMN {col}")
            except: pass
        # New driver profile columns
        for col in [
            "national_id TEXT DEFAULT ''",
            "birth_date TEXT DEFAULT ''",
            "driver_license_expiry TEXT DEFAULT ''",
            "vehicle_license_expiry TEXT DEFAULT ''",
        ]:
            try: cur.execute(f"ALTER TABLE drivers ADD COLUMN {col}")
            except: pass
        cur.execute("""CREATE TABLE IF NOT EXISTS workshop_records(
            id INTEGER PRIMARY KEY AUTOINCREMENT,driver_id INTEGER,type TEXT NOT NULL,
            quantity REAL,price REAL NOT NULL,notes TEXT DEFAULT '',created_at TEXT NOT NULL,
            operation_type TEXT DEFAULT '',vehicle_id INTEGER,odometer_reading REAL,
            description TEXT DEFAULT '',tire_action TEXT DEFAULT '',
            FOREIGN KEY(driver_id) REFERENCES drivers(id),FOREIGN KEY(vehicle_id) REFERENCES cars(id))""")
        for col in ["operation_type TEXT DEFAULT ''","vehicle_id INTEGER","odometer_reading REAL",
                    "description TEXT DEFAULT ''","tire_action TEXT DEFAULT ''","location TEXT DEFAULT ''"]:
            try: cur.execute(f"ALTER TABLE workshop_records ADD COLUMN {col}")
            except: pass
        cur.execute("""CREATE TABLE IF NOT EXISTS emergency_reports(
            id INTEGER PRIMARY KEY AUTOINCREMENT,driver_id INTEGER,car_id INTEGER,
            type TEXT NOT NULL,audio_data TEXT,notes TEXT DEFAULT '',
            created_at TEXT NOT NULL,is_read INTEGER DEFAULT 0,
            location TEXT DEFAULT '',
            is_handled INTEGER DEFAULT 0,
            action_taken TEXT DEFAULT '',
            handled_by TEXT DEFAULT '',
            FOREIGN KEY(driver_id) REFERENCES drivers(id),FOREIGN KEY(car_id) REFERENCES cars(id))""")
        for col in ["location TEXT DEFAULT ''","is_handled INTEGER DEFAULT 0",
                    "action_taken TEXT DEFAULT ''","handled_by TEXT DEFAULT ''"]:
            try: cur.execute(f"ALTER TABLE emergency_reports ADD COLUMN {col}")
            except: pass
        # Standalone garage records — NOT linked to trips
        cur.execute("""CREATE TABLE IF NOT EXISTS garage_records(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER NOT NULL,
            car_id INTEGER,
            location TEXT NOT NULL,
            recorded_at TEXT NOT NULL,
            notes TEXT DEFAULT '')
        """)
        # Global settings table (admin-only config)
        cur.execute("""CREATE TABLE IF NOT EXISTS app_settings(
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at TEXT NOT NULL)
        """)
        # Seed default global price if not set
        for ws_type in ['fuel','oil','filter','tire','battery','belt','other']:
            cur.execute("INSERT OR IGNORE INTO app_settings(key,value,updated_at) VALUES(?,?,?)",
                        (f'price_{ws_type}','0',datetime.now().isoformat()))

SECRET_KEY = os.environ.get("SECRET_KEY", "fleet-secret-key-2024-change-in-production")
ALGORITHM = "HS256"
TOKEN_EXP = 480

def create_token(data):
    d = data.copy()
    d["exp"] = datetime.utcnow() + timedelta(minutes=TOKEN_EXP)
    return jwt.encode(d, SECRET_KEY, algorithm=ALGORITHM)

def verify_token(token):
    try: return jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
    except jwt.ExpiredSignatureError: raise HTTPException(401,"Token expired")
    except: raise HTTPException(401,"Invalid token")

class LoginReq(BaseModel):
    username: str; password: str
class UserResp(BaseModel):
    id:int; username:str; role:str; driver_id:Optional[int]=None
class LoginResp(BaseModel):
    access_token:str; token_type:str; user:UserResp
class DriverCreate(BaseModel):
    name:str
    phone:str
    username:str=""
    password:str=""
    status:str="active"
    national_id:Optional[str]=""
    birth_date:Optional[str]=""
    driver_license_expiry:Optional[str]=""
    vehicle_license_expiry:Optional[str]=""
class DriverUpdate(BaseModel):
    name:str
    phone:str
    status:str
    national_id:Optional[str]=""
    birth_date:Optional[str]=""
    driver_license_expiry:Optional[str]=""
    vehicle_license_expiry:Optional[str]=""
class DriverResp(BaseModel):
    id:int
    name:str
    phone:str
    status:str
    user_id:Optional[int]=None
    username:Optional[str]=None
    national_id:Optional[str]=None
    birth_date:Optional[str]=None
    driver_license_expiry:Optional[str]=None
    vehicle_license_expiry:Optional[str]=None
class CarCreate(BaseModel):
    plate:str; model:str; status:str="available"
class CarUpdate(BaseModel):
    plate:str; model:str; status:str
class CarResp(BaseModel):
    id:int; plate:str; model:str; status:str
class PermCreate(BaseModel):
    driver_id:int; car_id:int
class PermResp(BaseModel):
    id:int; driver_id:int; car_id:int
class TripStart(BaseModel):
    driver_id:int; car_id:int; start_odometer:float
    start_location:str=""

class TripEnd(BaseModel):
    trip_id:int; end_odometer:float; end_location:str=""; notes:Optional[str]=""
class TripResp(BaseModel):
    id:int; driver_id:int; car_id:int; start_time:Optional[str]; end_time:Optional[str]
    start_odometer:float; end_odometer:Optional[float]; start_location:Optional[str]
    end_location:Optional[str]; garage_location:Optional[str]; notes:Optional[str]
class UserCreate(BaseModel):
    username:str; password:str; role:str
class WorkshopCreate(BaseModel):
    driver_id:int; type:str; quantity:Optional[float]=None; price:float
    notes:Optional[str]=""; operation_type:Optional[str]=""
    vehicle_id:Optional[int]=None; odometer_reading:Optional[float]=None
    description:Optional[str]=""; tire_action:Optional[str]=""
    location:Optional[str]=""
class EmergencyCreate(BaseModel):
    driver_id:int; car_id:Optional[int]=None; type:str
    audio_data:str; notes:Optional[str]=""
    location:Optional[str]=""

class EmergencyAction(BaseModel):
    is_handled:bool=True
    action_taken:str=""
    handled_by:str=""

class GarageRecordCreate(BaseModel):
    driver_id:int
    car_id:Optional[int]=None
    location:str
    notes:Optional[str]=""

class GarageRecordResp(BaseModel):
    id:int; driver_id:int; car_id:Optional[int]
    location:str; recorded_at:str; notes:Optional[str]
    driver_name:Optional[str]=None; car_plate:Optional[str]=None

app = FastAPI(title="Fleet Management v2")

@app.on_event("startup")
async def startup(): migrate_db()

@app.get("/")
async def root(): return FileResponse("index.html")

app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_credentials=True,
                   allow_methods=["*"],allow_headers=["*"])
security = HTTPBearer(auto_error=False)

async def get_user(cred:Optional[HTTPAuthorizationCredentials]=Depends(security))->Dict[str,Any]:
    if not cred: raise HTTPException(401,"Not authenticated")
    p = verify_token(cred.credentials)
    r = {"user_id":p.get("user_id"),"role":p.get("role"),"username":p.get("username")}
    if r["role"]=="driver":
        with get_db() as conn:
            c=conn.cursor(); c.execute("SELECT id FROM drivers WHERE user_id=?",(r["user_id"],))
            d=c.fetchone()
            if d: r["driver_id"]=d["id"]
    return r

def hp(pw): return bcrypt.hashpw(pw.encode(),bcrypt.gensalt()).decode()
def vp(plain,hashed): return bcrypt.checkpw(plain.encode(),hashed.encode())

def enrich(d):
    for k in ["start_location","end_location","garage_location"]: d.setdefault(k,None)
    return d

@app.post("/login",response_model=LoginResp)
async def login(data:LoginReq):
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT * FROM users WHERE username=?",(data.username,))
        u=c.fetchone()
        if not u or not vp(data.password,u["password"]): raise HTTPException(401,"بيانات خاطئة")
        did=None
        if u["role"]=="driver":
            c.execute("SELECT id FROM drivers WHERE user_id=?",(u["id"],)); d=c.fetchone(); did=d["id"] if d else None
        tok=create_token({"user_id":u["id"],"role":u["role"],"username":u["username"]})
        return LoginResp(access_token=tok,token_type="bearer",
                         user=UserResp(id=u["id"],username=u["username"],role=u["role"],driver_id=did))

@app.get("/drivers",response_model=List[DriverResp])
async def get_drivers(cu=Depends(get_user)):
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT d.*,u.username FROM drivers d LEFT JOIN users u ON d.user_id=u.id ORDER BY d.id DESC")
        return [dict(r) for r in c.fetchall()]

@app.post("/drivers",response_model=DriverResp)
async def create_driver(driver:DriverCreate,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    if not driver.username: raise HTTPException(400,"اسم المستخدم مطلوب")
    if not driver.password: raise HTTPException(400,"كلمة المرور مطلوبة")
    with get_db() as conn:
        c=conn.cursor()
        # Check username uniqueness
        c.execute("SELECT id FROM users WHERE username=?",(driver.username,))
        if c.fetchone(): raise HTTPException(400,"اسم المستخدم موجود مسبقاً — اختر اسماً آخر")
        c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                  (driver.username, hp(driver.password), "driver"))
        uid=c.lastrowid
        c.execute("""INSERT INTO drivers(name,phone,status,user_id,national_id,birth_date,driver_license_expiry,vehicle_license_expiry)
                     VALUES(?,?,?,?,?,?,?,?)""",
                  (driver.name,driver.phone,driver.status,uid,
                   driver.national_id or "",driver.birth_date or "",
                   driver.driver_license_expiry or "",driver.vehicle_license_expiry or ""))
        did=c.lastrowid
        return {"id":did,"name":driver.name,"phone":driver.phone,"status":driver.status,
                "user_id":uid,"username":driver.username,
                "national_id":driver.national_id,"birth_date":driver.birth_date,
                "driver_license_expiry":driver.driver_license_expiry,
                "vehicle_license_expiry":driver.vehicle_license_expiry}

@app.put("/drivers/{did}",response_model=DriverResp)
async def update_driver(did:int,driver:DriverUpdate,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor()
        c.execute("""UPDATE drivers SET name=?,phone=?,status=?,national_id=?,birth_date=?,
                     driver_license_expiry=?,vehicle_license_expiry=? WHERE id=?""",
                  (driver.name,driver.phone,driver.status,
                   driver.national_id or "",driver.birth_date or "",
                   driver.driver_license_expiry or "",driver.vehicle_license_expiry or "",did))
        if c.rowcount==0: raise HTTPException(404,"السائق غير موجود")
        c.execute("SELECT user_id FROM drivers WHERE id=?",(did,)); r=c.fetchone()
        return {"id":did,"name":driver.name,"phone":driver.phone,"status":driver.status,
                "user_id":r["user_id"] if r else None,
                "national_id":driver.national_id,"birth_date":driver.birth_date,
                "driver_license_expiry":driver.driver_license_expiry,
                "vehicle_license_expiry":driver.vehicle_license_expiry}

@app.put("/drivers/{did}/credentials")
async def update_driver_credentials(did:int,body:dict,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    new_username = (body.get("username") or "").strip()
    new_password = (body.get("password") or "").strip()
    if not new_username: raise HTTPException(400,"اسم المستخدم مطلوب")
    with get_db() as conn:
        c=conn.cursor()
        # Get driver's user_id
        c.execute("SELECT user_id FROM drivers WHERE id=?",(did,))
        drv=c.fetchone()
        if not drv or not drv["user_id"]: raise HTTPException(404,"المستخدم غير موجود")
        uid=drv["user_id"]
        # Check username not taken by another user
        c.execute("SELECT id FROM users WHERE username=? AND id!=?",(new_username,uid))
        if c.fetchone(): raise HTTPException(400,"اسم المستخدم موجود مسبقاً")
        if new_password:
            c.execute("UPDATE users SET username=?,password=? WHERE id=?",
                      (new_username,hp(new_password),uid))
        else:
            c.execute("UPDATE users SET username=? WHERE id=?",(new_username,uid))
        return {"message":"تم تحديث بيانات الدخول","username":new_username}

@app.delete("/drivers/{did}")
async def delete_driver(did:int,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT user_id FROM drivers WHERE id=?",(did,)); drv=c.fetchone()
        for t in ["driver_car_permissions","trips","workshop_records","emergency_reports"]:
            c.execute(f"DELETE FROM {t} WHERE driver_id=?",(did,))
        c.execute("DELETE FROM drivers WHERE id=?",(did,))
        if drv and drv["user_id"]: c.execute("DELETE FROM users WHERE id=?",(drv["user_id"],))
        return {"message":"تم الحذف"}

@app.get("/cars",response_model=List[CarResp])
async def get_cars(cu=Depends(get_user)):
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT * FROM cars ORDER BY id DESC"); return [dict(r) for r in c.fetchall()]

@app.post("/cars",response_model=CarResp)
async def create_car(car:CarCreate,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor(); c.execute("INSERT INTO cars(plate,model,status) VALUES(?,?,?)",(car.plate,car.model,car.status))
        return {"id":c.lastrowid,"plate":car.plate,"model":car.model,"status":car.status}

@app.put("/cars/{cid}",response_model=CarResp)
async def update_car(cid:int,car:CarUpdate,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor(); c.execute("UPDATE cars SET plate=?,model=?,status=? WHERE id=?",(car.plate,car.model,car.status,cid))
        if c.rowcount==0: raise HTTPException(404,"غير موجودة")
        return {"id":cid,"plate":car.plate,"model":car.model,"status":car.status}

@app.delete("/cars/{cid}")
async def delete_car(cid:int,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor(); c.execute("DELETE FROM driver_car_permissions WHERE car_id=?",(cid,)); c.execute("DELETE FROM cars WHERE id=?",(cid,))
        return {"message":"تم الحذف"}

@app.get("/permissions",response_model=List[PermResp])
async def get_perms(cu=Depends(get_user)):
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT * FROM driver_car_permissions ORDER BY id DESC"); return [dict(r) for r in c.fetchall()]

@app.post("/permissions",response_model=PermResp)
async def create_perm(perm:PermCreate,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT id FROM driver_car_permissions WHERE driver_id=? AND car_id=?",(perm.driver_id,perm.car_id))
        if c.fetchone(): raise HTTPException(400,"الصلاحية موجودة مسبقاً")
        c.execute("INSERT INTO driver_car_permissions(driver_id,car_id) VALUES(?,?)",(perm.driver_id,perm.car_id))
        return {"id":c.lastrowid,"driver_id":perm.driver_id,"car_id":perm.car_id}

@app.delete("/permissions/{pid}")
async def del_perm(pid:int,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn: conn.cursor().execute("DELETE FROM driver_car_permissions WHERE id=?",(pid,)); return {"ok":True}

@app.get("/trips",response_model=List[TripResp])
async def get_trips(cu=Depends(get_user)):
    with get_db() as conn:
        c=conn.cursor()
        if cu["role"]=="admin": c.execute("SELECT * FROM trips ORDER BY id DESC")
        else:
            did=cu.get("driver_id")
            if not did: return []
            c.execute("SELECT * FROM trips WHERE driver_id=? ORDER BY id DESC",(did,))
        return [enrich(dict(r)) for r in c.fetchall()]

@app.post("/trips/start",response_model=TripResp)
async def start_trip(trip:TripStart,cu=Depends(get_user)):
    if not trip.start_location: raise HTTPException(400,"موقع البداية مطلوب")
    if trip.start_odometer<0: raise HTTPException(400,"العداد يجب أن يكون موجباً")
    with get_db() as conn:
        c=conn.cursor()
        if cu["role"]!="admin":
            if trip.driver_id!=cu.get("driver_id"): raise HTTPException(403,"غير مصرح")
            c.execute("SELECT id FROM driver_car_permissions WHERE driver_id=? AND car_id=?",(trip.driver_id,trip.car_id))
            if not c.fetchone(): raise HTTPException(403,"غير مصرح بقيادة هذه المركبة")
        c.execute("SELECT id FROM trips WHERE driver_id=? AND end_time IS NULL",(trip.driver_id,))
        if c.fetchone(): raise HTTPException(400,"لديك رحلة نشطة، أنهها أولاً")
        now=datetime.now().isoformat()
        c.execute("INSERT INTO trips(driver_id,car_id,start_time,start_odometer,start_location,garage_location) VALUES(?,?,?,?,?,?)",
                  (trip.driver_id,trip.car_id,now,trip.start_odometer,trip.start_location,""))
        return {"id":c.lastrowid,"driver_id":trip.driver_id,"car_id":trip.car_id,"start_time":now,"end_time":None,
                "start_odometer":trip.start_odometer,"end_odometer":None,"start_location":trip.start_location,
                "end_location":None,"garage_location":"","notes":""}

@app.post("/trips/end",response_model=TripResp)
async def end_trip(te:TripEnd,cu=Depends(get_user)):
    if not te.end_location: raise HTTPException(400,"موقع النهاية مطلوب")
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT driver_id,start_odometer FROM trips WHERE id=? AND end_time IS NULL",(te.trip_id,))
        trip=c.fetchone()
        if not trip: raise HTTPException(404,"الرحلة غير موجودة")
        if cu["role"]!="admin" and trip["driver_id"]!=cu.get("driver_id"): raise HTTPException(403,"غير مصرح")
        if te.end_odometer<=trip["start_odometer"]: raise HTTPException(400,"عداد النهاية يجب أن يكون أكبر من البداية")
        now=datetime.now().isoformat()
        c.execute("UPDATE trips SET end_time=?,end_odometer=?,end_location=?,notes=? WHERE id=?",(now,te.end_odometer,te.end_location,te.notes,te.trip_id))
        c.execute("SELECT * FROM trips WHERE id=?",(te.trip_id,)); return enrich(dict(c.fetchone()))

# ── Garage Records (standalone, not linked to trips) ──────────────────────

@app.post("/garage/record")
async def create_garage_record(rec:GarageRecordCreate,cu=Depends(get_user)):
    if not rec.location.strip():
        raise HTTPException(400,"موقع الجراج مطلوب")
    if cu["role"]!="admin" and rec.driver_id!=cu.get("driver_id"):
        raise HTTPException(403,"غير مصرح")
    now=datetime.now().isoformat()
    with get_db() as conn:
        c=conn.cursor()
        c.execute("INSERT INTO garage_records(driver_id,car_id,location,recorded_at,notes) VALUES(?,?,?,?,?)",
                  (rec.driver_id,rec.car_id,rec.location,now,rec.notes or ""))
        rid=c.lastrowid
        plate=None
        if rec.car_id:
            c.execute("SELECT plate FROM cars WHERE id=?",(rec.car_id,))
            r=c.fetchone(); plate=r["plate"] if r else None
        c.execute("SELECT name FROM drivers WHERE id=?",(rec.driver_id,))
        dr=c.fetchone()
        return {"id":rid,"driver_id":rec.driver_id,"car_id":rec.car_id,
                "location":rec.location,"recorded_at":now,"notes":rec.notes or "",
                "driver_name":dr["name"] if dr else None,"car_plate":plate}

@app.get("/garage/records")
async def get_garage_records(cu=Depends(get_user)):
    with get_db() as conn:
        c=conn.cursor()
        q="""SELECT g.*,d.name as driver_name,c.plate as car_plate
             FROM garage_records g
             LEFT JOIN drivers d ON g.driver_id=d.id
             LEFT JOIN cars c ON g.car_id=c.id"""
        if cu["role"]=="admin":
            c.execute(q+" ORDER BY g.id DESC")
        else:
            did=cu.get("driver_id")
            if not did: return []
            c.execute(q+" WHERE g.driver_id=? ORDER BY g.id DESC",(did,))
        rows=[]
        for r in c.fetchall():
            d=dict(r); d.setdefault("notes",""); d.setdefault("driver_name",None); d.setdefault("car_plate",None)
            rows.append(d)
        return rows

@app.get("/garage/latest")
async def get_latest_garage_per_driver(cu=Depends(get_user)):
    """Returns the latest garage record per driver — used by admin map"""
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor()
        c.execute("""SELECT g.*,d.name as driver_name,c.plate as car_plate
                     FROM garage_records g
                     INNER JOIN (SELECT driver_id,MAX(id) as max_id FROM garage_records GROUP BY driver_id) latest
                       ON g.id=latest.max_id
                     LEFT JOIN drivers d ON g.driver_id=d.id
                     LEFT JOIN cars c ON g.car_id=c.id
                     ORDER BY g.recorded_at DESC""")
        rows=[]
        for r in c.fetchall():
            d=dict(r); d.setdefault("notes",""); d.setdefault("driver_name",None); d.setdefault("car_plate",None)
            rows.append(d)
        return rows

@app.post("/users")
async def create_user(user:UserCreate,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor()
        try:
            c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",(user.username,hp(user.password),user.role))
            uid=c.lastrowid
            if user.role=="driver": c.execute("INSERT INTO drivers(name,phone,status,user_id) VALUES(?,?,?,?)",(user.username,"","active",uid))
            return {"message":f"تم إنشاء {user.username}","user_id":uid}
        except sqlite3.IntegrityError: raise HTTPException(400,"اسم المستخدم موجود مسبقاً")

@app.get("/users/list")
async def list_users(cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT id,username,role FROM users ORDER BY id DESC"); return [dict(r) for r in c.fetchall()]

@app.delete("/users/{uid}")
async def del_user(uid:int,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT id,username FROM users WHERE id=?",(uid,)); u=c.fetchone()
        if not u: raise HTTPException(404,"غير موجود")
        if u["username"]=="admin": raise HTTPException(400,"لا يمكن حذف الحساب الرئيسي")
        c.execute("UPDATE drivers SET user_id=NULL WHERE user_id=?",(uid,)); c.execute("DELETE FROM users WHERE id=?",(uid,))
        return {"message":"تم الحذف"}

@app.get("/drivers/{did}/allowed-cars")
async def allowed_cars(did:int,cu=Depends(get_user)):
    if cu["role"]!="admin" and cu.get("driver_id")!=did: raise HTTPException(403,"غير مصرح")
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT c.* FROM cars c INNER JOIN driver_car_permissions p ON c.id=p.car_id WHERE p.driver_id=?",(did,))
        return [dict(r) for r in c.fetchall()]

@app.post("/trips/smart")
async def smart_trip(body:dict,cu=Depends(get_user)):
    """Auto-end active trip + start new one in a single call"""
    driver_id = cu.get("driver_id") if cu["role"]=="driver" else body.get("driver_id")
    if not driver_id: raise HTTPException(400,"driver_id مطلوب")
    car_id     = body.get("car_id")
    start_odo  = body.get("start_odometer",0)
    end_odo    = body.get("end_odometer")
    location   = body.get("location","")
    notes      = body.get("notes","")
    now = datetime.now().isoformat()
    with get_db() as conn:
        c = conn.cursor()
        # End any active trip for this driver
        c.execute("SELECT id,start_odometer FROM trips WHERE driver_id=? AND end_time IS NULL",(driver_id,))
        active = c.fetchone()
        if active:
            eo = end_odo if end_odo and end_odo > active["start_odometer"] else active["start_odometer"]+1
            c.execute("UPDATE trips SET end_time=?,end_odometer=?,end_location=?,notes=? WHERE id=?",
                      (now,eo,location,notes,active["id"]))
        # Validate car permission
        if cu["role"]!="admin":
            c.execute("SELECT id FROM driver_car_permissions WHERE driver_id=? AND car_id=?",(driver_id,car_id))
            if not c.fetchone(): raise HTTPException(403,"غير مصرح بهذه المركبة")
        # Start new trip
        c.execute("INSERT INTO trips(driver_id,car_id,start_time,start_odometer,start_location,garage_location) VALUES(?,?,?,?,?,'')",
                  (driver_id,car_id,now,start_odo,location))
        return {"id":c.lastrowid,"message":"تم إنهاء الرحلة السابقة وبدء رحلة جديدة" if active else "تم بدء رحلة جديدة",
                "ended_trip_id":active["id"] if active else None}

@app.post("/trips/end-current")
async def end_current_trip(body:dict,cu=Depends(get_user)):
    """End active trip for the current driver"""
    driver_id = cu.get("driver_id") if cu["role"]=="driver" else body.get("driver_id")
    if not driver_id: raise HTTPException(400,"driver_id مطلوب")
    end_odo  = body.get("end_odometer")
    location = body.get("location","")
    notes    = body.get("notes","")
    now = datetime.now().isoformat()
    with get_db() as conn:
        c = conn.cursor()
        c.execute("SELECT id,start_odometer FROM trips WHERE driver_id=? AND end_time IS NULL",(driver_id,))
        active = c.fetchone()
        if not active: raise HTTPException(404,"لا توجد رحلة نشطة")
        if end_odo is None or end_odo <= active["start_odometer"]:
            raise HTTPException(400,"عداد النهاية يجب أن يكون أكبر من البداية")
        c.execute("UPDATE trips SET end_time=?,end_odometer=?,end_location=?,notes=? WHERE id=?",
                  (now,end_odo,location,notes,active["id"]))
        c.execute("SELECT * FROM trips WHERE id=?",(active["id"],))
        return enrich(dict(c.fetchone()))

@app.post("/workshops")
async def create_workshop(rec:WorkshopCreate,cu=Depends(get_user)):
    if rec.type not in ["oil","filter","tire","battery","belt","other","fuel"]: raise HTTPException(400,"نوع غير صالح")
    if cu["role"]!="admin" and rec.driver_id!=cu.get("driver_id"): raise HTTPException(403,"غير مصرح")
    # For drivers, use per-type price set by admin (driver cannot override)
    final_price = rec.price if cu["role"]=="admin" else 0.0
    if cu["role"]=="driver":
        with get_db() as conn:
            c=conn.cursor()
            c.execute("SELECT value FROM app_settings WHERE key=?", (f'price_{rec.type}',))
            row=c.fetchone()
            final_price = float(row["value"]) if row else 0.0
    if final_price<0: raise HTTPException(400,"السعر لا يمكن أن يكون سالباً")
    now=datetime.now().isoformat()
    with get_db() as conn:
        c=conn.cursor()
        c.execute("""INSERT INTO workshop_records
                     (driver_id,type,quantity,price,notes,created_at,
                      operation_type,vehicle_id,odometer_reading,description,tire_action,location)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (rec.driver_id,rec.type,rec.quantity,final_price,rec.notes or "",now,
                   rec.operation_type or "",rec.vehicle_id,rec.odometer_reading,
                   rec.description or "",rec.tire_action or "",rec.location or ""))
        rid=c.lastrowid; vp=None
        if rec.vehicle_id:
            c.execute("SELECT plate FROM cars WHERE id=?",(rec.vehicle_id,)); cv=c.fetchone(); vp=cv["plate"] if cv else None
        return {"id":rid,"driver_id":rec.driver_id,"type":rec.type,"quantity":rec.quantity,"price":rec.price,
                "notes":rec.notes or "","created_at":now,"operation_type":rec.operation_type or "",
                "vehicle_id":rec.vehicle_id,"odometer_reading":rec.odometer_reading,"description":rec.description or "",
                "tire_action":rec.tire_action or "","vehicle_plate":vp,"location":rec.location or ""}

@app.get("/workshops")
async def get_workshops(cu=Depends(get_user)):
    with get_db() as conn:
        c=conn.cursor(); q="SELECT w.*,d.name as driver_name,c.plate as vehicle_plate FROM workshop_records w LEFT JOIN drivers d ON w.driver_id=d.id LEFT JOIN cars c ON w.vehicle_id=c.id"
        if cu["role"]=="admin": c.execute(q+" ORDER BY w.id DESC")
        else:
            did=cu.get("driver_id")
            if not did: return []
            c.execute(q+" WHERE w.driver_id=? ORDER BY w.id DESC",(did,))
        rows=[]
        for r in c.fetchall():
            d=dict(r)
            for k in ["operation_type","vehicle_id","odometer_reading","description","vehicle_plate","tire_action","location"]: d.setdefault(k,None)
            rows.append(d)
        return rows

@app.post("/emergency/report")
async def create_emergency(rep:EmergencyCreate,cu=Depends(get_user)):
    if rep.type not in ("emergency","accident"): raise HTTPException(400,"نوع غير صالح")
    if cu["role"]!="admin" and rep.driver_id!=cu.get("driver_id"): raise HTTPException(403,"غير مصرح")
    now=datetime.now().isoformat()
    with get_db() as conn:
        c=conn.cursor()
        c.execute("""INSERT INTO emergency_reports
                     (driver_id,car_id,type,audio_data,notes,created_at,is_read,location)
                     VALUES(?,?,?,?,?,?,0,?)""",
                  (rep.driver_id,rep.car_id,rep.type,rep.audio_data,rep.notes or "",now,rep.location or ""))
        return {"id":c.lastrowid,"created_at":now,"type":rep.type}

@app.get("/emergency/reports")
async def get_emergency_reports(cu=Depends(get_user)):
    with get_db() as conn:
        c=conn.cursor(); q="SELECT e.*,d.name as driver_name,c.plate as car_plate FROM emergency_reports e LEFT JOIN drivers d ON e.driver_id=d.id LEFT JOIN cars c ON e.car_id=c.id"
        if cu["role"]=="admin": c.execute(q+" ORDER BY e.id DESC")
        else:
            did=cu.get("driver_id")
            if not did: return []
            c.execute(q+" WHERE e.driver_id=? ORDER BY e.id DESC",(did,))
        rows=[]
        for r in c.fetchall():
            d=dict(r)
            d["is_read"]    =bool(d.get("is_read",0))
            d["is_handled"] =bool(d.get("is_handled",0))
            d.setdefault("driver_name",None); d.setdefault("car_plate",None)
            d.setdefault("location",""); d.setdefault("action_taken",""); d.setdefault("handled_by","")
            rows.append(d)
        return rows

@app.put("/emergency/reports/{rid}/read")
async def mark_read(rid:int,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn: conn.cursor().execute("UPDATE emergency_reports SET is_read=1 WHERE id=?",(rid,)); return {"ok":True}

@app.put("/emergency/reports/{rid}/action")
async def update_emergency_action(rid:int,body:EmergencyAction,cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    admin_name = cu.get("username","admin")
    with get_db() as conn:
        c=conn.cursor()
        c.execute("""UPDATE emergency_reports
                     SET is_handled=?,action_taken=?,handled_by=?,is_read=1
                     WHERE id=?""",
                  (1 if body.is_handled else 0, body.action_taken, body.handled_by or admin_name, rid))
        return {"ok":True,"handled_by":body.handled_by or admin_name}

@app.get("/emergency/unread-count")
async def unread_count(cu=Depends(get_user)):
    if cu["role"]!="admin": return {"count":0}
    with get_db() as conn:
        c=conn.cursor(); c.execute("SELECT COUNT(*) as cnt FROM emergency_reports WHERE is_read=0"); return {"count":c.fetchone()["cnt"]}

# ── App Settings (Global Config) ─────────────────────────────────

@app.get("/settings/prices")
async def get_prices(cu=Depends(get_user)):
    with get_db() as conn:
        c=conn.cursor()
        c.execute("SELECT key,value FROM app_settings WHERE key LIKE 'price_%'")
        rows=c.fetchall()
        return {r["key"].replace("price_",""):float(r["value"]) for r in rows}

@app.put("/settings/prices")
async def set_prices(body:dict, cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    valid_types=["fuel","oil","filter","tire","battery","belt","other"]
    now=datetime.now().isoformat()
    saved={}
    with get_db() as conn:
        c=conn.cursor()
        for ws_type,price in body.items():
            if ws_type not in valid_types: continue
            if not isinstance(price,(int,float)) or price<0: continue
            c.execute("INSERT OR REPLACE INTO app_settings(key,value,updated_at) VALUES(?,?,?)",
                      (f'price_{ws_type}',str(float(price)),now))
            saved[ws_type]=float(price)
    return {"saved":saved}

@app.post("/admin/reset-data")
async def reset_all_data(cu=Depends(get_user)):
    if cu["role"]!="admin": raise HTTPException(403,"Admin only")
    with get_db() as conn:
        c=conn.cursor()
        for tbl in ["emergency_reports","garage_records","workshop_records",
                    "trips","driver_car_permissions","drivers"]:
            c.execute(f"DELETE FROM {tbl}")
        c.execute("DELETE FROM users WHERE role != 'admin'")
        return {"message":"تم مسح جميع البيانات التشغيلية بنجاح"}

if __name__=="__main__":
    import uvicorn; uvicorn.run(app,host="127.0.0.1",port=8000,reload=True)
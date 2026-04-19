"""seed_data.py — wipes old data, imports 15 drivers & cars from Excel"""
import sqlite3, bcrypt, os
from datetime import datetime

DB = os.environ.get("DATABASE_PATH", "trip_tracker.db")
def hp(pw): return bcrypt.hashpw(str(pw).encode(), bcrypt.gensalt()).decode()

conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row
conn.execute("PRAGMA foreign_keys=ON")
c = conn.cursor()

# ══ WIPE ALL OLD DATA ══
print("Wiping old data...")
for tbl in ["driver_car_permissions","workshop_records","trips","emergency_reports",
            "garage_records","workshop_records"]:
    c.execute(f"DELETE FROM {tbl}")
c.execute("DELETE FROM drivers")
c.execute("DELETE FROM cars")
c.execute("DELETE FROM users WHERE username != 'admin'")
conn.commit()
print("✓ Old data cleared")

DATA = [
    {'sector': 'القطاع', 'car_type': 'سيارة ملاكي ( داستر )', 'plate': '528 ي ج ص', 'chassis': '9552776688', 'driver_name': 'عماد سيد طه بيومي', 'fixed_num': '205676', 'phone': '01060616646', 'drv_lic_exp': '2028-02-05', 'car_lic_exp': '2026-05-19', 'username': 'عماد', 'password': '205676'},
    {'sector': 'القطاع', 'car_type': 'سيارة بيك اب كابينة مفردة ( فورد )', 'plate': '382 ي و د', 'chassis': '738217', 'driver_name': 'عبد الرحمن حسن عبد الرحمن', 'fixed_num': '203172', 'phone': '01200315734', 'drv_lic_exp': '2027-10-13', 'car_lic_exp': '2026-12-06', 'username': 'عبد', 'password': '203172'},
    {'sector': 'القطاع', 'car_type': 'سيارة بيك اب كابينة مفردة ( ميتسوبيشي )', 'plate': '671 ل و د', 'chassis': '11199', 'driver_name': 'علاء فؤاد محمد', 'fixed_num': '196987', 'phone': '01121232561', 'drv_lic_exp': '2027-02-13', 'car_lic_exp': '2026-07-05', 'username': 'علاء', 'password': '196987'},
    {'sector': 'مدينة نصر', 'car_type': 'سيارة بيك اب كابينة مزدوجة ( تويوتا )', 'plate': '724 ا ل د', 'chassis': 'MROES12G903025849', 'driver_name': 'محمود صبحي', 'fixed_num': '792223', 'phone': '01003772514', 'drv_lic_exp': '2026-09-01', 'car_lic_exp': '2027-01-12', 'username': 'محمود', 'password': '792223'},
    {'sector': 'مدينة نصر', 'car_type': 'سيارة بيك اب كابينة مفردة ( ميتسوبيشي )', 'plate': '175 ص ق د', 'chassis': '00-4063', 'driver_name': 'تامر رمضان محمد مرسي', 'fixed_num': '205063', 'phone': '01115753157', 'drv_lic_exp': '2028-03-23', 'car_lic_exp': '2026-09-11', 'username': 'تامر', 'password': '205063'},
    {'sector': 'حلوان', 'car_type': 'سيارة بيك اب كابينة مزدوجة ( تويوتا )', 'plate': '854 ف و د', 'chassis': '600651043', 'driver_name': 'طاق محمد فتحي', 'fixed_num': '153728', 'phone': '01155411749', 'drv_lic_exp': '2026-11-08', 'car_lic_exp': '2026-10-16', 'username': 'طاق', 'password': '153728'},
    {'sector': 'صيانة القصور', 'car_type': 'سيارة بيك اب كابينة مزدوجة ( شيفروليه )', 'plate': '764 س ق د', 'chassis': '7128121', 'driver_name': 'شعبان حسين', 'fixed_num': '185590', 'phone': '01202266261', 'drv_lic_exp': '2027-07-30', 'car_lic_exp': '2026-05-15', 'username': 'شعبان', 'password': '185590'},
    {'sector': 'صيانة القصور', 'car_type': 'سيارة بيك اب كابينة مزدوجة ( نيسان )', 'plate': '6715 ص س', 'chassis': '00-4670', 'driver_name': 'عماد ربيع', 'fixed_num': '205675', 'phone': '01024349359', 'drv_lic_exp': '2027-08-01', 'car_lic_exp': '2026-12-09', 'username': 'عماد', 'password': '205675'},
    {'sector': 'المنشات المتميزة', 'car_type': 'سيارة بيك اب كابينة مزدوجة ( تويوتا )', 'plate': '732 ي و د', 'chassis': '3026119', 'driver_name': 'سعيد ابراهيم', 'fixed_num': '178604', 'phone': '01201211526', 'drv_lic_exp': '2028-08-26', 'car_lic_exp': '2026-12-18', 'username': 'سعيد', 'password': '178604'},
    {'sector': 'القاهرة', 'car_type': 'سيارة بيك اب مفردة ( شيفروليه )', 'plate': '4264 ا ج ف', 'chassis': '6160797', 'driver_name': 'محمد فوزي امين', 'fixed_num': '153559', 'phone': '01142038999', 'drv_lic_exp': '2026-10-22', 'car_lic_exp': '2026-04-28', 'username': 'محمد', 'password': '153559'},
    {'sector': 'القاهرة', 'car_type': 'سيارة ملاكي ( لانسر )', 'plate': '9723 ل ق', 'chassis': '701689', 'driver_name': 'ابراهيم عبد الفتاح', 'fixed_num': '199486', 'phone': '01001364517', 'drv_lic_exp': '2026-10-10', 'car_lic_exp': '2026-11-05', 'username': 'ابراهيم', 'password': '199486'},
    {'sector': 'صيانة القصور', 'car_type': 'سيارة بيك اب مفردة', 'plate': '849 ف و د', 'chassis': '7092', 'driver_name': 'مصطفى محمد عبد المحسن', 'fixed_num': '185705', 'phone': '01148795756', 'drv_lic_exp': '2026-07-05', 'car_lic_exp': '2026-10-07', 'username': 'مصطفى', 'password': '185705'},
    {'sector': 'صيانة القصور', 'car_type': 'سيارة قلاب', 'plate': '167 م ق د', 'chassis': '7996', 'driver_name': 'عمرو مجدي', 'fixed_num': '229663', 'phone': '01150381244', 'drv_lic_exp': '2028-09-06', 'car_lic_exp': '2026-09-06', 'username': 'عمرو', 'password': '229663'},
    {'sector': 'مدينة نصر', 'car_type': 'سيارة بيك اب كابينة مزدوجة ( ميتسوبيشي )', 'plate': '674 ج و د', 'chassis': '112304', 'driver_name': 'ايمن فرج فتحي', 'fixed_num': '178236', 'phone': '01211755649', 'drv_lic_exp': '2029-01-26', 'car_lic_exp': '2027-03-27', 'username': 'ايمن', 'password': '178236'},
    {'sector': 'مدينة نصر', 'car_type': 'سيارة بيك اب كابينة مزدوجة ( نيسان )', 'plate': '6658 ف ل', 'chassis': '9072', 'driver_name': 'محمد حمدي فؤاد', 'fixed_num': '791609', 'phone': '01124373036', 'drv_lic_exp': '2027-10-08', 'car_lic_exp': '2026-06-29', 'username': 'محمد', 'password': '791609'},
]

for row in DATA:
    # Insert car
    try:
        c.execute("""INSERT OR IGNORE INTO cars(plate,model,status,chassis,car_license_expiry,sector)
                     VALUES(?,?,?,?,?,?)""",
                  (row["plate"], row["car_type"], "available",
                   row["chassis"], row["car_lic_exp"], row["sector"]))
    except Exception as e: print(f"Car {row['plate']}: {e}")

    # Insert user
    uid = None
    try:
        c.execute("SELECT id FROM users WHERE username=?", (row["username"],))
        ex = c.fetchone()
        if not ex:
            c.execute("INSERT INTO users(username,password,role) VALUES(?,?,?)",
                      (row["username"], hp(row["password"]), "driver"))
            uid = c.lastrowid
        else:
            uid = ex["id"]
    except Exception as e: print(f"User {row['username']}: {e}"); continue

    # Insert driver
    did = None
    try:
        c.execute("""INSERT INTO drivers(name,phone,status,user_id,
                     driver_license_expiry,vehicle_license_expiry,national_id)
                     VALUES(?,?,?,?,?,?,?)""",
                  (row["driver_name"], row["phone"], "active", uid,
                   row["drv_lic_exp"], row["car_lic_exp"], row["fixed_num"]))
        did = c.lastrowid
    except Exception as e: print(f"Driver {row['driver_name']}: {e}"); continue

    # Link driver ↔ car permission
    try:
        c.execute("SELECT id FROM cars WHERE plate=?", (row["plate"],))
        car = c.fetchone()
        if car and did:
            c.execute("INSERT OR IGNORE INTO driver_car_permissions(driver_id,car_id) VALUES(?,?)",
                      (did, car["id"]))
    except Exception as e: print(f"Permission: {e}")

    print(f"✓ {row['driver_name']} | {row['plate']} | user:{row['username']} pw:{row['password']}")

conn.commit()
conn.close()
print(f"\n✅ Done! {len(DATA)} drivers & cars imported.")
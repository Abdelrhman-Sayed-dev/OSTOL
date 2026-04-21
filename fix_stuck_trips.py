"""fix_stuck_trips.py — fixes trips that have garage_location but no end_time"""
import sqlite3, os
from datetime import datetime

DB = os.environ.get("DATABASE_PATH", "trip_tracker.db")
conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row
c = conn.cursor()

# Find all active trips that have garage_location set
c.execute("""
    SELECT id, driver_id, car_id, start_odometer, garage_location
    FROM trips
    WHERE end_time IS NULL
    AND garage_location IS NOT NULL
    AND garage_location != ''
""")
stuck = c.fetchall()
print(f"Found {len(stuck)} stuck trips...")

now = datetime.utcnow().isoformat() + "Z"
fixed = 0
for t in stuck:
    # Try to get time from garage_records for this car
    c.execute("""
        SELECT recorded_at FROM garage_records
        WHERE car_id = ?
        ORDER BY id DESC LIMIT 1
    """, (t['car_id'],))
    gr = c.fetchone()
    end_time = gr['recorded_at'] if gr else now

    c.execute("""
        UPDATE trips
        SET end_time = ?,
            end_location = garage_location
        WHERE id = ?
    """, (end_time, t['id']))
    fixed += 1
    print(f"  - Trip #{t['id']} ended at {end_time} | loc: {t['garage_location']}")

conn.commit()
conn.close()
print(f"\nFixed {fixed} trips")

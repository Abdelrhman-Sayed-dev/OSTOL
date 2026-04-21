"""fix_stuck_trips.py — fixes trips that have garage_location but no end_time"""
import sqlite3, os
DB = os.environ.get("DATABASE_PATH","trip_tracker.db")
conn = sqlite3.connect(DB); conn.row_factory = sqlite3.Row; c = conn.cursor()

c.execute("""
    UPDATE trips
    SET end_time = (
        SELECT recorded_at FROM garage_records
        WHERE garage_records.car_id = trips.car_id
        ORDER BY id DESC LIMIT 1
    ),
    end_location = garage_location
    WHERE end_time IS NULL
    AND garage_location IS NOT NULL AND garage_location != ''
""")
fixed = conn.total_changes
conn.commit(); conn.close()
print(f"✅ Fixed {fixed} stuck trips")

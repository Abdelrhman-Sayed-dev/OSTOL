"""fix_stuck_trips.py — ends active trips that have a garage record"""
import sqlite3, os

DB = os.environ.get("DATABASE_PATH", "trip_tracker.db")
conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row
c = conn.cursor()

# Find active trips where garage_records has an entry for the same car
c.execute("""
    SELECT t.id as trip_id, t.car_id, t.start_odometer,
           g.location as garage_loc, g.recorded_at as garage_time
    FROM trips t
    INNER JOIN garage_records g ON g.car_id = t.car_id
    WHERE t.end_time IS NULL
    ORDER BY g.id DESC
""")
rows = c.fetchall()

# Keep only latest garage per trip
seen = set()
to_fix = []
for r in rows:
    if r['trip_id'] not in seen:
        seen.add(r['trip_id'])
        to_fix.append(r)

print(f"Found {len(to_fix)} trips to fix...")

for r in to_fix:
    c.execute("""
        UPDATE trips
        SET end_time     = ?,
            end_location = ?,
            garage_location = ?
        WHERE id = ?
    """, (r['garage_time'], r['garage_loc'], r['garage_loc'], r['trip_id']))
    print(f"  ✓ Trip #{r['trip_id']} ended | loc:{r['garage_loc']} | time:{r['garage_time']}")

conn.commit()
conn.close()
print(f"\n✅ Fixed {len(to_fix)} trips")

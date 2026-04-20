import sqlite3, os
conn = sqlite3.connect(os.environ.get('DATABASE_PATH', 'trip_tracker.db'))
for tbl in ['trips', 'workshop_records', 'emergency_reports', 'garage_records']:
    conn.execute(f'DELETE FROM {tbl}')
    print(f'✓ {tbl} cleared')
conn.commit()
conn.close()
print('Done!')

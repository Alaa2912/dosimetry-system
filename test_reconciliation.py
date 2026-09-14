import os
import sys
sys.path.insert(0, '/home/spark/osl_system')
import pandas as pd
from app.database import init_db, get_db
from app.services.importer import analyze_reader_file, commit_reader_import

init_db()
conn = get_db()
cursor = conn.cursor()

# Make sure we have 1 org, 1 dept, 1 subscriber with badge XA001, and 1 subscriber with double badge (XA002_OA and XA002_UA)
cursor.execute("INSERT OR IGNORE INTO organizations (id, name, code) VALUES (10, 'مستشفى الفحص', 'TEST-01')")
cursor.execute("INSERT OR IGNORE INTO departments (id, org_id, name, default_frequency) VALUES (20, 10, 'قسم الاختبار', 'MONTHLY')")
cursor.execute("INSERT OR IGNORE INTO subscribers (id, dept_id, name, has_double_badge) VALUES (30, 20, 'د. سمير', 0)")
cursor.execute("INSERT OR IGNORE INTO subscribers (id, dept_id, name, has_double_badge) VALUES (31, 20, 'د. رانية', 1)")

# Assign badges for 2026 Monthly 1
cursor.execute("INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position) VALUES (30, 2026, 'MONTHLY', 1, 'SN_SAMIR', 'CHEST')")
cursor.execute("INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position) VALUES (31, 2026, 'MONTHLY', 1, 'SN_RANIA_OA', 'OA')")
cursor.execute("INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position) VALUES (31, 2026, 'MONTHLY', 1, 'SN_RANIA_UA', 'UA')")
conn.commit()
conn.close()

# Create test file with:
# 1. SN_SAMIR (Single, matched)
# 2. SN_RANIA_OA (Appears TWICE - duplicate!)
# 3. SN_RANIA_UA (Single, matched)
# 4. UNKNOWN_999 (Unassigned / not registered!)
test_file = "/home/spark/osl_system/sample_data/test_recon.csv"
df = pd.DataFrame([
    {"Serial": "SN_SAMIR", "Date": "01/01/2026", "Hp10": 0.25, "Hp07": 0.25},
    {"Serial": "SN_RANIA_OA", "Date": "01/01/2026", "Hp10": 4.00, "Hp07": 4.10},
    {"Serial": "SN_RANIA_OA", "Date": "01/01/2026", "Hp10": 4.50, "Hp07": 4.60}, # Duplicate reading!
    {"Serial": "SN_RANIA_UA", "Date": "01/01/2026", "Hp10": 1.00, "Hp07": 1.00},
    {"Serial": "UNKNOWN_999", "Date": "01/01/2026", "Hp10": 0.80, "Hp07": 0.85}, # Unassigned!
])
df.to_csv(test_file, index=False)

print("Running analyze_reader_file...")
analysis = analyze_reader_file(
    file_path=test_file,
    period_year=2026,
    period_type='MONTHLY',
    period_number=1,
    col_serial="Serial",
    col_read_date="Date",
    col_hp10="Hp10",
    col_hp07="Hp07"
)

print("Summary:", analysis["summary"])
print("Duplicates found:", len(analysis["duplicates"]))
for d in analysis["duplicates"]:
    print(f"  Duplicate serial: {d['serial']} for worker: {d['subscriber_name']} ({d['readings_count']} readings)")

print("Unassigned found:", len(analysis["unassigned"]))
for u in analysis["unassigned"]:
    print(f"  Unassigned serial: {u['serial']} (Hp10={u['hp10']})")

# Now test commit with resolutions:
# Choose the second reading (row index 3) for SN_RANIA_OA
# Route UNKNOWN_999 to unassigned_log
print("Running commit_reader_import with user resolutions...")
commit_res = commit_reader_import(
    file_path=test_file,
    period_year=2026,
    period_type='MONTHLY',
    period_number=1,
    col_serial="Serial",
    col_read_date="Date",
    col_hp10="Hp10",
    col_hp07="Hp07",
    duplicate_resolutions={"SN_RANIA_OA": 3},
    unassigned_resolutions=[{"serial": "UNKNOWN_999", "action": "unassigned_log", "read_date": "01/01/2026", "hp10": 0.80, "hp07": 0.85, "notes": "وضحة مجهولة تم حفظها في السجل"}]
)

print("Commit status:", commit_res["status"])
print("Matched subscribers:", commit_res["total_subscribers_matched"])

conn = get_db()
cursor = conn.cursor()
cursor.execute("SELECT s.name, cd.hp10, cd.hp07, cd.remark FROM calculated_doses cd JOIN subscribers s ON cd.subscriber_id = s.id WHERE cd.period_year=2026")
for row in cursor.fetchall():
    print(f"Worker: {row['name']} -> Hp10: {row['hp10']} | Hp07: {row['hp07']} | Remark: {row['remark']}")

cursor.execute("SELECT * FROM unassigned_readings WHERE serial_number='UNKNOWN_999'")
un_row = cursor.fetchone()
print(f"Unassigned Table record: {un_row['serial_number']} | Notes: {un_row['notes']}")
conn.close()

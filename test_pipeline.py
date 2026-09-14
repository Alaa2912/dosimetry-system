import os
import sys
sys.path.insert(0, '/home/spark/osl_system')

from app.database import init_db, get_db
from app.services.importer import process_reader_import
from app.services.reporter import (
    generate_monthly_excel_report,
    generate_quarterly_excel_report,
    generate_monthly_word_report,
    generate_quarterly_word_report
)
import pandas as pd

# 1. Initialize fresh database
init_db()
conn = get_db()
cursor = conn.cursor()

# 2. Insert Organization & Department
cursor.execute("""
INSERT INTO organizations (code, name, address, contact, classification, agreement_date)
VALUES ('KH-01', 'مستشفى الخالدي الطبي', 'عمان - الدوار الرابع', '06-4644281', 'Medical Field', '01/01/2024')
""")
org_id = cursor.lastrowid

cursor.execute("""
INSERT INTO departments (org_id, name, default_frequency)
VALUES (?, 'قسم القسطرة القلبية (Cath Lab)', 'MONTHLY')
""", (org_id,))
cath_dept_id = cursor.lastrowid

cursor.execute("""
INSERT INTO departments (org_id, name, default_frequency)
VALUES (?, 'قسم الأشعة العامة (General Radiology)', 'QUARTERLY')
""", (org_id,))
rad_dept_id = cursor.lastrowid

# 3. Insert Subscribers
# Dr. Omar wears double badge (OA and UA) in Cath Lab
cursor.execute("""
INSERT INTO subscribers (dept_id, name, national_id, employee_no, job_title, started_date, has_double_badge)
VALUES (?, 'د. عمر خالد (استشاري قسطرة)', '9801020304', 'EMP-101', 'Interventional Cardiologist', '01/01/2024', 1)
""", (cath_dept_id,))
sub1_id = cursor.lastrowid

# Nurse Ahmad wears single badge in Cath Lab
cursor.execute("""
INSERT INTO subscribers (dept_id, name, national_id, employee_no, job_title, started_date, has_double_badge)
VALUES (?, 'أحمد يوسف (فني قسطرة)', '9912030405', 'EMP-102', 'Cath Lab Technologist', '01/03/2024', 0)
""", (cath_dept_id,))
sub2_id = cursor.lastrowid

# Dr. Ola in Radiology (Quarterly)
cursor.execute("""
INSERT INTO subscribers (dept_id, name, national_id, employee_no, job_title, started_date, has_double_badge)
VALUES (?, 'Dr. Ola Alomari', '9850102030', 'EMP-201', 'Radiologist', '21/09/2025', 0)
""", (rad_dept_id,))
sub3_id = cursor.lastrowid

# 4. Assign Badges for Month 12 / 2024
# Sub1: OA = XA02945251I, UA = XA03090321L
cursor.execute("""
INSERT INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
VALUES (?, 2024, 'MONTHLY', 12, 'XA02945251I', 'OA')
""", (sub1_id,))

cursor.execute("""
INSERT INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
VALUES (?, 2024, 'MONTHLY', 12, 'XA03090321L', 'UA')
""", (sub1_id,))

# Sub2: Single badge = XA030903865
cursor.execute("""
INSERT INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
VALUES (?, 2024, 'MONTHLY', 12, 'XA030903865', 'CHEST')
""", (sub2_id,))

# Sub3: Quarterly Q1 / 2026: XA033848307
cursor.execute("""
INSERT INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
VALUES (?, 2026, 'QUARTERLY', 1, 'XA033848307', 'CHEST')
""", (sub3_id,))

conn.commit()
conn.close()

# 5. Create Sample Reader File based on real cathkh12.xls data
sample_csv = "/home/spark/osl_system/sample_data/cath_reader_sample.csv"
df_sample = pd.DataFrame([
    {"Read ID": 1186.0, "Process Number": "cathKhaldi", "Serial Number": "XA02945251I", "Read Date/Time": 45638.4859, "Deep Dose": 5.49, "Shallow Dose": 5.50},
    {"Read ID": 1176.0, "Process Number": "cathKhaldi", "Serial Number": "XA03090321L", "Read Date/Time": 45638.4823, "Deep Dose": 4.78, "Shallow Dose": 4.93},
    {"Read ID": 1183.0, "Process Number": "cathKhaldi", "Serial Number": "XA030903865", "Read Date/Time": 45638.4849, "Deep Dose": 5.23, "Shallow Dose": 5.37},
])
df_sample.to_csv(sample_csv, index=False)

print("Sample reader file created. Now processing import...")

# 6. Run Import & Calculation
res = process_reader_import(
    file_path=sample_csv,
    period_year=2024,
    period_type='MONTHLY',
    period_number=12,
    col_serial="Serial Number",
    col_read_date="Read Date/Time",
    col_hp10="Deep Dose",
    col_hp07="Shallow Dose",
    reception_date="10/12/2024",
    report_date="15/12/2024",
    template_name_to_save="InLight Standard Reader"
)

print("Import Results:", res)

# Verify calculated doses
conn = get_db()
cursor = conn.cursor()
cursor.execute("""
SELECT s.name, s.has_double_badge, cd.hp10, cd.hp07, cd.accumulated_hp10, cd.remark
FROM calculated_doses cd
JOIN subscribers s ON cd.subscriber_id = s.id
""")
for r in cursor.fetchall():
    print(f"Worker: {r['name']} | Double: {r['has_double_badge']} | Hp10: {r['hp10']} | Hp07: {r['hp07']} | Remark: {r['remark']}")

# 7. Generate Excel & Word Reports
os.makedirs("/home/spark/osl_system/output", exist_ok=True)
excel_out = "/home/spark/osl_system/output/Report_Monthly_12_2024.xlsx"
word_out = "/home/spark/osl_system/output/Report_Monthly_12_2024.docx"

generate_monthly_excel_report(cath_dept_id, 2024, 12, excel_out)
generate_monthly_word_report(cath_dept_id, 2024, 12, word_out)

print(f"Generated Excel report: {excel_out} (Exists: {os.path.exists(excel_out)})")
print(f"Generated Word report: {word_out} (Exists: {os.path.exists(word_out)})")

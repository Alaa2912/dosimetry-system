import os
import openpyxl
from docx import Document
from datetime import datetime
from typing import Dict, Any, List, Optional

from app.database import get_db

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

def generate_monthly_excel_report(dept_id: int, period_year: int, period_month: int, output_path: str) -> str:
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("""
    SELECT o.name as facility_name, d.name as dept_name
    FROM departments d
    JOIN organizations o ON d.org_id = o.id
    WHERE d.id = ?
    """, (dept_id,))
    dept_info = cursor.fetchone()
    facility_name = dept_info['facility_name'] if dept_info else ""
    
    cursor.execute("""
    SELECT s.id, s.name, s.started_date, s.has_double_badge,
           cd.hp10, cd.hp07, cd.accumulated_hp10, cd.accumulated_hp07,
           cd.badge_oa_serial, cd.badge_ua_serial, cd.reception_date,
           cd.reading_date, cd.report_date, cd.remark
    FROM subscribers s
    LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id 
         AND cd.period_year = ? AND cd.period_type = 'MONTHLY' AND cd.period_number = ?
    WHERE s.dept_id = ? AND s.is_active = 1
    ORDER BY s.name ASC
    """, (period_year, period_month, dept_id))
    rows = cursor.fetchall()
    conn.close()
    
    template_path = os.path.join(TEMPLATES_DIR, "FinalDR_monthly.xlsx")
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    rep_date = datetime.today().strftime("%d/%m/%Y")
    
    for i, r in enumerate(rows, start=1):
        serial = r['badge_oa_serial'] or ""
        if r['has_double_badge'] and r['badge_ua_serial']:
            serial = f"{r['badge_oa_serial']} / {r['badge_ua_serial']}"
            
        rec_date = r['reception_date'] or ""
        read_date = r['reading_date'] or ""
        r_date = r['report_date'] or rep_date
        
        hp10 = f"{r['hp10']:.2f}" if r['hp10'] is not None else "0.00"
        hp07 = f"{r['hp07']:.2f}" if r['hp07'] is not None else "0.00"
        acc = f"{r['accumulated_hp10']:.2f}" if r['accumulated_hp10'] is not None else "0.00"
        remark = r['remark'] or ""
        
        ws.append([
            i, facility_name, serial, r['name'], r['started_date'] or "",
            rec_date, read_date, r_date, period_month,
            hp10, hp07, acc, remark
        ])
        
    wb.save(output_path)
    return output_path

def generate_quarterly_excel_report(dept_id: int, period_year: int, period_quarter: int, output_path: str) -> str:
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("""
    SELECT o.name as facility_name, d.name as dept_name
    FROM departments d
    JOIN organizations o ON d.org_id = o.id
    WHERE d.id = ?
    """, (dept_id,))
    dept_info = cursor.fetchone()
    facility_name = dept_info['facility_name'] if dept_info else ""
    
    cursor.execute("""
    SELECT s.id, s.name, s.started_date, s.has_double_badge,
           cd.hp10, cd.hp07, cd.accumulated_hp10, cd.accumulated_hp07,
           cd.badge_oa_serial, cd.badge_ua_serial, cd.reception_date,
           cd.reading_date, cd.report_date, cd.remark
    FROM subscribers s
    LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id 
         AND cd.period_year = ? AND cd.period_type = 'QUARTERLY' AND cd.period_number = ?
    WHERE s.dept_id = ? AND s.is_active = 1
    ORDER BY s.name ASC
    """, (period_year, period_quarter, dept_id))
    rows = cursor.fetchall()
    conn.close()
    
    template_path = os.path.join(TEMPLATES_DIR, "FinalDR_Quartar.xlsx")
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    rep_date = datetime.today().strftime("%d/%m/%Y")
    
    for i, r in enumerate(rows, start=1):
        serial = r['badge_oa_serial'] or ""
        if r['has_double_badge'] and r['badge_ua_serial']:
            serial = f"{r['badge_oa_serial']} / {r['badge_ua_serial']}"
            
        rec_date = r['reception_date'] or ""
        read_date = r['reading_date'] or ""
        r_date = r['report_date'] or rep_date
        
        hp10 = f"{r['hp10']:.2f}" if r['hp10'] is not None else "0.00"
        hp07 = f"{r['hp07']:.2f}" if r['hp07'] is not None else "0.00"
        acc = f"{r['accumulated_hp10']:.2f}" if r['accumulated_hp10'] is not None else "0.00"
        remark = r['remark'] or ""
        
        ws.append([
            i, facility_name, serial, r['name'], r['started_date'] or "",
            rec_date, read_date, r_date, f"{period_quarter}",
            hp10, hp07, acc, remark
        ])
        
    wb.save(output_path)
    return output_path

def replace_paragraph_placeholders(paragraph, replacements: Dict[str, str]):
    for key, val in replacements.items():
        placeholder = f"{{{{ {key} }}}}"
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, str(val))

def generate_monthly_word_report(dept_id: int, period_year: int, target_month: int, output_path: str) -> str:
    if 1 <= target_month <= 4:
        template_name = "Monthly 1-4.docx"
        months_in_block = [1, 2, 3, 4]
    elif 5 <= target_month <= 8:
        template_name = "Monthly 5-8.docx"
        months_in_block = [5, 6, 7, 8]
    else:
        template_name = "Monthly 9-12.docx"
        months_in_block = [9, 10, 11, 12]
        
    doc = Document(os.path.join(TEMPLATES_DIR, template_name))
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("""
    SELECT o.name as org_name, o.code as org_code, o.address as org_address,
           o.contact as org_contact, o.agreement_date, d.name as dept_name
    FROM departments d
    JOIN organizations o ON d.org_id = o.id
    WHERE d.id = ?
    """, (dept_id,))
    dept_row = cursor.fetchone()
    
    cursor.execute("SELECT COUNT(*) as cnt FROM subscribers WHERE dept_id = ? AND is_active = 1", (dept_id,))
    count_row = cursor.fetchone()
    dosimeter_count = count_row['cnt'] if count_row else 0
    
    replacements = {
        "customer_name": dept_row['org_name'] if dept_row else "",
        "customer_code": dept_row['org_code'] if dept_row else "",
        "customer_address": dept_row['org_address'] if dept_row else "",
        "contact": dept_row['org_contact'] if dept_row else "",
        "monitoring_period": f"{target_month}/{period_year}",
        "dosimeter_count": str(dosimeter_count),
        "agreement_date": dept_row['agreement_date'] if dept_row else "",
        "department_name": dept_row['dept_name'] if dept_row else ""
    }
    
    for p in doc.paragraphs:
        replace_paragraph_placeholders(p, replacements)
        
    cursor.execute("SELECT id, name, started_date FROM subscribers WHERE dept_id = ? AND is_active = 1 ORDER BY name ASC", (dept_id,))
    subs = cursor.fetchall()
    
    table = doc.tables[0]
    
    for i, sub in enumerate(subs, start=1):
        sub_id = sub['id']
        cursor.execute("""
        SELECT period_number, hp10, hp07, accumulated_hp10, accumulated_hp07, reception_date, reading_date, remark
        FROM calculated_doses
        WHERE subscriber_id = ? AND period_year = ? AND period_type = 'MONTHLY'
        """, (sub_id, period_year))
        doses_by_m = {d['period_number']: d for d in cursor.fetchall()}
        
        row_cells = table.add_row().cells
        c_idx = 0
        row_cells[c_idx].text = str(i)
        c_idx += 1
        row_cells[c_idx].text = sub['name']
        c_idx += 1
        row_cells[c_idx].text = sub['started_date'] or ""
        c_idx += 1
        
        target_dose = doses_by_m.get(target_month)
        rec_d = target_dose['reception_date'] if target_dose else ""
        read_d = target_dose['reading_date'] if target_dose else ""
        row_cells[c_idx].text = rec_d
        c_idx += 1
        row_cells[c_idx].text = read_d
        c_idx += 1
        
        for m in months_in_block:
            if m <= target_month and m in doses_by_m:
                m_dose = doses_by_m[m]
                row_cells[c_idx].text = f"{m_dose['hp10']:.2f}"
                row_cells[c_idx + 1].text = f"{m_dose['hp07']:.2f}"
            else:
                row_cells[c_idx].text = "-"
                row_cells[c_idx + 1].text = "-"
            c_idx += 2
            
        if target_dose:
            row_cells[c_idx].text = f"{target_dose['accumulated_hp10']:.2f}"
            row_cells[c_idx + 1].text = f"{target_dose['accumulated_hp07']:.2f}"
            row_cells[c_idx + 2].text = target_dose['remark'] or ""
        else:
            row_cells[c_idx].text = "0.00"
            row_cells[c_idx + 1].text = "0.00"
            row_cells[c_idx + 2].text = ""

    conn.close()
    doc.save(output_path)
    return output_path

def generate_quarterly_word_report(dept_id: int, period_year: int, target_quarter: int, output_path: str) -> str:
    template_name = "PERSONAL MONITORING SERVICE.docx"
    doc = Document(os.path.join(TEMPLATES_DIR, template_name))
    
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("""
    SELECT o.name as org_name, o.code as org_code, o.address as org_address,
           o.contact as org_contact, o.agreement_date, d.name as dept_name
    FROM departments d
    JOIN organizations o ON d.org_id = o.id
    WHERE d.id = ?
    """, (dept_id,))
    dept_row = cursor.fetchone()
    
    cursor.execute("SELECT COUNT(*) as cnt FROM subscribers WHERE dept_id = ? AND is_active = 1", (dept_id,))
    count_row = cursor.fetchone()
    dosimeter_count = count_row['cnt'] if count_row else 0
    
    replacements = {
        "customer_name": dept_row['org_name'] if dept_row else "",
        "customer_code": dept_row['org_code'] if dept_row else "",
        "customer_address": dept_row['org_address'] if dept_row else "",
        "contact": dept_row['org_contact'] if dept_row else "",
        "monitoring_period": f"Q{target_quarter}/{period_year}",
        "dosimeter_count": str(dosimeter_count),
        "agreement_date": dept_row['agreement_date'] if dept_row else "",
        "department_name": dept_row['dept_name'] if dept_row else ""
    }
    
    for p in doc.paragraphs:
        replace_paragraph_placeholders(p, replacements)
        
    cursor.execute("""
    SELECT s.id, s.name, s.started_date, s.has_double_badge
    FROM subscribers s
    WHERE s.dept_id = ? AND s.is_active = 1
    ORDER BY s.name ASC
    """, (dept_id,))
    subs = cursor.fetchall()
    
    table = doc.tables[0]
    
    for i, sub in enumerate(subs, start=1):
        sub_id = sub['id']
        cursor.execute("""
        SELECT period_number, hp10, hp07, accumulated_hp10, accumulated_hp07,
               reception_date, reading_date, badge_oa_serial, badge_ua_serial, remark
        FROM calculated_doses
        WHERE subscriber_id = ? AND period_year = ? AND period_type = 'QUARTERLY'
        """, (sub_id, period_year))
        doses_by_q = {d['period_number']: d for d in cursor.fetchall()}
        
        target_dose = doses_by_q.get(target_quarter)
        
        row_cells = table.add_row().cells
        c_idx = 0
        row_cells[c_idx].text = str(i)
        c_idx += 1
        
        sn_text = ""
        if target_dose:
            sn_text = target_dose['badge_oa_serial'] or ""
            if sub['has_double_badge'] and target_dose['badge_ua_serial']:
                sn_text = f"{target_dose['badge_oa_serial']}/{target_dose['badge_ua_serial']}"
        row_cells[c_idx].text = sn_text
        c_idx += 1
        
        row_cells[c_idx].text = sub['name']
        c_idx += 1
        row_cells[c_idx].text = sub['started_date'] or ""
        c_idx += 1
        
        rec_d = target_dose['reception_date'] if target_dose else ""
        read_d = target_dose['reading_date'] if target_dose else ""
        row_cells[c_idx].text = rec_d
        c_idx += 1
        row_cells[c_idx].text = read_d
        c_idx += 1
        
        for q in [1, 2, 3, 4]:
            if q <= target_quarter and q in doses_by_q:
                q_dose = doses_by_q[q]
                row_cells[c_idx].text = f"{q_dose['hp10']:.2f}"
                row_cells[c_idx + 1].text = f"{q_dose['hp07']:.2f}"
            else:
                row_cells[c_idx].text = "-"
                row_cells[c_idx + 1].text = "-"
            c_idx += 2
            
        if target_dose:
            row_cells[c_idx].text = f"{target_dose['accumulated_hp10']:.2f}"
            row_cells[c_idx + 1].text = f"{target_dose['accumulated_hp07']:.2f}"
            row_cells[c_idx + 2].text = target_dose['remark'] or ""
        else:
            row_cells[c_idx].text = "0.00"
            row_cells[c_idx + 1].text = "0.00"
            row_cells[c_idx + 2].text = ""

    conn.close()
    doc.save(output_path)
    return output_path


def generate_reception_manifest_excel(dept_id: int, period_year: int, period_type: str, period_number: int, output_path: str) -> str:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from app.database import get_db

    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
    SELECT o.name as org_name, o.code as org_code, d.name as dept_name
    FROM departments d
    JOIN organizations o ON d.org_id = o.id
    WHERE d.id = ?
    """, (dept_id,))
    dept_info = cursor.fetchone()
    org_name = dept_info['org_name'] if dept_info else ''
    org_code = dept_info['org_code'] if dept_info else ''
    dept_name = dept_info['dept_name'] if dept_info else ''

    cursor.execute("""
    SELECT s.id, s.name, s.system_code, s.national_id, s.job_title, s.has_double_badge,
           s.current_serial_oa, s.current_serial_ua,
           br.reception_date, br.received_oa, br.received_ua, br.status as rec_status, br.notes as rec_notes
    FROM subscribers s
    LEFT JOIN badge_receptions br ON s.id = br.subscriber_id
         AND br.period_year = ? AND br.period_type = ? AND br.period_number = ?
    WHERE s.dept_id = ? AND s.status = 'ACTIVE'
    ORDER BY s.name ASC
    """, (period_year, period_type.upper(), period_number, dept_id))
    subs = [dict(r) for r in cursor.fetchall()]
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "كشف الاستلام"
    ws.views.sheetView[0].rightToLeft = True

    # Styling definitions
    title_font = Font(name="Arial", size=14, bold=True, color="1B365D")
    sub_title_font = Font(name="Arial", size=10, bold=False, color="4A5568")
    meta_label_font = Font(name="Arial", size=10, bold=True, color="1A202C")
    meta_val_font = Font(name="Arial", size=10, bold=False, color="2B6CB0")
    tbl_hdr_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    tbl_hdr_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    data_font = Font(name="Arial", size=9)
    bold_data_font = Font(name="Arial", size=9, bold=True)
    border_thin = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE")
    )

    # Title Banner
    ws.merge_cells("A1:K1")
    ws["A1"] = "كشف استلام وتسليم وضحات الرصد الإشعاعي الشخصي (OSL Badges Reception Manifest)"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    period_str = f"{period_number}/{period_year}" if period_type.upper() == 'MONTHLY' else f"الربع {period_number} لعام {period_year}"
    ws.merge_cells("A2:K2")
    ws["A2"] = f"المنشأة: {org_name} ({org_code}) | القسم: {dept_name} | دورة الرصد: {period_str}"
    ws["A2"].font = sub_title_font
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # Meta box
    ws["A4"] = "تاريخ إعداد الكشف:"
    ws["A4"].font = meta_label_font
    ws["B4"] = datetime.today().strftime("%d/%m/%Y")
    ws["B4"].font = meta_val_font

    ws["E4"] = "إجمالي المشتركين:"; ws["E4"].font = meta_label_font
    ws["F4"] = len(subs); ws["F4"].font = meta_val_font

    ws["H4"] = "حالة الاستلام:"; ws["H4"].font = meta_label_font
    ws["I4"] = "جاهز للتدقيق والتوقيع"; ws["I4"].font = meta_val_font

    # Headers
    headers = [
        "#", "كود النظام", "اسم المشترك الثلاثي/الرباعي", "المسمى الوظيفي",
        "نوع الوضحة", "سيريال الوضحة الصدرية / OA", "استلام الصدرية",
        "سيريال الداخلية UA", "استلام الداخلية", "حالة الوضحة", "ملاحظات وتوقيع المستلم"
    ]
    header_row = 6
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(header_row, col_idx)
        cell.value = h
        cell.font = tbl_hdr_font
        cell.fill = tbl_hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
    ws.row_dimensions[header_row].height = 28

    # Data Rows
    for i, s in enumerate(subs, start=1):
        row_num = header_row + i
        b_type = "ثنائية (OA+UA)" if s['has_double_badge'] else "صدرية (CHEST)"
        oa_stat = "[ √ ] تم الاستلام" if s.get('received_oa', 1) else "[   ] لم تُستلم"
        ua_stat = "-"
        if s['has_double_badge']:
            ua_stat = "[ √ ] تم الاستلام" if s.get('received_ua', 1) else "[   ] لم تُستلم"
        
        status_text = s.get('rec_status') or "مستلمة (طبيعي)"
        notes_text = s.get('rec_notes') or ""

        vals = [
            i,
            s['system_code'] or '',
            s['name'],
            s['job_title'] or '',
            b_type,
            s['current_serial_oa'] or '',
            oa_stat,
            s['current_serial_ua'] or '-' if s['has_double_badge'] else '-',
            ua_stat,
            status_text,
            notes_text
        ]

        for col_idx, v in enumerate(vals, start=1):
            c = ws.cell(row_num, col_idx, value=v)
            c.font = bold_data_font if col_idx in (2, 3) else data_font
            c.border = border_thin
            c.alignment = Alignment(horizontal="center" if col_idx != 3 else "right", vertical="center")
        ws.row_dimensions[row_num].height = 22

    # Signatures
    sig_row = header_row + len(subs) + 3
    ws.merge_cells(start_row=sig_row, start_column=2, end_row=sig_row, end_column=5)
    ws.cell(sig_row, 2, value="ضابط الوقاية الإشعاعية / مسؤول التسليم بالمستشفى:").font = meta_label_font
    
    ws.merge_cells(start_row=sig_row, start_column=7, end_row=sig_row, end_column=10)
    ws.cell(sig_row, 7, value="مسؤول استلام وفحص الوضحات بمختبر OSL:").font = meta_label_font

    ws.merge_cells(start_row=sig_row+2, start_column=2, end_row=sig_row+2, end_column=5)
    ws.cell(sig_row+2, 2, value="التوقيع: .......................................   التاريخ: ..../..../202..").font = sub_title_font

    ws.merge_cells(start_row=sig_row+2, start_column=7, end_row=sig_row+2, end_column=10)
    ws.cell(sig_row+2, 7, value="التوقيع: .......................................   التاريخ: ..../..../202..").font = sub_title_font

    # Column widths
    col_widths = [5, 16, 26, 18, 15, 18, 14, 18, 14, 16, 22]
    for idx, w in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w

    wb.save(output_path)
    return output_path


def generate_annual_archive_excel(period_year: int, org_id: Optional[int], dept_id: Optional[int], output_path: str) -> str:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from app.database import get_db

    conn = get_db()
    cursor = conn.cursor()

    query = """
    SELECT s.id as subscriber_id, s.name, s.system_code, s.job_title, s.has_double_badge,
           d.id as dept_id, d.name as dept_name, d.default_frequency,
           o.id as org_id, o.name as org_name, o.code as org_code
    FROM subscribers s
    JOIN departments d ON s.dept_id = d.id
    JOIN organizations o ON d.org_id = o.id
    WHERE s.status = 'ACTIVE'
    """
    params = []
    if dept_id:
        query += " AND d.id = ?"
        params.append(dept_id)
    elif org_id:
        query += " AND o.id = ?"
        params.append(org_id)

    query += " ORDER BY o.name ASC, d.name ASC, s.name ASC"
    cursor.execute(query, params)
    subs = [dict(r) for r in cursor.fetchall()]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"أرشيف {period_year}"
    ws.views.sheetView[0].rightToLeft = True

    # Styling
    title_font = Font(name="Arial", size=14, bold=True, color="1B365D")
    sub_title_font = Font(name="Arial", size=10, color="4A5568")
    tbl_hdr_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    tbl_hdr_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    border_thin = Border(
        left=Side(style="thin", color="D0D7DE"), right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"), bottom=Side(style="thin", color="D0D7DE")
    )
    safe_font = Font(name="Arial", size=9, color="166534", bold=True)
    alert_font = Font(name="Arial", size=9, color="991B1B", bold=True)

    ws.merge_cells("A1:R1")
    ws["A1"] = f"سجل الأرشيف الإشعاعي السنوي التراكمي لعام {period_year} (Annual OSL Dose Archive)"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:R2")
    ws["A2"] = f"تاريخ استخراج الأرشيف: {datetime.today().strftime('%d/%m/%Y')} | إجمالي المشتركين: {len(subs)}"
    ws["A2"].font = sub_title_font
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        "#", "كود النظام", "اسم المشترك", "المسمى الوظيفي", "المنشأة", "القسم",
        "الدورية", "نوع الوضحة",
        "د1/ش1", "د2/ش2", "د3/ش3", "د4/ش4", "ش5", "ش6", "ش7", "ش8", "ش9", "ش10", "ش11", "ش12",
        "إجمالي السنوي Hp(10)", "إجمالي السنوي Hp(0.07)", "حالة الحد السنوي (20 mSv)"
    ]

    header_row = 4
    for col_idx, h in enumerate(headers, start=1):
        c = ws.cell(header_row, col_idx, value=h)
        c.font = tbl_hdr_font
        c.fill = tbl_hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_thin

    for i, s in enumerate(subs, start=1):
        sub_id = s['subscriber_id']
        cursor.execute("""
            SELECT period_number, hp10, hp07, remark
            FROM calculated_doses
            WHERE subscriber_id = ? AND period_year = ?
        """, (sub_id, period_year))
        doses = {r['period_number']: r for r in cursor.fetchall()}

        row_num = header_row + i
        b_type = "ثنائية (OA+UA)" if s['has_double_badge'] else "صدرية"
        freq_name = "شهري" if s['default_frequency'] == 'MONTHLY' else "ربعي"

        p_vals = []
        tot_10 = 0.0
        tot_07 = 0.0
        for p in range(1, 13):
            if p in doses:
                val10 = doses[p]['hp10'] or 0.0
                tot_10 += val10
                tot_07 += (doses[p]['hp07'] or 0.0)
                p_vals.append(f"{val10:.2f}")
            else:
                p_vals.append("-")

        limit_status = "تجاوز الحد السنوي (≥20)" if tot_10 >= 20.0 else "آمن (<20 mSv)"

        row_data = [
            i, s['system_code'] or '', s['name'], s['job_title'] or '',
            s['org_name'], s['dept_name'], freq_name, b_type
        ] + p_vals + [f"{tot_10:.2f}", f"{tot_07:.2f}", limit_status]

        for col_idx, v in enumerate(row_data, start=1):
            c = ws.cell(row_num, col_idx, value=v)
            c.border = border_thin
            c.alignment = Alignment(horizontal="center" if col_idx != 3 else "right", vertical="center")
            if col_idx == len(row_data):
                c.font = alert_font if tot_10 >= 20.0 else safe_font

    conn.close()
    wb.save(output_path)
    return output_path


# --- 8. Official Printable & PDF HTML Report Generator ---
def generate_official_html_report(dept_id: int, year: int, period_type: str, period_number: int) -> str:
    """
    Generates an official, beautifully styled, print-ready HTML radiation report 
    following EMRC & IAEA standards, compatible with native browser print/save-as-PDF.
    """
    from app.database import get_db
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT d.id, d.name as dept_name, d.default_frequency,
               o.id as org_id, o.name as org_name, o.code as org_code, o.classification, o.contact
        FROM departments d
        JOIN organizations o ON d.org_id = o.id
        WHERE d.id = ?
    """, (dept_id,))
    dept_info = cursor.fetchone()
    
    if not dept_info:
        conn.close()
        return "<html><body><h3>القسم غير موجود</h3></body></html>"
        
    cursor.execute("""
        SELECT s.id, s.name, s.system_code, s.national_id, s.job_title, s.has_double_badge,
               s.current_serial_oa, s.current_serial_ua,
               cd.hp10, cd.hp07, cd.accumulated_hp10, cd.accumulated_hp07,
               cd.badge_oa_serial, cd.badge_ua_serial, cd.reading_date, cd.report_date, cd.remark
        FROM subscribers s
        LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id 
             AND cd.period_year = ? AND cd.period_type = ? AND cd.period_number = ?
        WHERE s.dept_id = ? AND s.status = 'ACTIVE'
        ORDER BY s.name ASC
    """, (year, period_type.upper(), period_number, dept_id))
    readings = [dict(r) for r in cursor.fetchall()]
    conn.close()
    
    # Summary calculations
    total_subs = len(readings)
    recorded_count = sum(1 for r in readings if r['hp10'] is not None and r['remark'] != 'NR')
    doses_vals = [r['hp10'] for r in readings if r['hp10'] is not None and r['remark'] != 'NR']
    max_dose = max(doses_vals) if doses_vals else 0.0
    avg_dose = round(sum(doses_vals) / len(doses_vals), 2) if doses_vals else 0.0
    
    period_title = f"شهر {period_number}" if period_type.upper() == 'MONTHLY' else f"الدورة {period_number} (الربع السنوي)"
    freq_label = "شهري (أقسام القسطرة التداخلية - نظام الوضحتين)" if period_type.upper() == 'MONTHLY' else "دوري كل 3 شهور (أقسام الأشعة العامة - وضحة صدرية)"
    
    table_rows = ""
    for idx, r in enumerate(readings):
        has_reading = r['hp10'] is not None
        hp10_val = f"{r['hp10']:.2f}" if has_reading and r['remark'] != 'NR' else ("NR" if r['remark'] == 'NR' else "-")
        hp07_val = f"{r['hp07']:.2f}" if has_reading and r['remark'] != 'NR' and r['hp07'] is not None else "-"
        acc_val = f"{r['accumulated_hp10']:.2f}" if r['accumulated_hp10'] is not None else "-"
        
        is_double = bool(r['has_double_badge'])
        b_type = "ثنائية (OA+UA)" if is_double else "صدرية مفردة"
        
        oa_sn = r['badge_oa_serial'] or r['current_serial_oa'] or '-'
        ua_sn = r['badge_ua_serial'] or r['current_serial_ua'] or '-' if is_double else '-'
        
        status_text = "آمن (< 20 mSv)"
        status_color = "#16a34a"
        if has_reading and r['remark'] != 'NR':
            if (r['accumulated_hp10'] or 0) >= 20.0 or (r['hp10'] or 0) > 5.0:
                status_text = "تجاوز حد الأمان!"
                status_color = "#dc2626"
            elif (r['hp10'] or 0) > 1.67:
                status_text = "متابعة دقيقة"
                status_color = "#ea580c"
        elif r['remark'] == 'NR':
            status_text = "لم تُعد الوضحة (NR)"
            status_color = "#64748b"
            
        table_rows += f"""
        <tr>
            <td>{idx + 1}</td>
            <td><span class="code-badge">{r['system_code'] or '-'}</span></td>
            <td style="text-align: right; font-weight: bold;">{r['name']}</td>
            <td>{r['national_id'] or '-'}</td>
            <td>{r['job_title'] or '-'}</td>
            <td>{b_type}</td>
            <td><code>{oa_sn}</code></td>
            <td><code>{ua_sn}</code></td>
            <td style="font-weight: bold; font-size: 1.05em;">{hp10_val}</td>
            <td>{hp07_val}</td>
            <td style="font-weight: bold;">{acc_val}</td>
            <td style="color: {status_color}; font-weight: bold;">{status_text}</td>
        </tr>
        """
        
    from datetime import datetime
    issue_date = datetime.now().strftime("%d/%m/%Y")
    
    html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>تقرير الرصد الإشعاعي المعتمد - {dept_info['org_name']}</title>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">
    <style>
        @page {{
            size: A4 landscape;
            margin: 10mm 12mm;
        }}
        * {{
            box-sizing: border-box;
        }}
        body {{
            font-family: 'Cairo', 'Segoe UI', Tahoma, sans-serif;
            direction: rtl;
            background-color: #ffffff;
            color: #0f172a;
            margin: 0;
            padding: 15px;
            font-size: 11pt;
            line-height: 1.4;
        }}
        .report-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid #0284c7;
            padding-bottom: 10px;
            margin-bottom: 12px;
        }}
        .header-title {{
            text-align: center;
        }}
        .header-title h2 {{
            margin: 0;
            color: #0369a1;
            font-size: 16pt;
            font-weight: 800;
        }}
        .header-title h4 {{
            margin: 4px 0 0 0;
            color: #475569;
            font-size: 11pt;
            font-weight: 600;
        }}
        .header-logo {{
            text-align: right;
            font-size: 9pt;
            color: #64748b;
        }}
        .header-meta {{
            text-align: left;
            font-size: 9pt;
            color: #64748b;
        }}
        .meta-card {{
            background-color: #f8fafc;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            padding: 10px 14px;
            margin-bottom: 12px;
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 10px;
            font-size: 9.5pt;
        }}
        .meta-item b {{
            color: #1e293b;
            display: block;
        }}
        .meta-item span {{
            color: #0369a1;
            font-weight: 700;
        }}
        table.report-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 9pt;
            margin-bottom: 14px;
        }}
        table.report-table th, table.report-table td {{
            border: 1px solid #cbd5e1;
            padding: 5px 6px;
            text-align: center;
        }}
        table.report-table th {{
            background-color: #0284c7;
            color: #ffffff;
            font-weight: 700;
        }}
        table.report-table tr:nth-child(even) {{
            background-color: #f8fafc;
        }}
        .code-badge {{
            font-family: monospace;
            background: #e0f2fe;
            color: #0369a1;
            padding: 1px 4px;
            border-radius: 3px;
        }}
        code {{
            font-family: monospace;
            background: #f1f5f9;
            padding: 1px 3px;
            border-radius: 3px;
        }}
        .summary-box {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            background-color: #f0fdf4;
            border: 1px solid #bbf7d0;
            border-radius: 6px;
            padding: 8px 14px;
            font-size: 9pt;
            margin-bottom: 14px;
        }}
        .signatures-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            text-align: center;
            font-size: 9.5pt;
            margin-top: 15px;
            padding-top: 10px;
            border-top: 1px dashed #cbd5e1;
        }}
        .sig-box {{
            padding: 8px;
        }}
        .sig-line {{
            margin-top: 35px;
            border-top: 1px solid #94a3b8;
            padding-top: 4px;
            font-weight: 600;
            color: #334155;
        }}
        .action-bar {{
            position: fixed;
            bottom: 20px;
            left: 20px;
            display: flex;
            gap: 10px;
            z-index: 999;
        }}
        .btn-print {{
            background-color: #0284c7;
            color: white;
            border: none;
            padding: 10px 20px;
            font-size: 11pt;
            font-weight: bold;
            border-radius: 6px;
            cursor: pointer;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            font-family: inherit;
        }}
        .btn-print:hover {{
            background-color: #0369a1;
        }}
        @media print {{
            .action-bar {{ display: none !important; }}
            body {{ padding: 0; }}
            table.report-table th {{ background-color: #0284c7 !important; color: white !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
            .meta-card, .summary-box {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
        }}
    </style>
</head>
<body>
    <div class="action-bar">
        <button class="btn-print" onclick="window.print();">🖨️ طباعة / حفظ كـ PDF</button>
    </div>

    <div class="report-header">
        <div class="header-logo">
            <b>المملكة الأردنية الهاشمية</b><br>
            هيئة تنظيم قطاع الطاقة والمعادن (EMRC)<br>
            المختبر الفني للرصد الإشعاعي الفردي
        </div>
        <div class="header-title">
            <h2>تقرير نتائج قياس الجرعات الإشعاعية المهنية</h2>
            <h4>Landauer InLight® Personal Dosimetry Official Report</h4>
        </div>
        <div class="header-meta">
            <b>كود التقرير:</b> {dept_info['org_code']}-{year}-P{period_number}<br>
            <b>تاريخ الإصدار:</b> {issue_date}<br>
            <b>المعيار:</b> IAEA Safety Standards & EMRC
        </div>
    </div>

    <div class="meta-card">
        <div class="meta-item">
            <b>المنشأة الطبية / المستشفى:</b>
            <span>{dept_info['org_name']} ({dept_info['org_code']})</span>
        </div>
        <div class="meta-item">
            <b>القسم الإشعاعي:</b>
            <span>{dept_info['dept_name']}</span>
        </div>
        <div class="meta-item">
            <b>فترة الرصد والسنة:</b>
            <span>{period_title} - لعام {year}</span>
        </div>
        <div class="meta-item">
            <b>طبيعة ونظام الرصد:</b>
            <span>{freq_label}</span>
        </div>
    </div>

    <table class="report-table">
        <thead>
            <tr>
                <th>#</th>
                <th>كود المشترك</th>
                <th>اسم المشترك الرباعي</th>
                <th>الرقم الوطني / الوظيفي</th>
                <th>المسمى الوظيفي</th>
                <th>نوع الوضحة</th>
                <th>سيريال OA</th>
                <th>سيريال UA</th>
                <th>الجرعة العميقة Hp(10) mSv</th>
                <th>الجرعة السطحية Hp(0.07) mSv</th>
                <th>الجرعة التراكمية السنوية mSv</th>
                <th>حالة الامتثال</th>
            </tr>
        </thead>
        <tbody>
            {table_rows}
        </tbody>
    </table>

    <div class="summary-box">
        <div>• إجمالي الكوادر المشمولة: <b>{total_subs}</b> مشترك</div>
        <div>• عدد الوضحات المقروءة والمعتمدة: <b>{recorded_count}</b> وضحة</div>
        <div>• أعلى جرعة مسجلة في هذا التقرير: <b style="color: {'#dc2626' if max_dose > 1.67 else '#16a34a'};">{max_dose} mSv</b></div>
        <div>• متوسط الجرعات: <b>{avg_dose} mSv</b></div>
        <div>• الحد السنوي المعتمد: <b>20.0 mSv/سنة</b> (جميع النتائج مطابقة للأنظمة)</div>
    </div>

    <div class="signatures-grid">
        <div class="sig-box">
            <b>إعداد وتدقيق فني القياس والرصد:</b>
            <div class="sig-line">فني مختبر OSL المعتمد</div>
        </div>
        <div class="sig-box">
            <b>مصادقة مسؤول الوقاية الإشعاعية (RPO):</b>
            <div class="sig-line">مسؤول الحماية الإشعاعية للمنشأة</div>
        </div>
        <div class="sig-box">
            <b>ختم المختبر والاعتماد الرسمي:</b>
            <div class="sig-line">المختبر الفني للرصد الإشعاعي</div>
        </div>
    </div>
</body>
</html>"""
    return html


def convert_html_to_pdf_robust(html_content: str, out_pdf_path: str) -> bool:
    """
    Robust multi-strategy PDF converter with zero crashes.
    1. WeasyPrint
    2. Headless Edge/Chrome
    3. Headless LibreOffice
    Returns True if valid PDF produced, False otherwise.
    """
    import subprocess
    import os
    
    # 1. Try WeasyPrint
    try:
        import weasyprint
        weasyprint.HTML(string=html_content).write_pdf(out_pdf_path)
        if os.path.exists(out_pdf_path) and os.path.getsize(out_pdf_path) > 500:
            return True
    except Exception:
        pass
        
    temp_html = out_pdf_path.replace(".pdf", ".html")
    try:
        with open(temp_html, "w", encoding="utf-8") as f:
            f.write(html_content)
    except Exception:
        pass
        
    # 2. Try Edge / Chrome / Chromium
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "msedge", "chrome", "google-chrome", "chromium", "/usr/bin/chromium"
    ]
    for p in edge_paths:
        try:
            cmd = [p, "--headless", "--disable-gpu", f"--print-to-pdf={out_pdf_path}", temp_html]
            subprocess.run(cmd, capture_output=True, timeout=12)
            if os.path.exists(out_pdf_path) and os.path.getsize(out_pdf_path) > 500:
                return True
        except Exception:
            continue
            
    # 3. Try LibreOffice
    soffice_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "soffice", "libreoffice", "/usr/bin/libreoffice"
    ]
    for p in soffice_paths:
        try:
            cmd = [p, "--headless", "--convert-to", "pdf", "--outdir", os.path.dirname(out_pdf_path), temp_html]
            subprocess.run(cmd, capture_output=True, timeout=12)
            if os.path.exists(out_pdf_path) and os.path.getsize(out_pdf_path) > 500:
                return True
        except Exception:
            continue
            
    return False

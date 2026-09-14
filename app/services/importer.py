import os
import pandas as pd
import sqlite3
from typing import Dict, Any, List, Optional
from datetime import datetime

from app.database import get_db
from app.services.calculator import (
    parse_excel_date,
    calculate_single_badge_dose,
    calculate_double_badge_dose,
    compute_accumulated_dose
)

def universal_read_file(file_path: str, nrows: Optional[int] = None) -> pd.DataFrame:
    """
    Robust reader for Excel & Reader files.
    Accurately detects:
    1. True binary Excel 97-2003 (.xls) via OLE2 signature (D0 CF 11 E0) -> uses xlrd
    2. True modern Excel (.xlsx) via ZIP signature (PK 03 04) -> uses openpyxl
    3. InLight reader files exported as CSV text (even if named .xls) -> uses text parser
    NEVER falls back to reading binary files as text!
    """
    with open(file_path, 'rb') as f:
        header = f.read(8)
        
    # 1. Real Binary Excel 97-2003 (.xls) - OLE2 BIFF8
    if header.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1') or header.startswith(b'\xd0\xcf\x11\xe0'):
        try:
            return pd.read_excel(file_path, engine='xlrd', nrows=nrows)
        except ImportError:
            raise ValueError("الملف المرفوع هو ملف إكسل ثنائي (.xls) ويتطلب مكتبة xlrd. يرجى تنفيذ الأمر: pip install xlrd في موجه الأوامر، أو حفظ الملف بصيغة .xlsx أو .csv.")
        except Exception as e:
            raise ValueError(f"خطأ في قراءة ملف الإكسل الثنائي (.xls): {str(e)}")

    # 2. Modern Excel (.xlsx) - ZIP Container
    if header.startswith(b'PK\x03\x04'):
        try:
            return pd.read_excel(file_path, engine='openpyxl', nrows=nrows)
        except Exception as e:
            raise ValueError(f"خطأ في قراءة ملف الإكسل الحديث (.xlsx): {str(e)}")

    # 3. Plain text / CSV files (including Landauer InLight .xls which are actually CSV text)
    encodings = ['utf-8', 'utf-8-sig', 'cp1256', 'latin1', 'iso-8859-1']
    for enc in encodings:
        for sep in [',', '\t', ';', None]:
            try:
                df = pd.read_csv(file_path, sep=sep, nrows=nrows, engine='python', encoding=enc)
                if len(df.columns) > 1:
                    return df
            except Exception:
                continue

    try:
        return pd.read_csv(file_path, nrows=nrows, on_bad_lines='skip')
    except Exception as e:
        raise ValueError(f"تعذر تحديد ترميز الملف: {str(e)}")

def preview_file(file_path: str, max_rows: int = 5) -> Dict[str, Any]:
    df = universal_read_file(file_path, nrows=max_rows)
    columns = [str(c).strip() for c in df.columns]
    preview_data = df.fillna('').astype(str).to_dict(orient='records')
    
    return {
        'columns': columns,
        'preview': preview_data,
        'total_columns': len(columns)
    }

def get_assigned_serials_for_period(conn, period_year: int, period_type: str, period_number: int) -> Dict[str, Dict[str, Any]]:
    """
    Retrieves all assigned serials for a given period:
    1. From explicit period assignments in badge_assignments
    2. Fallback to active subscribers' registered serials (current_serial_oa, current_serial_ua)
    """
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ba.serial_number, ba.badge_position, ba.subscriber_id,
               s.name as subscriber_name, s.system_code, s.has_double_badge, s.status as subscriber_status,
               d.id as dept_id, d.name as dept_name, d.default_frequency as dept_freq, o.name as org_name
        FROM badge_assignments ba
        JOIN subscribers s ON ba.subscriber_id = s.id
        JOIN departments d ON s.dept_id = d.id
        JOIN organizations o ON d.org_id = o.id
        WHERE ba.period_year = ? AND ba.period_type = ? AND ba.period_number = ?
    """, (period_year, period_type.upper(), period_number))
    assignments = [dict(a) for a in cursor.fetchall()]
    assigned_serials_map = {a['serial_number']: a for a in assignments}
    
    # Also check active subscribers with default serials
    cursor.execute("""
        SELECT s.id as subscriber_id, s.name as subscriber_name, s.system_code, s.has_double_badge,
               s.current_serial_oa, s.current_serial_ua,
               d.id as dept_id, d.name as dept_name, d.default_frequency as dept_freq, o.name as org_name
        FROM subscribers s
        JOIN departments d ON s.dept_id = d.id
        JOIN organizations o ON d.org_id = o.id
        WHERE s.status = 'ACTIVE'
    """)
    all_active_subs = cursor.fetchall()
    
    for s in all_active_subs:
        sub_id = s['subscriber_id']
        is_double = bool(s['has_double_badge'])
        freq = s['dept_freq'] or 'QUARTERLY'
        
        oa_sn = (s['current_serial_oa'] or '').strip()
        if oa_sn and oa_sn not in assigned_serials_map:
            pos = 'OA' if is_double else 'CHEST'
            assigned_serials_map[oa_sn] = {
                "serial_number": oa_sn,
                "badge_position": pos,
                "subscriber_id": sub_id,
                "subscriber_name": s['subscriber_name'],
                "system_code": s['system_code'],
                "has_double_badge": is_double,
                "dept_id": s['dept_id'],
                "dept_name": s['dept_name'],
                "dept_freq": freq,
                "org_name": s['org_name']
            }
            
        ua_sn = (s['current_serial_ua'] or '').strip()
        if ua_sn and is_double and ua_sn not in assigned_serials_map:
            assigned_serials_map[ua_sn] = {
                "serial_number": ua_sn,
                "badge_position": 'UA',
                "subscriber_id": sub_id,
                "subscriber_name": s['subscriber_name'],
                "system_code": s['system_code'],
                "has_double_badge": is_double,
                "dept_id": s['dept_id'],
                "dept_name": s['dept_name'],
                "dept_freq": freq,
                "org_name": s['org_name']
            }
            
    return assigned_serials_map

def analyze_reader_file(
    file_path: str,
    period_year: int,
    period_type: str,
    period_number: int,
    col_serial: str,
    col_read_date: Optional[str],
    col_hp10: str,
    col_hp07: str
) -> Dict[str, Any]:
    df = universal_read_file(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    
    for col in [col_serial, col_hp10, col_hp07]:
        if col not in df.columns:
            raise ValueError(f"العمود '{col}' غير موجود في الملف المرفوع.")
            
    readings_grouped: Dict[str, List[Dict[str, Any]]] = {}
    
    for idx, row in df.iterrows():
        serial = str(row[col_serial]).strip()
        if not serial or serial.lower() in ('nan', 'none', ''):
            continue
            
        raw_date = row[col_read_date] if col_read_date and col_read_date in df.columns else ''
        formatted_date = parse_excel_date(raw_date)
        
        try:
            raw_10 = float(row[col_hp10])
        except (ValueError, TypeError):
            raw_10 = 0.0
            
        try:
            raw_07 = float(row[col_hp07])
        except (ValueError, TypeError):
            raw_07 = 0.0
            
        entry = {
            "row_index": int(idx) + 1,
            "serial": serial,
            "read_date": formatted_date,
            "hp10": round(raw_10, 2),
            "hp07": round(raw_07, 2)
        }
        
        if serial not in readings_grouped:
            readings_grouped[serial] = []
        readings_grouped[serial].append(entry)
        
    conn = get_db()
    assigned_serials_map = get_assigned_serials_for_period(conn, period_year, period_type, period_number)
    conn.close()
    
    matched_list = []
    duplicate_list = []
    unassigned_list = []
    
    for serial, read_list in readings_grouped.items():
        is_assigned = serial in assigned_serials_map
        assignment_info = assigned_serials_map.get(serial, {})
        
        if len(read_list) > 1:
            duplicate_list.append({
                "serial": serial,
                "readings_count": len(read_list),
                "is_assigned": is_assigned,
                "subscriber_id": assignment_info.get("subscriber_id"),
                "subscriber_name": assignment_info.get("subscriber_name", "غير مسجل لمشترك"),
                "system_code": assignment_info.get("system_code", ""),
                "dept_id": assignment_info.get("dept_id"),
                "dept_name": assignment_info.get("dept_name", ""),
                "dept_freq": assignment_info.get("dept_freq", "QUARTERLY"),
                "badge_position": assignment_info.get("badge_position", "CHEST"),
                "readings": read_list
            })
        else:
            single_reading = read_list[0]
            if is_assigned:
                matched_list.append({
                    "serial": serial,
                    "subscriber_id": assignment_info.get("subscriber_id"),
                    "subscriber_name": assignment_info.get("subscriber_name"),
                    "system_code": assignment_info.get("system_code", ""),
                    "dept_id": assignment_info.get("dept_id"),
                    "dept_name": assignment_info.get("dept_name"),
                    "dept_freq": assignment_info.get("dept_freq", "QUARTERLY"),
                    "badge_position": assignment_info.get("badge_position"),
                    "has_double_badge": bool(assignment_info.get("has_double_badge")),
                    "read_date": single_reading["read_date"],
                    "hp10": single_reading["hp10"],
                    "hp07": single_reading["hp07"]
                })
            else:
                unassigned_list.append({
                    "serial": serial,
                    "read_date": single_reading["read_date"],
                    "hp10": single_reading["hp10"],
                    "hp07": single_reading["hp07"],
                    "row_index": single_reading["row_index"]
                })
                
    missing_list = []
    for sn, info in assigned_serials_map.items():
        if sn not in readings_grouped:
            missing_list.append({
                "serial": sn,
                "subscriber_id": info["subscriber_id"],
                "subscriber_name": info["subscriber_name"],
                "dept_name": info["dept_name"],
                "dept_freq": info.get("dept_freq", "QUARTERLY"),
                "badge_position": info["badge_position"]
            })

    # Smart detection of mixed Monthly & Quarterly badges in the same file
    monthly_matched = [m for m in matched_list if m.get("dept_freq") == "MONTHLY"]
    quarterly_matched = [m for m in matched_list if m.get("dept_freq") != "MONTHLY"]
    
    mixed_detected = False
    suggested_months = []
    default_suggested_month = 1
    
    if period_type.upper() == "QUARTERLY":
        start_m = (period_number - 1) * 3 + 1
        suggested_months = [start_m, start_m + 1, start_m + 2]
        default_suggested_month = start_m + 2 # Default to 3rd month of the quarter
        if len(monthly_matched) > 0:
            mixed_detected = True
    elif period_type.upper() == "MONTHLY":
        suggested_months = [period_number]
        default_suggested_month = period_number
        if len(quarterly_matched) > 0:
            mixed_detected = True
            
    monthly_depts = sorted(list(set(m["dept_name"] for m in monthly_matched if m.get("dept_name"))))
    quarterly_depts = sorted(list(set(q["dept_name"] for q in quarterly_matched if q.get("dept_name"))))
            
    return {
        "status": "success",
        "summary": {
            "total_serials_in_file": len(readings_grouped),
            "matched_count": len(matched_list),
            "duplicate_serials_count": len(duplicate_list),
            "unassigned_serials_count": len(unassigned_list),
            "missing_count": len(missing_list),
            "mixed_frequencies_detected": mixed_detected,
            "monthly_count": len(monthly_matched),
            "quarterly_count": len(quarterly_matched),
            "monthly_depts": monthly_depts,
            "quarterly_depts": quarterly_depts,
            "suggested_months": suggested_months,
            "default_suggested_month": default_suggested_month
        },
        "matched": matched_list,
        "duplicates": duplicate_list,
        "unassigned": unassigned_list,
        "missing": missing_list
    }

def commit_reader_import(
    file_path: str,
    period_year: int,
    period_type: str,
    period_number: int,
    col_serial: str,
    col_read_date: Optional[str],
    col_hp10: str,
    col_hp07: str,
    duplicate_resolutions: Dict[str, int],
    unassigned_resolutions: List[Dict[str, Any]],
    reception_date: Optional[str] = None,
    report_date: Optional[str] = None,
    monthly_period_number: Optional[int] = None
) -> Dict[str, Any]:
    conn = get_db()
    cursor = conn.cursor()
    
    for un in (unassigned_resolutions or []):
        sn = un.get("serial")
        act = un.get("action")
        if act == "assign" and un.get("subscriber_id"):
            sub_id = un["subscriber_id"]
            pos = un.get("position", "CHEST").upper()
            
            # Check subscriber dept frequency
            cursor.execute("""
                SELECT d.default_frequency 
                FROM subscribers s 
                JOIN departments d ON s.dept_id = d.id 
                WHERE s.id = ?
            """, (sub_id,))
            row_freq = cursor.fetchone()
            sub_freq = (row_freq['default_frequency'] if row_freq else 'QUARTERLY') or 'QUARTERLY'
            
            if monthly_period_number and sub_freq == 'MONTHLY':
                assign_type = 'MONTHLY'
                assign_num = int(monthly_period_number)
            else:
                assign_type = period_type.upper()
                assign_num = period_number
                
            cursor.execute("""
                INSERT OR REPLACE INTO badge_assignments 
                (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (sub_id, period_year, assign_type, assign_num, sn, pos))
            # Also update subscriber current serial
            if pos == 'UA':
                cursor.execute("UPDATE subscribers SET current_serial_ua = ? WHERE id = ?", (sn, sub_id))
            else:
                cursor.execute("UPDATE subscribers SET current_serial_oa = ? WHERE id = ?", (sn, sub_id))
        elif act == "unassigned_log":
            cursor.execute("""
                INSERT INTO unassigned_readings 
                (serial_number, read_date, hp10, hp07, period_year, period_type, period_number, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (sn, un.get("read_date", ""), un.get("hp10", 0.0), un.get("hp07", 0.0), period_year, period_type.upper(), period_number, un.get("notes", "وضحة مجهولة تم تحويلها لسجل غير المسجلين")))
    conn.commit()

    df = universal_read_file(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    
    rows_by_serial: Dict[str, List[Dict[str, Any]]] = {}
    for idx, row in df.iterrows():
        serial = str(row[col_serial]).strip()
        if not serial or serial.lower() in ('nan', 'none', ''):
            continue
        raw_date = row[col_read_date] if col_read_date and col_read_date in df.columns else ''
        formatted_date = parse_excel_date(raw_date)
        try:
            raw_10 = float(row[col_hp10])
        except (ValueError, TypeError):
            raw_10 = 0.0
        try:
            raw_07 = float(row[col_hp07])
        except (ValueError, TypeError):
            raw_07 = 0.0
            
        item = {
            "row_index": int(idx) + 1,
            "read_date": formatted_date,
            "hp10": raw_10,
            "hp07": raw_07
        }
        if serial not in rows_by_serial:
            rows_by_serial[serial] = []
        rows_by_serial[serial].append(item)
        
    final_readings_map: Dict[str, Dict[str, Any]] = {}
    for serial, items in rows_by_serial.items():
        if len(items) == 1:
            final_readings_map[serial] = items[0]
        else:
            chosen_row = (duplicate_resolutions or {}).get(serial)
            chosen_item = None
            if chosen_row is not None:
                for it in items:
                    if it["row_index"] == chosen_row:
                        chosen_item = it
                        break
            final_readings_map[serial] = chosen_item or items[-1]
            
        it = final_readings_map[serial]
        cursor.execute("""
            INSERT INTO raw_readings (serial_number, read_date, raw_hp10, raw_hp07)
            VALUES (?, ?, ?, ?)
        """, (serial, it["read_date"], it["hp10"], it["hp07"]))
    conn.commit()

    # Retrieve all assigned serials (including from subscriber profile defaults)
    assigned_serials_map = get_assigned_serials_for_period(conn, period_year, period_type, period_number)
    
    # Group by subscriber
    subscribers_map: Dict[int, Dict[str, Any]] = {}
    for sn, info in assigned_serials_map.items():
        sub_id = info['subscriber_id']
        if sub_id not in subscribers_map:
            subscribers_map[sub_id] = {
                'name': info['subscriber_name'],
                'system_code': info['system_code'],
                'has_double_badge': bool(info['has_double_badge']),
                'dept_freq': info.get('dept_freq', 'QUARTERLY'),
                'dept_name': info.get('dept_name', ''),
                'badges': {}
            }
        pos = info['badge_position'].upper()
        subscribers_map[sub_id]['badges'][pos] = sn
        
    calculated_results = []
    matched_count = 0
    matched_monthly_count = 0
    matched_quarterly_count = 0
    
    for sub_id, sub_info in subscribers_map.items():
        badges = sub_info['badges']
        is_double = sub_info['has_double_badge']
        dept_freq = sub_info.get('dept_freq', 'QUARTERLY')
        
        # Route to appropriate period based on frequency
        if monthly_period_number and dept_freq == 'MONTHLY':
            target_period_type = 'MONTHLY'
            target_period_number = int(monthly_period_number)
        else:
            target_period_type = period_type.upper()
            target_period_number = period_number

        reading_date_val = ''
        has_reading = False
        
        if is_double:
            oa_sn = badges.get('OA') or badges.get('CHEST')
            ua_sn = badges.get('UA')
            oa_data = final_readings_map.get(oa_sn) if oa_sn else None
            ua_data = final_readings_map.get(ua_sn) if ua_sn else None
            
            if oa_data or ua_data:
                has_reading = True
                matched_count += 1
                if target_period_type == 'MONTHLY':
                    matched_monthly_count += 1
                else:
                    matched_quarterly_count += 1
                    
                reading_date_val = (oa_data['read_date'] if oa_data else '') or (ua_data['read_date'] if ua_data else '')
                final_hp10, final_hp07, remark = calculate_double_badge_dose(
                    oa_data['hp10'] if oa_data else 0.0,
                    oa_data['hp07'] if oa_data else 0.0,
                    ua_data['hp10'] if ua_data else 0.0,
                    ua_data['hp07'] if ua_data else 0.0
                )
            else:
                # If subscriber is monthly and this is primarily quarterly without monthly selected, skip
                if dept_freq == 'MONTHLY' and not monthly_period_number:
                    continue
                final_hp10, final_hp07, remark = 0.0, 0.0, 'NR'
                
            badge_oa = oa_sn or ''
            badge_ua = ua_sn or ''
        else:
            sn = badges.get('CHEST') or badges.get('OA') or (list(badges.values())[0] if badges else None)
            data = final_readings_map.get(sn) if sn else None
            
            if data:
                has_reading = True
                matched_count += 1
                if target_period_type == 'MONTHLY':
                    matched_monthly_count += 1
                else:
                    matched_quarterly_count += 1
                    
                reading_date_val = data['read_date']
                final_hp10, final_hp07, remark = calculate_single_badge_dose(data['hp10'], data['hp07'])
            else:
                if dept_freq == 'MONTHLY' and not monthly_period_number:
                    continue
                final_hp10, final_hp07, remark = 0.0, 0.0, 'NR'
                
            badge_oa = sn or ''
            badge_ua = ''

        # Compute accumulated dose using target_period_type and target_period_number
        acc_hp10, acc_hp07 = compute_accumulated_dose(
            conn, sub_id, period_year, target_period_type, target_period_number, final_hp10, final_hp07
        )
        
        rep_date = report_date or datetime.today().strftime('%d/%m/%Y')
        rec_date = reception_date or reading_date_val
        
        cursor.execute("""
            INSERT INTO calculated_doses (
                subscriber_id, period_year, period_type, period_number,
                hp10, hp07, accumulated_hp10, accumulated_hp07,
                is_double_badge, badge_oa_serial, badge_ua_serial,
                reception_date, reading_date, report_date, remark
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(subscriber_id, period_year, period_type, period_number) DO UPDATE SET
                hp10 = excluded.hp10,
                hp07 = excluded.hp07,
                accumulated_hp10 = excluded.accumulated_hp10,
                accumulated_hp07 = excluded.accumulated_hp07,
                is_double_badge = excluded.is_double_badge,
                badge_oa_serial = excluded.badge_oa_serial,
                badge_ua_serial = excluded.badge_ua_serial,
                reception_date = excluded.reception_date,
                reading_date = excluded.reading_date,
                report_date = excluded.report_date,
                remark = excluded.remark
        """, (
            sub_id, period_year, target_period_type, target_period_number,
            final_hp10, final_hp07, acc_hp10, acc_hp07,
            1 if is_double else 0, badge_oa, badge_ua,
            rec_date, reading_date_val, rep_date, remark
        ))
        
        # Also ensure badge_assignments has record for this period
        if badge_oa:
            cursor.execute("""
                INSERT OR IGNORE INTO badge_assignments 
                (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (sub_id, period_year, target_period_type, target_period_number, badge_oa, 'OA' if is_double else 'CHEST'))
        if badge_ua and is_double:
            cursor.execute("""
                INSERT OR IGNORE INTO badge_assignments 
                (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (sub_id, period_year, target_period_type, target_period_number, badge_ua, 'UA'))
        
        calculated_results.append({
            'subscriber_id': sub_id,
            'name': sub_info['name'],
            'system_code': sub_info['system_code'],
            'period_type': target_period_type,
            'period_number': target_period_number,
            'hp10': final_hp10,
            'hp07': final_hp07,
            'accumulated_hp10': acc_hp10,
            'accumulated_hp07': acc_hp07,
            'remark': remark
        })
        
    conn.commit()
    conn.close()
    
    return {
        'status': 'success',
        'total_readings_in_file': len(rows_by_serial),
        'total_subscribers_matched': matched_count,
        'matched_monthly_count': matched_monthly_count,
        'matched_quarterly_count': matched_quarterly_count,
        'monthly_period_number': monthly_period_number,
        'period_number': period_number,
        'period_type': period_type.upper(),
        'results': calculated_results
    }

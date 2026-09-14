import os
import shutil
import json
from typing import Optional, List, Dict, Any
from datetime import datetime
import hmac, hashlib, base64, time
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Body, Request, Response
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd

from app.database import get_db, init_db, DB_PATH
from app.services.importer import preview_file, analyze_reader_file, commit_reader_import, universal_read_file
from app.services.calculator import compute_accumulated_dose, calculate_single_badge_dose, calculate_double_badge_dose
from app.services.name_matcher import calculate_name_similarity, find_duplicate_names
from app.services.reporter import (
    generate_monthly_excel_report,
    generate_quarterly_excel_report,
    generate_monthly_word_report,
    generate_quarterly_word_report,
    generate_reception_manifest_excel,
    generate_annual_archive_excel
)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
TEMPLATES_DIR = os.path.join(BASE_DIR, "app", "templates")
STATIC_HTML_PATH = os.path.join(BASE_DIR, "app", "static", "index.html")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

app = FastAPI(title="OSL Dosimetry Management System", version="3.5.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.on_event("startup")
def startup_event():
    init_db()

def generate_new_system_code(cursor) -> str:
    year = datetime.now().year
    cursor.execute("SELECT COUNT(*) FROM subscribers")
    count = cursor.fetchone()[0] + 1
    return f"OSL-{year}-{count:04d}"

# --- 0. Database Backup & Restore ---
@app.get("/api/database/backup")
def backup_database():
    if not os.path.exists(DB_PATH):
        raise HTTPException(status_code=404, detail="قاعدة البيانات غير موجودة")
    date_str = datetime.now().strftime("%Y%m%d_%H%M")
    return FileResponse(DB_PATH, filename=f"osl_database_backup_{date_str}.db", media_type="application/octet-stream")

@app.post("/api/database/restore")
async def restore_database(file: UploadFile = File(...)):
    try:
        # Overwrite DB_PATH
        with open(DB_PATH, "wb") as f:
            shutil.copyfileobj(file.file, f)
        init_db()
        return {"status": "success", "message": "تمت استعادة قاعدة البيانات بنجاح!"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"فشل استعادة قاعدة البيانات: {str(e)}")

# --- 1. System Stats ---

@app.post("/api/database/reset")
def reset_database(confirmation: str = Form(...)):
    clean_conf = confirmation.strip().lower()
    if clean_conf not in ("confirm", "تأكيد"):
        raise HTTPException(status_code=400, detail="كلمة التأكيد غير صحيحة. يرجى كتابة 'تأكيد' أو 'CONFIRM' للمتابعة")
        
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM badge_receptions")
    cursor.execute("DELETE FROM calculated_doses")
    cursor.execute("DELETE FROM raw_readings")
    cursor.execute("DELETE FROM unassigned_readings")
    cursor.execute("DELETE FROM badge_assignments")
    cursor.execute("DELETE FROM subscribers")
    cursor.execute("DELETE FROM departments")
    cursor.execute("DELETE FROM organizations")
    cursor.execute("DELETE FROM import_templates")
    try:
        cursor.execute("DELETE FROM sqlite_sequence")
    except Exception:
        pass
    conn.commit()
    conn.close()
    return {"status": "success", "message": "تم إفراغ كافة بيانات قاعدة البيانات بنجاح"}


# --- 0.5 System Settings & Active Year APIs ---
@app.get("/api/settings")
def get_settings():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT key, value FROM system_settings")
    settings = {r['key']: r['value'] for r in cursor.fetchall()}
    conn.close()
    return {
        "active_year": int(settings.get('active_year', 2026)),
        "theme": settings.get('theme', 'dark'),
        "language": settings.get('language', 'ar')
    }

@app.post("/api/settings")
def update_settings(
    active_year: Optional[int] = Form(None),
    theme: Optional[str] = Form(None),
    language: Optional[str] = Form(None)
):
    conn = get_db()
    cursor = conn.cursor()
    if active_year:
        cursor.execute("""
            INSERT INTO system_settings (key, value) VALUES ('active_year', ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = CURRENT_TIMESTAMP
        """, (str(active_year),))
    if theme:
        cursor.execute("""
            INSERT INTO system_settings (key, value) VALUES ('theme', ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = CURRENT_TIMESTAMP
        """, (theme,))
    if language:
        cursor.execute("""
            INSERT INTO system_settings (key, value) VALUES ('language', ?)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = CURRENT_TIMESTAMP
        """, (language,))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "تم حفظ الإعدادات بنجاح"}

@app.post("/api/settings/new-year")
def open_new_year(new_year: int = Form(...)):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO system_settings (key, value) VALUES ('active_year', ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = CURRENT_TIMESTAMP
    """, (str(new_year),))
    conn.commit()
    conn.close()
    return {"status": "success", "new_year": new_year, "message": f"تم تفعيل وافتتاح السنة التشغيلية الجديدة {new_year} بنجاح!"}

@app.get("/api/stats")
def get_stats():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM organizations")
    orgs_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM departments")
    depts_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM subscribers WHERE status = 'ACTIVE'")
    active_subs_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM subscribers WHERE status = 'RESIGNED'")
    resigned_subs_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM subscribers WHERE has_double_badge = 1 AND status = 'ACTIVE'")
    double_subs_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM calculated_doses")
    doses_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM unassigned_readings")
    unassigned_count = cursor.fetchone()[0]
    conn.close()
    return {
        "organizations": orgs_count,
        "departments": depts_count,
        "subscribers": active_subs_count,
        "resigned_subscribers": resigned_subs_count,
        "double_badge_subscribers": double_subs_count,
        "calculated_doses": doses_count,
        "unassigned_readings": unassigned_count
    }



# =========================================================================
# --- 0. MULTI-TENANT AUTHENTICATION & ACCESS CONTROL SYSTEM ---
# =========================================================================
AUTH_SECRET_KEY = b"osl_inlight_radiation_auth_secret_key_2026_secure"

def create_access_token(user_dict: dict, expires_in_hours: int = 72) -> str:
    payload = {
        "user_id": user_dict["id"],
        "username": user_dict["username"],
        "role": user_dict["role"],
        "org_id": user_dict.get("org_id"),
        "exp": int(time.time()) + (expires_in_hours * 3600)
    }
    payload_bytes = json.dumps(payload).encode("utf-8")
    sig = hmac.new(AUTH_SECRET_KEY, payload_bytes, hashlib.sha256).digest()
    token = base64.urlsafe_b64encode(payload_bytes).decode("utf-8") + "." + base64.urlsafe_b64encode(sig).decode("utf-8")
    return token

def verify_access_token(token: str):
    try:
        parts = token.split(".")
        if len(parts) != 2:
            return None
        payload_bytes = base64.urlsafe_b64decode(parts[0].encode("utf-8"))
        sig = base64.urlsafe_b64decode(parts[1].encode("utf-8"))
        expected_sig = hmac.new(AUTH_SECRET_KEY, payload_bytes, hashlib.sha256).digest()
        if not hmac.compare_digest(sig, expected_sig):
            return None
        payload = json.loads(payload_bytes.decode("utf-8"))
        if payload.get("exp", 0) < int(time.time()):
            return None
        return payload
    except Exception:
        return None

import hashlib

def hash_pw(password: str) -> str:
    return hashlib.sha256(("osl_salt_2026_" + password).encode('utf-8')).hexdigest()

def verify_pw(entered: str, stored: str) -> bool:
    if not stored or not entered:
        return False
    if stored == entered:
        return True
    return stored == hash_pw(entered)

def resolve_user(request: Request = None, token: Optional[str] = None):
    if not token and request:
        auth_header = request.headers.get("Authorization")
        if auth_header and auth_header.startswith("Bearer "):
            token = auth_header.split(" ")[1]
        elif "osl_token" in request.cookies:
            token = request.cookies.get("osl_token")
        elif "token" in request.query_params:
            token = request.query_params.get("token")
            
    if not token:
        return None
        
    payload = verify_access_token(token)
    if not payload:
        return None
        
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT u.id, u.username, u.email, u.full_name, u.role, u.org_id,
               o.name as org_name, o.code as org_code
        FROM users u
        LEFT JOIN organizations o ON u.org_id = o.id
        WHERE u.id = ?
    """, (payload["user_id"],))
    user = cursor.fetchone()
    conn.close()
    return dict(user) if user else None


@app.post("/api/auth/login")
def auth_login(username_or_email: str = Form(...), password: str = Form(...), response: Response = None):
    login_id = username_or_email.strip()
    pw = password.strip()
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT u.id, u.username, u.email, u.full_name, u.role, u.password, u.org_id,
               o.name as org_name, o.code as org_code
        FROM users u
        LEFT JOIN organizations o ON u.org_id = o.id
        WHERE LOWER(u.username) = LOWER(?) OR (u.email IS NOT NULL AND LOWER(u.email) = LOWER(?))
    """, (login_id, login_id))
    user = cursor.fetchone()
    conn.close()
    
    if not user or not verify_pw(pw, user["password"]):
        raise HTTPException(status_code=401, detail="اسم المستخدم أو كلمة المرور غير صحيحة")
        
    user_dict = dict(user)
    user_dict.pop("password", None)
    
    token = create_access_token(user_dict)
    
    if response:
        response.set_cookie("osl_token", token, max_age=72*3600, httponly=False)
        
    return {
        "status": "success",
        "message": f"مرحباً بك، {user_dict['full_name']}",
        "token": token,
        "user": user_dict
    }

@app.get("/api/auth/me")
def auth_me(request: Request):
    user = resolve_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="غير مسجل الدخول")
    return {
        "status": "success",
        "user": user
    }

@app.post("/api/auth/logout")
def auth_logout(response: Response):
    if response:
        response.delete_cookie("osl_token")
    return {"status": "success", "message": "تم تسجيل الخروج بنجاح"}

@app.get("/api/admin/hospital-users")
def list_hospital_users(request: Request):
    user = resolve_user(request)
    if not user or user.get("role") != "ADMIN":
        raise HTTPException(status_code=403, detail="هذه الصلاحية متاحة لمدير المنظومة فقط")
        
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT u.id, u.username, u.email, u.full_name, u.role, u.password, u.org_id, u.created_at,
               o.name as org_name, o.code as org_code
        FROM users u
        LEFT JOIN organizations o ON u.org_id = o.id
        WHERE u.role = 'HOSPITAL'
        ORDER BY o.name ASC, u.username ASC
    """)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows

@app.post("/api/admin/hospital-users")
def create_hospital_user(
    org_id: int = Form(...),
    username: str = Form(...),
    email: str = Form(""),
    full_name: str = Form(...),
    password: str = Form(...),
    request: Request = None
):
    user = resolve_user(request)
    if not user or user.get("role") != "ADMIN":
        raise HTTPException(status_code=403, detail="هذه الصلاحية متاحة لمدير المنظومة فقط")
        
    clean_un = username.strip()
    clean_em = email.strip() or None
    clean_pw = password.strip()
    clean_fn = full_name.strip()
    
    if not clean_un or not clean_pw:
        raise HTTPException(status_code=400, detail="يرجى إدخال اسم مستخدم وكلمة مرور صالحة")
        
    conn = get_db()
    cursor = conn.cursor()
    
    # Check uniqueness
    cursor.execute("SELECT id FROM users WHERE LOWER(username) = LOWER(?)", (clean_un,))
    if cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=400, detail=f"اسم المستخدم '{clean_un}' مسجل مسبقاً")
        
    if clean_em:
        cursor.execute("SELECT id FROM users WHERE LOWER(email) = LOWER(?)", (clean_em,))
        if cursor.fetchone():
            conn.close()
            raise HTTPException(status_code=400, detail=f"البريد الإلكتروني '{clean_em}' مسجل مسبقاً")
            
    hashed_pw = hash_pw(clean_pw)
    cursor.execute("""
        INSERT INTO users (username, email, full_name, role, password, org_id)
        VALUES (?, ?, ?, 'HOSPITAL', ?, ?)
    """, (clean_un, clean_em, clean_fn, hashed_pw, org_id))
    conn.commit()
    new_id = cursor.lastrowid
    conn.close()
    
    return {"status": "success", "message": f"تم إنشاء حساب للمنشأة بنجاح", "user_id": new_id}

@app.delete("/api/admin/hospital-users/{user_id}")
def delete_hospital_user(user_id: int, request: Request = None):
    user = resolve_user(request)
    if not user or user.get("role") != "ADMIN":
        raise HTTPException(status_code=403, detail="هذه الصلاحية متاحة لمدير المنظومة فقط")
        
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM users WHERE id = ? AND role = 'HOSPITAL'", (user_id,))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "تم حذف حساب المنشأة بنجاح"}

# --- 1.1 Comprehensive Dashboard Stats ---
@app.get("/api/dashboard/stats")
def get_dashboard_stats(request: Request = None):
    user = resolve_user(request)
    user_org_id = user["org_id"] if (user and user.get("role") == "HOSPITAL") else None

    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("SELECT value FROM system_settings WHERE key = 'active_year'")
    row_y = cursor.fetchone()
    year = int(row_y[0]) if row_y and row_y[0] else 2026
    
    if user_org_id:
        orgs_count = 1
        cursor.execute("SELECT COUNT(*) FROM departments WHERE org_id = ?", (user_org_id,))
        depts_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM subscribers s JOIN departments d ON s.dept_id = d.id WHERE d.org_id = ? AND s.status = 'ACTIVE'", (user_org_id,))
        active_subs_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM subscribers s JOIN departments d ON s.dept_id = d.id WHERE d.org_id = ? AND s.status = 'RESIGNED'", (user_org_id,))
        resigned_subs_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM subscribers s JOIN departments d ON s.dept_id = d.id WHERE d.org_id = ? AND s.has_double_badge = 1 AND s.status = 'ACTIVE'", (user_org_id,))
        double_subs_count = cursor.fetchone()[0]
        total_single_subs = max(0, active_subs_count - double_subs_count)

        cursor.execute("""
            SELECT COUNT(cd.id) 
            FROM calculated_doses cd
            JOIN subscribers s ON cd.subscriber_id = s.id
            JOIN departments d ON s.dept_id = d.id
            WHERE cd.period_year = ? AND d.org_id = ?
        """, (year, user_org_id))
        total_doses = cursor.fetchone()[0]

        cursor.execute("""
            SELECT COUNT(cd.id) 
            FROM calculated_doses cd
            JOIN subscribers s ON cd.subscriber_id = s.id
            JOIN departments d ON s.dept_id = d.id
            WHERE cd.period_year = ? AND cd.hp10 > 1.67 AND d.org_id = ?
        """, (year, user_org_id))
        total_alerts = cursor.fetchone()[0]

        cursor.execute("""
            SELECT MAX(cd.hp10), ROUND(AVG(CASE WHEN cd.remark != 'NR' THEN cd.hp10 END), 2)
            FROM calculated_doses cd
            JOIN subscribers s ON cd.subscriber_id = s.id
            JOIN departments d ON s.dept_id = d.id
            WHERE cd.period_year = ? AND d.org_id = ?
        """, (year, user_org_id))
        max_avg = cursor.fetchone()
        overall_max_dose = max_avg[0] or 0.0
        overall_avg_dose = max_avg[1] or 0.0

        cursor.execute("""
            SELECT s.id, s.has_double_badge,
                   COALESCE(SUM(cd.hp10), 0.0) as cum_hp10,
                   COUNT(cd.id) as readings_count
            FROM subscribers s
            JOIN departments d ON s.dept_id = d.id
            LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id AND cd.period_year = ?
            WHERE s.status = 'ACTIVE' AND d.org_id = ?
            GROUP BY s.id
        """, (year, user_org_id))
        subs_year_data = cursor.fetchall()

        cursor.execute("""
            SELECT cd.period_number, cd.period_type,
                   d.default_frequency,
                   ROUND(AVG(cd.hp10), 2) as avg_hp10,
                   ROUND(MAX(cd.hp10), 2) as max_hp10,
                   COUNT(cd.id) as count_readings
            FROM calculated_doses cd
            JOIN subscribers s ON cd.subscriber_id = s.id
            JOIN departments d ON s.dept_id = d.id
            WHERE cd.period_year = ? AND cd.remark != 'NR' AND d.org_id = ?
            GROUP BY cd.period_number, cd.period_type, d.default_frequency
            ORDER BY cd.period_type DESC, cd.period_number ASC
        """, (year, user_org_id))
        trend_rows = cursor.fetchall()

        cursor.execute("""
            SELECT cd.hp10
            FROM calculated_doses cd
            JOIN subscribers s ON cd.subscriber_id = s.id
            JOIN departments d ON s.dept_id = d.id
            WHERE cd.period_year = ? AND cd.remark != 'NR' AND d.org_id = ?
            ORDER BY cd.id DESC
            LIMIT 20
        """, (year, user_org_id))
        pulse_bars = [round(r['hp10'], 2) for r in cursor.fetchall()]

        cursor.execute("""
            SELECT o.id, o.name, o.code, o.classification, o.contact,
                   COUNT(DISTINCT d.id) as depts_count,
                   COUNT(DISTINCT CASE WHEN s.status = 'ACTIVE' THEN s.id END) as active_subs,
                   COUNT(DISTINCT CASE WHEN s.status = 'RESIGNED' THEN s.id END) as resigned_subs,
                   COUNT(DISTINCT CASE WHEN s.has_double_badge = 1 AND s.status = 'ACTIVE' THEN s.id END) as double_subs,
                   COUNT(DISTINCT cd.id) as total_doses,
                   MAX(cd.hp10) as max_dose,
                   ROUND(AVG(CASE WHEN cd.hp10 IS NOT NULL AND cd.remark != 'NR' THEN cd.hp10 END), 2) as avg_dose,
                   COUNT(DISTINCT CASE WHEN cd.hp10 > 1.67 THEN cd.id END) as alerts_count
            FROM organizations o
            LEFT JOIN departments d ON o.id = d.org_id
            LEFT JOIN subscribers s ON d.id = s.dept_id
            LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id
            WHERE o.id = ?
            GROUP BY o.id, o.name, o.code, o.classification, o.contact
        """, (user_org_id,))
        org_rows = [dict(r) for r in cursor.fetchall()]

        cursor.execute("""
            SELECT cd.period_year, cd.period_type, cd.period_number, COUNT(cd.id) as doses_count,
                   ROUND(AVG(CASE WHEN cd.remark != 'NR' THEN cd.hp10 END), 2) as avg_hp10,
                   MAX(cd.hp10) as max_hp10,
                   COUNT(CASE WHEN cd.hp10 > 1.67 THEN 1 END) as alerts_count
            FROM calculated_doses cd
            JOIN subscribers s ON cd.subscriber_id = s.id
            JOIN departments d ON s.dept_id = d.id
            WHERE d.org_id = ?
            GROUP BY cd.period_year, cd.period_type, cd.period_number
            ORDER BY cd.period_year DESC, cd.period_number DESC
        """, (user_org_id,))
        periods_summary = [dict(p) for p in cursor.fetchall()]

    else:
        cursor.execute("SELECT COUNT(*) FROM organizations")
        orgs_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM departments")
        depts_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM subscribers WHERE status = 'ACTIVE'")
        active_subs_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM subscribers WHERE status = 'RESIGNED'")
        resigned_subs_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM subscribers WHERE has_double_badge = 1 AND status = 'ACTIVE'")
        double_subs_count = cursor.fetchone()[0]
        total_single_subs = max(0, active_subs_count - double_subs_count)
        
        cursor.execute("SELECT COUNT(*) FROM calculated_doses WHERE period_year = ?", (year,))
        total_doses = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM calculated_doses WHERE period_year = ? AND hp10 > 1.67", (year,))
        total_alerts = cursor.fetchone()[0]
        cursor.execute("SELECT MAX(hp10), ROUND(AVG(CASE WHEN remark != 'NR' THEN hp10 END), 2) FROM calculated_doses WHERE period_year = ?", (year,))
        max_avg = cursor.fetchone()
        overall_max_dose = max_avg[0] or 0.0
        overall_avg_dose = max_avg[1] or 0.0

        cursor.execute("""
            SELECT s.id, s.has_double_badge,
                   COALESCE(SUM(cd.hp10), 0.0) as cum_hp10,
                   COUNT(cd.id) as readings_count
            FROM subscribers s
            LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id AND cd.period_year = ?
            WHERE s.status = 'ACTIVE'
            GROUP BY s.id
        """, (year,))
        subs_year_data = cursor.fetchall()

        cursor.execute("""
            SELECT cd.period_number, cd.period_type,
                   d.default_frequency,
                   ROUND(AVG(cd.hp10), 2) as avg_hp10,
                   ROUND(MAX(cd.hp10), 2) as max_hp10,
                   COUNT(cd.id) as count_readings
            FROM calculated_doses cd
            JOIN subscribers s ON cd.subscriber_id = s.id
            JOIN departments d ON s.dept_id = d.id
            WHERE cd.period_year = ? AND cd.remark != 'NR'
            GROUP BY cd.period_number, cd.period_type, d.default_frequency
            ORDER BY cd.period_type DESC, cd.period_number ASC
        """, (year,))
        trend_rows = cursor.fetchall()

        cursor.execute("""
            SELECT cd.hp10
            FROM calculated_doses cd
            WHERE cd.period_year = ? AND cd.remark != 'NR'
            ORDER BY cd.id DESC
            LIMIT 20
        """, (year,))
        pulse_bars = [round(r['hp10'], 2) for r in cursor.fetchall()]

        cursor.execute("""
            SELECT o.id, o.name, o.code, o.classification, o.contact,
                   COUNT(DISTINCT d.id) as depts_count,
                   COUNT(DISTINCT CASE WHEN s.status = 'ACTIVE' THEN s.id END) as active_subs,
                   COUNT(DISTINCT CASE WHEN s.status = 'RESIGNED' THEN s.id END) as resigned_subs,
                   COUNT(DISTINCT CASE WHEN s.has_double_badge = 1 AND s.status = 'ACTIVE' THEN s.id END) as double_subs,
                   COUNT(DISTINCT cd.id) as total_doses,
                   MAX(cd.hp10) as max_dose,
                   ROUND(AVG(CASE WHEN cd.hp10 IS NOT NULL AND cd.remark != 'NR' THEN cd.hp10 END), 2) as avg_dose,
                   COUNT(DISTINCT CASE WHEN cd.hp10 > 1.67 THEN cd.id END) as alerts_count
            FROM organizations o
            LEFT JOIN departments d ON o.id = d.org_id
            LEFT JOIN subscribers s ON d.id = s.dept_id
            LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id
            GROUP BY o.id, o.name, o.code, o.classification, o.contact
            ORDER BY active_subs DESC, o.name ASC
        """)
        org_rows = [dict(r) for r in cursor.fetchall()]

        cursor.execute("""
            SELECT period_year, period_type, period_number, COUNT(id) as doses_count,
                   ROUND(AVG(CASE WHEN remark != 'NR' THEN hp10 END), 2) as avg_hp10,
                   MAX(hp10) as max_hp10,
                   COUNT(CASE WHEN hp10 > 1.67 THEN 1 END) as alerts_count
            FROM calculated_doses
            GROUP BY period_year, period_type, period_number
            ORDER BY period_year DESC, period_number DESC
        """)
        periods_summary = [dict(p) for p in cursor.fetchall()]

    zero_dose_subs = sum(1 for r in subs_year_data if r['cum_hp10'] <= 0.1)
    adherent_subs = sum(1 for r in subs_year_data if r['readings_count'] > 0)
    high_dose_subs = sum(1 for r in subs_year_data if r['cum_hp10'] > 1.67)
    exceeded_subs = sum(1 for r in subs_year_data if r['cum_hp10'] >= 20.0)

    zero_dose_pct = round((zero_dose_subs / active_subs_count * 100), 1) if active_subs_count > 0 else 0.0
    double_badge_pct = round((double_subs_count / active_subs_count * 100), 1) if active_subs_count > 0 else 0.0
    single_badge_pct = round((total_single_subs / active_subs_count * 100), 1) if active_subs_count > 0 else 0.0
    adherence_pct = round((adherent_subs / active_subs_count * 100), 1) if active_subs_count > 0 else 0.0
    safe_compliance_pct = round(((active_subs_count - exceeded_subs) / active_subs_count * 100), 1) if active_subs_count > 0 else 100.0

    cathlab_points = {}
    radiology_points = {}
    recorded_periods = []
    
    for tr in trend_rows:
        p_num = tr['period_number']
        p_type = tr['period_type']
        freq = tr['default_frequency']
        avg_v = tr['avg_hp10']
        max_v = tr['max_hp10']
        cnt = tr['count_readings']
        
        is_cath = (freq == 'MONTHLY' or p_type == 'MONTHLY')
        if is_cath:
            cathlab_points[p_num] = avg_v
            label = f"شهر {p_num} (قسطرة)"
            color = "#ff9100"
        else:
            radiology_points[p_num] = avg_v
            label = f"دورة {p_num} (أشعة)"
            color = "#f50057"
            
        recorded_periods.append({
            "period_number": p_num,
            "period_type": p_type,
            "is_cathlab": is_cath,
            "label": label,
            "color": color,
            "avg_hp10": avg_v,
            "max_hp10": max_v,
            "count": cnt
        })
            
    has_trend_data = len(trend_rows) > 0

    cursor.execute("SELECT COUNT(*) FROM raw_readings")
    total_raw = cursor.fetchone()[0]
    matched_pct = min(100.0, round((total_doses / max(1, total_raw)) * 100, 1)) if total_raw > 0 else (100.0 if total_doses > 0 else 0.0)
    
    for org in org_rows:
        org['pct_of_total_subs'] = round((org['active_subs'] / active_subs_count * 100), 1) if active_subs_count > 0 else 0.0
        org['max_dose'] = round(org['max_dose'] or 0.0, 2)
        org['avg_dose'] = round(org['avg_dose'] or 0.0, 2)
        
        cursor.execute("""
            SELECT d.id, d.name,
                   COUNT(DISTINCT CASE WHEN s.status = 'ACTIVE' THEN s.id END) as active_subs,
                   COUNT(DISTINCT CASE WHEN s.has_double_badge = 1 AND s.status = 'ACTIVE' THEN s.id END) as double_subs,
                   COUNT(DISTINCT cd.id) as doses_count
            FROM departments d
            LEFT JOIN subscribers s ON d.id = s.dept_id
            LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id
            WHERE d.org_id = ?
            GROUP BY d.id, d.name
            ORDER BY active_subs DESC
        """, (org['id'],))
        org['departments'] = [dict(d) for d in cursor.fetchall()]
        
    conn.close()
    
    return {
        "summary": {
            "active_year": year,
            "total_organizations": orgs_count,
            "total_departments": depts_count,
            "total_active_subscribers": active_subs_count,
            "total_resigned_subscribers": resigned_subs_count,
            "total_double_badge_subscribers": double_subs_count,
            "total_single_badge_subscribers": total_single_subs,
            "total_recorded_doses": total_doses,
            "total_high_dose_alerts": total_alerts,
            "overall_max_dose": overall_max_dose,
            "overall_avg_dose": overall_avg_dose,
            "compliance": {
                "zero_dose_subs": zero_dose_subs,
                "zero_dose_pct": zero_dose_pct,
                "double_subs": double_subs_count,
                "double_badge_pct": double_badge_pct,
                "single_subs": total_single_subs,
                "single_badge_pct": single_badge_pct,
                "adherent_subs": adherent_subs,
                "adherence_pct": adherence_pct,
                "high_dose_subs": high_dose_subs,
                "exceeded_subs": exceeded_subs,
                "safe_compliance_pct": safe_compliance_pct
            },
            "operational": {
                "matched_pct": matched_pct,
                "reception_pct": adherence_pct,
                "double_pct": double_badge_pct
            },
            "dose_trends": {
                "has_data": has_trend_data,
                "periods": recorded_periods,
                "cathlab": [cathlab_points.get(m, 0.0) for m in range(1, 13)],
                "radiology": [radiology_points.get(q, 0.0) for q in range(1, 5)],
                "max_dose": overall_max_dose,
                "avg_dose": overall_avg_dose
            },
            "pulse_bars": pulse_bars
        },
        "hospital_comparison": org_rows,
        "periods_summary": periods_summary
    }

@app.get("/api/analytics/hierarchy")
def get_analytics_hierarchy(year: Optional[int] = None, request: Request = None):
    user = resolve_user(request)
    user_org_id = user["org_id"] if (user and user.get("role") == "HOSPITAL") else None
    
    conn = get_db()
    cursor = conn.cursor()
    
    if not year:
        cursor.execute("SELECT value FROM system_settings WHERE key = 'active_year'")
        row_y = cursor.fetchone()
        year = int(row_y[0]) if row_y and row_y[0] else 2026
        
    if user_org_id:
        cursor.execute("""
            SELECT COUNT(DISTINCT s.id) as total_subs,
                   1 as total_orgs,
                   COUNT(DISTINCT d.id) as total_depts
            FROM organizations o
            JOIN departments d ON o.id = d.org_id
            LEFT JOIN subscribers s ON d.id = s.dept_id AND s.status = 'ACTIVE'
            WHERE o.id = ?
        """, (user_org_id,))
    else:
        cursor.execute("""
            SELECT COUNT(DISTINCT s.id) as total_subs,
                   COUNT(DISTINCT o.id) as total_orgs,
                   COUNT(DISTINCT d.id) as total_depts
            FROM organizations o
            JOIN departments d ON o.id = d.org_id
            LEFT JOIN subscribers s ON d.id = s.dept_id AND s.status = 'ACTIVE'
        """)
    row_sum = cursor.fetchone()
    
    if user_org_id:
        cursor.execute("SELECT id, name, code, classification FROM organizations WHERE id = ?", (user_org_id,))
    else:
        cursor.execute("SELECT id, name, code, classification FROM organizations ORDER BY name ASC")
    orgs = cursor.fetchall()
    
    colors_palette = ['#00e5ff', '#ff9100', '#00e676', '#f50057', '#b388ff', '#ffd600', '#26c6da', '#ff7043', '#38bdf8', '#a855f7']
    
    hospitals_data = []
    overall_cum_hp10 = 0.0
    overall_max_dose = 0.0
    
    for idx, org in enumerate(orgs):
        org_id, org_name, org_code, org_class = org
        org_color = colors_palette[idx % len(colors_palette)]
        
        cursor.execute("""
            SELECT d.id, d.name, d.default_frequency, d.has_double_badge_default
            FROM departments d
            WHERE d.org_id = ?
            ORDER BY d.name ASC
        """, (org_id,))
        depts = cursor.fetchall()
        
        depts_data = []
        org_total_subs = 0
        org_cum_hp10 = 0.0
        org_cum_hp07 = 0.0
        org_max_hp10 = 0.0
        
        for jdx, dept in enumerate(depts):
            d_id, d_name, d_freq, d_double = dept
            dept_color = colors_palette[(idx * 3 + jdx + 1) % len(colors_palette)]
            
            cursor.execute("""
                SELECT s.id, s.name, s.system_code, s.job_title, s.has_double_badge,
                       s.current_serial_oa, s.current_serial_ua,
                       ROUND(COALESCE(SUM(cd.hp10), 0.0), 2) as cum_hp10,
                       ROUND(COALESCE(SUM(cd.hp07), 0.0), 2) as cum_hp07,
                       COALESCE(MAX(cd.hp10), 0.0) as max_hp10,
                       COUNT(cd.id) as readings_count
                FROM subscribers s
                LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id AND cd.period_year = ?
                WHERE s.dept_id = ? AND s.status = 'ACTIVE'
                GROUP BY s.id
                ORDER BY cum_hp10 DESC, s.name ASC
            """, (year, d_id))
            
            employees = []
            dept_cum_hp10 = 0.0
            dept_cum_hp07 = 0.0
            dept_max_hp10 = 0.0
            
            for emp in cursor.fetchall():
                e_id, e_name, e_code, e_job, e_double, e_oa, e_ua, e_cum10, e_cum07, e_max10, e_cnt = emp
                employees.append({
                    'id': e_id,
                    'name': e_name,
                    'system_code': e_code or '',
                    'job_title': e_job or '',
                    'has_double_badge': bool(e_double),
                    'current_serial_oa': e_oa or '',
                    'current_serial_ua': e_ua or '',
                    'cumulative_hp10': round(e_cum10, 2),
                    'cumulative_hp07': round(e_cum07, 2),
                    'max_hp10': round(e_max10, 2),
                    'readings_count': e_cnt,
                    'is_safe': e_cum10 <= 20.0
                })
                dept_cum_hp10 += e_cum10
                dept_cum_hp07 += e_cum07
                if e_max10 > dept_max_hp10:
                    dept_max_hp10 = e_max10
                    
            dept_cum_hp10 = round(dept_cum_hp10, 2)
            dept_cum_hp07 = round(dept_cum_hp07, 2)
            org_cum_hp10 += dept_cum_hp10
            org_cum_hp07 += dept_cum_hp07
            org_total_subs += len(employees)
            if dept_max_hp10 > org_max_hp10:
                org_max_hp10 = dept_max_hp10
                
            depts_data.append({
                'id': d_id,
                'name': d_name,
                'frequency': d_freq or 'QUARTERLY',
                'has_double_badge': bool(d_double),
                'color': dept_color,
                'subscribers_count': len(employees),
                'cumulative_hp10': dept_cum_hp10,
                'cumulative_hp07': dept_cum_hp07,
                'max_hp10': round(dept_max_hp10, 2),
                'employees': employees
            })
            
        org_cum_hp10 = round(org_cum_hp10, 2)
        org_cum_hp07 = round(org_cum_hp07, 2)
        overall_cum_hp10 += org_cum_hp10
        if org_max_hp10 > overall_max_dose:
            overall_max_dose = org_max_hp10
            
        hospitals_data.append({
            'id': org_id,
            'name': org_name,
            'code': org_code or '',
            'classification': org_class or '',
            'color': org_color,
            'subscribers_count': org_total_subs,
            'cumulative_hp10': org_cum_hp10,
            'cumulative_hp07': org_cum_hp07,
            'max_hp10': round(org_max_hp10, 2),
            'departments': depts_data
        })
        
    conn.close()
    
    return {
        'summary': {
            'active_year': year,
            'total_subscribers': row_sum[0] if row_sum else 0,
            'total_hospitals': row_sum[1] if row_sum else 0,
            'total_departments': row_sum[2] if row_sum else 0,
            'overall_cumulative_hp10': round(overall_cum_hp10, 2),
            'overall_max_dose': round(overall_max_dose, 2)
        },
        'hospitals': hospitals_data
    }

# --- 2. Organizations & Departments APIs ---
@app.get("/api/organizations")
def get_organizations(request: Request = None):
    user = resolve_user(request)
    conn = get_db()
    cursor = conn.cursor()
    if user and user.get("role") == "HOSPITAL":
        cursor.execute("SELECT * FROM organizations WHERE id = ? ORDER BY name ASC", (user["org_id"],))
    else:
        cursor.execute("SELECT * FROM organizations ORDER BY name ASC")
    orgs = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return orgs

@app.post("/api/organizations")
def create_organization(name: str = Form(...), code: str = Form(...), address: str = Form(""), contact: str = Form(""), agreement_date: str = Form("")):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO organizations (name, code, address, contact, agreement_date) VALUES (?, ?, ?, ?, ?)", (name, code, address, contact, agreement_date))
    conn.commit()
    new_id = cursor.lastrowid
    conn.close()
    return {"status": "success", "id": new_id}

@app.get("/api/departments")
def get_departments(org_id: Optional[int] = None, request: Request = None):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        org_id = user["org_id"]
    conn = get_db()
    cursor = conn.cursor()
    if org_id:
        cursor.execute("SELECT d.*, o.name as org_name, o.code as org_code FROM departments d JOIN organizations o ON d.org_id = o.id WHERE d.org_id = ? ORDER BY d.name ASC", (org_id,))
    else:
        cursor.execute("SELECT d.*, o.name as org_name, o.code as org_code FROM departments d JOIN organizations o ON d.org_id = o.id ORDER BY o.name, d.name ASC")
    depts = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return depts


@app.put("/api/departments/{dept_id}")
def update_department(
    dept_id: int,
    name: str = Form(...),
    default_frequency: str = Form("QUARTERLY"),
    has_double_badge_default: bool = Form(False)
):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE departments 
        SET name = ?, default_frequency = ?, has_double_badge_default = ?
        WHERE id = ?
    """, (name.strip(), default_frequency.upper(), 1 if has_double_badge_default else 0, dept_id))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "تم تحديث إعدادات القسم بنجاح"}

@app.post("/api/departments")
def create_department(org_id: int = Form(...), name: str = Form(...), default_frequency: str = Form("QUARTERLY")):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("INSERT INTO departments (org_id, name, default_frequency) VALUES (?, ?, ?)", (org_id, name, default_frequency))
    conn.commit()
    new_id = cursor.lastrowid
    conn.close()
    return {"status": "success", "id": new_id}

# --- 3. Subscribers APIs (CRUD, Duplicate Check, Resignation) ---
@app.get("/api/subscribers")
def get_subscribers(
    org_id: Optional[int] = None,
    dept_id: Optional[int] = None,
    status: Optional[str] = "ACTIVE",
    search: Optional[str] = None,
    request: Request = None
):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        org_id = user["org_id"]
    conn = get_db()
    cursor = conn.cursor()
    query = """
    SELECT s.*, d.name as dept_name, o.name as org_name, o.id as org_id
    FROM subscribers s 
    JOIN departments d ON s.dept_id = d.id 
    JOIN organizations o ON d.org_id = o.id
    WHERE 1=1
    """
    params = []
    if org_id:
        query += " AND o.id = ?"
        params.append(org_id)
    if dept_id:
        query += " AND d.id = ?"
        params.append(dept_id)
    if status and status.upper() != "ALL":
        query += " AND s.status = ?"
        params.append(status.upper())
    if search:
        s_term = f"%{search.strip()}%"
        query += " AND (s.name LIKE ? OR s.national_id LIKE ? OR s.system_code LIKE ? OR s.current_serial_oa LIKE ? OR s.current_serial_ua LIKE ?)"
        params.extend([s_term, s_term, s_term, s_term, s_term])
        
    query += " ORDER BY s.id DESC"
    cursor.execute(query, params)
    subs = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return subs

@app.post("/api/subscribers/check-duplicate")
def check_duplicate_subscriber(name: str = Form(...)):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT s.id, s.name, s.system_code, s.status, d.name as dept_name, o.name as org_name
    FROM subscribers s
    JOIN departments d ON s.dept_id = d.id
    JOIN organizations o ON d.org_id = o.id
    """)
    all_subs = [dict(r) for r in cursor.fetchall()]
    conn.close()
    
    matches = find_duplicate_names(name, all_subs, threshold=0.82)
    return {
        "has_duplicates": len(matches) > 0,
        "matches": matches
    }

@app.post("/api/subscribers")
def create_subscriber(
    dept_id: int = Form(...),
    name: str = Form(...),
    national_id: str = Form(""),
    job_title: str = Form(""),
    started_date: str = Form(""),
    has_double_badge: bool = Form(False),
    current_serial_oa: str = Form(""),
    current_serial_ua: str = Form("")
):
    conn = get_db()
    cursor = conn.cursor()
    sys_code = generate_new_system_code(cursor)
    
    cursor.execute("""
    INSERT INTO subscribers (dept_id, name, system_code, national_id, job_title, started_date, has_double_badge, current_serial_oa, current_serial_ua, status)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'ACTIVE')
    """, (dept_id, name.strip(), sys_code, national_id.strip(), job_title.strip(), started_date.strip(), 1 if has_double_badge else 0, current_serial_oa.strip(), current_serial_ua.strip()))
    new_id = cursor.lastrowid
    
    # Auto-assign serials to current year/period if provided
    now = datetime.now()
    year = now.year
    month = now.month
    if current_serial_oa.strip():
        pos = "OA" if has_double_badge else "CHEST"
        cursor.execute("""
        INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
        VALUES (?, ?, 'MONTHLY', ?, ?, ?)
        """, (new_id, year, month, current_serial_oa.strip(), pos))
        cursor.execute("""
        INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
        VALUES (?, ?, 'QUARTERLY', ?, ?, ?)
        """, (new_id, year, (month - 1) // 3 + 1, current_serial_oa.strip(), pos))
        
    if current_serial_ua.strip() and has_double_badge:
        cursor.execute("""
        INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
        VALUES (?, ?, 'MONTHLY', ?, ?, 'UA')
        """, (new_id, year, month, current_serial_ua.strip()))
        cursor.execute("""
        INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
        VALUES (?, ?, 'QUARTERLY', ?, ?, 'UA')
        """, (new_id, year, (month - 1) // 3 + 1, current_serial_ua.strip()))

    conn.commit()
    conn.close()
    return {"status": "success", "id": new_id, "system_code": sys_code}


@app.post("/api/subscribers/bulk-delete")
def bulk_delete_subscribers(subscriber_ids: List[int] = Body(...)):
    if not subscriber_ids:
        raise HTTPException(status_code=400, detail="لم يتم تحديد أي مشتركين للحذف")
    conn = get_db()
    cursor = conn.cursor()
    placeholders = ','.join(['?'] * len(subscriber_ids))
    cursor.execute(f"DELETE FROM calculated_doses WHERE subscriber_id IN ({placeholders})", subscriber_ids)
    cursor.execute(f"DELETE FROM badge_assignments WHERE subscriber_id IN ({placeholders})", subscriber_ids)
    cursor.execute(f"DELETE FROM badge_receptions WHERE subscriber_id IN ({placeholders})", subscriber_ids)
    cursor.execute(f"DELETE FROM subscribers WHERE id IN ({placeholders})", subscriber_ids)
    conn.commit()
    conn.close()
    return {"status": "success", "message": f"تم حذف {len(subscriber_ids)} مشترك بنجاح مع كافة سجلاتهم"}

@app.delete("/api/subscribers/{sub_id}")
def delete_single_subscriber(sub_id: int):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM calculated_doses WHERE subscriber_id = ?", (sub_id,))
    cursor.execute("DELETE FROM badge_assignments WHERE subscriber_id = ?", (sub_id,))
    cursor.execute("DELETE FROM badge_receptions WHERE subscriber_id = ?", (sub_id,))
    cursor.execute("DELETE FROM subscribers WHERE id = ?", (sub_id,))
    conn.commit()
    conn.close()
    return {"status": "success", "message": f"تم حذف المشترك رقم {sub_id} بنجاح"}

@app.put("/api/subscribers/{sub_id}")
def update_subscriber(
    sub_id: int,
    dept_id: int = Form(...),
    name: str = Form(...),
    national_id: str = Form(""),
    job_title: str = Form(""),
    started_date: str = Form(""),
    has_double_badge: bool = Form(False),
    current_serial_oa: str = Form(""),
    current_serial_ua: str = Form("")
):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    UPDATE subscribers 
    SET dept_id = ?, name = ?, national_id = ?, job_title = ?, started_date = ?, has_double_badge = ?, current_serial_oa = ?, current_serial_ua = ?
    WHERE id = ?
    """, (dept_id, name.strip(), national_id.strip(), job_title.strip(), started_date.strip(), 1 if has_double_badge else 0, current_serial_oa.strip(), current_serial_ua.strip(), sub_id))
    
    # Also update assignments
    now = datetime.now()
    year = now.year
    month = now.month
    if current_serial_oa.strip():
        pos = "OA" if has_double_badge else "CHEST"
        cursor.execute("""
        INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
        VALUES (?, ?, 'MONTHLY', ?, ?, ?)
        """, (sub_id, year, month, current_serial_oa.strip(), pos))
        cursor.execute("""
        INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
        VALUES (?, ?, 'QUARTERLY', ?, ?, ?)
        """, (sub_id, year, (month - 1) // 3 + 1, current_serial_oa.strip(), pos))
        
    if current_serial_ua.strip() and has_double_badge:
        cursor.execute("""
        INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
        VALUES (?, ?, 'MONTHLY', ?, ?, 'UA')
        """, (sub_id, year, month, current_serial_ua.strip()))
        cursor.execute("""
        INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
        VALUES (?, ?, 'QUARTERLY', ?, ?, 'UA')
        """, (sub_id, year, (month - 1) // 3 + 1, current_serial_ua.strip()))

    conn.commit()
    conn.close()
    return {"status": "success"}

@app.post("/api/subscribers/{sub_id}/quick-serial")
def quick_update_subscriber_serial(
    sub_id: int,
    serial_type: str = Form(...), # "OA" or "UA"
    new_serial: str = Form(...)
):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, has_double_badge, current_serial_oa, current_serial_ua FROM subscribers WHERE id = ?", (sub_id,))
    sub = cursor.fetchone()
    if not sub:
        conn.close()
        raise HTTPException(status_code=404, detail="المشترك غير موجود")
        
    clean_sn = new_serial.strip()
    if serial_type.upper() == "UA":
        cursor.execute("UPDATE subscribers SET current_serial_ua = ? WHERE id = ?", (clean_sn, sub_id))
    else:
        cursor.execute("UPDATE subscribers SET current_serial_oa = ? WHERE id = ?", (clean_sn, sub_id))
        
    conn.commit()
    conn.close()
    return {
        "status": "success",
        "message": f"تم تحديث سيريال {serial_type.upper()} للمشترك بنجاح",
        "subscriber_id": sub_id,
        "serial_type": serial_type.upper(),
        "new_serial": clean_sn
    }

@app.post("/api/doses/{dose_id}/quick-serial")
def quick_update_dose_serial(
    dose_id: int,
    serial_type: str = Form(...), # "OA" or "UA"
    new_serial: str = Form(...),
    request: Request = None
):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        raise HTTPException(status_code=403, detail="غير مصرح لحساب المنشأة بتعديل سيريال الوضحة في سجل القراءات.")
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, subscriber_id FROM calculated_doses WHERE id = ?", (dose_id,))
    dose = cursor.fetchone()
    if not dose:
        conn.close()
        raise HTTPException(status_code=404, detail="سجل القراءة غير موجود")
        
    clean_sn = new_serial.strip()
    sub_id = dose["subscriber_id"]
    if serial_type.upper() == "UA":
        cursor.execute("UPDATE calculated_doses SET badge_ua_serial = ? WHERE id = ?", (clean_sn, dose_id))
        cursor.execute("UPDATE subscribers SET current_serial_ua = ? WHERE id = ?", (clean_sn, sub_id))
    else:
        cursor.execute("UPDATE calculated_doses SET badge_oa_serial = ? WHERE id = ?", (clean_sn, dose_id))
        cursor.execute("UPDATE subscribers SET current_serial_oa = ? WHERE id = ?", (clean_sn, sub_id))
        
    conn.commit()
    conn.close()
    return {
        "status": "success",
        "message": "تم تحديث السيريال في سجل القراءة وملف المشترك بنجاح",
        "dose_id": dose_id,
        "serial_type": serial_type.upper(),
        "new_serial": clean_sn
    }


@app.post("/api/subscribers/{sub_id}/resign")
def mark_subscriber_resigned(sub_id: int, resigned_date: str = Form(...), resignation_notes: str = Form("")):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
    UPDATE subscribers 
    SET status = 'RESIGNED', resigned_date = ?, resignation_notes = ?
    WHERE id = ?
    """, (resigned_date, resignation_notes, sub_id))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.post("/api/subscribers/{sub_id}/reinstate")
def reinstate_subscriber(sub_id: int):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("UPDATE subscribers SET status = 'ACTIVE', resigned_date = NULL WHERE id = ?", (sub_id,))
    conn.commit()
    conn.close()
    return {"status": "success"}

# --- 4. Custom Excel Import for Subscribers (معاينة وتعيين أعمدة حر) ---
@app.post("/api/subscribers/preview-custom-excel")
async def preview_custom_subs_excel(file: UploadFile = File(...)):
    dest_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(dest_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    try:
        preview = preview_file(dest_path)
        preview["saved_filename"] = file.filename
        return preview
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"خطأ في قراءة ملف المشتركين: {str(e)}")

@app.post("/api/subscribers/import-custom-excel")
def import_custom_subs_excel(
    filename: str = Form(...),
    default_dept_id: Optional[int] = Form(None),
    col_name: str = Form(...),
    col_national_id: Optional[str] = Form(None),
    col_job: Optional[str] = Form(None),
    col_hospital: Optional[str] = Form(None),
    col_dept: Optional[str] = Form(None),
    col_badge_type: Optional[str] = Form(None),
    col_serial_oa: Optional[str] = Form(None),
    col_serial_ua: Optional[str] = Form(None),
    col_started_date: Optional[str] = Form(None)
):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        raise HTTPException(status_code=403, detail="غير مصرح لحساب المنشأة باعتماد واستيراد ملفات القارئ.")

    file_path = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="الملف غير موجود")
        
    df = universal_read_file(file_path)
    df.columns = [str(c).strip() for c in df.columns]

    conn = get_db()
    cursor = conn.cursor()

    imported = 0
    assigned = 0
    now = datetime.now()
    year = now.year
    month = now.month
    quarter = (month - 1) // 3 + 1

    for _, row in df.iterrows():
        name = str(row.get(col_name, "")).strip()
        if not name or name.lower() in ('nan', 'none', ''):
            continue

        nat_id = str(row.get(col_national_id, "")).strip() if col_national_id else ""
        if nat_id.lower() == 'nan': nat_id = ''
        
        job = str(row.get(col_job, "")).strip() if col_job else ""
        if job.lower() == 'nan': job = ''

        start_date = str(row.get(col_started_date, "")).strip() if col_started_date else ""
        if start_date.lower() == 'nan': start_date = ''

        type_str = str(row.get(col_badge_type, "")).lower() if col_badge_type else ""
        has_double = 1 if "double" in type_str or "وضحتين" in type_str or "2" in type_str or "خارجية" in type_str else 0

        # Hospital & Dept
        target_dept_id = default_dept_id
        if col_hospital and col_dept and str(row.get(col_hospital, '')).strip() and str(row.get(col_dept, '')).strip():
            hosp_name = str(row[col_hospital]).strip()
            dept_name = str(row[col_dept]).strip()
            cursor.execute("SELECT id FROM organizations WHERE name = ?", (hosp_name,))
            h_row = cursor.fetchone()
            if h_row:
                h_id = h_row['id']
            else:
                cursor.execute("INSERT INTO organizations (name, code) VALUES (?, ?)", (hosp_name, f"ORG-{imported+1}"))
                h_id = cursor.lastrowid

            cursor.execute("SELECT id FROM departments WHERE org_id = ? AND name = ?", (h_id, dept_name))
            d_row = cursor.fetchone()
            if d_row:
                target_dept_id = d_row['id']
            else:
                cursor.execute("INSERT INTO departments (org_id, name) VALUES (?, ?)", (h_id, dept_name))
                target_dept_id = cursor.lastrowid
                
        if not target_dept_id:
            # Fallback to first department
            cursor.execute("SELECT id FROM departments LIMIT 1")
            first_d = cursor.fetchone()
            target_dept_id = first_d['id'] if first_d else 1

        sn_oa = str(row.get(col_serial_oa, "")).strip() if col_serial_oa else ""
        if sn_oa.lower() == 'nan': sn_oa = ''
        
        sn_ua = str(row.get(col_serial_ua, "")).strip() if col_serial_ua else ""
        if sn_ua.lower() == 'nan': sn_ua = ''
        if sn_ua: has_double = 1

        # Check existing
        cursor.execute("SELECT id FROM subscribers WHERE dept_id = ? AND name = ?", (target_dept_id, name))
        existing_sub = cursor.fetchone()
        if existing_sub:
            sub_id = existing_sub['id']
            cursor.execute("""
            UPDATE subscribers SET national_id = ?, job_title = ?, started_date = ?, has_double_badge = ?, current_serial_oa = ?, current_serial_ua = ?
            WHERE id = ?
            """, (nat_id, job, start_date, has_double, sn_oa, sn_ua, sub_id))
        else:
            sys_code = generate_new_system_code(cursor)
            cursor.execute("""
            INSERT INTO subscribers (dept_id, name, system_code, national_id, job_title, started_date, has_double_badge, current_serial_oa, current_serial_ua, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'ACTIVE')
            """, (target_dept_id, name, sys_code, nat_id, job, start_date, has_double, sn_oa, sn_ua))
            sub_id = cursor.lastrowid
            imported += 1

        # Badge assignments
        if sn_oa:
            pos = "OA" if has_double else "CHEST"
            cursor.execute("""
            INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
            VALUES (?, ?, 'MONTHLY', ?, ?, ?)
            """, (sub_id, year, month, sn_oa, pos))
            cursor.execute("""
            INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
            VALUES (?, ?, 'QUARTERLY', ?, ?, ?)
            """, (sub_id, year, quarter, sn_oa, pos))
            assigned += 1
            
        if sn_ua and has_double:
            cursor.execute("""
            INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
            VALUES (?, ?, 'MONTHLY', ?, ?, 'UA')
            """, (sub_id, year, month, sn_ua))
            cursor.execute("""
            INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
            VALUES (?, ?, 'QUARTERLY', ?, ?, 'UA')
            """, (sub_id, year, quarter, sn_ua))
            assigned += 1

    conn.commit()
    conn.close()
    return {"status": "success", "added_count": imported, "imported_count": imported, "imported_subscribers": imported, "assigned_count": assigned, "message": f"تم استيراد {imported} مشترك بنجاح"}

# --- 5. Download & Upload Standard Template ---
@app.get("/api/templates/download-subscribers-excel")
def download_subscribers_template():
    template_file = os.path.join(TEMPLATES_DIR, "نموذج_تعبئة_المشتركين.xlsx")
    return FileResponse(template_file, filename="نموذج_تعبئة_المشتركين.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.post("/api/subscribers/upload-template-excel")
async def upload_subscribers_template(file: UploadFile = File(...)):
    dest_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(dest_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
        
    df = universal_read_file(dest_path)
    df.columns = [str(c).strip() for c in df.columns]

    conn = get_db()
    cursor = conn.cursor()

    imported_count = 0
    assigned_count = 0
    now = datetime.now()
    year = now.year
    month = now.month
    quarter = (month - 1) // 3 + 1

    for _, row in df.iterrows():
        name = str(row.get("اسم المشترك (رباعي)", row.get("اسم المشترك", row.get("الاسم", "")))).strip()
        if not name or name.lower() in ('nan', 'none', ''):
            continue

        nat_id = str(row.get("الرقم الوطني / الوظيفي", row.get("الرقم الوطني", ""))).strip()
        if nat_id.lower() == 'nan': nat_id = ''
        job = str(row.get("المسمى الوظيفي", row.get("الوظيفة", ""))).strip()
        if job.lower() == 'nan': job = ''
        org_name = str(row.get("اسم المستشفى / المركز", row.get("المركز", "المركز الرئيسي"))).strip()
        if org_name.lower() == 'nan' or not org_name: org_name = 'المركز الرئيسي'
        org_code = str(row.get("كود المركز", "C-01")).strip()
        if org_code.lower() == 'nan' or not org_code: org_code = 'C-01'
        dept_name = str(row.get("اسم القسم", "القسم العام")).strip()
        if dept_name.lower() == 'nan' or not dept_name: dept_name = 'القسم العام'
        freq_str = str(row.get("نوع دورة القسم (شهري / دوري)", "دوري")).strip()
        default_freq = "MONTHLY" if "شهر" in freq_str else "QUARTERLY"
        type_str = str(row.get("نوع الارتداء (وضحة واحدة / وضحتين)", "وضحة واحدة")).strip()
        has_double = 1 if "وضحتين" in type_str or "2" in type_str or "خارجية" in type_str else 0
        start_date = str(row.get("تاريخ بدء الاشتراك (DD/MM/YYYY)", "")).strip()
        if start_date.lower() == 'nan': start_date = ''

        cursor.execute("SELECT id FROM organizations WHERE name = ? OR code = ?", (org_name, org_code))
        org_row = cursor.fetchone()
        org_id = org_row['id'] if org_row else cursor.execute("INSERT INTO organizations (name, code) VALUES (?, ?)", (org_name, org_code)).lastrowid

        cursor.execute("SELECT id FROM departments WHERE org_id = ? AND name = ?", (org_id, dept_name))
        dept_row = cursor.fetchone()
        dept_id = dept_row['id'] if dept_row else cursor.execute("INSERT INTO departments (org_id, name, default_frequency) VALUES (?, ?, ?)", (org_id, dept_name, default_freq)).lastrowid

        sn_oa = str(row.get("الرقم التسلسلي للوضحة الصدرية أو الخارجية (OA)", "")).strip()
        if sn_oa.lower() == 'nan': sn_oa = ''
        sn_ua = str(row.get("الرقم التسلسلي للوضحة الداخلية (UA إن وجدت)", "")).strip()
        if sn_ua.lower() == 'nan': sn_ua = ''
        if sn_ua: has_double = 1

        cursor.execute("SELECT id FROM subscribers WHERE dept_id = ? AND name = ?", (dept_id, name))
        sub_row = cursor.fetchone()
        if sub_row:
            sub_id = sub_row['id']
            cursor.execute("""
            UPDATE subscribers SET national_id = ?, job_title = ?, started_date = ?, has_double_badge = ?, current_serial_oa = ?, current_serial_ua = ? WHERE id = ?
            """, (nat_id, job, start_date, has_double, sn_oa, sn_ua, sub_id))
        else:
            sys_code = generate_new_system_code(cursor)
            cursor.execute("""
            INSERT INTO subscribers (dept_id, name, system_code, national_id, job_title, started_date, has_double_badge, current_serial_oa, current_serial_ua, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'ACTIVE')
            """, (dept_id, name, sys_code, nat_id, job, start_date, has_double, sn_oa, sn_ua))
            sub_id = cursor.lastrowid
            imported_count += 1

        if sn_oa:
            pos = "OA" if has_double else "CHEST"
            cursor.execute("""
            INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
            VALUES (?, ?, 'MONTHLY', ?, ?, ?)
            """, (sub_id, year, month, sn_oa, pos))
            cursor.execute("""
            INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
            VALUES (?, ?, 'QUARTERLY', ?, ?, ?)
            """, (sub_id, year, quarter, sn_oa, pos))
            assigned_count += 1

        if sn_ua and has_double:
            cursor.execute("""
            INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
            VALUES (?, ?, 'MONTHLY', ?, ?, 'UA')
            """, (sub_id, year, month, sn_ua))
            cursor.execute("""
            INSERT OR REPLACE INTO badge_assignments (subscriber_id, period_year, period_type, period_number, serial_number, badge_position)
            VALUES (?, ?, 'QUARTERLY', ?, ?, 'UA')
            """, (sub_id, year, quarter, sn_ua))
            assigned_count += 1

    conn.commit()
    conn.close()

    return {
        "status": "success",
        "added_count": imported_count,
        "imported_subscribers": imported_count,
        "imported_count": imported_count,
        "assigned_badges": assigned_count,
        "assigned_count": assigned_count,
        "message": f"تم استيراد {imported_count} مشترك بنجاح"
    }


# --- 5.5 Badge Reception Manifest APIs (كشف استلام وتسليم الوضحات) ---
@app.get("/api/reception/manifest")
def get_reception_manifest(
    dept_id: int,
    period_year: int = 2024,
    period_type: str = "MONTHLY",
    period_number: int = 1,
    request: Request = None
):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        chk_conn = get_db()
        chk_c = chk_conn.cursor()
        chk_c.execute("SELECT org_id FROM departments WHERE id = ?", (dept_id,))
        dept_row = chk_c.fetchone()
        chk_conn.close()
        if not dept_row or dept_row["org_id"] != user["org_id"]:
            raise HTTPException(status_code=403, detail="غير مصرح لك بعرض كشف قسم تابع لمنشأة أخرى")

    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT o.name as org_name, o.code as org_code, d.name as dept_name
        FROM departments d
        JOIN organizations o ON d.org_id = o.id
        WHERE d.id = ?
    """, (dept_id,))
    dept_info = cursor.fetchone()
    if not dept_info:
        conn.close()
        raise HTTPException(status_code=404, detail="القسم غير موجود")
        
    cursor.execute("""
        SELECT s.id as subscriber_id, s.name, s.system_code, s.national_id, s.job_title,
               s.started_date, s.has_double_badge, s.current_serial_oa, s.current_serial_ua,
               br.id as reception_id, br.reception_date, br.received_oa, br.received_ua,
               br.status as rec_status, br.notes as rec_notes, br.receiver_name, br.delivered_by
        FROM subscribers s
        LEFT JOIN badge_receptions br ON s.id = br.subscriber_id
             AND br.period_year = ? AND br.period_type = ? AND br.period_number = ?
        WHERE s.dept_id = ? AND s.status = 'ACTIVE'
        ORDER BY s.name ASC
    """, (period_year, period_type.upper(), period_number, dept_id))
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    
    for r in rows:
        if r['reception_id'] is None:
            r['received_oa'] = 1
            r['received_ua'] = 1 if r['has_double_badge'] else 0
            r['rec_status'] = 'RECEIVED'
            r['reception_date'] = datetime.today().strftime('%d/%m/%Y')
            r['rec_notes'] = ''
            
    return {
        "org_name": dept_info['org_name'],
        "org_code": dept_info['org_code'],
        "dept_name": dept_info['dept_name'],
        "period_year": period_year,
        "period_type": period_type.upper(),
        "period_number": period_number,
        "subscribers": rows
    }

@app.post("/api/reception/save")
def save_reception_manifest(
    dept_id: int = Form(...),
    period_year: int = Form(...),
    period_type: str = Form(...),
    period_number: int = Form(...),
    reception_date: str = Form(...),
    receiver_name: Optional[str] = Form(None),
    delivered_by: Optional[str] = Form(None),
    items_json: str = Form(...)
):
    conn = get_db()
    cursor = conn.cursor()
    items = json.loads(items_json) if items_json else []
    
    for it in items:
        sub_id = it['subscriber_id']
        rec_oa = 1 if it.get('received_oa', True) else 0
        rec_ua = 1 if it.get('received_ua', True) else 0
        stat = it.get('status', 'RECEIVED')
        notes = it.get('notes', '')
        
        cursor.execute("""
            INSERT INTO badge_receptions (
                subscriber_id, dept_id, period_year, period_type, period_number,
                reception_date, received_oa, received_ua, status, notes, receiver_name, delivered_by
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(subscriber_id, period_year, period_type, period_number) DO UPDATE SET
                reception_date = excluded.reception_date,
                received_oa = excluded.received_oa,
                received_ua = excluded.received_ua,
                status = excluded.status,
                notes = excluded.notes,
                receiver_name = excluded.receiver_name,
                delivered_by = excluded.delivered_by
        """, (
            sub_id, dept_id, period_year, period_type.upper(), period_number,
            reception_date, rec_oa, rec_ua, stat, notes, receiver_name or '', delivered_by or ''
        ))
        
        # Also update reception_date in calculated_doses if a dose record exists
        cursor.execute("""
            UPDATE calculated_doses
            SET reception_date = ?
            WHERE subscriber_id = ? AND period_year = ? AND period_type = ? AND period_number = ?
        """, (reception_date, sub_id, period_year, period_type.upper(), period_number))
        
    conn.commit()
    conn.close()
    return {"status": "success", "message": f"تم حفظ كشف الاستلام لـ {len(items)} مشترك بنجاح"}

@app.get("/api/reception/export-excel")
def export_reception_excel(dept_id: int, period_year: int = 2024, period_type: str = "MONTHLY", period_number: int = 1):
    filename = f"Reception_Manifest_{period_year}_{period_type}_{period_number}_dept_{dept_id}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, filename)
    generate_reception_manifest_excel,
    generate_annual_archive_excel(dept_id, period_year, period_type, period_number, out_path)
    return FileResponse(out_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- 6. Reader Import & Smart Reconciliation ---
@app.post("/api/readings/preview")
async def preview_reader_file(file: UploadFile = File(...)):
    dest_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(dest_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
        
    try:
        preview = preview_file(dest_path)
        preview["saved_filename"] = file.filename
        return preview
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"خطأ في قراءة ملف الجهاز: {str(e)}")

@app.post("/api/readings/analyze")
def analyze_readings(
    filename: str = Form(...),
    period_year: int = Form(...),
    period_type: str = Form(...),
    period_number: int = Form(...),
    col_serial: str = Form(...),
    col_read_date: Optional[str] = Form(None),
    col_hp10: str = Form(...),
    col_hp07: str = Form(...)
):
    file_path = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="الملف غير موجود")
        
    try:
        analysis = analyze_reader_file(
            file_path=file_path,
            period_year=period_year,
            period_type=period_type.upper(),
            period_number=period_number,
            col_serial=col_serial,
            col_read_date=col_read_date,
            col_hp10=col_hp10,
            col_hp07=col_hp07
        )
        return analysis
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"خطأ أثناء فحص البيانات: {str(e)}")

@app.post("/api/readings/commit")
def commit_readings(
    filename: str = Form(...),
    request: Request = None,
    period_year: int = Form(...),
    period_type: str = Form(...),
    period_number: int = Form(...),
    col_serial: str = Form(...),
    col_read_date: Optional[str] = Form(None),
    col_hp10: str = Form(...),
    col_hp07: str = Form(...),
    duplicate_resolutions: str = Form("{}"),
    unassigned_resolutions: str = Form("[]"),
    reception_date: Optional[str] = Form(None),
    report_date: Optional[str] = Form(None),
    monthly_period_number: Optional[int] = Form(None)
):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        raise HTTPException(status_code=403, detail="غير مصرح لحساب المنشأة باعتماد واستيراد ملفات القارئ.")

    file_path = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="الملف غير موجود")
        
    dup_res = json.loads(duplicate_resolutions) if duplicate_resolutions else {}
    un_res = json.loads(unassigned_resolutions) if unassigned_resolutions else []
    
    try:
        result = commit_reader_import(
            file_path=file_path,
            period_year=period_year,
            period_type=period_type.upper(),
            period_number=period_number,
            col_serial=col_serial,
            col_read_date=col_read_date,
            col_hp10=col_hp10,
            col_hp07=col_hp07,
            duplicate_resolutions=dup_res,
            unassigned_resolutions=un_res,
            reception_date=reception_date,
            report_date=report_date,
            monthly_period_number=monthly_period_number
        )
        return result
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


# --- 6.5 Dose Records Review & Management APIs ---
@app.get("/api/doses")
def get_doses(
    dept_id: Optional[int] = None,
    org_id: Optional[int] = None,
    period_year: int = 2024,
    period_type: str = "MONTHLY",
    period_number: int = 1,
    request: Request = None
):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        org_id = user["org_id"]
    conn = get_db()
    cursor = conn.cursor()
    
    query = """
    SELECT s.id as subscriber_id, s.name, s.system_code, s.national_id, s.job_title,
           s.started_date, s.has_double_badge, s.status, s.current_serial_oa, s.current_serial_ua,
           d.id as dept_id, d.name as dept_name, o.id as org_id, o.name as org_name,
           cd.id as dose_id, cd.hp10, cd.hp07, cd.accumulated_hp10, cd.accumulated_hp07,
           cd.is_double_badge as dose_is_double, cd.badge_oa_serial, cd.badge_ua_serial,
           cd.reception_date, cd.reading_date, cd.report_date, cd.remark,
           (CASE WHEN cd.id IS NOT NULL THEN 1 ELSE 0 END) as is_recorded
    FROM subscribers s
    JOIN departments d ON s.dept_id = d.id
    JOIN organizations o ON d.org_id = o.id
    LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id 
         AND cd.period_year = ? AND cd.period_type = ? AND cd.period_number = ?
    WHERE s.status = 'ACTIVE'
    """
    params = [period_year, period_type.upper(), period_number]
    
    if dept_id:
        query += " AND s.dept_id = ?"
        params.append(dept_id)
    elif org_id:
        query += " AND o.id = ?"
        params.append(org_id)
        
    query += " ORDER BY o.name ASC, d.name ASC, s.name ASC"
    
    cursor.execute(query, params)
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    
    total_subs = len(rows)
    recorded_count = sum(1 for r in rows if r['is_recorded'] and r['remark'] != 'NR')
    nr_count = sum(1 for r in rows if (not r['is_recorded']) or (r['remark'] and 'NR' in r['remark']))
    doses_vals = [r['hp10'] for r in rows if r['is_recorded'] and r['hp10'] is not None and r['remark'] != 'NR']
    max_hp10 = max(doses_vals) if doses_vals else 0.0
    avg_hp10 = round(sum(doses_vals) / len(doses_vals), 2) if doses_vals else 0.0
    threshold = 1.67 if period_type.upper() == 'MONTHLY' else 5.0
    high_dose_count = sum(1 for r in rows if r['is_recorded'] and r['hp10'] is not None and r['hp10'] > threshold)
    
    return {
        "period_year": period_year,
        "period_type": period_type.upper(),
        "period_number": period_number,
        "summary": {
            "total_subscribers": total_subs,
            "recorded_count": recorded_count,
            "nr_count": nr_count,
            "max_hp10": round(max_hp10, 2),
            "avg_hp10": avg_hp10,
            "high_dose_alerts": high_dose_count,
            "threshold": threshold
        },
        "readings": rows
    }

@app.post("/api/doses/manual")
def save_manual_dose(
    subscriber_id: int = Form(...),
    request: Request = None,
    period_year: int = Form(...),
    period_type: str = Form(...),
    period_number: int = Form(...),
    hp10: float = Form(...),
    hp07: float = Form(...),
    badge_oa_serial: Optional[str] = Form(None),
    badge_ua_serial: Optional[str] = Form(None),
    reading_date: Optional[str] = Form(None),
    reception_date: Optional[str] = Form(None),
    report_date: Optional[str] = Form(None),
    remark: Optional[str] = Form(None)
):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        raise HTTPException(status_code=403, detail="غير مصرح لحساب المنشأة بتعديل أو إدخال القراءات الإشعاعية. إدخال وتعديل القراءات محصور بمدير المختبر المركزي.")

    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute("SELECT has_double_badge, current_serial_oa, current_serial_ua FROM subscribers WHERE id = ?", (subscriber_id,))
    sub = cursor.fetchone()
    if not sub:
        conn.close()
        raise HTTPException(status_code=404, detail="المشترك غير موجود")
        
    is_double = bool(sub['has_double_badge'])
    oa_sn = (badge_oa_serial if badge_oa_serial is not None else sub['current_serial_oa']) or ''
    ua_sn = (badge_ua_serial if badge_ua_serial is not None else sub['current_serial_ua']) or ''
    
    acc_hp10, acc_hp07 = compute_accumulated_dose(
        conn, subscriber_id, period_year, period_type.upper(), period_number, hp10, hp07
    )
    rep_date = report_date or datetime.today().strftime('%d/%m/%Y')
    read_date = reading_date or datetime.today().strftime('%d/%m/%Y')
    rec_date = reception_date or read_date
    
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
        subscriber_id, period_year, period_type.upper(), period_number,
        hp10, hp07, acc_hp10, acc_hp07,
        1 if is_double else 0, oa_sn, ua_sn,
        rec_date, read_date, rep_date, remark or ''
    ))
    
    conn.commit()
    conn.close()
    return {"status": "success", "message": "تم حفظ القراءة بنجاح"}

@app.delete("/api/doses/{dose_id}")
def delete_dose(dose_id: int, request: Request = None):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        raise HTTPException(status_code=403, detail="غير مصرح لحساب المنشأة بحذف القراءات الإشعاعية.")
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM calculated_doses WHERE id = ?", (dose_id,))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "تم حذف القراءة بنجاح"}

@app.get("/api/unassigned-readings")
def get_unassigned_readings():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM unassigned_readings ORDER BY id DESC")
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return rows

# --- 7. Report Exports (with Department and Hospital Filter) ---
@app.get("/api/reports/monthly/excel")
def export_monthly_excel(dept_id: int, year: int, month: int):
    filename = f"Report_Monthly_{year}_{month}_dept_{dept_id}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, filename)
    generate_monthly_excel_report(dept_id, year, month, out_path)
    return FileResponse(out_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/reports/quarterly/excel")
def export_quarterly_excel(dept_id: int, year: int, quarter: int):
    filename = f"Report_Quarterly_{year}_Q{quarter}_dept_{dept_id}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, filename)
    generate_quarterly_excel_report(dept_id, year, quarter, out_path)
    return FileResponse(out_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.get("/api/reports/monthly/word")
def export_monthly_word(dept_id: int, year: int, month: int):
    filename = f"Report_Monthly_{year}_{month}_dept_{dept_id}.docx"
    out_path = os.path.join(OUTPUT_DIR, filename)
    generate_monthly_word_report(dept_id, year, month, out_path)
    return FileResponse(out_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.get("/api/reports/quarterly/word")
def export_quarterly_word(dept_id: int, year: int, quarter: int):
    filename = f"Report_Quarterly_{year}_Q{quarter}_dept_{dept_id}.docx"
    out_path = os.path.join(OUTPUT_DIR, filename)
    generate_quarterly_word_report(dept_id, year, quarter, out_path)
    return FileResponse(out_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.get("/api/reports/monthly/pdf")
def export_monthly_pdf(dept_id: int, year: int, month: int):
    from app.services.reporter import generate_official_html_report, convert_html_to_pdf_robust
    html_content = generate_official_html_report(dept_id, year, "MONTHLY", month)
    filename_pdf = f"Report_Monthly_{year}_M{month}_dept_{dept_id}.pdf"
    pdf_path = os.path.join(OUTPUT_DIR, filename_pdf)
    
    success = convert_html_to_pdf_robust(html_content, pdf_path)
    if success and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 500:
        return FileResponse(pdf_path, filename=filename_pdf, media_type="application/pdf")
    else:
        return HTMLResponse(content=html_content + "<script>window.onload = function() { window.print(); };</script>")

@app.get("/api/reports/quarterly/pdf")
def export_quarterly_pdf(dept_id: int, year: int, quarter: int):
    from app.services.reporter import generate_official_html_report, convert_html_to_pdf_robust
    html_content = generate_official_html_report(dept_id, year, "QUARTERLY", quarter)
    filename_pdf = f"Report_Quarterly_{year}_Q{quarter}_dept_{dept_id}.pdf"
    pdf_path = os.path.join(OUTPUT_DIR, filename_pdf)
    
    success = convert_html_to_pdf_robust(html_content, pdf_path)
    if success and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 500:
        return FileResponse(pdf_path, filename=filename_pdf, media_type="application/pdf")
    else:
        return HTMLResponse(content=html_content + "<script>window.onload = function() { window.print(); };</script>")

@app.get("/api/reports/printable")
def export_printable_report(dept_id: int, year: int, period_type: str, period_number: int, auto_print: int = 1):
    from app.services.reporter import generate_official_html_report
    html_content = generate_official_html_report(dept_id, year, period_type.upper(), period_number)
    if auto_print:
        html_content += "<script>window.onload = function() { window.print(); };</script>"
    return HTMLResponse(content=html_content)

# --- 7.5 Historical Archive APIs ---
@app.get("/api/archive/years")
def get_archive_years():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT period_year FROM calculated_doses ORDER BY period_year DESC")
    years = [r[0] for r in cursor.fetchall()]
    cursor.execute("SELECT value FROM system_settings WHERE key = 'active_year'")
    row = cursor.fetchone()
    active_y = int(row[0]) if row else 2026
    if active_y not in years:
        years.insert(0, active_y)
    if not years:
        years = [2026, 2024]
    conn.close()
    return {"years": years, "active_year": active_y}

@app.get("/api/archive/summary")
def get_archive_summary(
    year: int,
    org_id: Optional[int] = None,
    dept_id: Optional[int] = None,
    search: Optional[str] = None,
    request: Request = None
):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL":
        org_id = user["org_id"]
    conn = get_db()
    cursor = conn.cursor()
    
    query = """
    SELECT s.id as subscriber_id, s.name, s.system_code, s.job_title, s.has_double_badge,
           s.current_serial_oa, s.current_serial_ua,
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
        
    if search:
        query += " AND (s.name LIKE ? OR s.system_code LIKE ? OR s.current_serial_oa LIKE ?)"
        term = f"%{search.strip()}%"
        params.extend([term, term, term])
        
    query += " ORDER BY o.name ASC, d.name ASC, s.name ASC"
    cursor.execute(query, params)
    subs = [dict(r) for r in cursor.fetchall()]
    
    total_subscribers = len(subs)
    total_annual_dose_all = 0.0
    max_annual_dose = 0.0
    exceeded_limit_count = 0
    
    results = []
    for s in subs:
        sub_id = s['subscriber_id']
        cursor.execute("""
            SELECT period_number, hp10, hp07, remark, reading_date
            FROM calculated_doses
            WHERE subscriber_id = ? AND period_year = ?
            ORDER BY period_number ASC
        """, (sub_id, year))
        doses_map = {r['period_number']: dict(r) for r in cursor.fetchall()}
        
        annual_hp10 = sum(d['hp10'] for d in doses_map.values() if d['hp10'] is not None)
        annual_hp07 = sum(d['hp07'] for d in doses_map.values() if d['hp07'] is not None)
        
        annual_hp10 = round(annual_hp10, 2)
        annual_hp07 = round(annual_hp07, 2)
        
        if annual_hp10 > max_annual_dose:
            max_annual_dose = annual_hp10
        total_annual_dose_all += annual_hp10
        
        if annual_hp10 >= 20.0:
            exceeded_limit_count += 1
            
        results.append({
            "subscriber_id": sub_id,
            "name": s['name'],
            "system_code": s['system_code'],
            "job_title": s['job_title'],
            "org_name": s['org_name'],
            "org_code": s['org_code'],
            "dept_name": s['dept_name'],
            "default_frequency": s['default_frequency'],
            "has_double_badge": bool(s['has_double_badge']),
            "current_serial_oa": s['current_serial_oa'],
            "current_serial_ua": s['current_serial_ua'],
            "periods_doses": doses_map,
            "periods_count": len(doses_map),
            "annual_hp10": annual_hp10,
            "annual_hp07": annual_hp07,
            "is_exceeded": (annual_hp10 >= 20.0),
            "status_label": "تجاوز الحد (≥20)" if annual_hp10 >= 20.0 else "آمن"
        })
        
    conn.close()
    
    return {
        "year": year,
        "summary": {
            "total_subscribers": total_subscribers,
            "total_annual_dose": round(total_annual_dose_all, 2),
            "avg_annual_dose": round(total_annual_dose_all / total_subscribers, 2) if total_subscribers > 0 else 0.0,
            "max_annual_dose": max_annual_dose,
            "exceeded_limit_count": exceeded_limit_count
        },
        "subscribers": results
    }

@app.get("/api/archive/export-excel")
def export_archive_excel(year: int, dept_id: Optional[int] = None, org_id: Optional[int] = None):
    filename = f"Annual_Dose_Archive_{year}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, filename)
    generate_annual_archive_excel(year, org_id, dept_id, out_path)
    return FileResponse(out_path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- 7.6 Facility Summary, Drilldown & Periodic Schedule Matrix APIs ---

def parse_date_for_sort(date_str):
    import datetime
    if not date_str or not str(date_str).strip():
        return datetime.date.min
    ds = str(date_str).strip()
    parts = ds.split('/')
    if len(parts) == 3:
        try:
            return datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
        except Exception:
            pass
    parts = ds.split('-')
    if len(parts) == 3:
        try:
            return datetime.date(int(parts[0]), int(parts[1]), int(parts[2]))
        except Exception:
            pass
    return datetime.date.min

@app.get("/api/facilities/summary-readings")
def get_facilities_summary_readings(
    year: Optional[int] = None,
    sort_by: str = "latest_reading",
    request: Request = None
):
    user = resolve_user(request)
    user_org_id = user["org_id"] if (user and user.get("role") == "HOSPITAL") else None

    conn = get_db()
    cursor = conn.cursor()

    if not year:
        cursor.execute("SELECT value FROM system_settings WHERE key = 'active_year'")
        row_y = cursor.fetchone()
        year = int(row_y[0]) if row_y and row_y[0] else 2026

    where_clause = ""
    params = [year]
    if user_org_id:
        where_clause = "WHERE o.id = ?"
        params.append(user_org_id)

    query = f"""
        SELECT o.id, o.name, o.code, o.classification, o.contact, o.address,
               COUNT(DISTINCT s.id) as total_subscribers,
               COUNT(DISTINCT CASE WHEN s.status = 'ACTIVE' THEN s.id END) as active_subscribers,
               COUNT(DISTINCT CASE WHEN s.has_double_badge = 1 AND s.status = 'ACTIVE' THEN s.id END) as double_badge_subscribers,
               COUNT(DISTINCT d.id) as departments_count,
               MAX(cd.reading_date) as latest_reading_date,
               COUNT(DISTINCT cd.id) as total_doses_in_year,
               ROUND(AVG(CASE WHEN cd.remark != 'NR' THEN cd.hp10 END), 2) as avg_dose_in_year,
               MAX(cd.hp10) as max_dose_in_year,
               COUNT(DISTINCT CASE WHEN cd.hp10 > 1.67 THEN cd.id END) as high_dose_alerts
        FROM organizations o
        LEFT JOIN departments d ON o.id = d.org_id
        LEFT JOIN subscribers s ON d.id = s.dept_id
        LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id AND cd.period_year = ?
        {where_clause}
        GROUP BY o.id, o.name, o.code, o.classification, o.contact, o.address
    """
    cursor.execute(query, tuple(params))
    org_rows = [dict(r) for r in cursor.fetchall()]

    for org in org_rows:
        cursor.execute("""
            SELECT d.id, d.name, d.default_frequency,
                   COUNT(DISTINCT CASE WHEN s.status = 'ACTIVE' THEN s.id END) as active_subs,
                   MAX(cd.reading_date) as dept_latest_reading
            FROM departments d
            LEFT JOIN subscribers s ON d.id = s.dept_id
            LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id AND cd.period_year = ?
            WHERE d.org_id = ?
            GROUP BY d.id, d.name, d.default_frequency
            ORDER BY d.name ASC
        """, (year, org["id"]))
        org["departments"] = [dict(d) for d in cursor.fetchall()]

        cursor.execute("""
            SELECT cd.period_type, cd.period_number, cd.reading_date
            FROM calculated_doses cd
            JOIN subscribers s ON cd.subscriber_id = s.id
            JOIN departments d ON s.dept_id = d.id
            WHERE d.org_id = ? AND cd.period_year = ? AND cd.reading_date IS NOT NULL AND cd.reading_date != ''
            ORDER BY cd.id DESC
            LIMIT 1
        """, (org["id"], year))
        latest_row = cursor.fetchone()
        if latest_row:
            p_type = latest_row["period_type"]
            p_num = latest_row["period_number"]
            if p_type == "MONTHLY":
                org["latest_period_label"] = f"شهر {p_num} / {year}"
                org["latest_period_en"] = f"Month {p_num} / {year}"
            else:
                org["latest_period_label"] = f"الربع {p_num} / {year}"
                org["latest_period_en"] = f"Quarter {p_num} / {year}"
        else:
            org["latest_period_label"] = "لا توجد قراءات بعد"
            org["latest_period_en"] = "No readings yet"

        read_date_str = org.get("latest_reading_date")
        org["_sort_date"] = parse_date_for_sort(read_date_str)

    if sort_by == "latest_reading":
        org_rows.sort(key=lambda x: (x["_sort_date"], x["active_subscribers"]), reverse=True)
    elif sort_by == "subscribers_desc":
        org_rows.sort(key=lambda x: x["active_subscribers"], reverse=True)
    elif sort_by == "name_asc":
        org_rows.sort(key=lambda x: x["name"])

    for org in org_rows:
        org.pop("_sort_date", None)

    conn.close()
    return {
        "year": year,
        "total_facilities": len(org_rows),
        "facilities": org_rows
    }

@app.get("/api/facilities/{org_id}/subscribers-readings")
def get_facility_subscribers_readings(
    org_id: int,
    year: Optional[int] = None,
    dept_id: Optional[int] = None,
    request: Request = None
):
    user = resolve_user(request)
    if user and user.get("role") == "HOSPITAL" and user.get("org_id") != org_id:
        raise HTTPException(status_code=403, detail="غير مصرح لك بالاطلاع على كوادر منشأة أخرى")

    conn = get_db()
    cursor = conn.cursor()

    if not year:
        cursor.execute("SELECT value FROM system_settings WHERE key = 'active_year'")
        row_y = cursor.fetchone()
        year = int(row_y[0]) if row_y and row_y[0] else 2026

    cursor.execute("SELECT id, name, code, classification, contact, address FROM organizations WHERE id = ?", (org_id,))
    org = cursor.fetchone()
    if not org:
        conn.close()
        raise HTTPException(status_code=404, detail="المنشأة غير موجودة")

    params = [year, org_id]
    dept_filter = ""
    if dept_id:
        dept_filter = "AND d.id = ?"
        params.append(dept_id)

    cursor.execute(f"""
        SELECT s.id, s.name, s.system_code, s.national_id, s.job_title, s.status,
               s.has_double_badge, s.current_serial_oa, s.current_serial_ua,
               s.started_date, s.resigned_date,
               d.id as dept_id, d.name as dept_name, d.default_frequency,
               COALESCE(SUM(cd.hp10), 0.0) as annual_hp10,
               COALESCE(SUM(cd.hp07), 0.0) as annual_hp07,
               COUNT(cd.id) as readings_count,
               MAX(cd.reading_date) as last_sub_reading_date
        FROM subscribers s
        JOIN departments d ON s.dept_id = d.id
        LEFT JOIN calculated_doses cd ON s.id = cd.subscriber_id AND cd.period_year = ?
        WHERE d.org_id = ?
        {dept_filter}
        GROUP BY s.id
        ORDER BY d.name ASC, s.name ASC
    """, tuple(params))
    subs = [dict(s) for s in cursor.fetchall()]

    for s in subs:
        cursor.execute("""
            SELECT id, period_type, period_number, hp10, hp07,
                   badge_oa_serial, badge_ua_serial, reading_date, reception_date,
                   remark
            FROM calculated_doses
            WHERE subscriber_id = ? AND period_year = ?
            ORDER BY period_type DESC, period_number ASC
        """, (s["id"], year))
        s["readings"] = [dict(r) for r in cursor.fetchall()]
        s["annual_hp10"] = round(s["annual_hp10"], 2)
        s["annual_hp07"] = round(s["annual_hp07"], 2)

    conn.close()
    return {
        "organization": dict(org),
        "year": year,
        "subscribers_count": len(subs),
        "subscribers": subs
    }

@app.get("/api/schedule/matrix")
def get_schedule_matrix(
    year: Optional[int] = None,
    frequency: Optional[str] = None,
    request: Request = None
):
    user = resolve_user(request)
    user_org_id = user["org_id"] if (user and user.get("role") == "HOSPITAL") else None

    conn = get_db()
    cursor = conn.cursor()

    if not year:
        cursor.execute("SELECT value FROM system_settings WHERE key = 'active_year'")
        row_y = cursor.fetchone()
        year = int(row_y[0]) if row_y and row_y[0] else 2026

    where_clauses = ["1=1"]
    params = []
    if user_org_id:
        where_clauses.append("o.id = ?")
        params.append(user_org_id)
    if frequency and frequency.upper() in ["MONTHLY", "QUARTERLY"]:
        where_clauses.append("d.default_frequency = ?")
        params.append(frequency.upper())

    where_str = " AND ".join(where_clauses)

    cursor.execute(f"""
        SELECT d.id as dept_id, d.name as dept_name, d.default_frequency,
               o.id as org_id, o.name as org_name, o.code as org_code,
               COUNT(DISTINCT CASE WHEN s.status = 'ACTIVE' THEN s.id END) as active_subs,
               COUNT(DISTINCT CASE WHEN s.has_double_badge = 1 AND s.status = 'ACTIVE' THEN s.id END) as double_subs
        FROM departments d
        JOIN organizations o ON d.org_id = o.id
        LEFT JOIN subscribers s ON d.id = s.dept_id
        WHERE {where_str}
        GROUP BY d.id, d.name, d.default_frequency, o.id, o.name, o.code
        ORDER BY o.name ASC, d.name ASC
    """, tuple(params))
    departments = [dict(d) for d in cursor.fetchall()]

    cursor.execute("""
        SELECT s.dept_id, cd.period_type, cd.period_number,
               COUNT(cd.id) as readings_count,
               MAX(cd.reading_date) as latest_reading_date,
               ROUND(AVG(CASE WHEN cd.remark != 'NR' THEN cd.hp10 END), 2) as avg_hp10,
               MAX(cd.hp10) as max_hp10,
               COUNT(CASE WHEN cd.remark = 'NR' THEN 1 END) as nr_count,
               COUNT(CASE WHEN cd.hp10 > 1.67 THEN 1 END) as alerts_count
        FROM calculated_doses cd
        JOIN subscribers s ON cd.subscriber_id = s.id
        WHERE cd.period_year = ?
        GROUP BY s.dept_id, cd.period_type, cd.period_number
    """, (year,))
    readings_map = {}
    for r in cursor.fetchall():
        key = (r["dept_id"], r["period_type"], r["period_number"])
        readings_map[key] = dict(r)

    import datetime
    today = datetime.date.today()
    current_year = today.year
    current_month = today.month
    current_quarter = (current_month - 1) // 3 + 1

    total_expected_slots = 0
    total_completed_slots = 0
    total_overdue_slots = 0
    total_due_slots = 0

    matrix_rows = []
    for dept in departments:
        freq = dept["default_frequency"]
        dept_id = dept["dept_id"]
        row = {
            "dept_id": dept_id,
            "dept_name": dept["dept_name"],
            "org_id": dept["org_id"],
            "org_name": dept["org_name"],
            "org_code": dept["org_code"],
            "default_frequency": freq,
            "active_subs": dept["active_subs"],
            "double_subs": dept["double_subs"],
            "monthly_periods": {},
            "quarterly_periods": {}
        }

        # Monthly slots (1..12)
        for m in range(1, 13):
            key = (dept_id, "MONTHLY", m)
            is_req = (freq == "MONTHLY")
            read_info = readings_map.get(key)
            has_reading = read_info is not None and read_info["readings_count"] > 0

            status = "NOT_APPLICABLE"
            if is_req:
                total_expected_slots += 1
                if has_reading:
                    status = "COMPLETED"
                    total_completed_slots += 1
                else:
                    if year < current_year or (year == current_year and m < current_month):
                        status = "OVERDUE"
                        total_overdue_slots += 1
                    elif year == current_year and m == current_month:
                        status = "DUE"
                        total_due_slots += 1
                    else:
                        status = "UPCOMING"

            row["monthly_periods"][str(m)] = {
                "period_number": m,
                "is_required": is_req,
                "status": status,
                "reading_date": read_info["latest_reading_date"] if read_info else None,
                "readings_count": read_info["readings_count"] if read_info else 0,
                "avg_hp10": read_info["avg_hp10"] if read_info else None,
                "max_hp10": read_info["max_hp10"] if read_info else None,
                "alerts_count": read_info["alerts_count"] if read_info else 0,
                "nr_count": read_info["nr_count"] if read_info else 0
            }

        # Quarterly slots (1..4)
        for q in range(1, 5):
            key = (dept_id, "QUARTERLY", q)
            is_req = (freq == "QUARTERLY")
            read_info = readings_map.get(key)
            has_reading = read_info is not None and read_info["readings_count"] > 0

            status = "NOT_APPLICABLE"
            if is_req:
                total_expected_slots += 1
                if has_reading:
                    status = "COMPLETED"
                    total_completed_slots += 1
                else:
                    if year < current_year or (year == current_year and q < current_quarter):
                        status = "OVERDUE"
                        total_overdue_slots += 1
                    elif year == current_year and q == current_quarter:
                        status = "DUE"
                        total_due_slots += 1
                    else:
                        status = "UPCOMING"

            row["quarterly_periods"][str(q)] = {
                "period_number": q,
                "is_required": is_req,
                "status": status,
                "reading_date": read_info["latest_reading_date"] if read_info else None,
                "readings_count": read_info["readings_count"] if read_info else 0,
                "avg_hp10": read_info["avg_hp10"] if read_info else None,
                "max_hp10": read_info["max_hp10"] if read_info else None,
                "alerts_count": read_info["alerts_count"] if read_info else 0,
                "nr_count": read_info["nr_count"] if read_info else 0
            }

        matrix_rows.append(row)

    conn.close()

    compliance_pct = round((total_completed_slots / max(1, (total_completed_slots + total_overdue_slots))) * 100, 1) if (total_completed_slots + total_overdue_slots) > 0 else 100.0

    return {
        "year": year,
        "summary": {
            "total_departments": len(departments),
            "total_expected_slots": total_expected_slots,
            "total_completed_slots": total_completed_slots,
            "total_overdue_slots": total_overdue_slots,
            "total_due_slots": total_due_slots,
            "compliance_percentage": compliance_pct
        },
        "matrix": matrix_rows
    }


# --- 8. Integrated Web UI ---
@app.get("/", response_class=HTMLResponse)
def index_page():
    if not os.path.exists(STATIC_HTML_PATH):
        raise HTTPException(status_code=404, detail="ملف الواجهة غير موجود")
    with open(STATIC_HTML_PATH, "r", encoding="utf-8") as f:
        return f.read()


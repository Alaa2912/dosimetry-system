import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.environ.get("DB_PATH", os.path.join(BASE_DIR, "osl_service.db"))

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA synchronous=NORMAL;")
    conn.execute("PRAGMA cache_size=-64000;")
    return conn

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    
    # 0. System Users / Staff
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        email TEXT UNIQUE,
        full_name TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'TECHNICIAN',
        password TEXT NOT NULL DEFAULT '123456',
        org_id INTEGER,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(org_id) REFERENCES organizations(id)
    )
    """)
    
    cursor.execute("SELECT id FROM users WHERE username = 'admin'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO users (username, email, full_name, role, password) VALUES ('admin', '3la2al7abees@gmail.com', 'Alaa Alhabees', 'ADMIN', 'admin123')")
    
    # 1. Organizations
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS organizations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE,
        name TEXT NOT NULL,
        address TEXT,
        contact TEXT,
        classification TEXT DEFAULT 'Medical Field',
        agreement_date TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # 2. Departments
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS departments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        org_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        default_frequency TEXT DEFAULT 'QUARTERLY',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(org_id) REFERENCES organizations(id)
    )
    """)
    
    # 3. Subscribers (المشتركين مع السيريال المباشر والرقم الإشعاعي الدائم)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS subscribers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        dept_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        system_code TEXT UNIQUE,
        national_id TEXT,
        employee_no TEXT,
        job_title TEXT,
        started_date TEXT,
        has_double_badge BOOLEAN DEFAULT 0,
        current_serial_oa TEXT, -- السيريال الافتراضي للوضحة الصدرية أو الخارجية
        current_serial_ua TEXT, -- السيريال الافتراضي للوضحة الداخلية
        status TEXT DEFAULT 'ACTIVE', -- 'ACTIVE' or 'RESIGNED'
        resigned_date TEXT,
        resignation_notes TEXT,
        is_active BOOLEAN DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(dept_id) REFERENCES departments(id)
    )
    """)
    
    # 4. Badge Assignments
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS badge_assignments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        subscriber_id INTEGER NOT NULL,
        period_year INTEGER NOT NULL,
        period_type TEXT NOT NULL,
        period_number INTEGER NOT NULL,
        serial_number TEXT NOT NULL,
        badge_position TEXT DEFAULT 'CHEST',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(subscriber_id) REFERENCES subscribers(id),
        UNIQUE(serial_number, period_year, period_type, period_number)
    )
    """)
    
    # 5. Import Templates
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS import_templates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        file_format TEXT DEFAULT 'excel',
        delimiter TEXT DEFAULT ',',
        skip_rows INTEGER DEFAULT 0,
        col_serial TEXT NOT NULL,
        col_read_date TEXT,
        col_hp10 TEXT NOT NULL,
        col_hp07 TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # 6. Raw Readings
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS raw_readings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_id TEXT,
        serial_number TEXT NOT NULL,
        read_date TEXT,
        raw_hp10 REAL,
        raw_hp07 REAL,
        raw_data TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # 7. Calculated Doses
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS calculated_doses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        subscriber_id INTEGER NOT NULL,
        period_year INTEGER NOT NULL,
        period_type TEXT NOT NULL,
        period_number INTEGER NOT NULL,
        hp10 REAL DEFAULT 0.0,
        hp07 REAL DEFAULT 0.0,
        accumulated_hp10 REAL DEFAULT 0.0,
        accumulated_hp07 REAL DEFAULT 0.0,
        is_double_badge BOOLEAN DEFAULT 0,
        badge_oa_serial TEXT,
        badge_ua_serial TEXT,
        reception_date TEXT,
        reading_date TEXT,
        report_date TEXT,
        remark TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(subscriber_id) REFERENCES subscribers(id),
        UNIQUE(subscriber_id, period_year, period_type, period_number)
    )
    """)
    
    # 8. Unassigned Readings
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS unassigned_readings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        serial_number TEXT NOT NULL,
        read_date TEXT,
        hp10 REAL,
        hp07 REAL,
        period_year INTEGER,
        period_type TEXT,
        period_number INTEGER,
        status TEXT DEFAULT 'UNASSIGNED',
        assigned_subscriber_id INTEGER,
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    
    # 9. Badge Receptions (كشف استلام الوضحات)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS badge_receptions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        subscriber_id INTEGER NOT NULL,
        dept_id INTEGER NOT NULL,
        period_year INTEGER NOT NULL,
        period_type TEXT NOT NULL,
        period_number INTEGER NOT NULL,
        reception_date TEXT,
        received_oa BOOLEAN DEFAULT 1,
        received_ua BOOLEAN DEFAULT 1,
        status TEXT DEFAULT 'RECEIVED',
        notes TEXT,
        receiver_name TEXT,
        delivered_by TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(subscriber_id) REFERENCES subscribers(id),
        FOREIGN KEY(dept_id) REFERENCES departments(id),
        UNIQUE(subscriber_id, period_year, period_type, period_number)
    )
    """)

    
    # 10. System Settings
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS system_settings (
        key TEXT PRIMARY KEY,
        value TEXT NOT NULL,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    cursor.execute("INSERT OR IGNORE INTO system_settings (key, value) VALUES ('active_year', '2026')")
    cursor.execute("INSERT OR IGNORE INTO system_settings (key, value) VALUES ('theme', 'dark')")
    cursor.execute("INSERT OR IGNORE INTO system_settings (key, value) VALUES ('language', 'ar')")

    
    # 11. High-Performance B-Tree Indexes (Optimized for 2,000,000+ records)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_calc_doses_sub_year ON calculated_doses (subscriber_id, period_year)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_calc_doses_period ON calculated_doses (period_year, period_type, period_number)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_calc_doses_hp10 ON calculated_doses (hp10)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_subs_dept ON subscribers (dept_id)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_subs_status ON subscribers (status)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_subs_serial_oa ON subscribers (current_serial_oa)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_subs_serial_ua ON subscribers (current_serial_ua)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_badge_assign_period ON badge_assignments (period_year, period_type, period_number)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_raw_readings_sn ON raw_readings (serial_number)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_receptions_dept_period ON badge_receptions (dept_id, period_year, period_type, period_number)")

    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("Database successfully initialized at", DB_PATH)

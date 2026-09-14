from datetime import datetime, timedelta
from typing import Optional, Tuple, Dict, Any

LLD_THRESHOLD = 0.01  # Lower Limit of Detection in mSv

def parse_excel_date(date_val) -> str:
    """
    Parses date from string, int, or float (Excel serial date number).
    Returns formatted date YYYY-MM-DD or DD/MM/YYYY.
    """
    if date_val is None or date_val == "":
        return ""
    
    # Check if numeric Excel serial date
    try:
        float_val = float(date_val)
        # Excel date epoch starts Dec 30, 1899
        base_date = datetime(1899, 12, 30)
        dt = base_date + timedelta(days=float_val)
        return dt.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        pass
    
    # Try standard string date parsing
    str_val = str(date_val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try:
            dt = datetime.strptime(str_val.split()[0], fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue
            
    return str_val

def sanitize_dose(raw_dose: Optional[float]) -> Tuple[float, str]:
    """
    Sanitizes raw dose from reader.
    If dose is below LLD (e.g. negative or < 0.01 mSv), treat as 0.00 with 'LLD' remark.
    """
    if raw_dose is None:
        return 0.0, "NI"  # Not Issued / Missing
    
    try:
        val = float(raw_dose)
    except (ValueError, TypeError):
        return 0.0, "ES"
        
    if val < LLD_THRESHOLD:
        return 0.0, "LLD"
        
    return round(val, 2), ""

def calculate_single_badge_dose(raw_hp10: float, raw_hp07: float) -> Tuple[float, float, str]:
    """
    Calculates dose for single dosimeter.
    """
    hp10, rem1 = sanitize_dose(raw_hp10)
    hp07, rem2 = sanitize_dose(raw_hp07)
    remark = rem1 or rem2
    return hp10, hp07, remark

def calculate_double_badge_dose(
    oa_hp10: float, oa_hp07: float,
    ua_hp10: float, ua_hp07: float
) -> Tuple[float, float, str]:
    """
    Double Dosimetry (وضحة خارجية Over Apron + وضحة داخلية Under Apron)
    Formula:
      Effective Hp(10) = (External * 0.025) + (Internal * 0.5)
      Effective Hp(0.07) = (External * 0.025) + (Internal * 0.5)
    """
    # Sanitize inputs (if below LLD, consider 0 for formula calculation)
    ext_10 = max(0.0, float(oa_hp10 or 0.0))
    int_10 = max(0.0, float(ua_hp10 or 0.0))
    ext_07 = max(0.0, float(oa_hp07 or 0.0))
    int_07 = max(0.0, float(ua_hp07 or 0.0))
    
    calc_hp10 = (ext_10 * 0.025) + (int_10 * 0.5)
    calc_hp07 = (ext_07 * 0.025) + (int_07 * 0.5)
    
    final_hp10, rem10 = sanitize_dose(calc_hp10)
    final_hp07, rem07 = sanitize_dose(calc_hp07)
    
    remark = "OA+UA"
    if rem10 == "LLD" and rem07 == "LLD":
        remark = "LLD (OA+UA)"
        
    return final_hp10, final_hp07, remark

def compute_accumulated_dose(
    conn, subscriber_id: int, period_year: int, period_type: str, current_period_number: int,
    current_hp10: float, current_hp07: float
) -> Tuple[float, float]:
    """
    Sums doses from period 1 up to current_period_number in the same year.
    """
    cursor = conn.cursor()
    cursor.execute("""
    SELECT SUM(hp10) as total_hp10, SUM(hp07) as total_hp07
    FROM calculated_doses
    WHERE subscriber_id = ? AND period_year = ? AND period_type = ? AND period_number < ?
    """, (subscriber_id, period_year, period_type, current_period_number))
    row = cursor.fetchone()
    
    prev_hp10 = row['total_hp10'] if row and row['total_hp10'] is not None else 0.0
    prev_hp07 = row['total_hp07'] if row and row['total_hp07'] is not None else 0.0
    
    acc_hp10 = round(prev_hp10 + current_hp10, 2)
    acc_hp07 = round(prev_hp07 + current_hp07, 2)
    
    return acc_hp10, acc_hp07

if __name__ == "__main__":
    print("Testing parse_excel_date:", parse_excel_date(45638.4859))
    print("Testing single badge:", calculate_single_badge_dose(0.35, 0.40))
    print("Testing double badge (Ext=10.0, Int=1.0):", calculate_double_badge_dose(10.0, 10.0, 1.0, 1.0))
    # Ext*0.025 + Int*0.5 = 10*0.025 (0.25) + 1*0.5 (0.5) = 0.75

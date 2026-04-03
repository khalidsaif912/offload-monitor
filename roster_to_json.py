"""
roster_to_json.py
=================
يقرأ ملف إكسل الروستر من OneDrive ويحوّله إلى roster.json
متوافق مع البنية المطلوبة من offload_monitor.py

الاستخدام:
    python roster_to_json.py

متغيرات البيئة المطلوبة:
    ONEDRIVE_ROSTER_URL  : رابط مشاركة OneDrive لملف الإكسل
    ROSTER_JSON_PATH     : مسار حفظ الملف (افتراضي: docs/data/roster.json)

الشيتات التي يقرأها:
    Supervisors, Load Control, Export Checker, Export Operators, Officers, Acceptance
"""

from __future__ import annotations

import os
import re
import io
import json
import hashlib
import calendar
from datetime import datetime, date
from pathlib import Path

import requests
import pandas as pd

# ── إعدادات ──────────────────────────────────────────────────────────────────

ROSTER_JSON_PATH = Path(os.getenv("ROSTER_JSON_PATH", "docs/data/roster.json"))
STATE_FILE       = Path("roster_state.txt")

# أكواد المناوبات → فئة
# MN / ME = Morning, AN / AE = Afternoon, NN / NE = Night
# OFF, O1-O9 = Off, LV = Leave, TR = Training, STAN/STMN... = Standby
SHIFT_MORNING   = re.compile(r"^(ST)?M[ENM]?\d*$", re.I)
SHIFT_AFTERNOON = re.compile(r"^(ST)?A[ENM]?\d*$", re.I)
SHIFT_NIGHT     = re.compile(r"^(ST)?N[ENM]?\d*$", re.I)
SHIFT_OFF       = re.compile(r"^O\d*$", re.I)
LEAVE_CODES     = {"LV", "TR", "SICK", "EL", "OFF"}

SHIFT_LABELS = {
    "morning":   "Morning",
    "afternoon": "Afternoon",
    "night":     "Night",
    "off":       "Off Day",
    "leave":     "Annual Leave",
    "training":  "Training",
}

# أسماء الشيتات ودوراتها في الـ JSON
DEPT_SHEETS = {
    "Supervisors":       "Supervisors",
    "Load Control":      "Load Control",
    "Export Checker":    "Export Checker",
    "Export Operators":  "Export Operators",
    "Officers":          "Officers",
    "Acceptance":        "Acceptance",
}


# ── تحميل الملف ──────────────────────────────────────────────────────────────

def download_excel(url: str) -> bytes:
    url = url.strip()
    sep = "&" if "?" in url else "?"
    if "download=1" not in url:
        url += f"{sep}download=1"
    url += f"&__ts={int(datetime.now().timestamp())}"
    r = requests.get(url, timeout=30, headers={
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache",
    })
    r.raise_for_status()
    return r.content


def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


# ── تحليل كود المناوبة ───────────────────────────────────────────────────────

def classify_code(code: str) -> str:
    """Return shift category from a roster cell code."""
    c = str(code or "").strip().upper()
    if not c or c in ("NAN", "-", ""):
        return ""
    if c == "TR":
        return "training"
    if c == "LV":
        return "leave"
    if SHIFT_OFF.match(c):
        return "off"
    if SHIFT_MORNING.match(c):
        return "morning"
    if SHIFT_AFTERNOON.match(c):
        return "afternoon"
    if SHIFT_NIGHT.match(c):
        return "night"
    # fallback: if starts with known leave codes
    if any(c.startswith(lv) for lv in LEAVE_CODES):
        return "off"
    return ""


# ── قراءة شيت الروستر ────────────────────────────────────────────────────────

def parse_sheet(df: pd.DataFrame, dept_name: str, year: int, month: int) -> dict[str, list[dict]]:
    """
    Parse one roster sheet.
    Returns: { "YYYY-MM-DD": [ {name, sn, dept, shift_label}, ... ] }
    """
    # Row 3 (index 3) = day numbers
    # Row 4+ (index 4+) = employee rows
    # Col 5+ (index 5+) = day columns

    days_in_month = calendar.monthrange(year, month)[1]

    # Find day-number columns: row index 3, col 5 onward
    day_row = df.iloc[3]
    col_for_day: dict[int, int] = {}   # day_number → col_index
    for col_i, val in enumerate(day_row):
        if isinstance(val, (int, float)) and not pd.isna(val):
            d = int(val)
            if 1 <= d <= days_in_month:
                col_for_day[d] = col_i

    if not col_for_day:
        print(f"  [parse_sheet] {dept_name}: no day columns found")
        return {}

    # Collect results per date
    result: dict[str, list[dict]] = {}

    # Employee rows start at index 4
    for row_i in range(4, len(df)):
        row = df.iloc[row_i]

        # Employee name/SN is in col 2
        raw_name = str(row.iloc[2] if len(row) > 2 else "").strip()
        if not raw_name or raw_name.upper() in ("NAN", "NONE", ""):
            continue
        # Skip summary rows (Morning Shift, Afternoon Shift, etc.)
        if any(kw in raw_name for kw in ("Morning", "Afternoon", "Night", "Offday", "Off Day",
                                          "Training", "Standby", "Extra", "On duty", "Employee")):
            continue

        # Extract name and SN: "Mohamed Al Amri - 81404"
        sn = ""
        name = raw_name
        m = re.match(r"^(.+?)\s*[-–]\s*(\d+)\s*$", raw_name)
        if m:
            name = m.group(1).strip()
            sn   = m.group(2).strip()

        # For each day, read the shift code
        for day_num, col_i in col_for_day.items():
            if col_i >= len(row):
                continue
            code = str(row.iloc[col_i] if not pd.isna(row.iloc[col_i]) else "").strip()
            category = classify_code(code)
            if not category:
                continue

            date_str = f"{year:04d}-{month:02d}-{day_num:02d}"
            if date_str not in result:
                result[date_str] = []

            result[date_str].append({
                "name":        name,
                "sn":          sn,
                "dept":        dept_name,
                "shift_label": SHIFT_LABELS.get(category, category),
                "code":        code.upper(),
            })

    return result


# ── بناء cards_html ───────────────────────────────────────────────────────────

def build_cards_html(employees: list[dict]) -> str:
    """
    Build cards_html HTML string compatible with offload_monitor.py parser.
    Structure:
      <div class="deptCard">
        <div class="deptTitle">DEPT</div>
        <div class="shiftCard">
          <div class="shiftLabel">Morning</div>
          <div class="empRow"><span class="empName">Name - SN</span></div>
        </div>
      </div>
    """
    # Group by dept → shift_label → employees
    from collections import defaultdict
    by_dept: dict[str, dict[str, list]] = defaultdict(lambda: defaultdict(list))

    for emp in employees:
        by_dept[emp["dept"]][emp["shift_label"]].append(emp)

    html_parts = []
    shift_order = ["Morning", "Afternoon", "Night", "Annual Leave", "Training", "Off Day"]

    for dept, shifts in by_dept.items():
        dept_html = f'<div class="deptCard"><div class="deptTitle">{dept}</div>'
        for shift_label in shift_order:
            if shift_label not in shifts:
                continue
            shift_html = f'<div class="shiftCard"><div class="shiftLabel">{shift_label}</div>'
            for emp in shifts[shift_label]:
                display = f'{emp["name"]} - {emp["sn"]}' if emp["sn"] else emp["name"]
                shift_html += f'<div class="empRow"><span class="empName">{display}</span></div>'
            shift_html += '</div>'
            dept_html += shift_html
        dept_html += '</div>'
        html_parts.append(dept_html)

    return "".join(html_parts)


# ── استخراج الشهر والسنة ────────────────────────────────────────────────────

MONTH_NAMES = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}

def extract_month_year(df: pd.DataFrame) -> tuple[int, int]:
    """Try to extract month and year from the title row (row index 1)."""
    title = ""
    for val in df.iloc[1]:
        if isinstance(val, str) and len(val) > 5:
            title = val.upper()
            break

    year = datetime.now().year
    month = datetime.now().month

    # Try to find year
    m_year = re.search(r"\b(202\d)\b", title)
    if m_year:
        year = int(m_year.group(1))

    # Try to find month name
    for name, num in MONTH_NAMES.items():
        if name.upper() in title:
            month = num
            break

    return year, month


# ── البرنامج الرئيسي ─────────────────────────────────────────────────────────

def main():
    roster_url = os.environ.get("ONEDRIVE_ROSTER_URL", "").strip()
    if not roster_url:
        print("ERROR: ONEDRIVE_ROSTER_URL environment variable not set.")
        return

    print(f"[roster] Downloading Excel from OneDrive...")
    try:
        excel_bytes = download_excel(roster_url)
    except Exception as e:
        print(f"[roster] Download failed: {e}")
        return

    new_hash = sha256_bytes(excel_bytes)
    print(f"[roster] Excel size: {len(excel_bytes):,} bytes | sha256: {new_hash[:16]}")

    # Check if file changed
    force = os.getenv("FORCE_REBUILD", "").strip().lower() in ("1", "true", "yes")
    if STATE_FILE.exists() and not force:
        old_hash = STATE_FILE.read_text(encoding="utf-8").strip()
        if old_hash == new_hash:
            print("[roster] No change in Excel file — skipping rebuild.")
            return

    print("[roster] Change detected. Parsing sheets...")

    try:
        xl = pd.ExcelFile(io.BytesIO(excel_bytes))
    except Exception as e:
        print(f"[roster] Failed to open Excel: {e}")
        return

    # All data per date
    from collections import defaultdict
    all_by_date: dict[str, list[dict]] = defaultdict(list)

    year = month = None

    for sheet_name, dept_name in DEPT_SHEETS.items():
        if sheet_name not in xl.sheet_names:
            print(f"  [roster] Sheet '{sheet_name}' not found — skipping")
            continue

        try:
            df = xl.parse(sheet_name, header=None)
        except Exception as e:
            print(f"  [roster] Error reading sheet '{sheet_name}': {e}")
            continue

        # Extract month/year from first sheet that has it
        if year is None:
            year, month = extract_month_year(df)
            print(f"  [roster] Detected: {calendar.month_name[month]} {year}")

        sheet_data = parse_sheet(df, dept_name, year, month)
        for date_str, emps in sheet_data.items():
            all_by_date[date_str].extend(emps)

        total = sum(len(v) for v in sheet_data.values())
        print(f"  [roster] {dept_name}: {total} employee-day records across {len(sheet_data)} days")

    if not all_by_date:
        print("[roster] WARNING: No data extracted.")
        return

    # Build JSON structure
    days_json: dict[str, dict] = {}
    for date_str in sorted(all_by_date.keys()):
        emps = all_by_date[date_str]
        cards = build_cards_html(emps)
        days_json[date_str] = {
            "cards_html": cards,
            "employee_count": len(emps),
        }

    output = {
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "month": f"{calendar.month_name[month]} {year}",
        "source": "onedrive-excel",
        "days": days_json,
    }

    # Save
    ROSTER_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    ROSTER_JSON_PATH.write_text(
        json.dumps(output, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    STATE_FILE.write_text(new_hash, encoding="utf-8")

    total_days = len(days_json)
    print(f"[roster] ✓ Saved {total_days} days to {ROSTER_JSON_PATH}")


if __name__ == "__main__":
    main()

"""
╔══════════════════════════════════════════════════════════════════╗
║              OFFLOAD MONITOR — Cargo Tracking System             ║
║              Automated HTML Report Generator via GitHub Actions  ║
╚══════════════════════════════════════════════════════════════════╝

المؤلف   : نظام متابعة الشحنات المُفرَّغة
المنطقة  : Asia/Muscat
التخزين  : GitHub Pages ← /docs
"""

from __future__ import annotations

import os
import re
import json
import hashlib
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup


# ══════════════════════════════════════════════════════════════════
#  الإعدادات العامة
# ══════════════════════════════════════════════════════════════════

ONEDRIVE_URL: str = os.environ["ONEDRIVE_FILE_URL"]
TIMEZONE: str     = "Asia/Muscat"

DATA_DIR:   Path = Path("data")
STATE_FILE: Path = Path("state.txt")
DOCS_DIR:   Path = Path("docs")   # GitHub Pages تخدم من /docs


# ══════════════════════════════════════════════════════════════════
#  الدوال المساعدة
# ══════════════════════════════════════════════════════════════════

def download_file() -> str:
    """تحميل ملف HTML من OneDrive وإضافة معامل التنزيل إن لم يكن موجوداً."""
    url = ONEDRIVE_URL.strip()
    separator = "&" if "?" in url else "?"
    if "download=1" not in url:
        url += f"{separator}download=1"

    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return response.text


def compute_sha256(content: str) -> str:
    """إنتاج بصمة SHA-256 للمحتوى النصي."""
    return hashlib.sha256(content.encode("utf-8", errors="ignore")).hexdigest()


def get_shift(now: datetime) -> str:
    """
    تحديد الوردية بناءً على الوقت الحالي:
        shift1 → 06:00 – 14:29
        shift2 → 14:30 – 21:29
        shift3 → 21:30 – 05:59
    """
    total_minutes = now.hour * 60 + now.minute

    if 6 * 60 <= total_minutes < 14 * 60 + 30:
        return "shift1"
    if 14 * 60 + 30 <= total_minutes < 21 * 60 + 30:
        return "shift2"
    return "shift3"


def slugify(text: str, max_length: int = 80) -> str:
    """
    تحويل النص إلى معرّف آمن للاستخدام في أسماء الملفات:
    المسافات → شرطة سفلية، إزالة الأحرف الخاصة، اقتصاص الطول.
    """
    text = (text or "UNKNOWN").strip()
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^A-Za-z0-9_-]", "_", text)
    return (text or "UNKNOWN")[:max_length]


def load_json(path: Path, default):
    """قراءة ملف JSON مع إرجاع قيمة افتراضية عند الغياب أو الخطأ."""
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return default


def write_json(path: Path, data) -> None:
    """كتابة بيانات JSON مع إنشاء المجلدات تلقائياً."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ══════════════════════════════════════════════════════════════════
#  تحليل جدول HTML واستخراج البيانات المنظّمة
# ══════════════════════════════════════════════════════════════════

EXPECTED_HEADERS: list[str] = [
    "ITEM",
    "DATE",
    "FLIGHT",
    "STD/ETD",
    "DEST",
    "Email Received Time",
    "Physical Cargo received from Ramp",
    "Trolley/ ULD Number",
    "Offloading Process Completed in CMS",
    "Offloading Pieces Verification",
    "Offloading Reason",
    "Remarks/Additional Information",
]

# تعيين رؤوس الأعمدة إلى مفاتيح داخلية، مع دعم التسميات البديلة
COLUMN_ALIASES: dict[str, list[str]] = {
    "ITEM":                                    ["ITEM"],
    "DATE":                                    ["DATE"],
    "FLIGHT":                                  ["FLIGHT", "FLIGHT #"],
    "STD/ETD":                                 ["STD/ETD", "STD/ATD"],
    "DEST":                                    ["DEST", "DESTINATION"],
    "Email Received Time":                     ["Email Received Time", "Email"],
    "Physical Cargo received from Ramp":       ["Physical Cargo received from Ramp",
                                                "Physical Cargo received", "Physical"],
    "Trolley/ ULD Number":                     ["Trolley/ ULD Number",
                                                "Trolley/ULD Number", "Trolley"],
    "Offloading Process Completed in CMS":     ["Offloading Process Completed in CMS",
                                                "CMS", "Offloading Process Completed"],
    "Offloading Pieces Verification":          ["Offloading Pieces Verification",
                                                "Pieces", "Offloading Pieces"],
    "Offloading Reason":                       ["Offloading Reason", "Reason"],
    "Remarks/Additional Information":          ["Remarks/Additional Information", "Remarks"],
}


def normalize_header(raw: str) -> str:
    """تنظيف رأس العمود وتوحيد المسافات والأشكال المختلفة."""
    header = re.sub(r"\s+", " ", (raw or "").strip())
    header = header.replace("STD/ ETD", "STD/ETD")
    header = header.replace("Trolley/ULD", "Trolley/ ULD")
    return header


def find_main_table(soup: BeautifulSoup):
    """
    البحث عن أفضل جدول في HTML يحتوي على أكبر عدد من رؤوس الأعمدة المتوقعة.
    يُعيد الجدول الأعلى تطابقاً، أو None إن لم يُعثر على شيء.
    """
    best_table = None
    best_score = 0

    for table in soup.find_all("table"):
        # محاولة قراءة الرؤوس من <th> أو من الصف الأول
        th_elements = table.find_all("th")
        if th_elements:
            headers = [normalize_header(th.get_text(" ", strip=True)) for th in th_elements]
        else:
            first_row = table.find("tr")
            if not first_row:
                continue
            headers = [
                normalize_header(cell.get_text(" ", strip=True))
                for cell in first_row.find_all(["td", "th"])
            ]

        header_text = " | ".join(headers).lower()
        score = sum(1 for h in EXPECTED_HEADERS if h.lower() in header_text)

        if score > best_score:
            best_score = score
            best_table = table

    return best_table


def _resolve_column_index(headers: list[str], column_key: str) -> int | None:
    """إيجاد فهرس العمود في الرؤوس عبر البحث في الأسماء البديلة."""
    for alias in COLUMN_ALIASES.get(column_key, [column_key]):
        if alias in headers:
            return headers.index(alias)
    return None


def parse_rows_from_table(table) -> list[dict]:
    """
    استخراج جميع صفوف البيانات من الجدول بعد تجاهل الهيدر.
    يُعيد قائمة من القواميس بمفاتيح موحّدة.
    """
    rows: list[dict] = []
    all_rows = table.find_all("tr")
    if not all_rows:
        return rows

    # استخراج رؤوس الأعمدة من الصف الأول
    header_cells = all_rows[0].find_all(["th", "td"])
    headers = [normalize_header(cell.get_text(" ", strip=True)) for cell in header_cells]

    # بناء خريطة الفهارس
    col = {key: _resolve_column_index(headers, key) for key in COLUMN_ALIASES}

    def get_cell(row_values: list[str], column_key: str) -> str:
        idx = col.get(column_key)
        if idx is None or idx >= len(row_values):
            return ""
        return row_values[idx].strip()

    # معالجة صفوف البيانات (تخطّي الصف الأول = الهيدر)
    for tr in all_rows[1:]:
        cells = tr.find_all(["td", "th"])
        if not cells:
            continue

        values = [cell.get_text(" ", strip=True) for cell in cells]

        # تجاهل الصفوف الفارغة كلياً
        if all(not v.strip() for v in values):
            continue

        row = {
            "item":       get_cell(values, "ITEM"),
            "date":       get_cell(values, "DATE"),
            "flight":     get_cell(values, "FLIGHT"),
            "std_etd":    get_cell(values, "STD/ETD"),
            "dest":       get_cell(values, "DEST"),
            "email_time": get_cell(values, "Email Received Time"),
            "physical":   get_cell(values, "Physical Cargo received from Ramp"),
            "trolley":    get_cell(values, "Trolley/ ULD Number"),
            "cms":        get_cell(values, "Offloading Process Completed in CMS"),
            "pieces":     get_cell(values, "Offloading Pieces Verification"),
            "reason":     get_cell(values, "Offloading Reason"),
            "remarks":    get_cell(values, "Remarks/Additional Information"),
        }

        # تجاهل الصفوف التي تفتقر إلى أي بيانات رئيسية
        if not any([row["flight"], row["date"], row["dest"], row["trolley"]]):
            continue

        rows.append(row)

    return rows


def extract_structured_data(html: str) -> list[dict]:
    """نقطة الدخول الرئيسية لتحليل HTML واستخراج بيانات الشحنات."""
    soup = BeautifulSoup(html, "html.parser")
    table = find_main_table(soup)
    if not table:
        return []
    return parse_rows_from_table(table)


def group_rows_by_flight(rows: list[dict]) -> dict[str, dict]:
    """
    تجميع صفوف الشحنات حسب الرحلة الجوية.
    مفتاح التجميع: (DATE, FLIGHT, STD/ETD, DEST).
    """
    grouped: dict[str, dict] = {}

    for row in rows:
        key = slugify(
            f"{row.get('date', '')}_{row.get('flight', '')}"
            f"_{row.get('std_etd', '')}_{row.get('dest', '')}"
        )

        if key not in grouped:
            grouped[key] = {
                "date":    row.get("date", ""),
                "flight":  row.get("flight", ""),
                "std_etd": row.get("std_etd", ""),
                "dest":    row.get("dest", ""),
                "items":   [],
            }

        grouped[key]["items"].append(row)

    return grouped


# ══════════════════════════════════════════════════════════════════
#  التخزين وتتبع التحديثات
# ══════════════════════════════════════════════════════════════════

def save_or_update_flights(
    grouped: dict[str, dict],
    now: datetime,
) -> tuple[str, str, dict]:
    """
    حفظ بيانات كل رحلة في ملف JSON منفصل،
    وتحديث ملف meta.json بعدد مرات التحديث وآخر توقيت.

    يُعيد: (date_dir, shift, meta)
    """
    date_dir = now.strftime("%Y-%m-%d")
    shift    = get_shift(now)
    folder   = DATA_DIR / date_dir / shift
    folder.mkdir(parents=True, exist_ok=True)

    meta_path = folder / "meta.json"
    meta      = load_json(meta_path, {"flights": {}})

    for key, flight_info in grouped.items():
        filename  = (
            f"{slugify(flight_info['flight'])}"
            f"_{slugify(flight_info['date'])}"
            f"_{slugify(flight_info['dest'])}.json"
        )
        file_path = folder / filename
        existed   = file_path.exists()

        payload = {
            "date":     flight_info["date"],
            "flight":   flight_info["flight"],
            "std_etd":  flight_info["std_etd"],
            "dest":     flight_info["dest"],
            "items":    flight_info["items"],
            "saved_at": now.isoformat(),
        }
        file_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        entry = meta["flights"].get(filename, {
            "flight":     flight_info["flight"],
            "date":       flight_info["date"],
            "dest":       flight_info["dest"],
            "created_at": now.isoformat(),
            "updated_at": now.isoformat(),
            "updates":    0,
        })

        if existed:
            entry["updates"] = int(entry.get("updates", 0)) + 1

        entry["updated_at"]          = now.isoformat()
        meta["flights"][filename]    = entry

    write_json(meta_path, meta)
    return date_dir, shift, meta


# ══════════════════════════════════════════════════════════════════
#  إنشاء تقارير HTML
# ══════════════════════════════════════════════════════════════════

def _build_css() -> str:
    return """
    :root {
        --bg-page:    #eef2ff;
        --bg-card:    #ffffff;
        --color-text: #0f172a;
        --color-muted:#64748b;
        --color-line: #e2e8f0;
        --color-blue: #0b5ed7;
        --color-blue2:#0a3f9c;
        --color-badge:#f59e0b;
    }

    * { box-sizing: border-box; }

    body {
        margin: 0;
        background: var(--bg-page);
        font-family: Calibri, Arial, sans-serif;
        color: var(--color-text);
    }

    .wrap {
        max-width: 1400px;
        margin: 0 auto;
        padding: 24px 18px;
    }

    /* ── شريط العنوان ── */
    .topbar {
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        gap: 12px;
        flex-wrap: wrap;
        margin-bottom: 20px;
    }

    h1 { margin: 0; font-size: 26px; }

    .subtitle {
        color: var(--color-muted);
        font-size: 13px;
        margin-top: 4px;
    }

    /* ── بطاقة الرحلة ── */
    .flight-card {
        background: var(--bg-card);
        border: 1px solid var(--color-line);
        border-radius: 14px;
        overflow: hidden;
        margin: 14px 0;
        box-shadow: 0 2px 10px rgba(2, 6, 23, .07);
    }

    .flight-head {
        background: linear-gradient(90deg, var(--color-blue), var(--color-blue2));
        color: #fff;
        padding: 11px 16px;
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
        align-items: center;
    }

    .flight-name {
        font-weight: 900;
        font-size: 16px;
        letter-spacing: .4px;
    }

    .pill {
        background: rgba(255, 255, 255, .18);
        border: 1px solid rgba(255, 255, 255, .25);
        padding: 2px 11px;
        border-radius: 999px;
        font-size: 12px;
        white-space: nowrap;
    }

    .pill-updated {
        background: var(--color-badge);
        color: #111827;
        font-weight: 900;
        border: none;
    }

    /* ── جدول الشحنات ── */
    table { width: 100%; border-collapse: collapse; }

    th, td {
        border-top: 1px solid var(--color-line);
        padding: 9px 11px;
        font-size: 13px;
        vertical-align: top;
    }

    th {
        background: #f8fafc;
        text-align: left;
        color: #334155;
        font-weight: 800;
        white-space: nowrap;
    }

    td { color: var(--color-text); }

    .mono { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; }
    .muted { color: var(--color-muted); }

    .footer {
        margin-top: 28px;
        color: var(--color-muted);
        font-size: 12px;
        text-align: center;
    }
    """


def _render_flight_card(payload: dict, meta_entry: dict) -> str:
    """بناء HTML لبطاقة رحلة واحدة مع جدول بنودها."""
    flight   = payload.get("flight", "")
    date     = payload.get("date", "")
    std_etd  = payload.get("std_etd", "")
    dest     = payload.get("dest", "")
    items    = payload.get("items", [])
    updates  = int(meta_entry.get("updates", 0))
    upd_at   = meta_entry.get("updated_at", "")

    update_badge = (
        f'<span class="pill pill-updated">UPDATED ×{updates}</span>'
        if updates > 0 else ""
    )

    header = f"""
    <div class="flight-head">
        <span class="flight-name">{flight}</span>
        <span class="pill">DATE: <b>{date}</b></span>
        <span class="pill">STD/ETD: <b>{std_etd}</b></span>
        <span class="pill">DEST: <b>{dest}</b></span>
        {update_badge}
        <span class="pill">Last update: <b class="mono">{upd_at}</b></span>
    </div>
    """

    rows_html = "".join(f"""
        <tr>
            <td class="mono">{r.get('item', '')}</td>
            <td class="mono">{r.get('email_time', '')}</td>
            <td class="mono">{r.get('physical', '')}</td>
            <td>{r.get('trolley', '')}</td>
            <td class="mono">{r.get('cms', '')}</td>
            <td class="mono">{r.get('pieces', '')}</td>
            <td>{r.get('reason', '')}</td>
            <td>{r.get('remarks', '')}</td>
        </tr>
    """ for r in items)

    table = f"""
    <table>
        <thead>
            <tr>
                <th>ITEM</th>
                <th>Email Time</th>
                <th>Physical</th>
                <th>Trolley / ULD</th>
                <th>CMS</th>
                <th>Pieces</th>
                <th>Reason</th>
                <th>Remarks</th>
            </tr>
        </thead>
        <tbody>
            {rows_html}
        </tbody>
    </table>
    """

    return f'<div class="flight-card">{header}{table}</div>'


def build_shift_report(date_dir: str, shift: str) -> None:
    """إنشاء تقرير HTML لوردية محددة وحفظه في docs/."""
    folder = DATA_DIR / date_dir / shift
    if not folder.exists():
        return

    meta          = load_json(folder / "meta.json", {"flights": {}})
    flights_meta  = meta.get("flights", {})
    flight_files  = sorted(p for p in folder.glob("*.json") if p.name != "meta.json")

    cards = []
    for path in flight_files:
        payload     = json.loads(path.read_text(encoding="utf-8"))
        meta_entry  = flights_meta.get(path.name, {})
        cards.append(_render_flight_card(payload, meta_entry))

    body_content = "\n".join(cards) if cards else "<p class='muted'>لا توجد بيانات بعد.</p>"

    shift_label = {
        "shift1": "Shift 1 — 06:00 to 14:30",
        "shift2": "Shift 2 — 14:30 to 21:30",
        "shift3": "Shift 3 — 21:30 to 06:00",
    }.get(shift, shift)

    html = f"""<!doctype html>
<html lang="ar" dir="ltr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Offload Monitor — {date_dir} — {shift}</title>
    <style>{_build_css()}</style>
</head>
<body>
<div class="wrap">
    <div class="topbar">
        <div>
            <h1>✈ Offload Monitor Report</h1>
            <div class="subtitle">{date_dir} — {shift_label}</div>
        </div>
    </div>

    {body_content}

    <div class="footer">
        Generated automatically by GitHub Actions · {date_dir}
    </div>
</div>
</body>
</html>"""

    out_dir = DOCS_DIR / date_dir / shift
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "index.html").write_text(html, encoding="utf-8")


def build_root_index() -> None:
    """إنشاء الصفحة الرئيسية docs/index.html بروابط لجميع التقارير."""
    if not DOCS_DIR.exists():
        return

    day_dirs = sorted(
        (p for p in DOCS_DIR.iterdir()
         if p.is_dir() and re.match(r"\d{4}-\d{2}-\d{2}", p.name)),
        reverse=True,
    )

    links: list[tuple[str, str, str]] = []
    for day_dir in day_dirs:
        for shift in ("shift1", "shift2", "shift3"):
            if (day_dir / shift / "index.html").exists():
                links.append((day_dir.name, shift, f"{day_dir.name}/{shift}/"))

    list_items = "".join(
        f"<li><a href='{href}'>{day} — {shift}</a></li>"
        for day, shift, href in links
    ) or "<li>No reports yet.</li>"

    html = f"""<!doctype html>
<html lang="ar" dir="ltr">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Offload Monitor</title>
    <style>
        body {{
            font-family: Calibri, Arial, sans-serif;
            background: #eef2ff;
            margin: 0;
        }}
        .wrap {{
            max-width: 900px;
            margin: 0 auto;
            padding: 24px 18px;
        }}
        .card {{
            background: #fff;
            border: 1px solid #e2e8f0;
            border-radius: 14px;
            padding: 20px 24px;
            box-shadow: 0 2px 10px rgba(2, 6, 23, .06);
        }}
        h1 {{ margin: 0 0 10px 0; font-size: 22px; }}
        ul {{ margin: 12px 0 0 20px; line-height: 1.9; }}
        a {{
            text-decoration: none;
            color: #0b5ed7;
            font-weight: 600;
        }}
        a:hover {{ text-decoration: underline; }}
    </style>
</head>
<body>
<div class="wrap">
    <div class="card">
        <h1>✈ Offload Monitor</h1>
        <p style="color:#64748b;margin:0 0 6px 0">Select a shift report to view:</p>
        <ul>{list_items}</ul>
    </div>
</div>
</body>
</html>"""

    (DOCS_DIR / "index.html").write_text(html, encoding="utf-8")


# ══════════════════════════════════════════════════════════════════
#  نقطة الدخول الرئيسية
# ══════════════════════════════════════════════════════════════════

def main() -> None:
    now = datetime.now(ZoneInfo(TIMEZONE))

    print(f"[{now.isoformat()}] Downloading file…")
    html     = download_file()
    new_hash = compute_sha256(html)

    # التحقق من وجود تغيير في المحتوى
    if STATE_FILE.exists():
        old_hash = STATE_FILE.read_text(encoding="utf-8").strip()
        if old_hash == new_hash:
            print("No change detected. Exiting.")
            return

    print("Change detected. Parsing table data…")
    rows = extract_structured_data(html)

    if not rows:
        print("WARNING: No table rows extracted. Check HTML structure/headers.")
        # حفظ الـ hash لتجنب إعادة المعالجة على نفس المحتوى الخاطئ
        STATE_FILE.write_text(new_hash, encoding="utf-8")
        return

    print(f"Extracted {len(rows)} row(s). Grouping by flight…")
    grouped = group_rows_by_flight(rows)

    date_dir, shift, _ = save_or_update_flights(grouped, now)

    print(f"Building shift report: {date_dir} / {shift}…")
    build_shift_report(date_dir, shift)

    print("Rebuilding root index…")
    build_root_index()

    STATE_FILE.write_text(new_hash, encoding="utf-8")
    print("Done. Docs report updated successfully. ✓")


if __name__ == "__main__":
    main()

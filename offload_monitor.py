"""
╔══════════════════════════════════════════════════════════════════╗
║              OFFLOAD MONITOR — Cargo Tracking System             ║
║              Automated HTML Report Generator via GitHub Actions  ║
╚══════════════════════════════════════════════════════════════════╝
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
DOCS_DIR:   Path = Path("docs")


# ══════════════════════════════════════════════════════════════════
#  الدوال المساعدة
# ══════════════════════════════════════════════════════════════════

def download_file() -> str:
    url = ONEDRIVE_URL.strip()
    separator = "&" if "?" in url else "?"
    if "download=1" not in url:
        url += f"{separator}download=1"
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return response.text


def compute_sha256(content: str) -> str:
    return hashlib.sha256(content.encode("utf-8", errors="ignore")).hexdigest()


def get_shift(now: datetime) -> str:
    mins = now.hour * 60 + now.minute
    if 6 * 60 <= mins < 14 * 60 + 30:
        return "shift1"
    if 14 * 60 + 30 <= mins < 21 * 60 + 30:
        return "shift2"
    return "shift3"


def slugify(text: str, max_length: int = 80) -> str:
    text = (text or "UNKNOWN").strip()
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^A-Za-z0-9_-]", "_", text)
    return (text or "UNKNOWN")[:max_length]


def load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return default


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def cell_text(element) -> str:
    if element is None:
        return ""
    text = element.get_text(" ", strip=True)
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace("\xa0", "").strip()
    return text


def row_texts(tr) -> list[str]:
    return [cell_text(c) for c in tr.find_all(["td", "th"])]


def _get(row: list[str], idx: int | None) -> str:
    if idx is None or idx >= len(row):
        return ""
    return row[idx]


def _find_value_after(row: list[str], keys: list[str]) -> str:
    for i, cell in enumerate(row):
        if cell.upper().strip() in [k.upper() for k in keys]:
            if i + 1 < len(row):
                return row[i + 1]
    return ""


def _find_index(row: list[str], keys: list[str]) -> int | None:
    for i, cell in enumerate(row):
        if any(k.upper() in cell.upper() for k in keys):
            return i
    return None


# ══════════════════════════════════════════════════════════════════
#  تحليل HTML
#
#  النوع A (الحالي): جدول أفقي
#    Row: FLIGHT # | WY223 | DATE | 18.JUL | DESTINATION | COK
#    Row: AWB | PCS | KGS | DESCRIPTION | REASON
#    Row: 910... | 35 | 781 | COURIER | SPACE
#    Row: TOTAL | 35 | 781 | ...
#
#  النوع B (مستقبلي): جدول عمودي
#    Header: ITEM | DATE | FLIGHT | STD/ETD | DEST | Email | Physical | ...
#    Row:    1    | 26FEB| WY251  | 0120    | MAA  | 0124  | 0134     | ...
# ══════════════════════════════════════════════════════════════════

def extract_flights(html: str) -> list[dict]:
    soup   = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")
    if not tables:
        return []

    best: list[dict] = []

    for table in tables:
        rows = table.find_all("tr")
        if len(rows) < 3:
            continue
        all_rows = [row_texts(tr) for tr in rows]

        # جرّب النوع A أولاً (الأولوية لأنه التنسيق الحالي)
        result_a = _parse_type_a(all_rows)

        # جرّب النوع B فقط إذا فشل النوع A
        if result_a:
            if len(result_a) > len(best):
                best = result_a
        else:
            result_b = _parse_type_b(all_rows)
            if result_b and len(result_b) > len(best):
                best = result_b

    return best


def _parse_type_a(all_rows: list[list[str]]) -> list[dict]:
    """
    يبحث عن صفوف تحتوي على FLIGHT # + DATE + DESTINATION في صف واحد أفقي
    بنمط key-value (الخلية الزوجية = مفتاح، الخلية الفردية = قيمة).
    ثم يقرأ AWB/PCS/KGS/DESCRIPTION/REASON من الصفوف التالية.
    """
    flights = []
    i = 0
    while i < len(all_rows):
        row    = all_rows[i]
        joined = " ".join(row).upper()

        # الشرط: الصف يحتوي كلمة FLIGHT (كمفتاح) وDATE وDESTINATION
        # ويجب أن تكون القيم في الخلايا المجاورة (ليس كهيدر عمودي)
        if ("FLIGHT" in joined and "DATE" in joined and
                ("DESTINATION" in joined or "DEST" in joined)):

            flight_num  = _find_value_after(row, ["FLIGHT #", "FLIGHT#", "FLIGHT"])
            date        = _find_value_after(row, ["DATE"])
            destination = _find_value_after(row, ["DESTINATION", "DEST"])

            # تجاهل الصف إذا كانت القيم كلها فارغة (يعني هيدر وليس بيانات)
            if not flight_num and not destination:
                i += 1
                continue

            # تجاهل إذا كانت القيم نفسها كلمات هيدر (النوع B)
            if flight_num.upper() in ("ITEM", "AWB", "PCS", ""):
                i += 1
                continue

            # الصف التالي: هيدر الشحنات
            cargo_header = all_rows[i + 1] if i + 1 < len(all_rows) else []
            awb_idx  = _find_index(cargo_header, ["AWB"])
            pcs_idx  = _find_index(cargo_header, ["PCS", "PIECES"])
            kgs_idx  = _find_index(cargo_header, ["KGS", "KG"])
            desc_idx = _find_index(cargo_header, ["DESCRIPTION", "DESC"])
            rsn_idx  = _find_index(cargo_header, ["REASON"])

            items = []
            j = i + 2
            while j < len(all_rows):
                dr     = all_rows[j]
                dr_str = " ".join(dr).upper()
                if "TOTAL" in dr_str:
                    j += 1
                    break
                if "FLIGHT" in dr_str and "DATE" in dr_str:
                    break

                awb  = _get(dr, awb_idx)
                pcs  = _get(dr, pcs_idx)
                kgs  = _get(dr, kgs_idx)
                desc = _get(dr, desc_idx)
                rsn  = _get(dr, rsn_idx)

                if not any([awb, pcs, kgs, desc, rsn]):
                    j += 1
                    continue

                items.append({"awb": awb, "pcs": pcs, "kgs": kgs,
                               "description": desc, "reason": rsn})
                j += 1

            flights.append({
                "flight":      flight_num,
                "date":        date,
                "std_etd":     "",
                "destination": destination,
                "format":      "A",
                "items":       items,
            })
            i = j
        else:
            i += 1

    return flights


def _parse_type_b(all_rows: list[list[str]]) -> list[dict]:
    """
    يبحث عن صف هيدر عمودي (ITEM, DATE, FLIGHT, DEST...)
    ثم يقرأ بيانات الشحنات صفاً صفاً.
    """
    header_idx = None
    headers    = []
    for i, row in enumerate(all_rows):
        joined = " ".join(row).upper()
        hits   = sum(1 for kw in ["ITEM", "DATE", "FLIGHT", "DEST"] if kw in joined)
        if hits >= 3:
            header_idx = i
            headers    = [h.upper().strip() for h in row]
            break

    if header_idx is None:
        return []

    def col(names: list[str]) -> int | None:
        for name in names:
            for j, h in enumerate(headers):
                if name in h:
                    return j
        return None

    c_item  = col(["ITEM"])
    c_date  = col(["DATE"])
    c_flt   = col(["FLIGHT"])
    c_std   = col(["STD", "ETD"])
    c_dest  = col(["DEST"])
    c_email = col(["EMAIL"])
    c_phys  = col(["PHYSICAL"])
    c_trol  = col(["TROLLEY", "ULD"])
    c_cms   = col(["CMS", "OFFLOADING PROCESS"])
    c_pcs   = col(["PIECES", "VERIFICATION"])
    c_rsn   = col(["REASON"])
    c_rmk   = col(["REMARKS"])

    flights: list[dict] = []
    current: dict | None = None

    for row in all_rows[header_idx + 1:]:
        if all(not v for v in row):
            continue

        flt  = _get(row, c_flt)
        date = _get(row, c_date)
        dest = _get(row, c_dest)

        is_new = bool(flt or date) and (
            current is None
            or flt  != current.get("flight", "")
            or dest != current.get("destination", "")
        )

        if is_new:
            current = {
                "flight":      flt,
                "date":        date,
                "std_etd":     _get(row, c_std),
                "destination": dest,
                "format":      "B",
                "items":       [],
            }
            flights.append(current)

        if current is not None:
            item = {
                "item":     _get(row, c_item),
                "email":    _get(row, c_email),
                "physical": _get(row, c_phys),
                "trolley":  _get(row, c_trol),
                "cms":      _get(row, c_cms),
                "pieces":   _get(row, c_pcs),
                "reason":   _get(row, c_rsn),
                "remarks":  _get(row, c_rmk),
            }
            if any(item.values()):
                current["items"].append(item)

    return flights


# ══════════════════════════════════════════════════════════════════
#  التخزين
# ══════════════════════════════════════════════════════════════════

def save_flights(flights: list[dict], now: datetime) -> tuple[str, str, dict]:
    date_dir = now.strftime("%Y-%m-%d")
    shift    = get_shift(now)
    folder   = DATA_DIR / date_dir / shift
    folder.mkdir(parents=True, exist_ok=True)

    meta_path = folder / "meta.json"
    meta      = load_json(meta_path, {"flights": {}})

    for flight in flights:
        filename  = slugify(f"{flight['flight']}_{flight.get('date','')}_{flight['destination']}") + ".json"
        file_path = folder / filename
        existed   = file_path.exists()

        payload = {**flight, "saved_at": now.isoformat()}
        file_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        entry = meta["flights"].get(filename, {
            "flight":     flight["flight"],
            "date":       flight.get("date", ""),
            "dest":       flight["destination"],
            "created_at": now.isoformat(),
            "updated_at": now.isoformat(),
            "updates":    0,
        })
        if existed:
            entry["updates"] = int(entry.get("updates", 0)) + 1
        entry["updated_at"]       = now.isoformat()
        meta["flights"][filename] = entry

    write_json(meta_path, meta)
    return date_dir, shift, meta


# ══════════════════════════════════════════════════════════════════
#  CSS
# ══════════════════════════════════════════════════════════════════

def _css() -> str:
    return """
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@400;600;700&display=swap');

    :root {
        --bg:       #f0f4f8;
        --card:     #ffffff;
        --line:     #dde3ec;
        --blue:     #1a56db;
        --blue-dk:  #1240a8;
        --blue-hd:  #0d2d6e;
        --gray-hd:  #e8edf5;
        --text:     #0f1f3d;
        --muted:    #6b7a99;
        --badge:    #f59e0b;
        --row-alt:  #f7f9fc;
        --total-bg: #e8f0fe;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
        background: var(--bg);
        font-family: 'IBM Plex Sans', Calibri, sans-serif;
        color: var(--text);
        font-size: 13px;
    }

    .page-header {
        background: linear-gradient(135deg, var(--blue-hd) 0%, var(--blue-dk) 100%);
        color: #fff;
        padding: 18px 32px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        flex-wrap: wrap;
        box-shadow: 0 2px 12px rgba(13,45,110,.25);
    }

    .page-header h1   { font-size: 19px; font-weight: 700; }
    .page-header .sub { font-size: 12px; opacity: .75; margin-top: 3px; }

    .stat-box {
        background: rgba(255,255,255,.13);
        border: 1px solid rgba(255,255,255,.22);
        border-radius: 8px;
        padding: 7px 18px;
        text-align: center;
        font-size: 12px;
    }
    .stat-box strong { display: block; font-size: 22px; font-weight: 700; }

    .wrap { max-width: 1280px; margin: 0 auto; padding: 24px 20px 48px; }

    /* ── الجدول الرئيسي ── */
    .report-table {
        width: 100%;
        border-collapse: collapse;
        background: var(--card);
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 2px 14px rgba(15,31,61,.08);
        border: 1px solid var(--line);
    }

    .report-table thead th {
        background: var(--blue-hd);
        color: #fff;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: .5px;
        padding: 10px 11px;
        border-right: 1px solid rgba(255,255,255,.1);
        white-space: nowrap;
    }
    .report-table thead th:last-child { border-right: none; }

    /* ══ الصف الأول: معلومات الرحلة ══ */
    tr.row-flight td {
        background: var(--gray-hd);
        border-top: 3px solid var(--blue);
        border-bottom: 1px solid var(--line);
        padding: 10px 12px;
        vertical-align: middle;
    }

    .flt-num {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 16px;
        font-weight: 700;
        color: var(--blue-hd);
        white-space: nowrap;
    }

    .lbl {
        font-size: 10px;
        font-weight: 600;
        color: var(--muted);
        text-transform: uppercase;
        letter-spacing: .4px;
        display: block;
        margin-bottom: 1px;
    }

    .val { font-weight: 700; color: var(--blue-dk); }

    .badge-upd {
        background: var(--badge);
        color: #111;
        font-size: 11px;
        font-weight: 800;
        padding: 2px 9px;
        border-radius: 999px;
        margin-right: 6px;
    }

    /* ══ الصف الثاني: هيدر الشحنة ══ */
    tr.row-subhead td {
        background: #dce8ff;
        padding: 6px 11px;
        font-size: 11px;
        font-weight: 700;
        color: var(--blue-hd);
        text-transform: uppercase;
        letter-spacing: .4px;
        border-bottom: 1px solid #c3d4f5;
    }

    /* ══ الصف الثاني: بيانات الشحنة ══ */
    tr.row-cargo td {
        padding: 8px 11px;
        border-bottom: 1px solid #edf0f7;
        border-right: 1px solid #edf0f7;
        vertical-align: top;
    }
    tr.row-cargo:nth-child(odd)  td { background: var(--card); }
    tr.row-cargo:nth-child(even) td { background: var(--row-alt); }
    tr.row-cargo td:last-child { border-right: none; }

    .mono { font-family: 'IBM Plex Mono', monospace; font-size: 12px; }
    .dim  { color: var(--muted); }

    .reason-tag {
        display: inline-block;
        background: #fff3e0;
        color: #b45309;
        border: 1px solid #fcd34d;
        border-radius: 5px;
        padding: 1px 8px;
        font-size: 11px;
        font-weight: 600;
    }

    /* ══ الصف الثالث: التعبئة ══ */
    tr.row-totals td {
        background: var(--total-bg);
        border-top: 2px solid var(--blue);
        border-bottom: 3px solid var(--blue-hd);
        padding: 8px 12px;
        font-weight: 700;
        color: var(--blue-hd);
        font-size: 12px;
    }
    .tlbl { font-size: 10px; font-weight: 400; color: var(--muted); display: block; }

    .footer { text-align:center; color:var(--muted); font-size:12px; margin-top:28px; }
    """


# ══════════════════════════════════════════════════════════════════
#  تصيير الرحلات
# ══════════════════════════════════════════════════════════════════

def _render_flight(flight: dict, meta_entry: dict) -> str:
    items     = flight.get("items", [])
    fmt       = flight.get("format", "A")
    updates   = int(meta_entry.get("updates", 0))
    upd_at    = meta_entry.get("updated_at", "")[:16].replace("T", " ")
    upd_badge = f'<span class="badge-upd">UPDATED ×{updates}</span>' if updates > 0 else ""

    def safe_sum(key: str) -> int:
        total = 0
        for it in items:
            try:
                total += int(re.sub(r"[^\d]", "", it.get(key, "") or "0") or 0)
            except Exception:
                pass
        return total

    # ══════════════════════════════════════════════
    #  النوع A: AWB / PCS / KGS / DESCRIPTION / REASON
    # ══════════════════════════════════════════════
    if fmt == "A":
        total_pcs = safe_sum("pcs")
        total_kgs = safe_sum("kgs")

        row_flight = f"""
        <tr class="row-flight">
            <td colspan="2" class="flt-num">✈ {flight['flight'] or '—'}</td>
            <td><span class="lbl">DATE</span><span class="val">{flight.get('date','') or '—'}</span></td>
            <td colspan="3"><span class="lbl">DESTINATION</span><span class="val">{flight['destination'] or '—'}</span></td>
            <td colspan="2">
                {upd_badge}
                <span class="dim" style="font-size:11px">Last: {upd_at}</span>
            </td>
        </tr>"""

        row_subhead = """
        <tr class="row-subhead">
            <td colspan="2">AWB</td>
            <td>PCS</td>
            <td>KGS</td>
            <td colspan="2">DESCRIPTION</td>
            <td colspan="2">REASON</td>
        </tr>"""

        cargo_rows = ""
        for it in items:
            rsn = it.get("reason", "")
            reason_cell = f'<span class="reason-tag">{rsn}</span>' if rsn else '<span class="dim">—</span>'
            cargo_rows += f"""
            <tr class="row-cargo">
                <td colspan="2" class="mono">{it.get('awb','') or '—'}</td>
                <td class="mono">{it.get('pcs','') or '—'}</td>
                <td class="mono">{it.get('kgs','') or '—'}</td>
                <td colspan="2">{it.get('description','') or '—'}</td>
                <td colspan="2">{reason_cell}</td>
            </tr>"""

        if not cargo_rows:
            cargo_rows = '<tr class="row-cargo"><td colspan="8" class="dim" style="text-align:center;padding:14px">No cargo data</td></tr>'

        row_totals = f"""
        <tr class="row-totals">
            <td colspan="2"><span class="tlbl">TOTAL SHIPMENTS</span>{len(items)}</td>
            <td><span class="tlbl">TOTAL PCS</span>{total_pcs or '—'}</td>
            <td><span class="tlbl">TOTAL KGS</span>{total_kgs or '—'}</td>
            <td colspan="4"></td>
        </tr>"""

        return row_flight + row_subhead + cargo_rows + row_totals

    # ══════════════════════════════════════════════
    #  النوع B: ITEM / EMAIL / PHYSICAL / TROLLEY / CMS / PIECES / REASON / REMARKS
    # ══════════════════════════════════════════════
    total_pcs = safe_sum("pieces")

    row_flight = f"""
    <tr class="row-flight">
        <td class="flt-num">✈ {flight['flight'] or '—'}</td>
        <td><span class="lbl">DATE</span><span class="val">{flight.get('date','') or '—'}</span></td>
        <td><span class="lbl">STD/ETD</span><span class="val">{flight.get('std_etd','') or '—'}</span></td>
        <td><span class="lbl">DEST</span><span class="val">{flight['destination'] or '—'}</span></td>
        <td colspan="4">
            {upd_badge}
            <span class="dim" style="font-size:11px">Last: {upd_at}</span>
        </td>
    </tr>"""

    row_subhead = """
    <tr class="row-subhead">
        <td>ITEM</td>
        <td>EMAIL TIME</td>
        <td>PHYSICAL</td>
        <td>TROLLEY / ULD</td>
        <td>CMS</td>
        <td>PIECES</td>
        <td>REASON</td>
        <td>REMARKS</td>
    </tr>"""

    cargo_rows = ""
    for it in items:
        rsn = it.get("reason", "")
        reason_cell = f'<span class="reason-tag">{rsn}</span>' if rsn else '<span class="dim">—</span>'
        cargo_rows += f"""
        <tr class="row-cargo">
            <td class="mono">{it.get('item','') or '—'}</td>
            <td class="mono">{it.get('email','') or '—'}</td>
            <td class="mono">{it.get('physical','') or '—'}</td>
            <td>{it.get('trolley','') or '—'}</td>
            <td class="mono">{it.get('cms','') or '—'}</td>
            <td class="mono">{it.get('pieces','') or '—'}</td>
            <td>{reason_cell}</td>
            <td>{it.get('remarks','') or '—'}</td>
        </tr>"""

    if not cargo_rows:
        cargo_rows = '<tr class="row-cargo"><td colspan="8" class="dim" style="text-align:center;padding:14px">No cargo data</td></tr>'

    row_totals = f"""
    <tr class="row-totals">
        <td colspan="5"><span class="tlbl">TOTAL SHIPMENTS</span>{len(items)}</td>
        <td><span class="tlbl">TOTAL PIECES</span>{total_pcs or '—'}</td>
        <td colspan="2"></td>
    </tr>"""

    return row_flight + row_subhead + cargo_rows + row_totals


# ══════════════════════════════════════════════════════════════════
#  بناء صفحات HTML
# ══════════════════════════════════════════════════════════════════

def build_shift_report(date_dir: str, shift: str) -> None:
    folder = DATA_DIR / date_dir / shift
    if not folder.exists():
        return

    meta         = load_json(folder / "meta.json", {"flights": {}})
    flight_files = sorted(p for p in folder.glob("*.json") if p.name != "meta.json")
    flights      = [json.loads(p.read_text(encoding="utf-8")) for p in flight_files]

    shift_labels = {
        "shift1": "Shift 1 — 06:00 to 14:30",
        "shift2": "Shift 2 — 14:30 to 21:30",
        "shift3": "Shift 3 — 21:30 to 06:00",
    }
    shift_label   = shift_labels.get(shift, shift)
    total_flights = len(flights)
    total_items   = sum(len(f.get("items", [])) for f in flights)

    tbody = ""
    for flight in flights:
        fname  = slugify(f"{flight['flight']}_{flight.get('date','')}_{flight['destination']}") + ".json"
        meta_e = meta.get("flights", {}).get(fname, {})
        tbody += _render_flight(flight, meta_e)

    if not tbody:
        tbody = '<tr><td colspan="8" style="text-align:center;padding:36px;color:#6b7a99">No offload data recorded for this shift.</td></tr>'

    html = f"""<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Offload Monitor — {date_dir} — {shift}</title>
    <style>{_css()}</style>
</head>
<body>
<div class="page-header">
    <div>
        <h1>✈ Offload Monitor Report</h1>
        <div class="sub">{date_dir} &nbsp;·&nbsp; {shift_label}</div>
    </div>
    <div style="display:flex;gap:10px;flex-wrap:wrap">
        <div class="stat-box"><strong>{total_flights}</strong>Flights</div>
        <div class="stat-box"><strong>{total_items}</strong>Shipments</div>
    </div>
</div>
<div class="wrap">
    <table class="report-table">
        <thead>
            <tr>
                <th>FLIGHT / ITEM</th>
                <th>DATE / EMAIL</th>
                <th>STD·ETD / PHYSICAL</th>
                <th>DEST / TROLLEY·ULD</th>
                <th>AWB · DESC / CMS</th>
                <th>PCS</th>
                <th>KGS / REASON</th>
                <th>REMARKS / STATUS</th>
            </tr>
        </thead>
        <tbody>{tbody}</tbody>
    </table>
    <div class="footer">Generated automatically by GitHub Actions &nbsp;·&nbsp; {date_dir}</div>
</div>
</body>
</html>"""

    out_dir = DOCS_DIR / date_dir / shift
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "index.html").write_text(html, encoding="utf-8")


def build_root_index() -> None:
    if not DOCS_DIR.exists():
        return

    day_dirs = sorted(
        (p for p in DOCS_DIR.iterdir()
         if p.is_dir() and re.match(r"\d{4}-\d{2}-\d{2}", p.name)),
        reverse=True,
    )

    links = []
    for day_dir in day_dirs:
        for shift in ("shift1", "shift2", "shift3"):
            if (day_dir / shift / "index.html").exists():
                links.append((day_dir.name, shift, f"{day_dir.name}/{shift}/"))

    items_html = "".join(
        f'<li><a href="{href}">✈ {day} &nbsp;·&nbsp; {s}</a></li>'
        for day, s, href in links
    ) or "<li style='color:#6b7a99;padding:12px 4px'>No reports yet.</li>"

    html = f"""<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Offload Monitor</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600;700&display=swap');
        body {{ font-family:'IBM Plex Sans',Calibri,sans-serif; background:#f0f4f8; margin:0; }}
        .wrap {{ max-width:640px; margin:60px auto; padding:0 20px; }}
        .card {{ background:#fff; border:1px solid #dde3ec; border-radius:16px;
                 padding:32px 36px; box-shadow:0 4px 20px rgba(15,31,61,.09); }}
        h1 {{ font-size:22px; color:#0d2d6e; margin-bottom:6px; }}
        p  {{ color:#6b7a99; font-size:13px; margin-bottom:20px; }}
        ul {{ list-style:none; padding:0; }}
        li {{ border-bottom:1px solid #edf0f7; }}
        li:last-child {{ border-bottom:none; }}
        a  {{ display:block; padding:12px 6px; text-decoration:none;
              color:#1a56db; font-weight:600; font-size:14px;
              transition:all .15s; border-radius:8px; }}
        a:hover {{ background:#f0f4ff; padding-left:14px; color:#0d2d6e; }}
    </style>
</head>
<body>
<div class="wrap">
    <div class="card">
        <h1>✈ Offload Monitor</h1>
        <p>Select a shift report to view cargo offload details.</p>
        <ul>{items_html}</ul>
    </div>
</div>
</body>
</html>"""

    (DOCS_DIR / "index.html").write_text(html, encoding="utf-8")


# ══════════════════════════════════════════════════════════════════
#  نقطة الدخول
# ══════════════════════════════════════════════════════════════════

def main() -> None:
    now = datetime.now(ZoneInfo(TIMEZONE))
    print(f"[{now.isoformat()}] Downloading file…")

    html     = download_file()
    new_hash = compute_sha256(html)

    if STATE_FILE.exists():
        if STATE_FILE.read_text(encoding="utf-8").strip() == new_hash:
            print("No change detected. Exiting.")
            return

    print("Change detected. Parsing…")
    flights = extract_flights(html)

    if not flights:
        print("WARNING: No flights extracted. Check HTML structure.")
        STATE_FILE.write_text(new_hash, encoding="utf-8")
        return

    print(f"Extracted {len(flights)} flight(s). Saving…")
    date_dir, shift, _ = save_flights(flights, now)

    print(f"Building report: {date_dir}/{shift}…")
    build_shift_report(date_dir, shift)
    build_root_index()

    STATE_FILE.write_text(new_hash, encoding="utf-8")
    print(f"Done. ✓  ({len(flights)} flights saved)")


if __name__ == "__main__":
    main()

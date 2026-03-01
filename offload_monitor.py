"""
╔══════════════════════════════════════════════════════════════════╗
║              OFFLOAD MONITOR — Cargo Tracking System             ║
║              Automated HTML Report Generator via GitHub Actions  ║
╚══════════════════════════════════════════════════════════════════╝

  النوع A : جدول HTML أفقي  (FLIGHT# | DATE | DESTINATION → AWB/PCS/KGS)
  النوع B : جدول HTML عمودي (ITEM | DATE | FLIGHT | STD/ETD | DEST ...)
  النوع C : نص عادي         (OFFLOADED CARGO ON OV237/27FEB + سطر بيانات)
"""

from __future__ import annotations

import os
import re
import json
import hashlib
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup


# ══════════════════════════════════════════════════════════════════
#  الإعدادات العامة
# ══════════════════════════════════════════════════════════════════

ONEDRIVE_URL: str = os.environ["ONEDRIVE_FILE_URL"]
TIMEZONE: str     = "Asia/Muscat"

# إذا كنت تريد إجبار بناء التقرير حتى لو الـ hash لم يتغير (مفيد للتجارب/التشخيص)
FORCE_REBUILD: bool = os.getenv("FORCE_REBUILD", "").strip().lower() in ("1", "true", "yes", "y")

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

    # OneDrive أحيانًا يعيد نسخة مخزّنة (cache) من رابط المشاركة.
    # لذلك نضيف باراميتر متغير + Headers لمنع الكاش.
    url += f"&__ts={int(datetime.now().timestamp())}"

    response = requests.get(
        url,
        timeout=30,
        headers={
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0",
        },
    )
    response.raise_for_status()
    return response.text


def compute_sha256(content: str) -> str:
    return hashlib.sha256(content.encode("utf-8", errors="ignore")).hexdigest()


def normalize_flight_date(date_str: str, now: datetime) -> str:
    """Try to convert email-style dates like '27FEB', '18.JUL', '27-FEB', '18 JUL' into 'YYYY-MM-DD'.
    Returns '' if parsing fails.
    """
    s = (date_str or "").strip().upper()
    if not s:
        return ""

    # Common separators
    s = s.replace("/", " ").replace("-", " ").replace(".", " ")
    s = re.sub(r"\s+", " ", s).strip()

    # Already ISO?
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return s

    m = re.fullmatch(r"(\d{1,2})\s*([A-Z]{3,})", s)
    if not m:
        # sometimes like '18 JUL 2026'
        m2 = re.fullmatch(r"(\d{1,2})\s*([A-Z]{3,})\s*(\d{4})", s)
        if m2:
            day = int(m2.group(1))
            mon = m2.group(2)[:3]
            year = int(m2.group(3))
        else:
            return ""
    else:
        day = int(m.group(1))
        mon = m.group(2)[:3]
        year = now.year

    months = {
        "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
        "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12
    }
    if mon not in months:
        return ""

    try:
        d = datetime(year, months[mon], day).date()
    except ValueError:
        return ""

    # If date looks far in the future relative to now (e.g., year rollover), adjust backward a year.
    if (d - now.date()).days > 180:
        try:
            d = datetime(year - 1, months[mon], day).date()
        except ValueError:
            pass

    return d.isoformat()




def _time_only(val: str) -> str:
    """Return HH:MM from a datetime-like string."""
    s = (val or "").strip()
    if not s:
        return ""
    # common patterns: '2026-03-01 15:00', '2026-03-01T15:00:00+04:00', '15:00'
    m = re.search(r"(\d{2}:\d{2})", s)
    return m.group(1) if m else ""


def fetch_std_etd_airlabs(
    flight_iata: str,
    *,
    flight_date: str | None = None,
    dep_iata: str | None = None,
    arr_iata: str | None = None,
) -> str:
    """Fetch STD/ETD from AirLabs if AIRLABS_API_KEY is set.

    Returns:
      - 'HH:MM'                 (STD only)
      - 'HH:MM|HH:MM'           (STD|ETD) if both available and different

    Note: We intentionally return *time only* (no date) because the report already shows the date.
    """
    api_key = os.environ.get("AIRLABS_API_KEY", "").strip()
    if not api_key:
        return ""

    flight_iata = (flight_iata or "").strip().upper()
    if not flight_iata:
        return ""

    url = "https://airlabs.co/api/v9/schedules"
    params: dict[str, str] = {"api_key": api_key, "flight_iata": flight_iata}
    if flight_date:
        # AirLabs supports date filters in some plans; if ignored, selection logic below still works.
        params["flight_date"] = flight_date
    if dep_iata:
        params["dep_iata"] = dep_iata.strip().upper()
    if arr_iata:
        params["arr_iata"] = arr_iata.strip().upper()

    try:
        r = requests.get(url, params=params, timeout=30)
        r.raise_for_status()
        payload = r.json()
    except Exception:
        return ""

    rows = payload.get("response") or payload.get("data") or []
    if not isinstance(rows, list) or not rows:
        return ""

    # Pick the best row:
    # 1) exact flight match
    # 2) if flight_date provided, prefer rows whose dep_time/departure date contains it
    candidates = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        if (row.get("flight_iata") or "").strip().upper() != flight_iata:
            continue
        candidates.append(row)

    if not candidates:
        candidates = [r for r in rows if isinstance(r, dict)] or []

    best = None
    if flight_date:
        for row in candidates:
            dep_time = str(row.get("dep_time") or row.get("dep_time_utc") or "")
            if flight_date in dep_time:
                best = row
                break

    if best is None:
        best = candidates[0] if candidates else None

    if not best:
        return ""

    std = _time_only(str(best.get("dep_time") or best.get("dep_time_utc") or ""))
    etd = _time_only(str(best.get("dep_estimated") or best.get("dep_estimated_utc") or ""))

    if std and etd and std != etd:
        return f"{std}|{etd}"
    if std:
        return std
    return ""


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
#  تحليل HTML / النص
# ══════════════════════════════════════════════════════════════════

def extract_flights(html: str) -> list[dict]:
    """
    يجرّب الأنواع الثلاثة ويعيد أفضل نتيجة.
    الأولوية: A → B → C
    """
    soup   = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")

    best: list[dict] = []

    for table in tables:
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        all_rows = [row_texts(tr) for tr in rows]

        result_a = _parse_type_a(all_rows)
        if result_a:
            if len(result_a) > len(best):
                best = result_a
            continue

        result_b = _parse_type_b(all_rows)
        if result_b and len(result_b) > len(best):
            best = result_b

    # النوع C: نص عادي (يُضاف فوق ما وجدناه من جداول)
    result_c = _parse_type_c(soup)
    best.extend(result_c)

    return best


# ────────────────────────────────────────────────────────────────
#  النوع A — جدول أفقي
#  Row: FLIGHT # | WY223 | DATE | 18.JUL | DESTINATION | COK
#  Row: AWB | PCS | KGS | DESCRIPTION | REASON
#  Row: 910... | 35 | 781 | COURIER | SPACE
# ────────────────────────────────────────────────────────────────
def _parse_type_a(all_rows: list[list[str]]) -> list[dict]:
    flights = []
    i = 0
    while i < len(all_rows):
        row    = all_rows[i]
        joined = " ".join(row).upper()

        if ("FLIGHT" in joined and "DATE" in joined and
                ("DESTINATION" in joined or "DEST" in joined)):

            flight_num  = _find_value_after(row, ["FLIGHT #", "FLIGHT#", "FLIGHT"])
            date        = _find_value_after(row, ["DATE"])
            destination = _find_value_after(row, ["DESTINATION", "DEST"])

            if not flight_num and not destination:
                i += 1
                continue
            if flight_num.upper() in ("ITEM", "AWB", "PCS", ""):
                i += 1
                continue

            cargo_header = all_rows[i + 1] if i + 1 < len(all_rows) else []
            awb_idx  = _find_index(cargo_header, ["AWB"])
            pcs_idx  = _find_index(cargo_header, ["PCS", "PIECES"])
            kgs_idx  = _find_index(cargo_header, ["KGS", "KG"])
            desc_idx = _find_index(cargo_header, ["DESCRIPTION", "DESC"])
            rsn_idx  = _find_index(cargo_header, ["REASON"])

            items = []
            pending_trolley = ""
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

                awb_clean = (awb or "").strip().upper()

                # إذا كان السطر عبارة عن ULD/TROLLEY فقط (مثل AKE/PMC/BT/CBT...)،
                # لا نُنشئ صف شحنة جديد؛ بل نربطه بآخر شحنة سبق إضافتها.
                if re.match(r"^(CBT|BT|AKE|PMC|PAG|ULD|AKH|RKN)\w*", awb_clean) and not any([pcs, kgs, desc, rsn]):
                    if items:
                        items[-1]["trolley"] = awb
                    else:
                        pending_trolley = awb
                    j += 1
                    continue

                if not any([awb, pcs, kgs, desc, rsn]):
                    j += 1
                    continue

                trolley = pending_trolley
                pending_trolley = ""

                items.append({
                    "awb": awb,
                    "pcs": pcs,
                    "kgs": kgs,
                    "description": desc,
                    "reason": rsn,
                    "class_": "",
                    "item": "",
                    "trolley": trolley,
                })
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


# ────────────────────────────────────────────────────────────────
#  النوع B — جدول عمودي
#  Header: ITEM | DATE | FLIGHT | STD/ETD | DEST | Email | Physical | ...
# ────────────────────────────────────────────────────────────────
def _parse_type_b(all_rows: list[list[str]]) -> list[dict]:
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
                "item":        _get(row, c_item),
                "awb":         "",
                "pcs":         _get(row, c_pcs),
                "kgs":         "",
                "description": "",
                "class_":      "",
                "reason":      _get(row, c_rsn),
                "email":       _get(row, c_email),
                "physical":    _get(row, c_phys),
                "trolley":     _get(row, c_trol),
                "cms":         _get(row, c_cms),
                "remarks":     _get(row, c_rmk),
            }
            if any(item.values()):
                current["items"].append(item)

    return flights


# ────────────────────────────────────────────────────────────────
#  النوع C — نص عادي
#  OFFLOADED CARGO ON OV237/27FEB
#  703 13436275   14   SPORTS WERAS   B   194.0   SKTDUS
#  CGO OFFLOADED DUE SPACE
# ────────────────────────────────────────────────────────────────
def _parse_type_c(soup: BeautifulSoup) -> list[dict]:
    """
    يستخرج الرحلات من النصوص الحرة في الإيميل (ليس جداول).
    العنوان: OFFLOADED CARGO ON <FLIGHT>/<DATE>
    السطر:   <AWB>   <PCS>   <DESC>   <CLASS>   <KGS>   <DEST>
    السبب:   CGO OFFLOADED DUE <REASON>
    """
    flights = []

    # استخرج كل النصوص من الصفحة
    full_text = soup.get_text("\n")
    lines     = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

    # نمط العنوان: OFFLOADED CARGO ON WY237/27FEB أو OFFLOADED CARGO ON OV237/27FEB
    title_pat  = re.compile(
        r"OFFLOAD(?:ED)?\s+CARGO\s+ON\s+([A-Z0-9]{2,6})\s*/\s*(\w+)",
        re.IGNORECASE,
    )
    # نمط سطر البيانات: رقم AWB ثم PCS ثم DESC ثم CLASS ثم KGS ثم DEST
    # مثال: 703 13436275   14   SPORTS WERAS   B   194.0   SKTDUS
    data_pat   = re.compile(
        r"^(\d[\d\s]{5,15})\s{2,}(\d+)\s{2,}(.+?)\s{2,}([A-Z])\s{2,}([\d.]+)\s{2,}([A-Z]{3,6})\s*$"
    )
    reason_pat = re.compile(r"CGO\s+OFFLOAD(?:ED)?\s+DUE\s+(.+)", re.IGNORECASE)

    i = 0
    while i < len(lines):
        m_title = title_pat.search(lines[i])
        if m_title:
            flight_num = m_title.group(1).upper()
            date       = m_title.group(2).upper()
            items      = []
            reason     = ""

            # ابحث في الأسطر التالية عن البيانات والسبب
            j = i + 1
            while j < len(lines) and j < i + 15:
                m_data = data_pat.match(lines[j])
                if m_data:
                    awb   = m_data.group(1).strip()
                    pcs   = m_data.group(2).strip()
                    desc  = m_data.group(3).strip()
                    cls_  = m_data.group(4).strip()
                    kgs   = m_data.group(5).strip()
                    dest  = m_data.group(6).strip()
                    items.append({
                        "item":        "",
                        "awb":         awb,
                        "pcs":         pcs,
                        "kgs":         kgs,
                        "description": desc,
                        "class_":      cls_,
                        "reason":      "",
                        "email":       "",
                        "physical":    "",
                        "trolley":     "",
                        "cms":         "",
                        "remarks":     "",
                    })

                m_rsn = reason_pat.search(lines[j])
                if m_rsn:
                    reason = m_rsn.group(1).strip()
                    # أضف السبب لكل الشحنات
                    for it in items:
                        if not it["reason"]:
                            it["reason"] = reason

                j += 1

            if items or flight_num:
                flights.append({
                    "flight":      flight_num,
                    "date":        date,
                    "std_etd":     "",
                    "destination": items[0]["reason"].split()[-1] if items and not items[0].get("class_") else "",
                    "format":      "C",
                    "reason":      reason,
                    "items":       items,
                })
            i = j
        else:
            i += 1

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
        filename  = slugify(
            f"{flight['flight']}_{flight.get('date','')}_{flight.get('destination','')}"
        ) + ".json"
        file_path = folder / filename
        existed   = file_path.exists()

        payload = {**flight, "saved_at": now.isoformat()}
        file_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        entry = meta["flights"].get(filename, {
            "flight":     flight["flight"],
            "date":       flight.get("date", ""),
            "dest":       flight.get("destination", ""),
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
        --empty-bg: #fffbeb;
        --empty-br: #fde68a;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
        background: var(--bg);
        font-family: 'IBM Plex Sans', Calibri, sans-serif;
        color: var(--text);
        font-size: 13px;
    }

    /* ── شريط الرأس ── */
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

.hdr-left { display:flex; flex-direction:column; gap:6px; }
.btn-back {
    display:inline-block;
    width: fit-content;
    background: rgba(255,255,255,.14);
    border: 1px solid rgba(255,255,255,.26);
    color:#fff;
    text-decoration:none;
    padding:6px 10px;
    border-radius:10px;
    font-size:12px;
    font-weight:800;
}
.btn-back:hover { background: rgba(255,255,255,.22); }
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

    .wrap { max-width: 1400px; margin: 0 auto; padding: 24px 20px 48px; }

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
        text-align: center;
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
        margin-bottom: 2px;
    }

    .val { font-weight: 700; color: var(--blue-dk); font-size: 14px; }

    /* حقل فارغ قابل للتعبئة يدوياً */
    .empty-field {
        background: var(--empty-bg);
        border: 1px dashed var(--empty-br);
        border-radius: 5px;
        padding: 3px 8px;
        font-size: 11px;
        color: #92400e;
        min-width: 70px;
        display: inline-block;
        font-style: italic;
    }

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
        padding: 5px 11px;
        font-size: 10px;
        font-weight: 700;
        color: var(--blue-hd);
        text-transform: uppercase;
        letter-spacing: .5px;
        border-bottom: 1px solid #c3d4f5;
        text-align: center;
    }

    /* ══ الصف الثاني: بيانات الشحنة ══ */
    tr.row-cargo td {
        padding: 8px 11px;
        border-bottom: 1px solid #edf0f7;
        border-right: 1px solid #edf0f7;
        vertical-align: middle;
        text-align: center;
    }
    tr.row-cargo:nth-child(odd)  td { background: var(--card); }
    tr.row-cargo:nth-child(even) td { background: var(--row-alt); }
    tr.row-cargo td:last-child { border-right: none; }
    tr.row-cargo td:first-child { text-align: left; }

    .mono { font-family: 'IBM Plex Mono', monospace; font-size: 12px; }
    .dim  { color: var(--muted); font-size: 12px; }

    .reason-tag {
        display: inline-block;
        background: #fff3e0;
        color: #b45309;
        border: 1px solid #fcd34d;
        border-radius: 5px;
        padding: 2px 9px;
        font-size: 11px;
        font-weight: 700;
        white-space: nowrap;
    }

    .class-tag {
        display: inline-block;
        background: #e0e7ff;
        color: #3730a3;
        border: 1px solid #a5b4fc;
        border-radius: 5px;
        padding: 2px 9px;
        font-size: 11px;
        font-weight: 700;
    }

    /* ══ الصف الثالث: الإجمالي ══ */
    tr.row-totals td {
        background: var(--total-bg);
        border-top: 2px solid var(--blue);
        border-bottom: 3px solid var(--blue-hd);
        padding: 8px 12px;
        font-weight: 700;
        color: var(--blue-hd);
        font-size: 12px;
        text-align: center;
    }
    .tlbl {
        font-size: 10px;
        font-weight: 400;
        color: var(--muted);
        display: block;
        margin-bottom: 1px;
    }

    /* نص الـ label داخل صف الرحلة (FLIGHT: DATE: إلخ) */
    .flt-label-inline {
        font-size: 10px;
        font-weight: 700;
        color: var(--muted);
        text-transform: uppercase;
        letter-spacing: .5px;
        margin-right: 4px;
    }

    /* ══ الصف الأول: معلومات الرحلة inline ══ */
    tr.row-flight-inline td {
        background: var(--blue-hd);
        border-top: 3px solid var(--blue);
        padding: 9px 14px;
        border-bottom: 1px solid #1a3a7a;
        color: #fff;
    }

    .il-lbl {
        font-size: 10px;
        font-weight: 700;
        color: rgba(255,255,255,.6);
        text-transform: uppercase;
        letter-spacing: .4px;
        margin-right: 3px;
    }

    .il-val {
        font-size: 13px;
        font-weight: 700;
        color: #fff;
        margin-right: 2px;
    }

/* STD/ETD times inside banner (keep time only; no repeated STD/ETD text) */
.stdtime, .etdtime { color: #fff; font-weight: 800; font-family: 'IBM Plex Mono', monospace; }
.stdsep { color: rgba(255,255,255,.35); margin: 0 8px; font-weight: 700; }

    .il-sep {
        color: rgba(255,255,255,.25);
        margin: 0 10px;
        font-size: 14px;
    }

    .footer { text-align:center; color:var(--muted); font-size:12px; margin-top:28px; }
    """


# ══════════════════════════════════════════════════════════════════
#  تصيير الرحلات — هيكل موحّد لجميع الأنواع
#
#  الصف الأول  : ITEM | FLIGHT | DATE | STD/ETD | DEST | EMAIL_TIME | PHYSICAL | حقول فارغة
#  الصف الثاني : AWB  | PCS    | KGS  | CLASS   | DESCRIPTION | REASON
#  الصف الثالث : TOTAL SHIPMENTS | TOTAL PCS | TOTAL KGS
# ══════════════════════════════════════════════════════════════════

# أعمدة الجدول الثابتة (8 أعمدة)
# [0] ITEM   [1] FLIGHT   [2] DATE   [3] STD/ETD
# [4] DEST   [5] AWB/PCS  [6] KGS   [7] REASON/CLASS

def _empty(label: str) -> str:
    """خلية فارغة بلون مميز تُشير إلى بيانات ناقصة."""
    return f'<span class="empty-field">{label}…</span>'


def _render_std_etd(value: str) -> str:
    """
    value formats:
      - ''          -> empty
      - '15:00'     -> STD only
      - '15:00|14:55' -> STD|ETD
    We do NOT repeat 'STD'/'ETD' text because the label is already shown in the banner.
    """
    v = (value or "").strip()
    if not v:
        return ""
    if "|" in v:
        std, etd = (p.strip() for p in v.split("|", 1))
        if std and etd and std != etd:
            return f'<span class="stdtime">{std}</span><span class="stdsep">·</span><span class="etdtime">{etd}</span>'
        return f'<span class="stdtime">{std or etd}</span>'
    return f'<span class="stdtime">{v}</span>'



def _render_flight(flight: dict, meta_entry: dict) -> str:
    items     = flight.get("items", [])
    fmt       = flight.get("format", "A")
    updates   = int(meta_entry.get("updates", 0))
    upd_at    = meta_entry.get("updated_at", "")[:16].replace("T", " ")
    upd_badge = f'<span class="badge-upd">UPDATED ×{updates}</span>' if updates > 0 else ""

    # (type badge removed — B means Priority, not email type)

    def safe_sum(key: str) -> int:
        total = 0
        for it in items:
            try:
                total += int(re.sub(r"[^\d]", "", it.get(key, "") or "0") or 0)
            except Exception:
                pass
        return total

    total_pcs = safe_sum("pcs")
    total_kgs = safe_sum("kgs")

    # ══ الصف الأول: معلومات الرحلة ══
    # يعرض الحقول المتوفرة + حقول فارغة للبيانات الناقصة حسب النوع

    flt  = flight.get("flight", "") or ""
    date = flight.get("date", "")   or ""
    std  = flight.get("std_etd", "") or ""
    dest = flight.get("destination", "") or ""

    # ══ صفوف الشحنات — كل بيانات الرحلة مدمجة في كل صف ══
    cargo_rows = ""

    # ── صف الرحلة مرة واحدة فقط ──
    cargo_rows += f"""
    <tr class="row-flight-inline">
        <td colspan="12">
            <span class="il-lbl">FLIGHT:</span><span class="il-val flt-num">✈ {flt or '—'}</span>
            <span class="il-sep">·</span>
            <span class="il-lbl">DATE:</span><span class="il-val">{date or '—'}</span>
            <span class="il-sep">·</span>
            <span class="il-lbl">STD/ETD:</span><span class="il-val">{_render_std_etd(std) or _empty("STD/ETD")}</span>
            <span class="il-sep">·</span>
            <span class="il-lbl">DEST:</span><span class="il-val">{dest or '—'}</span>
            <span style="float:right">{upd_badge}<span class="dim" style="font-size:11px">Last: {upd_at}</span></span>
        </td>
    </tr>"""

    # ── هيدر الشحنات بعد صف الرحلة مباشرة ──
    cargo_rows += """
    <tr class="row-subhead">
        <td>ITEM</td>
        <td colspan="2">AWB</td>
        <td>PCS</td>
        <td>KGS</td>
        <td>PRIORITY</td>
        <td colspan="3">DESCRIPTION</td>
        <td colspan="2">TROLLEY / ULD</td>
        <td>REASON</td>
    </tr>"""

    # ── صفوف الشحنات ──
    for idx, it in enumerate(items, 1):
        rsn   = it.get("reason", "")
        cls_  = it.get("class_", "")
        trol  = it.get("trolley", "")
        reason_cell  = f'<span class="reason-tag">{rsn}</span>'  if rsn  else '<span class="dim">—</span>'
        class_cell   = f'<span class="class-tag">{cls_}</span>'  if cls_ else '<span class="dim">—</span>'
        trolley_cell = trol if trol else '<span class="dim">—</span>'

        cargo_rows += f"""
        <tr class="row-cargo">
            <td class="mono">{it.get('item','') or idx}</td>
            <td colspan="2" class="mono" style="text-align:left">{it.get('awb','') or '—'}</td>
            <td class="mono"><b>{it.get('pcs','') or '—'}</b></td>
            <td class="mono">{it.get('kgs','') or '—'}</td>
            <td>{class_cell}</td>
            <td colspan="3" style="text-align:left">{it.get('description','') or '—'}</td>
            <td colspan="2" class="mono">{trolley_cell}</td>
            <td>{reason_cell}</td>
        </tr>"""

    if not items:
        cargo_rows += '<tr class="row-cargo"><td colspan="12" class="dim" style="text-align:center;padding:14px">No cargo data</td></tr>'

    # ══ الصف الثالث: الإجمالي / التعبئة ══
    row_totals = f"""
    <tr class="row-totals">
        <td colspan="6" style="text-align:left">
            <span class="tlbl">TOTAL SHIPMENTS</span>{len(items)}
        </td>
        <td><span class="tlbl">TOTAL PCS</span>{total_pcs or '—'}</td>
        <td><span class="tlbl">TOTAL KGS</span>{total_kgs or '—'}</td>
        <td colspan="3">
            {upd_badge}
            <span class="dim" style="font-size:11px">Last: {upd_at}</span>
        </td>
        <td></td>
    </tr>"""

    return cargo_rows + row_totals


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
        "shift1": "صباح — 06:00 إلى 14:30",
        "shift2": "ظهر — 14:30 إلى 21:30",
        "shift3": "ليل — 21:30 إلى 06:00",
    }
    shift_label   = shift_labels.get(shift, shift)
    total_flights = len(flights)
    total_items   = sum(len(f.get("items", [])) for f in flights)

    tbody = ""
    for flight in flights:
        fname  = slugify(
            f"{flight['flight']}_{flight.get('date','')}_{flight.get('destination','')}"
        ) + ".json"
        meta_e = meta.get("flights", {}).get(fname, {})
        tbody += _render_flight(flight, meta_e)

    if not tbody:
        tbody = '<tr><td colspan="12" style="text-align:center;padding:36px;color:#6b7a99">No offload data recorded for this shift.</td></tr>'

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
    <div class="hdr-left">
        <a class="btn-back" href="../../">⬅ العودة للرئيسية</a>
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
        <tbody>{tbody}</tbody>
    </table>
    <div class="footer">Generated automatically by GitHub Actions &nbsp;·&nbsp; {date_dir}</div>
</div>
</body>
</html>"""

    out_dir = DOCS_DIR / date_dir / shift
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "index.html").write_text(html, encoding="utf-8")



def _shift_window(date_dir: str, shift: str) -> tuple[str, str]:
    """Return (start_iso, end_iso) in local time (Asia/Muscat) for display in the index."""
    # Shift boundaries (local time)
    if shift == "shift1":
        start_hm, end_hm, add_day = "06:00", "14:30", 0
    elif shift == "shift2":
        start_hm, end_hm, add_day = "14:30", "21:30", 0
    else:
        start_hm, end_hm, add_day = "21:30", "06:00", 1  # crosses midnight

    start = f"{date_dir} {start_hm}"
    # end date might be next day for shift3
    try:
        d = datetime.fromisoformat(date_dir)
        end_date = (d.replace(tzinfo=None) if isinstance(d, datetime) else d)
    except Exception:
        end_date = None

    if end_date is not None and add_day:
        # date_dir is YYYY-MM-DD
        d = datetime.fromisoformat(date_dir)
        end = (d + timedelta(days=1)).strftime("%Y-%m-%d") + f" {end_hm}"
    else:
        end = f"{date_dir} {end_hm}"
    return start, end


def build_root_index(now: datetime) -> None:
    """Beautiful Arabic home page with accordion days; current day opened by default."""
    if not DOCS_DIR.exists():
        return

    today = now.strftime("%Y-%m-%d")

    day_dirs = sorted(
        (p for p in DOCS_DIR.iterdir()
         if p.is_dir() and re.match(r"\d{4}-\d{2}-\d{2}", p.name)),
        reverse=True,
    )

    shift_meta = {
        "shift1": {"ar": "صباح", "time": "06:00–14:30"},
        "shift2": {"ar": "ظهر",  "time": "14:30–21:30"},
        "shift3": {"ar": "ليل",  "time": "21:30–06:00"},
    }

    days_html = ""
    for day_dir in day_dirs:
        day = day_dir.name
        open_attr = " open" if day == today else ""

        rows = ""
        for shift in ("shift1", "shift2", "shift3"):
            if not (day_dir / shift / "index.html").exists():
                continue

            start_dt, end_dt = _shift_window(day, shift)
            meta = shift_meta.get(shift, {"ar": shift, "time": ""})
            rows += f"""
            <a class="shift-card" href="{day}/{shift}/">
                <div class="shift-left">
                    <div class="shift-title">{meta['ar']} <span class="shift-time">({meta['time']})</span></div>
                    <div class="shift-sub">
                        <span class="pill">يبدأ</span> <span class="mono">{start_dt}</span>
                        <span class="dot">•</span>
                        <span class="pill">ينتهي</span> <span class="mono">{end_dt}</span>
                    </div>
                </div>
                <div class="shift-right">
                    <span class="chev">›</span>
                </div>
            </a>
            """

        if not rows:
            rows = '<div class="empty">لا توجد تقارير لهذا اليوم.</div>'

        days_html += f"""
        <details class="day-accordion"{open_attr}>
            <summary>
                <span class="day-title">📅 {day}</span>
                <span class="day-hint">{'اليوم' if day == today else 'اضغط للعرض'}</span>
            </summary>
            <div class="day-body">
                {rows}
            </div>
        </details>
        """

    if not days_html:
        days_html = "<div class='empty'>لا توجد تقارير بعد.</div>"

    html = f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Offload Monitor</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');
        :root {{
            --bg:#f0f4f8; --card:#fff; --line:#dde3ec;
            --blue:#1a56db; --blue2:#0d2d6e; --muted:#6b7a99; --shadow:0 6px 22px rgba(15,31,61,.10);
        }}
        * {{ box-sizing:border-box; }}
        body {{ margin:0; background:var(--bg); font-family:'IBM Plex Sans',Calibri,sans-serif; color:#0f1f3d; }}
        .top {{
            background: linear-gradient(135deg, var(--blue2) 0%, #1240a8 100%);
            color:#fff; padding:22px 26px; box-shadow:0 2px 12px rgba(13,45,110,.25);
        }}
        .top h1 {{ margin:0; font-size:22px; }}
        .top p {{ margin:6px 0 0; opacity:.8; font-size:13px; }}
        .wrap {{ max-width:980px; margin:18px auto 60px; padding:0 16px; }}
        .day-accordion {{
            background:var(--card); border:1px solid var(--line); border-radius:14px;
            box-shadow:var(--shadow); margin:14px 0; overflow:hidden;
        }}
        summary {{
            list-style:none; cursor:pointer; padding:14px 16px; display:flex; align-items:center; justify-content:space-between;
            background: #f7f9fc;
            border-bottom:1px solid #edf0f7;
        }}
        summary::-webkit-details-marker {{ display:none; }}
        .day-title {{ font-weight:800; color:var(--blue2); }}
        .day-hint {{ color:rgba(13,45,110,.65); font-size:12px; font-weight:700; }}
        .day-body {{ padding:14px 14px 6px; display:grid; gap:10px; }}
        .shift-card {{
            display:flex; justify-content:space-between; align-items:center; gap:14px;
            text-decoration:none; color:inherit;
            border:1px solid #edf0f7; border-radius:12px; padding:12px 14px;
            transition:transform .08s ease, border-color .12s ease, background .12s ease;
            background:#fff;
        }}
        .shift-card:hover {{ transform: translateY(-1px); border-color:#c9dbff; background:#f6f9ff; }}
        .shift-title {{ font-weight:800; color:#0d2d6e; }}
        .shift-time {{ font-weight:700; color:var(--muted); font-size:12px; }}
        .shift-sub {{ margin-top:6px; font-size:12px; color:var(--muted); display:flex; flex-wrap:wrap; gap:8px; align-items:center; }}
        .pill {{ background:#e8f0fe; color:#0d2d6e; border:1px solid #cfe0ff; border-radius:999px; padding:2px 8px; font-weight:800; font-size:11px; }}
        .mono {{ font-family:'IBM Plex Mono', monospace; }}
        .dot {{ opacity:.5; }}
        .chev {{ font-size:24px; color:#1a56db; font-weight:900; }}
        .empty {{ padding:18px; color:var(--muted); text-align:center; }}
        .footer {{ text-align:center; color:var(--muted); font-size:12px; margin-top:18px; }}
    </style>
</head>
<body>
    <div class="top">
        <h1>✈ Offload Monitor</h1>
        <p>اختر المناوبة لعرض تفاصيل الـ Offload. الأيام السابقة مطوية تلقائيًا.</p>
    </div>
    <div class="wrap">
        {days_html}
        <div class="footer">Generated automatically by GitHub Actions · {today}</div>
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

    # تشخيص سريع (يساعدك تتأكد أن OneDrive يرسل النسخة الجديدة)
    print(f"HTML length: {len(html)}")
    print(f"HTML sha256: {new_hash[:16]}")

    if STATE_FILE.exists():
        old_hash = STATE_FILE.read_text(encoding="utf-8").strip()
        if old_hash == new_hash and not FORCE_REBUILD:
            print("No change detected. Exiting. (Set FORCE_REBUILD=1 to force rebuild)")
            return
        if old_hash == new_hash and FORCE_REBUILD:
            print("No change detected, but FORCE_REBUILD=1 → continuing to rebuild.")

    print("Change detected. Parsing…")
    flights = extract_flights(html)

    if not flights:
        print("WARNING: No flights extracted. Check HTML structure.")
        STATE_FILE.write_text(new_hash, encoding="utf-8")
        return


    # ── Enrich flights with STD/ETD from an online source (AirLabs) if configured ──
    # Set AIRLABS_API_KEY in GitHub Secrets to enable.
    if os.environ.get("AIRLABS_API_KEY", "").strip():
        enriched = 0
        for f in flights:
            flt = (f.get("flight") or "").strip()
            if not flt:
                continue
            std_etd = fetch_std_etd_airlabs(flt, flight_date=normalize_flight_date(f.get('date',''), now))
            if std_etd:
                f["std_etd"] = std_etd
                enriched += 1
        print(f"STD/ETD enriched for {enriched} flight(s) via AirLabs.")
    else:
        print("AIRLABS_API_KEY not set — STD/ETD will remain blank.")

    print(f"Extracted {len(flights)} flight(s). Saving…")
    date_dir, shift, _ = save_flights(flights, now)

    print(f"Building report: {date_dir}/{shift}…")
    build_shift_report(date_dir, shift)
    build_root_index(now)

    STATE_FILE.write_text(new_hash, encoding="utf-8")
    print(f"Done. ✓  ({len(flights)} flights saved)")


if __name__ == "__main__":
    main()

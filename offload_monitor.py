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

# روابط الروستر (بدون نسخ أي ملف داخل صفحة الروستر)
ROSTER_PAGE_URL: str = "https://khalidsaif912.github.io/roster-site/"
ROSTER_JSON_RAW: str = "https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/data/roster.json"
ROSTER_JSON_GH:  str = "https://github.com/khalidsaif912/roster-site/blob/main/docs/data/roster.json"


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



def _airlabs_best_row(rows: list, flight_iata: str, flight_date: str | None) -> dict | None:
    """Pick the best matching row from AirLabs API response.

    Priority:
      1. Exact flight_iata match + date found in dep_time field
      2. Exact flight_iata match (first candidate)
      3. Any row (fallback)
    """
    if not rows:
        return None

    # Filter to exact flight matches
    candidates = [
        row for row in rows
        if isinstance(row, dict)
        and (row.get("flight_iata") or "").strip().upper() == flight_iata
    ]
    if not candidates:
        candidates = [row for row in rows if isinstance(row, dict)]

    if not candidates:
        return None

    # If we have a date, prefer the row whose dep_time contains that date
    if flight_date:
        for row in candidates:
            dep_time = str(
                row.get("dep_time")
                or row.get("dep_time_utc")
                or row.get("dep_scheduled")
                or ""
            )
            if flight_date in dep_time:
                return row

    # If only one candidate, return it; otherwise warn about ambiguity
    if len(candidates) == 1:
        return candidates[0]

    # Multiple candidates and no date match — log and return first
    print(
        f"  [AirLabs] WARNING: {len(candidates)} rows for {flight_iata}"
        f" and none matched date={flight_date!r}. Using first row."
        f" arr_iata values: {[c.get('arr_iata') for c in candidates[:5]]}"
    )
    return candidates[0]


def fetch_flight_info_airlabs(
    flight_iata: str,
    *,
    flight_date: str | None = None,
    dep_iata: str | None = None,
    arr_iata: str | None = None,
) -> dict | None:
    """Fetch flight info (STD/ETD + DEST) from AirLabs if AIRLABS_API_KEY is set.

    Strategy:
      1. Try /flights endpoint (live/real-time data) — gives actual dep per flight number+date.
      2. Fall back to /schedules if /flights returns nothing.

    Returns a dict like:
      {"std":"HH:MM", "etd":"HH:MM", "dest":"ADD"}
    Times are returned as *time only* (HH:MM), because the report already shows the date.
    """
    api_key = os.environ.get("AIRLABS_API_KEY", "").strip()
    if not api_key:
        return None

    flight_iata = (flight_iata or "").strip().upper()
    if not flight_iata:
        return None

    def _fetch(endpoint: str, extra_params: dict) -> list:
        base_params: dict[str, str] = {"api_key": api_key, "flight_iata": flight_iata}
        base_params.update(extra_params)
        try:
            resp = requests.get(
                f"https://airlabs.co/api/v9/{endpoint}",
                params=base_params,
                timeout=30,
            )
            resp.raise_for_status()
            payload = resp.json()
            rows = payload.get("response") or payload.get("data") or []
            return rows if isinstance(rows, list) else []
        except Exception as exc:
            print(f"  [AirLabs] {endpoint} error: {exc}")
            return []

    # ── 1) Try /flights (real-time) ───────────────────────────────
    extra: dict[str, str] = {}
    if dep_iata:
        extra["dep_iata"] = dep_iata.strip().upper()
    if arr_iata:
        extra["arr_iata"] = arr_iata.strip().upper()

    rows = _fetch("flights", extra)
    best = _airlabs_best_row(rows, flight_iata, flight_date)

    # ── 2) Fall back to /schedules ────────────────────────────────
    if not best:
        sched_extra: dict[str, str] = {}
        if flight_date:
            sched_extra["flight_date"] = flight_date
        if dep_iata:
            sched_extra["dep_iata"] = dep_iata.strip().upper()
        if arr_iata:
            sched_extra["arr_iata"] = arr_iata.strip().upper()
        rows = _fetch("schedules", sched_extra)
        best = _airlabs_best_row(rows, flight_iata, flight_date)

    if not best:
        return None

    std  = _time_only(str(best.get("dep_time")      or best.get("dep_time_utc")      or best.get("dep_scheduled") or ""))
    etd  = _time_only(str(best.get("dep_estimated")  or best.get("dep_estimated_utc") or ""))
    dest = str(best.get("arr_iata") or "").strip().upper()

    print(f"  [AirLabs] {flight_iata} → dest={dest!r}, std={std!r}, etd={etd!r}")
    return {"std": std, "etd": etd, "dest": dest}


def fetch_std_etd_airlabs(
    flight_iata: str,
    *,
    flight_date: str | None = None,
    dep_iata: str | None = None,
    arr_iata: str | None = None,
) -> str:
    """Backward-compatible helper: returns 'STD' and 'ETD' as:
      - 'HH:MM' (STD only)
      - 'HH:MM|HH:MM' (STD|ETD)
    """
    info = fetch_flight_info_airlabs(
        flight_iata,
        flight_date=flight_date,
        dep_iata=dep_iata,
        arr_iata=arr_iata,
    )
    if not info:
        return ""
    std = info.get("std") or ""
    etd = info.get("etd") or ""
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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

    :root {
        --bg:        #f1f5fb;
        --card:      #ffffff;
        --border:    #e2e8f4;
        --blue:      #2563eb;
        --blue-dk:   #1d4ed8;
        --blue-hd:   #0f2660;
        --navy:      #0a1f52;
        --sky:       #e8f0fe;
        --text:      #0f1f3d;
        --muted:     #64748b;
        --amber:     #f59e0b;
        --amber-bg:  #fffbeb;
        --amber-br:  #fde68a;
        --row-alt:   #f8fafd;
        --total-bg:  #eff6ff;
        --green:     #059669;
        --green-bg:  #ecfdf5;
        --red:       #dc2626;
        --shadow-sm: 0 1px 4px rgba(15,31,61,.07);
        --shadow-md: 0 4px 20px rgba(15,31,61,.10);
        --shadow-lg: 0 8px 32px rgba(15,31,61,.13);
        --radius:    14px;
    }

    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
        background: var(--bg);
        font-family: 'Inter', system-ui, sans-serif;
        color: var(--text);
        font-size: 13px;
        line-height: 1.5;
        -webkit-font-smoothing: antialiased;
    }

    /* ─── TOP HEADER ─── */
    .page-header {
        background: linear-gradient(135deg, var(--navy) 0%, #1a3a8f 55%, #2251c9 100%);
        color: #fff;
        padding: 16px 20px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        flex-wrap: wrap;
        box-shadow: 0 2px 16px rgba(10,31,82,.35);
        position: sticky;
        top: 0;
        z-index: 100;
    }

    .hdr-left { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }

    .btn-back {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        background: rgba(255,255,255,.12);
        border: 1px solid rgba(255,255,255,.22);
        color: #fff;
        text-decoration: none;
        padding: 7px 12px;
        border-radius: 9px;
        font-size: 12px;
        font-weight: 700;
        white-space: nowrap;
        transition: background .15s;
    }
    .btn-back:hover { background: rgba(255,255,255,.22); }

    .btn-link {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(255,255,255,.10);
        border: 1px solid rgba(255,255,255,.18);
        color: #fff;
        text-decoration: none;
        padding: 7px 12px;
        border-radius: 9px;
        font-size: 12px;
        font-weight: 700;
        white-space: nowrap;
        transition: background .15s;
    }
    .btn-link:hover { background: rgba(255,255,255,.18); }

    .share-box {
        margin: 14px 0 0;
        background: rgba(255,255,255,.08);
        border: 1px solid rgba(255,255,255,.14);
        border-radius: 14px;
        padding: 12px 14px;
        display: flex;
        justify-content: space-between;
        gap: 10px;
        flex-wrap: wrap;
        align-items: center;
    }
    .share-title { font-size: 12px; font-weight: 800; opacity: .95; }
    .share-sub { font-size: 11px; opacity: .7; margin-top: 2px; }
    .share-actions { display: flex; gap: 8px; flex-wrap: wrap; }

    .page-header h1   { font-size: 17px; font-weight: 800; letter-spacing: -.2px; }
    .page-header .sub { font-size: 11.5px; opacity: .7; margin-top: 2px; }

    .hdr-stats { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }

    .stat-box {
        background: rgba(255,255,255,.12);
        border: 1px solid rgba(255,255,255,.2);
        border-radius: 10px;
        padding: 6px 16px;
        text-align: center;
        font-size: 11px;
        font-weight: 600;
        color: rgba(255,255,255,.85);
        white-space: nowrap;
    }
    .stat-box strong { display: block; font-size: 20px; font-weight: 800; color: #fff; line-height: 1.2; }

    /* ─── MAIN WRAP ─── */
    .wrap {
        max-width: 1300px;
        margin: 0 auto;
        padding: 20px 14px 40px;
        display: flex;
        flex-direction: column;
        gap: 18px;
    }

    /* ─── FLIGHT CARD ─── */
    .flight-card {
        background: var(--card);
        border-radius: var(--radius);
        border: 1px solid var(--border);
        box-shadow: var(--shadow-md);
        overflow: hidden;
    }

    /* ─── FLIGHT BANNER ─── */
    .flt-banner {
        background: linear-gradient(90deg, var(--navy) 0%, #1a3d8f 100%);
        color: #fff;
        padding: 11px 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        flex-wrap: wrap;
    }

    .flt-banner-left {
        display: flex;
        align-items: center;
        gap: 14px;
        flex-wrap: wrap;
    }

    .flt-number {
        font-family: 'JetBrains Mono', monospace;
        font-size: 16px;
        font-weight: 700;
        letter-spacing: .3px;
        white-space: nowrap;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .flt-number::before { content: "✈"; font-size: 13px; opacity: .8; }

    .flt-chip {
        background: rgba(255,255,255,.13);
        border: 1px solid rgba(255,255,255,.2);
        border-radius: 7px;
        padding: 4px 10px;
        font-size: 11.5px;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 5px;
        white-space: nowrap;
    }
    .flt-chip .lbl {
        color: rgba(255,255,255,.55);
        font-size: 10px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: .5px;
    }

    .dest-chip {
        background: rgba(37,99,235,.35);
        border-color: rgba(96,165,250,.4);
        font-size: 13px;
        font-weight: 800;
        letter-spacing: .5px;
    }

    .flt-banner-right {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
    }

    .upd-flash {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        background: rgba(239,68,68,.15);
        color: #fca5a5;
        border: 1px solid rgba(239,68,68,.35);
        font-size: 10px;
        font-weight: 800;
        padding: 3px 9px;
        border-radius: 999px;
        letter-spacing: .5px;
        animation: updBlink 2.4s ease-in-out infinite;
    }
    @keyframes updBlink {
        0%,25%  { opacity:0; }
        35%,70% { opacity:1; }
        80%,100%{ opacity:0; }
    }

    .last-upd {
        font-size: 10px;
        color: rgba(255,255,255,.45);
        white-space: nowrap;
    }

    .stdtime { font-family: 'JetBrains Mono', monospace; font-weight: 700; color: #fff; }
    .etdsep  { color: rgba(255,255,255,.4); margin: 0 2px; }

    .empty-field {
        background: rgba(245,158,11,.15);
        border: 1px dashed rgba(245,158,11,.5);
        border-radius: 5px;
        padding: 2px 7px;
        font-size: 10.5px;
        color: #fcd34d;
        font-style: italic;
    }

    /* ─── CARGO TABLE ─── */
    .cargo-table {
        width: 100%;
        border-collapse: collapse;
    }

    .cargo-table thead tr {
        background: #f0f4fb;
        border-bottom: 2px solid var(--border);
    }
    .cargo-table thead th {
        padding: 7px 10px;
        font-size: 10px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: .6px;
        color: var(--muted);
        text-align: center;
        white-space: nowrap;
        border-right: 1px solid var(--border);
    }
    .cargo-table thead th:last-child { border-right: none; }
    .cargo-table thead th:first-child { text-align: left; }

    .cargo-table tbody tr.row-cargo td {
        padding: 8px 10px;
        border-bottom: 1px solid #eef1f8;
        border-right: 1px solid #eef1f8;
        vertical-align: middle;
        text-align: center;
    }
    .cargo-table tbody tr.row-cargo:last-child td { border-bottom: none; }
    .cargo-table tbody tr.row-cargo td:last-child  { border-right: none; }
    .cargo-table tbody tr.row-cargo td:first-child { text-align: left; padding-left: 14px; }
    .cargo-table tbody tr.row-cargo:nth-child(even) td { background: var(--row-alt); }

    /* ─── TOTALS ROW ─── */
    .row-totals td {
        background: var(--total-bg);
        border-top: 2px solid #bfdbfe;
        padding: 9px 14px;
        font-weight: 700;
        color: var(--blue-dk);
        font-size: 12px;
    }
    .tot-stat {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        margin-right: 18px;
    }
    .tot-lbl { color: var(--muted); font-weight: 600; font-size: 11px; }
    .tot-val { font-weight: 800; color: var(--blue-hd); }

    /* ─── TAGS ─── */
    .mono { font-family: 'JetBrains Mono', monospace; font-size: 12px; }
    .dim  { color: var(--muted); }

    .reason-tag {
        display: inline-block;
        background: #fff7ed;
        color: #c2410c;
        border: 1px solid #fed7aa;
        border-radius: 6px;
        padding: 2px 8px;
        font-size: 11px;
        font-weight: 700;
        white-space: nowrap;
    }

    .class-tag {
        display: inline-block;
        background: #ede9fe;
        color: #5b21b6;
        border: 1px solid #c4b5fd;
        border-radius: 6px;
        padding: 2px 8px;
        font-size: 11px;
        font-weight: 700;
    }

    /* ─── FOOTER ─── */
    .footer {
        text-align: center;
        color: var(--muted);
        font-size: 11.5px;
        padding-top: 4px;
    }

    /* ─── RESPONSIVE ─── */
    @media (max-width: 700px) {
        .page-header { padding: 12px 14px; }
        .page-header h1 { font-size: 15px; }
        .wrap { padding: 12px 10px 28px; gap: 14px; }
        .flt-banner { padding: 10px 12px; }
        .flt-number { font-size: 14px; }
        .cargo-table thead { display: none; }
        .cargo-table tbody tr.row-cargo {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 0;
            border-bottom: 2px solid var(--border);
        }
        .cargo-table tbody tr.row-cargo td {
            text-align: left !important;
            padding: 5px 10px;
            border-right: none;
        }
        .cargo-table tbody tr.row-cargo td::before {
            content: attr(data-label) " ";
            font-size: 9px;
            font-weight: 700;
            color: var(--muted);
            text-transform: uppercase;
            display: block;
        }
        .row-totals td { padding: 8px 12px; }
        .hdr-stats { display: none; }
    }
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


def _format_std_etd(raw: str) -> tuple[str, str]:
    """Normalize STD/ETD into pure HH:MM (or H:MM) strings.
    Accepts inputs like:
      - '14:25·14:18'
      - '14:25 · 14:18'
      - 'STD 14:25 · ETD 14:18'
      - ISO timestamps (we extract the time part)
    Returns (std, etd) where either may be ''.
    """
    s = (raw or "").strip()
    if not s:
        return "", ""

    # Grab time tokens HH:MM (or H:MM)
    times = re.findall(r"\b(\d{1,2}:\d{2})\b", s)
    if len(times) >= 2:
        return times[0], times[1]
    if len(times) == 1:
        return times[0], ""

    # Fallback: handle 3-4 digit times like 1425
    nums = re.findall(r"\b(\d{3,4})\b", s)
    def to_hhmm(n: str) -> str:
        n = n.zfill(4)
        return f"{n[:2]}:{n[2:]}"
    if len(nums) >= 2:
        return to_hhmm(nums[0]), to_hhmm(nums[1])
    if len(nums) == 1:
        return to_hhmm(nums[0]), ""
    return "", ""


def _std_etd_display(raw: str) -> str:
    """Return display string after 'STD:' label.
    - If both STD/ETD exist => 'HH:MM/ETD:HH:MM'
    - If only STD => 'HH:MM'
    """
    std, etd = _format_std_etd(raw)
    if std and etd:
        return f"{std}/ETD:{etd}"
    if std:
        return std
    return ""



def _render_std_etd(value: str) -> str:
    """Render STD/ETD in banner without dates.

    Banner already shows label 'STD:' so this function returns:
      - ''                          (no data)
      - '<span class="stdtime">14:25</span>'                                  (STD only)
      - '<span class="stdtime">14:25</span><span class="stdslash">/</span><span class="stdtime">ETD:14:18</span>'  (STD + ETD)

    Accepts raw values from:
      - Type B column (e.g. '14:25·14:18')
      - AirLabs enrich (e.g. 'STD 14:25 · ETD 14:18')
      - Legacy 'STD 14:25 · ETD 14:18' / ISO timestamps, etc.
    """
    v = (value or "").strip()
    if not v:
        return ""

    std, etd = _format_std_etd(v)

    if std and etd and std != etd:
        return (
            f'<span class="stdtime">{std}</span>'
            f'<span class="stdslash">/</span>'
            f'<span class="stdtime">ETD:{etd}</span>'
        )
    if std:
        return f'<span class="stdtime">{std}</span>'
    if etd:
        return f'<span class="stdtime">ETD:{etd}</span>'
    return ""



def _render_flight(flight: dict, meta_entry: dict) -> str:
    """Render a single flight card matching the Export Warehouse Activity Report style."""
    items   = flight.get("items", [])
    updates = int(meta_entry.get("updates", 0))
    upd_at  = meta_entry.get("updated_at", "")[:16].replace("T", " ")

    def safe_sum(key: str) -> float:
        total = 0.0
        for it in items:
            try:
                val = re.sub(r"[^\d.]", "", it.get(key, "") or "0")
                total += float(val or 0)
            except Exception:
                pass
        return total

    real_shipments    = sum(1 for it in items if (it.get("awb") or "").strip())
    total_pcs         = int(safe_sum("pcs"))
    total_kgs         = safe_sum("kgs")
    total_kgs_display = f"{total_kgs:g}"

    flt  = flight.get("flight", "") or "—"
    date = flight.get("date", "")   or "—"
    dest = flight.get("destination", "") or "—"

    # STD / ETD
    std_raw        = flight.get("std_etd", "") or ""
    std_val, etd_val = _format_std_etd(std_raw)
    std_display    = std_val or "—"
    etd_display    = etd_val or "—"

    # ops-row fields
    email    = (flight.get("email_time")   or "").strip() or "Pending"
    physical = (flight.get("physical")     or "").strip() or "Pending"
    cms      = (flight.get("cms")          or "").strip() or "Pending"
    verified = (flight.get("trolley")      or "").strip() or "Pending"
    remarks  = (flight.get("remarks")      or "").strip() or "—"

    def _status_span(val: str) -> str:
        if val and val not in ("Pending", "—"):
            return f'<span style="font-weight:700; color:#1a7a3c;">{val}</span>'
        if val == "—":
            return f'<span style="font-weight:700; color:#1b1f2a;">—</span>'
        return f'<span style="font-weight:700; color:#b26a00;">Pending</span>'

    upd_label = (
        f'Updated: <span style="color:#ffffff; font-weight:700;">{upd_at}</span>'
        if upd_at else ""
    )
    upd_flash = (
        '<span style="display:inline-block;background:rgba(239,68,68,.2);color:#fca5a5;'
        'border:1px solid rgba(239,68,68,.4);font-size:10px;font-weight:800;padding:2px 8px;'
        'border-radius:999px;margin-right:8px;">⟳ UPDATE</span>'
        if updates > 0 else ""
    )

    # Cargo rows
    cargo_rows = ""
    for idx, it in enumerate(items, 1):
        rsn  = it.get("reason", "")
        cls_ = it.get("class_", "")
        trol = it.get("trolley", "")
        bg   = "#ffffff" if idx % 2 == 1 else "#f8faff"

        reason_cell = (
            f'<span style="display:inline-block;padding:2px 6px;border-radius:6px;'
            f'background-color:#fff7ed;border:1px solid #fed7aa;color:#c2410c;'
            f'font-size:10.5px;font-weight:700;white-space:nowrap;">{rsn}</span>'
            if rsn else "—"
        )
        class_cell = (
            f'<span style="display:inline-block;padding:2px 6px;border-radius:6px;'
            f'background-color:#ede9fe;border:1px solid #c4b5fd;color:#5b21b6;'
            f'font-size:10.5px;font-weight:700;">{cls_}</span>'
            if cls_ else "—"
        )

        cargo_rows += f"""
            <tr data-offload-item="{idx}" style="background-color:{bg};">
              <td style="padding:7px 6px; border-bottom:1px solid #e4e9f5;">{it.get('item','') or idx}</td>
              <td style="padding:7px 6px; border-bottom:1px solid #e4e9f5; font-family:Calibri,Arial,sans-serif;">{it.get('awb','') or '—'}</td>
              <td style="padding:7px 6px; border-bottom:1px solid #e4e9f5; font-weight:700;">{it.get('pcs','') or '—'}</td>
              <td style="padding:7px 6px; border-bottom:1px solid #e4e9f5;">{it.get('kgs','') or '—'}</td>
              <td style="padding:7px 6px; border-bottom:1px solid #e4e9f5;">{class_cell}</td>
              <td style="padding:7px 6px; border-bottom:1px solid #e4e9f5;">{it.get('description','') or '—'}</td>
              <td style="padding:7px 6px; border-bottom:1px solid #e4e9f5; font-family:Calibri,Arial,sans-serif;">{trol or '—'}</td>
              <td style="padding:7px 6px; border-bottom:1px solid #e4e9f5;">{reason_cell}</td>
            </tr>"""

    if not cargo_rows:
        cargo_rows = (
            '<tr><td colspan="8" style="padding:14px 6px;color:#64748b;text-align:center;">'
            'No cargo data</td></tr>'
        )

    # Totals row
    totals_row = f"""
            <tr style="background-color:#eef3fc;">
              <td colspan="8" style="padding:7px 10px; border-top:2px solid #0b3a78; font-family:Calibri,Arial,sans-serif; font-size:11.5px; font-weight:700; color:#0b3a78;">
                Total Shipments: <span style="font-size:13px;">{real_shipments}</span>
                &nbsp;&nbsp;|&nbsp;&nbsp;
                Total PCS: <span style="font-size:13px;">{total_pcs or '—'}</span>
                &nbsp;&nbsp;|&nbsp;&nbsp;
                Total KGS: <span style="font-size:13px;">{total_kgs_display or '—'}</span>
              </td>
            </tr>"""

    return f"""
<!-- ===== Flight Card: {flt} ===== -->
<table width="100%" cellpadding="0" cellspacing="0" border="0"
  style="border-collapse:collapse; margin-bottom:14px;" data-offload-card="1" data-offload-flight="{flt}">

  <!-- Header row -->
  <tr>
    <td style="background-color:#0b3a78; padding:8px 10px; border:1px solid #0a3166;
               font-family:Calibri,Arial,sans-serif; color:#ffffff; font-size:12px;
               line-height:1.2; white-space:nowrap;">
      <span style="font-weight:700;">✈ {flt}</span>
      <span style="color:#a8c4f0;">&nbsp;|&nbsp;</span>
      Date: <span style="font-weight:700;">{date}</span>
      <span style="color:#a8c4f0;">&nbsp;|&nbsp;</span>
      STD: <span style="font-weight:700;">{std_display}</span>
      <span style="color:#a8c4f0;">&nbsp;|&nbsp;</span>
      ETD: <span style="font-weight:700;">{etd_display}</span>
      <span style="color:#a8c4f0;">&nbsp;|&nbsp;</span>
      DEST: <span style="font-weight:700;">{dest}</span>
    </td>
    <td align="right" style="background-color:#0b3a78; padding:8px 10px; border:1px solid #0a3166;
               font-family:Calibri,Arial,sans-serif; color:#a8c4f0; font-size:11px; white-space:nowrap;">
      {upd_flash}{upd_label}
    </td>
  </tr>

  <!-- Ops status row -->
  <tr>
    <td colspan="2" style="background-color:#f8faff; padding:6px 10px;
        border-left:1px solid #d0d9ee; border-right:1px solid #d0d9ee; border-bottom:1px solid #d0d9ee;
        font-family:Calibri,Arial,sans-serif; font-size:10.5px; color:#1b1f2a;
        white-space:nowrap; overflow:hidden;">
      Email: {_status_span(email)}
      <span style="color:#8aa2c7;">&nbsp;&nbsp;|&nbsp;&nbsp;</span>
      Ramp Received: {_status_span(physical)}
      <span style="color:#8aa2c7;">&nbsp;&nbsp;|&nbsp;&nbsp;</span>
      CMS Completed: {_status_span(cms)}
      <span style="color:#8aa2c7;">&nbsp;&nbsp;|&nbsp;&nbsp;</span>
      Pieces Verified: {_status_span(verified)}
      <span style="color:#8aa2c7;">&nbsp;&nbsp;|&nbsp;&nbsp;</span>
      Remarks: <span style="font-weight:700; color:#1b1f2a;">{remarks}</span>
    </td>
  </tr>

  <!-- Shipments table -->
  <tr>
    <td colspan="2" style="padding:0; border-left:1px solid #d0d9ee; border-right:1px solid #d0d9ee;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0"
             style="border-collapse:collapse; font-family:Calibri,Arial,sans-serif; font-size:11.2px;">
        <tr style="background-color:#eef3fc;">
          <td style="padding:6px 6px; border-bottom:1px solid #d0d9ee; color:#0b3a78; font-weight:700;">Item</td>
          <td style="padding:6px 6px; border-bottom:1px solid #d0d9ee; color:#0b3a78; font-weight:700;">AWB</td>
          <td style="padding:6px 6px; border-bottom:1px solid #d0d9ee; color:#0b3a78; font-weight:700;">PCS</td>
          <td style="padding:6px 6px; border-bottom:1px solid #d0d9ee; color:#0b3a78; font-weight:700;">KGS</td>
          <td style="padding:6px 6px; border-bottom:1px solid #d0d9ee; color:#0b3a78; font-weight:700;">Priority</td>
          <td style="padding:6px 6px; border-bottom:1px solid #d0d9ee; color:#0b3a78; font-weight:700;">Description</td>
          <td style="padding:6px 6px; border-bottom:1px solid #d0d9ee; color:#0b3a78; font-weight:700;">ULD</td>
          <td style="padding:6px 6px; border-bottom:1px solid #d0d9ee; color:#0b3a78; font-weight:700;">Offloading Reason</td>
        </tr>
        {cargo_rows}
        {totals_row}
      </table>
    </td>
  </tr>

  <!-- Bottom border -->
  <tr>
    <td colspan="2" style="border-left:1px solid #d0d9ee; border-right:1px solid #d0d9ee;
        border-bottom:1px solid #d0d9ee; font-size:1px; line-height:1px;">&nbsp;</td>
  </tr>
</table>
<!-- ===== /Flight Card: {flt} ===== -->"""


# ══════════════════════════════════════════════════════════════════
#  الروستر — جلب الموظفين وعرضهم
# ══════════════════════════════════════════════════════════════════

# Mapping: shift key → label used in roster.json cards_html
_SHIFT_TO_ROSTER_LABEL = {
    "shift1": "Morning",
    "shift2": "Afternoon",
    "shift3": "Night",
}

# Leave/off labels to separate from on-duty
_LEAVE_LABELS = {"Annual Leave", "Sick Leave", "Emergency Leave", "Off Day", "Training"}


def fetch_roster_staff(date_dir: str, shift: str) -> dict:
    """Fetch roster.json and return staff on duty for the given date/shift.

    Returns:
        {
          "on_duty":  [{"name": str, "dept": str, "sn": str}, ...],
          "on_leave": [{"name": str, "dept": str, "status": str}, ...],
        }
    """
    try:
        r = requests.get(ROSTER_JSON_RAW, timeout=15)
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"  [roster] Failed to fetch: {e}")
        return {"on_duty": [], "on_leave": []}

    days = data.get("days", {})
    day_data = days.get(date_dir)
    if not day_data:
        print(f"  [roster] No data for date {date_dir}")
        return {"on_duty": [], "on_leave": []}

    cards_html = day_data.get("cards_html", "")
    if not cards_html:
        return {"on_duty": [], "on_leave": []}

    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(cards_html, "html.parser")
    except Exception as e:
        print(f"  [roster] BeautifulSoup error: {e}")
        return {"on_duty": [], "on_leave": []}

    target_shift = _SHIFT_TO_ROSTER_LABEL.get(shift, "")
    on_duty:  list[dict] = []
    on_leave: list[dict] = []

    for dept_card in soup.find_all(class_="deptCard"):
        dept_title_el = dept_card.find(class_="deptTitle")
        dept = dept_title_el.text.strip() if dept_title_el else "Unknown"

        for shift_card in dept_card.find_all(class_="shiftCard"):
            shift_label_el = shift_card.find(class_="shiftLabel")
            shift_label = shift_label_el.text.strip() if shift_label_el else ""

            for emp_row in shift_card.find_all(class_="empRow"):
                name_el = emp_row.find(class_="empName")
                if not name_el:
                    continue
                raw = name_el.text.strip()

                # Separate name and SN: "Saleh Al Rashdi - 82546"
                sn = ""
                name = raw
                m = re.match(r"^(.+?)\s*[-–]\s*(\d+)\s*$", raw)
                if m:
                    name = m.group(1).strip()
                    sn   = m.group(2).strip()

                if shift_label in _LEAVE_LABELS:
                    on_leave.append({"name": name, "sn": sn, "dept": dept, "status": shift_label})
                elif shift_label == target_shift:
                    on_duty.append({"name": name, "sn": sn, "dept": dept})

    print(f"  [roster] {date_dir}/{shift}: {len(on_duty)} on duty, {len(on_leave)} on leave")
    return {"on_duty": on_duty, "on_leave": on_leave}


def _render_manpower(roster: dict) -> str:
    """Render Section 6 — MANPOWER in the same style as the report template."""
    on_duty  = roster.get("on_duty",  [])
    on_leave = roster.get("on_leave", [])

    if not on_duty and not on_leave:
        return ""

    # Group on-duty by department
    from collections import OrderedDict
    by_dept: dict[str, list] = OrderedDict()
    for emp in on_duty:
        by_dept.setdefault(emp["dept"], []).append(emp)

    # Build left column: on-duty grouped by dept
    left_html = ""
    for dept, emps in by_dept.items():
        items_html = ""
        for emp in emps:
            sn_part = f"SN {emp['sn']} " if emp["sn"] else ""
            items_html += f'<li>{sn_part}{emp["name"]} — <em>{dept}</em></li>\n'
        left_html += f"""
          <strong style="color:#0b3a78;">{dept}:</strong>
          <ul style="margin:4px 0 10px 20px; padding:0; font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; line-height:1.7;">
            {items_html}
          </ul>"""

    if not left_html:
        left_html = '<span style="color:#64748b;font-size:13px;">NIL</span>'

    # Build right column: on leave / off
    leave_items = ""
    for emp in on_leave:
        sn_part  = f"SN {emp['sn']} " if emp["sn"] else ""
        status   = emp.get("status", "Leave")
        leave_items += f'<li>{sn_part}{emp["name"]} — <em>{status}</em></li>\n'

    right_html = f"""
      <strong style="color:#0b3a78;">Leave / Off Day:</strong>
      <ul style="margin:4px 0 10px 20px; padding:0; font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; line-height:1.7;">
        {leave_items if leave_items else '<li>NIL</li>'}
      </ul>"""

    return f"""
  <!-- ═══ SECTION: MANPOWER ═══ -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700;
                       color:#0b3a78; letter-spacing:1px;">MANPOWER</span>
        </td>
      </tr>
    </table>

    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:10px;">
      <tr>
        <td width="55%" valign="top" style="padding-right:12px;
            font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; line-height:1.7;">
          {left_html}
        </td>
        <td width="45%" valign="top"
            style="font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; line-height:1.7;">
          {right_html}
        </td>
      </tr>
    </table>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>"""


# ══════════════════════════════════════════════════════════════════
#  بناء صفحات HTML
# ══════════════════════════════════════════════════════════════════

def _render_offload_table(flights: list[dict], meta: dict) -> str:
    """Render offload as flight cards (same style as the monitor page).
    Only rows with actual AWB or reason are shown — no empty rows.
    """
    if not flights:
        return """
    <div style="margin-top:10px; font-family:Calibri,Arial,sans-serif; font-size:12.5px; color:#64748b;">
      NIL — No offload data recorded for this shift.
    </div>"""

    cards = ""
    for flight in flights:
        flt     = flight.get("flight", "") or "—"
        date    = flight.get("date", "")   or "—"
        dest    = flight.get("destination", "") or "—"
        std_raw = flight.get("std_etd", "") or ""
        std_val, etd_val = _format_std_etd(std_raw)
        std = std_val or "—"
        etd = etd_val or "—"

        # ops fields
        email    = (flight.get("email_time") or flight.get("email") or "").strip() or "Pending"
        physical = (flight.get("physical") or "").strip() or "Pending"
        cms      = (flight.get("cms") or "").strip() or "Pending"
        verified = (flight.get("trolley") or "").strip() or "Pending"
        remarks  = (flight.get("remarks") or "").strip() or "—"

        def status_span(v):
            if v and v not in ("Pending", "—"):
                return f'<span style="font-weight:700;color:#1a7a3c;">{v}</span>'
            if v == "—":
                return f'<span style="font-weight:700;color:#1b1f2a;">—</span>'
            return f'<span style="font-weight:700;color:#b26a00;">Pending</span>'

        # Only real rows (have AWB or reason)
        items = [it for it in flight.get("items", [])
                 if (it.get("awb") or "").strip() or (it.get("reason") or "").strip()]

        rows = ""
        for idx, it in enumerate(items, 1):
            bg  = "#ffffff" if idx % 2 == 1 else "#f8faff"
            rsn = it.get("reason", "")
            cls_ = it.get("class_", "")
            reason_cell = (
                f'<span style="display:inline-block;padding:2px 6px;border-radius:6px;'
                f'background:#fff7ed;border:1px solid #fed7aa;color:#c2410c;'
                f'font-size:10.5px;font-weight:700;white-space:nowrap;">{rsn}</span>'
                if rsn else "—"
            )
            class_cell = (
                f'<span style="display:inline-block;padding:2px 6px;border-radius:6px;'
                f'background:#ede9fe;border:1px solid #c4b5fd;color:#5b21b6;'
                f'font-size:10.5px;font-weight:700;">{cls_}</span>'
                if cls_ else "—"
            )
            rows += f"""
            <tr style="background-color:{bg};">
              <td style="padding:7px 6px;border-bottom:1px solid #e4e9f5;">{idx}</td>
              <td style="padding:7px 6px;border-bottom:1px solid #e4e9f5;">{it.get('awb','') or '—'}</td>
              <td style="padding:7px 6px;border-bottom:1px solid #e4e9f5;font-weight:700;">{it.get('pcs','') or '—'}</td>
              <td style="padding:7px 6px;border-bottom:1px solid #e4e9f5;">{it.get('kgs','') or '—'}</td>
              <td style="padding:7px 6px;border-bottom:1px solid #e4e9f5;">{class_cell}</td>
              <td style="padding:7px 6px;border-bottom:1px solid #e4e9f5;">{it.get('description','') or '—'}</td>
              <td style="padding:7px 6px;border-bottom:1px solid #e4e9f5;">{it.get('trolley','') or '—'}</td>
              <td style="padding:7px 6px;border-bottom:1px solid #e4e9f5;">{reason_cell}</td>
            </tr>"""

        if not rows:
            rows = '<tr><td colspan="8" style="padding:10px 6px;color:#64748b;text-align:center;">No shipment data</td></tr>'

        # totals
        def safe_sum(key):
            t = 0.0
            for it in items:
                try: t += float(re.sub(r"[^\d.]","", it.get(key,"") or "0") or 0)
                except: pass
            return t

        real_shp = sum(1 for it in items if (it.get("awb") or "").strip())
        tot_pcs  = int(safe_sum("pcs"))
        tot_kgs  = safe_sum("kgs")

        cards += f"""
<table width="100%" cellpadding="0" cellspacing="0" border="0"
       style="border-collapse:collapse; margin-bottom:12px;">
  <!-- Flight header -->
  <tr>
    <td style="background-color:#0b3a78; padding:8px 10px; border:1px solid #0a3166;
               font-family:Calibri,Arial,sans-serif; color:#fff; font-size:12px; white-space:nowrap;">
      <span style="font-weight:700;">✈ {flt}</span>
      <span style="color:#a8c4f0;">&nbsp;|&nbsp;</span>
      Date: <span style="font-weight:700;">{date}</span>
      <span style="color:#a8c4f0;">&nbsp;|&nbsp;</span>
      STD: <span style="font-weight:700;">{std}</span>
      <span style="color:#a8c4f0;">&nbsp;|&nbsp;</span>
      ETD: <span style="font-weight:700;">{etd}</span>
      <span style="color:#a8c4f0;">&nbsp;|&nbsp;</span>
      DEST: <span style="font-weight:700;">{dest}</span>
    </td>
  </tr>
  <!-- Ops row -->
  <tr>
    <td style="background-color:#f8faff; padding:6px 10px;
               border-left:1px solid #d0d9ee; border-right:1px solid #d0d9ee; border-bottom:1px solid #d0d9ee;
               font-family:Calibri,Arial,sans-serif; font-size:10.5px; color:#1b1f2a;">
      Email: {status_span(email)}
      <span style="color:#8aa2c7;">&nbsp;&nbsp;|&nbsp;&nbsp;</span>
      Ramp Received: {status_span(physical)}
      <span style="color:#8aa2c7;">&nbsp;&nbsp;|&nbsp;&nbsp;</span>
      CMS Completed: {status_span(cms)}
      <span style="color:#8aa2c7;">&nbsp;&nbsp;|&nbsp;&nbsp;</span>
      Pieces Verified: {status_span(verified)}
      <span style="color:#8aa2c7;">&nbsp;&nbsp;|&nbsp;&nbsp;</span>
      Remarks: <span style="font-weight:700;">{remarks}</span>
    </td>
  </tr>
  <!-- Shipments -->
  <tr>
    <td style="padding:0; border-left:1px solid #d0d9ee; border-right:1px solid #d0d9ee;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0"
             style="border-collapse:collapse; font-family:Calibri,Arial,sans-serif; font-size:11.2px;">
        <tr style="background-color:#eef3fc;">
          <td style="padding:6px;border-bottom:1px solid #d0d9ee;color:#0b3a78;font-weight:700;">Item</td>
          <td style="padding:6px;border-bottom:1px solid #d0d9ee;color:#0b3a78;font-weight:700;">AWB</td>
          <td style="padding:6px;border-bottom:1px solid #d0d9ee;color:#0b3a78;font-weight:700;">PCS</td>
          <td style="padding:6px;border-bottom:1px solid #d0d9ee;color:#0b3a78;font-weight:700;">KGS</td>
          <td style="padding:6px;border-bottom:1px solid #d0d9ee;color:#0b3a78;font-weight:700;">Priority</td>
          <td style="padding:6px;border-bottom:1px solid #d0d9ee;color:#0b3a78;font-weight:700;">Description</td>
          <td style="padding:6px;border-bottom:1px solid #d0d9ee;color:#0b3a78;font-weight:700;">ULD</td>
          <td style="padding:6px;border-bottom:1px solid #d0d9ee;color:#0b3a78;font-weight:700;">Offloading Reason</td>
        </tr>
        {rows}
        <tr style="background-color:#eef3fc;">
          <td colspan="8" style="padding:7px 10px;border-top:2px solid #0b3a78;
              font-family:Calibri,Arial,sans-serif;font-size:11.5px;font-weight:700;color:#0b3a78;">
            Total Shipments: {real_shp}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            Total PCS: {tot_pcs or '—'}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            Total KGS: {f"{tot_kgs:g}" or '—'}
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="border-left:1px solid #d0d9ee;border-right:1px solid #d0d9ee;
               border-bottom:1px solid #d0d9ee;font-size:1px;line-height:1px;">&nbsp;</td>
  </tr>
</table>"""

    return f"""
    <div style="margin-top:12px; font-family:Calibri,Arial,sans-serif;
                font-size:12px; font-weight:700; color:#0b3a78; margin-bottom:8px;">
      Offload Record:
    </div>
    {cards}"""


def _render_manpower_section(roster: dict) -> str:
    """Render Section 6 MANPOWER exactly like index1.html."""
    on_duty  = roster.get("on_duty",  [])
    on_leave = roster.get("on_leave", [])

    if not on_duty and not on_leave:
        return """
          <strong style="color:#0b3a78;">On Duty:</strong>
          <ul style="margin:4px 0 10px 20px; padding:0;">
            <li style="color:#64748b;">No roster data available for this date.</li>
          </ul>"""

    # Group on-duty by dept
    from collections import OrderedDict
    by_dept: dict[str, list] = OrderedDict()
    for emp in on_duty:
        by_dept.setdefault(emp["dept"], []).append(emp)

    left_html = ""
    for dept, emps in by_dept.items():
        items_li = ""
        for emp in emps:
            sn = f"SN {emp['sn']} " if emp["sn"] else ""
            items_li += f"<li>{sn}{emp['name']} — <em>{dept}</em></li>\n"
        left_html += f"""
          <strong style="color:#0b3a78;">{dept}:</strong>
          <ul style="margin:4px 0 10px 20px; padding:0;">
            {items_li}
          </ul>"""

    if not left_html:
        left_html = '<p style="color:#64748b; margin:4px 0;">NIL</p>'

    leave_li = ""
    for emp in on_leave:
        sn = f"SN {emp['sn']} " if emp["sn"] else ""
        leave_li += f"<li>{sn}{emp['name']} — <em>{emp.get('status','Leave')}</em></li>\n"

    right_html = f"""
          <strong style="color:#0b3a78;">Leave / Off Day:</strong>
          <ul style="margin:4px 0 10px 20px; padding:0;">
            {leave_li if leave_li else '<li>NIL</li>'}
          </ul>"""

    return f"""
        <td width="50%" valign="top" style="padding-right:10px; font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; line-height:1.7;">
          {left_html}
        </td>
        <td width="50%" valign="top" style="font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; line-height:1.7;">
          {right_html}
        </td>"""


def build_shift_report(date_dir: str, shift: str) -> None:
    folder = DATA_DIR / date_dir / shift
    if not folder.exists():
        return

    meta         = load_json(folder / "meta.json", {"flights": {}})
    flight_files = sorted(p for p in folder.glob("*.json") if p.name != "meta.json")
    flights      = [json.loads(p.read_text(encoding="utf-8")) for p in flight_files]

    shift_labels = {
        "shift1": {"ar": "صباح",     "en": "Morning",   "time": "06:00 – 15:00"},
        "shift2": {"ar": "ظهر/مساء", "en": "Afternoon", "time": "15:00 – 22:00"},
        "shift3": {"ar": "ليل",      "en": "Night",      "time": "22:00 – 06:00"},
    }
    sl            = shift_labels.get(shift, {"ar": shift, "en": shift, "time": ""})
    shift_label   = f"{sl['en']} Shift — {sl['time']}"
    total_flights = len(flights)
    total_items   = sum(len(f.get("items", [])) for f in flights)

    # Build offload table (Section 4)
    offload_table_html = _render_offload_table(flights, meta)

    # Offload summary text for Shift Summary card
    flt_names = list({f.get("flight","") for f in flights if f.get("flight","")})
    if total_items:
        offload_summary = f"{total_items} offloaded shipment{'s' if total_items!=1 else ''} across {len(flt_names)} flight{'s' if len(flt_names)!=1 else ''} ({', '.join(flt_names)})."
    else:
        offload_summary = "NIL"

    # Fetch roster
    roster        = fetch_roster_staff(date_dir, shift)
    manpower_cols = _render_manpower_section(roster)

    # Format date for display
    try:
        from datetime import datetime as _dt
        d = _dt.strptime(date_dir, "%Y-%m-%d")
        date_display = d.strftime("%d %b %Y").upper()
    except Exception:
        date_display = date_dir

    html = f"""<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <title>Export Warehouse Activity Report — {date_display}</title>
</head>
<body style="margin:0; padding:0; background-color:#eef1f7; font-family:Calibri, Arial, sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#eef1f7; padding:20px 0;">
<tr><td align="left" style="padding:0 10px;">

<table width="760" cellpadding="0" cellspacing="0" border="0"
       style="width:760px; max-width:760px; background-color:#ffffff; border:1px solid #d0d5e8;">

  <!-- ═══ HEADER ═══ -->
  <tr>
    <td style="background-color:#0b3a78; padding:0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="6" style="background-color:#1a6ecf; font-size:1px; line-height:1px;">&nbsp;</td>
          <td style="padding:20px 22px 18px 20px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td>
                  <div style="font-family:Calibri,Arial,sans-serif; font-size:20px; font-weight:700; color:#ffffff; letter-spacing:0.5px; line-height:1.2;">
                    ✈&nbsp; Export Warehouse Activity Report
                  </div>
                  <div style="font-family:Calibri,Arial,sans-serif; font-size:13px; color:#a8c4f0; margin-top:5px; letter-spacing:0.3px;">
                    Shift Date: <strong style="color:#d4e6ff;">{date_display}</strong>
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    Time: <strong style="color:#d4e6ff;">{sl['time']} LT</strong>
                    &nbsp;&nbsp;|&nbsp;&nbsp;
                    <strong style="color:#d4e6ff;">{sl['en']} Shift</strong>
                  </div>
                </td>
                <td align="right" valign="middle" style="padding-right:4px;">
                  <div style="font-family:Calibri,Arial,sans-serif; font-size:11px; color:#6b9fd4; line-height:1.4;">
                    Oman SATS LLC<br>
                    <strong style="color:#a8c4f0;">Export Operations</strong>
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ═══ BACK LINK ═══ -->
  <tr>
    <td style="padding:10px 24px 8px 24px; background-color:#f8faff; border-bottom:1px solid #e4e9f5;">
      <a href="../../index.html"
         style="font-family:Calibri,Arial,sans-serif; font-size:12px; color:#0b3a78; font-weight:700; text-decoration:none;">
        ← Back to Index
      </a>
    </td>
  </tr>

  <!-- ═══ SHIFT SUMMARY ═══ -->
  <tr>
    <td style="padding:16px 24px 0 24px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
          <td style="padding:6px 10px; background-color:#eef3fc;">
            <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px; text-transform:uppercase;">Shift Summary</span>
          </td>
        </tr>
      </table>
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:10px; border:1px solid #e0e7f5;">
        <tr>
          <td width="50%" style="padding:10px 12px; border-right:1px solid #e0e7f5; border-bottom:1px solid #e0e7f5; font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; vertical-align:top; background-color:#f5f8ff;">
            <strong style="color:#0b3a78;">✅ Flight Performance:</strong><br>
            All flights departed on time (no cargo-related delay).
          </td>
          <td width="50%" style="padding:10px 12px; border-bottom:1px solid #e0e7f5; font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; vertical-align:top; background-color:#f5f8ff;">
            <strong style="color:#0b3a78;">🚫 Cargo Delay:</strong><br>
            NIL
          </td>
        </tr>
        <tr>
          <td style="padding:10px 12px; border-right:1px solid #e0e7f5; border-bottom:1px solid #e0e7f5; font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; vertical-align:top;">
            <strong style="color:#0b3a78;">⚠️ DG Irregularities:</strong><br>
            DG Embargo station check done.
          </td>
          <td style="padding:10px 12px; border-bottom:1px solid #e0e7f5; font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; vertical-align:top;">
            <strong style="color:#0b3a78;">🦺 Safety Incidents:</strong><br>
            Safety briefing done; PPE compliance confirmed.
          </td>
        </tr>
        <tr>
          <td style="padding:10px 12px; border-right:1px solid #e0e7f5; font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; vertical-align:top; background-color:#f5f8ff;">
            <strong style="color:#0b3a78;">📦 Offloaded / Missing AWB:</strong><br>
            {offload_summary}
          </td>
          <td style="padding:10px 12px; font-family:Calibri,Arial,sans-serif; font-size:13px; color:#1b1f2a; vertical-align:top; background-color:#f5f8ff;">
            <strong style="color:#0b3a78;">📋 UTL Report:</strong><br>
            NIL
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ═══ SECTION 1: OPERATIONAL ACTIVITIES ═══ -->
  <tr><td style="padding:18px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">1.&nbsp; OPERATIONAL ACTIVITIES</span>
        </td>
      </tr>
    </table>
    <div style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <strong style="color:#0b3a78;">Load Plan:</strong>
      <ul style="margin:4px 0 10px 22px; padding:0; color:#1b1f2a;"><li>&nbsp;</li></ul>
      <strong style="color:#0b3a78;">Advance Loading:</strong>
      <ul style="margin:4px 0 10px 22px; padding:0; color:#1b1f2a;"><li>&nbsp;</li></ul>
      <strong style="color:#0b3a78;">CSD Rescreening:</strong>
      <ul style="margin:4px 0 4px 22px; padding:0; color:#1b1f2a;"><li>&nbsp;</li></ul>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 2: BRIEFINGS ═══ -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">2.&nbsp; BRIEFINGS CONDUCTED</span>
        </td>
      </tr>
    </table>
    <div style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <ul style="margin:0 0 0 22px; padding:0; color:#1b1f2a;">
        <li>Safety tool box.</li>
        <li>ULD and net serviceability.</li>
        <li>Punctuality.</li>
        <li>Proper loading and counting of cargo.</li>
        <li>Not to use mobile phone while driving.</li>
      </ul>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 3: FLIGHT PERFORMANCE ═══ -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">3.&nbsp; FLIGHT PERFORMANCE</span>
        </td>
      </tr>
    </table>
    <div style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      ✅&nbsp; All flights departed on time; no delay related to Cargo.
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 4: CHECKS & COMPLIANCE + OFFLOAD ═══ -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">4.&nbsp; CHECKS &amp; COMPLIANCE</span>
        </td>
      </tr>
    </table>
    <div style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <strong style="color:#0b3a78;">DG Check:</strong> DG Embargo station check done. AWB left behind: NIL.
    </div>
    {offload_table_html}
    <div style="border-top:1px solid #e4e9f5; margin-top:16px;"></div>
  </td></tr>

  <!-- ═══ SECTION 5: SAFETY ═══ -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#1a7a3c;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#edf7f1;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#1a7a3c; letter-spacing:1px;">5.&nbsp; SAFETY</span>
        </td>
      </tr>
    </table>
    <div style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      Safety briefing conducted to all staff, drivers and porters. All checkers reminded to verify net expiration &amp; ULD serviceability.
      <br><br>
      <strong style="color:#1a7a3c;">✅ Note:</strong> All staff and drivers are wearing proper PPEs.
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 6: MANPOWER ═══ -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">6.&nbsp; MANPOWER</span>
        </td>
      </tr>
    </table>
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:10px;">
      <tr>
        {manpower_cols}
      </tr>
    </table>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 7: EQUIPMENT STATUS ═══ -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#1a7a3c;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#edf7f1;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#1a7a3c; letter-spacing:1px;">7.&nbsp; EQUIPMENT STATUS</span>
        </td>
      </tr>
    </table>
    <div style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      ✅&nbsp; <strong>ALL EQUIPMENT ARE OK.</strong>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 8: HANDOVER ═══ -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">8.&nbsp; HANDOVER DETAILS</span>
        </td>
      </tr>
    </table>
    <div style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <ul style="margin:0 0 10px 22px; padding:0;">
        <li>READ AND SIGN.</li>
        <li>Shell &amp; Al-Maha Card Fuel.</li>
        <li>DIP MAIL Cage Keys.</li>
        <li>Supervisor mobile phone H/O in good condition.</li>
        <li>DSE RADIO.</li>
        <li>All trolleys arranged for early flight.</li>
      </ul>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 9: OTHER ═══ -->
  <tr><td style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#7a5200;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#fdf6ec;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#7a5200; letter-spacing:1px;">9.&nbsp; OTHER</span>
        </td>
      </tr>
    </table>
    <div style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      &nbsp;
    </div>
  </td></tr>

  <!-- ═══ FOOTER ═══ -->
  <tr>
    <td style="padding:20px 24px 22px 24px; background-color:#f8faff; border-top:2px solid #0b3a78; margin-top:16px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td valign="top">
            <div style="font-family:Calibri,Arial,sans-serif; font-size:14px; color:#1b1f2a; line-height:1.7;">
              Best Regards,<br>
              <strong style="font-size:15px; color:#0b3a78;">Supervisor – Export Operations</strong><br>
              <span style="color:#444;">Oman SATS LLC</span>
            </div>
          </td>
          <td align="right" valign="bottom">
            <div style="font-family:Calibri,Arial,sans-serif; font-size:11px; color:#8a9ab5; font-style:italic;">
              Generated automatically · {date_dir}
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="background-color:#0b3a78; height:5px; font-size:1px; line-height:1px;">&nbsp;</td>
  </tr>

</table>
</td></tr>
</table>
</body>
</html>"""

    out_dir = DOCS_DIR / date_dir / shift
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "index.html").write_text(html, encoding="utf-8")



def _shift_window(date_dir: str, shift: str) -> tuple[str, str]:
    """Return (start_iso, end_iso) in local time (Asia/Muscat) for display in the index."""
    # Shift boundaries (local time)
    if shift == "shift1":
        start_hm, end_hm, add_day = "06:00", "15:00", 0
    elif shift == "shift2":
        start_hm, end_hm, add_day = "15:00", "22:00", 0
    else:
        start_hm, end_hm, add_day = "22:00", "06:00", 1  # crosses midnight

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


def _share_links_block(date_dir: str, shift: str) -> str:
    """روابط مشاركة (بدون لمس صفحة الروستر أو نسخ ملف داخلها)."""
    # قد لا تستخدم صفحة الروستر هذه الـ query params حالياً، لكنها مفيدة للمشاركة/التوثيق.
    roster_with_params = f"{ROSTER_PAGE_URL}?date={date_dir}&shift={shift}"
    return f"""
    <div class=\"share-box\">
      <div>
        <div class=\"share-title\">روابط سريعة</div>
        <div class=\"share-sub\">روابط مشاركة للروستر وملف الموظفين (بدون رفع أي ملف داخل صفحة الروستر)</div>
      </div>
      <div class=\"share-actions\">
        <a class=\"btn-link\" href=\"{roster_with_params}\" target=\"_blank\" rel=\"noopener\">👥 صفحة الروستر</a>
        <a class=\"btn-link\" href=\"{ROSTER_JSON_RAW}\" target=\"_blank\" rel=\"noopener\">🧾 roster.json (raw)</a>
        <a class=\"btn-link\" href=\"{ROSTER_JSON_GH}\" target=\"_blank\" rel=\"noopener\">✏️ GitHub</a>
      </div>
    </div>
    """


def build_root_index(now: datetime) -> None:
    """Modern home page with accordion days; current day opened by default."""
    if not DOCS_DIR.exists():
        return

    today = now.strftime("%Y-%m-%d")

    day_dirs = sorted(
        (p for p in DOCS_DIR.iterdir()
         if p.is_dir() and re.match(r"\d{4}-\d{2}-\d{2}", p.name)),
        reverse=True,
    )

    shift_meta = {
        "shift1": {"label": "Morning",   "ar": "صباح", "time": "06:00 – 15:00", "icon": "🌅"},
        "shift2": {"label": "Afternoon", "ar": "ظهر",  "time": "15:00 – 22:00", "icon": "☀️"},
        "shift3": {"label": "Night",     "ar": "ليل",  "time": "22:00 – 06:00", "icon": "🌙"},
    }

    # Count total days & flights for header
    total_days = len(day_dirs)

    days_html = ""
    for day_dir in day_dirs:
        day = day_dir.name
        is_today = day == today
        open_attr = " open" if is_today else ""

        # Count flights in this day
        day_flights = 0
        for shift in ("shift1", "shift2", "shift3"):
            shift_folder = DATA_DIR / day / shift
            if shift_folder.exists():
                day_flights += sum(1 for f in shift_folder.glob("*.json") if f.name != "meta.json")

        badge = '<span class="today-badge">TODAY</span>' if is_today else ""
        flights_pill = f'<span class="day-pill">{day_flights} flights</span>' if day_flights else ""

        rows = ""
        for shift in ("shift1", "shift2", "shift3"):
            if not (day_dir / shift / "index.html").exists():
                continue

            meta_s = shift_meta.get(shift, {"label": shift, "ar": shift, "time": "", "icon": "✈"})
            # Count flights for this specific shift
            shift_flt_count = 0
            shift_folder = DATA_DIR / day / shift
            if shift_folder.exists():
                shift_flt_count = sum(1 for f in shift_folder.glob("*.json") if f.name != "meta.json")
            flt_txt = f"{shift_flt_count} flight{'s' if shift_flt_count != 1 else ''}" if shift_flt_count else ""

            rows += f"""
            <a class="shift-card" href="{day}/{shift}/">
                <div class="sc-icon">{meta_s['icon']}</div>
                <div class="sc-body">
                    <div class="sc-title">{meta_s['ar']} <span class="sc-en">/ {meta_s['label']}</span></div>
                    <div class="sc-time">{meta_s['time']}</div>
                </div>
                <div class="sc-right">
                    {f'<span class="sc-count">{flt_txt}</span>' if flt_txt else ''}
                    <span class="sc-arrow">›</span>
                </div>
            </a>"""

        if not rows:
            rows = '<div class="empty-day">لا توجد تقارير لهذا اليوم.</div>'

        days_html += f"""
        <details class="day-accordion"{open_attr}>
            <summary class="day-summary">
                <div class="day-sum-left">
                    <span class="day-date">📅 {day}</span>
                    {badge}
                    {flights_pill}
                </div>
                <span class="day-chev">›</span>
            </summary>
            <div class="day-body">
                {rows}
            </div>
        </details>"""

    if not days_html:
        days_html = "<div class='empty-day' style='text-align:center;padding:48px'>لا توجد تقارير بعد.</div>"

    html = f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Offload Monitor</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');
        :root {{
            --bg:#f1f5fb; --card:#fff; --border:#e2e8f4;
            --navy:#0a1f52; --blue:#2563eb; --blue-dk:#1d4ed8;
            --muted:#64748b; --text:#0f1f3d;
            --shadow:0 4px 20px rgba(15,31,61,.10);
        }}
        *, *::before, *::after {{ box-sizing:border-box; margin:0; padding:0; }}
        body {{
            background:var(--bg);
            font-family:'Inter', system-ui, sans-serif;
            color:var(--text);
            -webkit-font-smoothing:antialiased;
        }}

        /* ── HEADER ── */
        .top {{
            background: linear-gradient(135deg, var(--navy) 0%, #1a3a8f 55%, #2251c9 100%);
            color:#fff;
            padding: 20px 24px 18px;
            box-shadow: 0 2px 16px rgba(10,31,82,.35);
        }}
        .top-inner {{ max-width:860px; margin:0 auto; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px; }}
        .top h1 {{ font-size:20px; font-weight:800; letter-spacing:-.2px; }}
        .top p  {{ font-size:12px; opacity:.7; margin-top:3px; }}
        .hdr-badge {{
            background:rgba(255,255,255,.12);
            border:1px solid rgba(255,255,255,.2);
            border-radius:10px;
            padding:7px 16px;
            font-size:12px; font-weight:600;
            color:rgba(255,255,255,.85);
            text-align:center;
        }}
        .hdr-badge strong {{ display:block; font-size:22px; font-weight:800; color:#fff; line-height:1.2; }}

        /* ── WRAP ── */
        .wrap {{ max-width:860px; margin:20px auto 60px; padding:0 14px; display:flex; flex-direction:column; gap:12px; }}

        /* ── ACCORDION ── */
        .day-accordion {{
            background:var(--card);
            border:1px solid var(--border);
            border-radius:16px;
            box-shadow:var(--shadow);
            overflow:hidden;
        }}
        .day-summary {{
            list-style:none;
            cursor:pointer;
            padding:14px 18px;
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:12px;
            background:#f7f9fd;
            border-bottom:1px solid var(--border);
            user-select:none;
            transition:background .12s;
        }}
        .day-summary:hover {{ background:#eff3fb; }}
        .day-accordion[open] .day-summary {{ background:#eff3fb; }}
        .day-summary::-webkit-details-marker {{ display:none; }}
        .day-sum-left {{ display:flex; align-items:center; gap:10px; flex-wrap:wrap; }}
        .day-date {{ font-weight:800; font-size:14px; color:var(--navy); font-family:'JetBrains Mono', monospace; }}
        .today-badge {{
            background: linear-gradient(90deg, #2563eb, #1d4ed8);
            color:#fff; font-size:10px; font-weight:800;
            padding:2px 8px; border-radius:999px;
            letter-spacing:.5px;
        }}
        .day-pill {{
            background:#e8f0fe; color:var(--navy);
            border:1px solid #bdd1ff;
            font-size:11px; font-weight:700;
            padding:2px 9px; border-radius:999px;
        }}
        .day-chev {{
            font-size:20px; color:var(--blue); font-weight:900;
            transition:transform .2s;
        }}
        .day-accordion[open] .day-chev {{ transform: rotate(90deg); }}

        .day-body {{ padding:12px 14px 14px; display:flex; flex-direction:column; gap:8px; }}

        /* ── SHIFT CARDS ── */
        .shift-card {{
            display:flex;
            align-items:center;
            gap:14px;
            text-decoration:none;
            color:inherit;
            background:#fff;
            border:1px solid var(--border);
            border-radius:12px;
            padding:12px 16px;
            transition:transform .1s, border-color .15s, box-shadow .15s;
        }}
        .shift-card:hover {{
            transform:translateY(-1px);
            border-color:#93c5fd;
            box-shadow:0 4px 14px rgba(37,99,235,.12);
        }}
        .sc-icon {{ font-size:22px; line-height:1; flex-shrink:0; }}
        .sc-body {{ flex:1; }}
        .sc-title {{ font-size:14px; font-weight:800; color:var(--navy); }}
        .sc-en {{ font-size:12px; font-weight:600; color:var(--muted); }}
        .sc-time {{ font-size:12px; color:var(--muted); margin-top:2px; font-family:'JetBrains Mono',monospace; }}
        .sc-right {{ display:flex; align-items:center; gap:10px; }}
        .sc-count {{
            background:#eff6ff; color:var(--blue-dk);
            border:1px solid #bfdbfe;
            font-size:11px; font-weight:700;
            padding:3px 10px; border-radius:999px;
        }}
        .sc-arrow {{ font-size:20px; color:var(--blue); font-weight:900; }}

        .empty-day {{ padding:18px; color:var(--muted); text-align:center; font-size:13px; }}
        .footer {{ text-align:center; color:var(--muted); font-size:11.5px; margin-top:6px; }}

        @media (max-width:600px) {{
            .top {{ padding:14px 14px 12px; }}
            .top h1 {{ font-size:17px; }}
            .wrap {{ padding:0 10px; margin-top:14px; }}
            .day-date {{ font-size:13px; }}
            .shift-card {{ padding:10px 12px; }}
            .hdr-badge {{ display:none; }}
        }}
    </style>
</head>
<body>
    <div class="top">
        <div class="top-inner">
            <div>
                <h1>✈ Offload Monitor</h1>
                <p>اختر المناوبة لعرض تفاصيل الـ Offload · الأيام السابقة مطوية تلقائيًا</p>
            </div>
            <div class="hdr-badge"><strong>{total_days}</strong>Days</div>
        </div>
    </div>
    <div class="wrap">
        {days_html}
        <div class="footer">Generated automatically by GitHub Actions · {today}</div>
    </div>
</body>
</html>"""

    (DOCS_DIR / "index.html").write_text(html, encoding="utf-8")


# ══════════════════════════════════════════════════════════════════
#  إعادة إثراء الرحلات السابقة بـ DEST الصحيح من AirLabs
# ══════════════════════════════════════════════════════════════════

def retroactive_enrich_all(now: datetime) -> None:
    """Go through ALL saved flight JSON files in data/ and update destination + STD/ETD
    from AirLabs for any flight where dest is missing or potentially wrong.

    Called only when AIRLABS_API_KEY is set AND RETRO_ENRICH=1 env var is present.
    """
    if not os.environ.get("AIRLABS_API_KEY", "").strip():
        print("[retroactive] AIRLABS_API_KEY not set — skipping.")
        return

    if not DATA_DIR.exists():
        print("[retroactive] data/ directory not found — skipping.")
        return

    updated_count = 0
    skipped_count = 0

    for json_file in sorted(DATA_DIR.rglob("*.json")):
        if json_file.name == "meta.json":
            continue

        try:
            flight = json.loads(json_file.read_text(encoding="utf-8"))
        except Exception:
            continue

        flt  = (flight.get("flight") or "").strip()
        date = (flight.get("date")   or "").strip()
        if not flt:
            skipped_count += 1
            continue

        flight_date_iso = normalize_flight_date(date, now) if date else None
        info = fetch_flight_info_airlabs(flt, flight_date=flight_date_iso)

        if not info:
            skipped_count += 1
            continue

        changed = False

        # Update destination
        new_dest = (info.get("dest") or "").strip()
        if new_dest and new_dest != (flight.get("destination") or "").strip():
            print(f"  [retro] {flt}/{date}: dest {flight.get('destination')!r} → {new_dest!r}")
            flight["destination"] = new_dest
            changed = True

        # Update STD/ETD only if missing
        std = (info.get("std") or "").strip()
        etd = (info.get("etd") or "").strip()
        new_std_etd = ""
        if std and etd and std != etd:
            new_std_etd = f"{std}|{etd}"
        elif std:
            new_std_etd = std

        if new_std_etd and not (flight.get("std_etd") or "").strip():
            flight["std_etd"] = new_std_etd
            changed = True

        if changed:
            flight["retro_enriched_at"] = now.isoformat()
            json_file.write_text(json.dumps(flight, ensure_ascii=False, indent=2), encoding="utf-8")
            updated_count += 1
        else:
            skipped_count += 1

    print(f"[retroactive] Done. Updated: {updated_count}, Skipped/unchanged: {skipped_count}")

    # Rebuild ALL shift reports
    print("[retroactive] Rebuilding HTML reports…")
    for date_dir_p in sorted(p for p in DATA_DIR.iterdir() if p.is_dir()):
        for shift in ("shift1", "shift2", "shift3"):
            if (date_dir_p / shift).exists():
                build_shift_report(date_dir_p.name, shift)
                print(f"  rebuilt: {date_dir_p.name}/{shift}")

    build_root_index(now)
    print("[retroactive] All reports rebuilt. ✓")


# ══════════════════════════════════════════════════════════════════
#  نقطة الدخول
# ══════════════════════════════════════════════════════════════════

def main() -> None:
    now = datetime.now(ZoneInfo(TIMEZONE))
    print(f"[{now.isoformat()}] Starting…")

    # ── وضع الإثراء الرجعي ──
    if os.getenv("RETRO_ENRICH", "").strip().lower() in ("1", "true", "yes", "y"):
        print("RETRO_ENRICH=1 detected. Running retroactive enrichment for all saved flights…")
        retroactive_enrich_all(now)
        return

    print(f"Downloading file…")
    html     = download_file()
    new_hash = compute_sha256(html)

    # تشخيص سريع
    print(f"HTML length: {len(html)}")
    print(f"HTML sha256: {new_hash[:16]}")

    if STATE_FILE.exists():
        old_hash = STATE_FILE.read_text(encoding="utf-8").strip()
        if old_hash == new_hash and not FORCE_REBUILD:
            print("No change detected, but rebuilding reports anyway…")
            build_root_index(now)
        if old_hash == new_hash and FORCE_REBUILD:
            print("No change detected, but FORCE_REBUILD=1 → continuing to rebuild.")

    print("Change detected. Parsing…")
    flights = extract_flights(html)

    if not flights:
        print("WARNING: No flights extracted. Check HTML structure.")
        STATE_FILE.write_text(new_hash, encoding="utf-8")
        return

    # ── Enrich with AirLabs ──
    if os.environ.get("AIRLABS_API_KEY", "").strip():
        enriched = 0
        for f in flights:
            flt = (f.get("flight") or "").strip()
            if not flt:
                continue
            info = fetch_flight_info_airlabs(flt, flight_date=normalize_flight_date(f.get('date',''), now))
            if info:
                std  = (info.get("std") or "").strip()
                etd  = (info.get("etd") or "").strip()
                if std and etd and std != etd:
                    f["std_etd"] = f"{std}|{etd}"
                elif std:
                    f["std_etd"] = std
                dest = (info.get("dest") or "").strip()
                if dest:
                    f["destination"] = dest
                enriched += 1
        print(f"Enriched {enriched} flight(s) via AirLabs.")
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

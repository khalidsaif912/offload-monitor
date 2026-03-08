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

RECIPIENTS_FILE: Path = DOCS_DIR / "data" / "email_recipients.json"

def ensure_email_recipients_file() -> None:
    try:
        if not RECIPIENTS_FILE.exists():
            RECIPIENTS_FILE.parent.mkdir(parents=True, exist_ok=True)
            RECIPIENTS_FILE.write_text(json.dumps({
                "recipients": [],
                "updated_at": "",
                "source": "offload-monitor"
            }, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        print(f"  [recipients] could not ensure file: {exc}")


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
    if 6 * 60 <= mins < 13 * 60:
        return "shift1"
    if 13 * 60 <= mins < 21 * 60:
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
    if not isinstance(meta, dict) or "flights" not in meta:
        meta = {"flights": {}}

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
            f'background-color:#e6f4ea;border:1px solid #a8d5ba;color:#1a7a3c;'
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
              <td style="padding:7px 6px; border-bottom:1px solid #e4e9f5;">{idx}</td>
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
          <td style="padding:6px 6px; border-bottom:1px solid #d0d9ee; color:#0b3a78; font-weight:700;">#</td>
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
    """Render offload — one card per flight matching screenshot design."""
    from datetime import datetime as _dt
    onedrive_url = os.environ.get("ONEDRIVE_FILE_URL","").strip()
    update_link  = onedrive_url if onedrive_url else "#"
    now_str      = _dt.now().strftime("%Y-%m-%d %H:%M")

    if not flights:
        return """
    <div style="margin-top:12px;overflow-x:auto;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"
           style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:12px;">
      <tr style="background:#0b3a78;">
        <td style="padding:7px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;text-align:center;width:30px;">#</td>
        <td style="padding:7px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;">AWB</td>
        <td style="padding:7px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;text-align:center;">PCS</td>
        <td style="padding:7px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;text-align:center;">KGS</td>
        <td style="padding:7px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;text-align:center;">Priority</td>
        <td style="padding:7px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;">Description</td>
        <td style="padding:7px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;text-align:center;">ULD</td>
        <td style="padding:7px 6px;color:#fff;font-weight:700;border:1px solid #0a3166;">Offloading Reason</td>
      </tr>
      <tr>
        <td colspan="8" style="padding:14px 6px;border:1px solid #dde6f5;
            color:#64748b;text-align:center;font-weight:600;">
          NIL — No offload data recorded for this shift.
        </td>
      </tr>
    </table>
    </div>"""

    cards = ""
    for flight in flights:
        flt     = flight.get("flight","")      or "—"
        date    = flight.get("date","")        or "—"
        dest    = flight.get("destination","") or "—"
        std_raw = flight.get("std_etd","")     or ""
        std_val, etd_val = _format_std_etd(std_raw)
        std = std_val or "—"
        etd = etd_val or "—"

        items = [it for it in flight.get("items",[]) if (it.get("awb","") or "").strip()]
        total_pcs = 0; total_kgs = 0.0
        for it in items:
            try: total_pcs += int(str(it.get("pcs","0") or "0").replace(",",""))
            except: pass
            try: total_kgs += float(str(it.get("kgs","0") or "0").replace(",",""))
            except: pass

        rows = ""
        for i, it in enumerate(items, 1):
            bg     = "#f3f4f6" if i%2 else "#ffffff"
            awb    = it.get("awb","")         or "—"
            pcs    = it.get("pcs","")         or "—"
            kgs    = it.get("kgs","")         or "—"
            desc   = it.get("description","") or "—"
            uld    = it.get("trolley","")     or "—"
            reason = (it.get("reason","")     or "").strip()
            pri    = it.get("class_","")      or "—"
            badge  = (
                f'<span style="display:inline-block;padding:2px 8px;border-radius:4px;'                f'background:#fff3e0;border:1px solid #ffb74d;color:#e65100;'                f'font-size:10.5px;font-weight:700;">{reason.upper()}</span>'
            ) if reason else "—"
            rows += f"""
        <tr style="background:{bg};">
          <td style="padding:9px 8px;border:1px solid #e5e7eb;text-align:center;font-weight:700;font-size:13px;">{i}</td>
          <td style="padding:9px 8px;border:1px solid #e5e7eb;font-family:'Courier New',monospace;font-size:12px;color:#1e40af;font-weight:600;">{awb}</td>
          <td style="padding:9px 8px;border:1px solid #e5e7eb;text-align:center;font-weight:700;font-size:13px;">{pcs}</td>
          <td style="padding:9px 8px;border:1px solid #e5e7eb;text-align:center;font-size:13px;">{kgs}</td>
          <td style="padding:9px 8px;border:1px solid #e5e7eb;text-align:center;font-size:13px;">{pri}</td>
          <td style="padding:9px 8px;border:1px solid #e5e7eb;font-size:13px;">{desc}</td>
          <td style="padding:9px 8px;border:1px solid #e5e7eb;text-align:center;font-size:13px;">{uld}</td>
          <td style="padding:9px 8px;border:1px solid #e5e7eb;font-size:13px;">{badge}</td>
        </tr>"""

        if not rows:
            rows = '<tr><td colspan="8" style="padding:12px;border:1px solid #dde6f5;color:#64748b;text-align:center;">No items recorded.</td></tr>'

        cards += f"""
    <div style="margin-top:14px;overflow:hidden;border:1px solid #c7d4f0;font-family:Calibri,Arial,sans-serif;">
      <!-- Flight header (table-based for Outlook) -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
        <tr>
          <td style="background-color:#1e40af;padding:10px 14px;color:#ffffff;font-size:13px;">
            <span style="color:#ffffff;font-weight:800;font-size:15px;">&#9992; {flt}</span>
            &nbsp;&nbsp;
            <span style="color:#93c5fd;font-size:12px;"><strong style="color:#ffffff;">Date:</strong> {date}</span>
            &nbsp;&nbsp;
            <span style="color:#93c5fd;font-size:12px;"><strong style="color:#ffffff;">STD:</strong> {std}</span>
            &nbsp;&nbsp;
            <span style="color:#93c5fd;font-size:12px;"><strong style="color:#ffffff;">ETD:</strong> {etd}</span>
            &nbsp;&nbsp;
            <span style="color:#93c5fd;font-size:12px;"><strong style="color:#ffffff;">DEST:</strong> <strong style="color:#fbbf24;font-size:13px;">{dest}</strong></span>
          </td>
          <td style="background-color:#1e40af;padding:10px 14px;text-align:right;color:#ffffff;font-size:12px;white-space:nowrap;">
            <span style="color:#b91c1c;font-weight:700;">&#10227; UPDATE</span>
            &nbsp;
            <span style="color:#bfdbfe;font-size:11px;">Updated: {now_str}</span>
          </td>
        </tr>
      </table>
      <!-- Status bar (table-based for Outlook) -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;background-color:#f0f5ff;border-bottom:1px solid #dde6f5;">
        <tr>
          <td style="padding:8px 14px;font-size:12px;color:#374151;">Email: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td style="padding:8px 14px;font-size:12px;color:#374151;">Ramp Received: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td style="padding:8px 14px;font-size:12px;color:#374151;">CMS Completed: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td style="padding:8px 14px;font-size:12px;color:#374151;">Pieces Verified: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td style="padding:8px 14px;font-size:12px;color:#374151;">Remarks: <strong style="color:#374151;">—</strong></td>
        </tr>
      </table>
      <!-- Table -->
      <div style="overflow-x:auto;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;table-layout:fixed;width:100%;font-size:13px;">
          <tr style="background:#1e40af;">
            <td style="padding:9px 8px;color:#ffffff;font-weight:700;border:1px solid #3b5fd9;text-align:center;width:42px;font-size:13px;white-space:nowrap;">#</td>
            <td style="padding:9px 8px;color:#ffffff;font-weight:700;border:1px solid #3b5fd9;font-size:13px;width:132px;white-space:nowrap;">AWB</td>
            <td style="padding:9px 8px;color:#ffffff;font-weight:700;border:1px solid #3b5fd9;text-align:center;font-size:13px;width:72px;white-space:nowrap;">PCS</td>
            <td style="padding:9px 8px;color:#ffffff;font-weight:700;border:1px solid #3b5fd9;text-align:center;font-size:13px;width:82px;white-space:nowrap;">KGS</td>
            <td style="padding:9px 8px;color:#ffffff;font-weight:700;border:1px solid #3b5fd9;text-align:center;font-size:13px;width:94px;white-space:nowrap;">Priority</td>
            <td style="padding:9px 8px;color:#ffffff;font-weight:700;border:1px solid #3b5fd9;font-size:13px;width:230px;">Description</td>
            <td style="padding:9px 8px;color:#ffffff;font-weight:700;border:1px solid #3b5fd9;text-align:center;font-size:13px;width:96px;white-space:nowrap;">ULD</td>
            <td style="padding:9px 8px;color:#ffffff;font-weight:700;border:1px solid #3b5fd9;font-size:13px;width:170px;">Offloading Reason</td>
          </tr>
          {rows}
        </table>
      </div>
      <!-- Footer -->
      <div style="background:#f8faff;padding:8px 14px;border-top:1px solid #dde6f5;font-size:12px;font-weight:600;color:#1b1f2a;">
        Total Shipments: <strong style="color:#0b3a78;">{len(items)}</strong>
        &nbsp;|&nbsp; Total PCS: <strong style="color:#0b3a78;">{total_pcs}</strong>
        &nbsp;|&nbsp; Total KGS: <strong style="color:#0b3a78;">{total_kgs:g}</strong>
      </div>
    </div>"""

    return cards

def _render_manpower_section(roster: dict, supervisor_display: str = "") -> str:
    """Render Section 6 MANPOWER — grouped by dept, sections B-G."""
    import re as _re
    on_duty = roster.get("on_duty", [])

    td_style  = "font-family:Calibri,Arial,sans-serif;font-size:13px;color:#1b1f2a;line-height:1.8;"
    hdr_style = "color:#0b3a78;font-weight:700;font-size:13px;"
    dept_hdr  = "color:#0b3a78;font-weight:700;font-size:12px;margin:8px 0 2px 0;"
    ul_style  = "margin:2px 0 10px 20px;padding:0;"
    nil_item  = '<li style="color:#64748b;">NIL</li>'

    EXCLUDED_DEPTS = {"officers"}
    EXCLUDED_SNS   = {"990737"}  # Said Al Amri

    def _is_support(e):
        return "support" in (e.get("name","") + e.get("dept","")).lower()

    def _is_excluded(e):
        dept = e.get("dept","").strip().lower()
        sn   = str(e.get("sn","")).strip()
        return dept in EXCLUDED_DEPTS or _is_support(e) or sn in EXCLUDED_SNS

    def _fmt_name(emp):
        raw  = emp.get("name","").strip()
        sn   = str(emp.get("sn") or "").strip()
        m = _re.match(r"^(.+?)\s*-\s*(\d{4,})\s*(?:\((.+?)\))?$", raw)
        if m:
            name_part = m.group(1).strip()
            sn_part   = m.group(2).strip()
            note      = m.group(3).strip() if m.group(3) else ""
            note_html = f' <em style="color:#888;font-size:11px;">({note})</em>' if note else ""
            return f"SN {sn_part} — {name_part}{note_html}"
        sn_html = f"SN {sn} — " if sn else ""
        return f"{sn_html}{raw}"

    # الموظفون الرئيسيون مجمّعون بالقسم
    from collections import OrderedDict
    main_emps = [e for e in on_duty if not _is_excluded(e)]
    by_dept = OrderedDict()
    for emp in main_emps:
        by_dept.setdefault(emp.get("dept","Other"), []).append(emp)

    # الأقسام التي تُعالج يدوياً — لا تُكرَّر في الـ loop أدناه
    MANUAL_DEPTS = {"supervisors"}

    grouped_html = ""
    # أولاً: قسم Supervisors — يُعرض دائماً في الأعلى (مع استثناء EXCLUDED_SNS)
    sup_in_roster = [
        e for e in on_duty
        if e.get("dept","").strip().lower() == "supervisors"
        and str(e.get("sn","")).strip() not in EXCLUDED_SNS
    ]
    if sup_in_roster:
        sup_li_roster = "".join(f"<li>{_fmt_name(e)}</li>\n" for e in sup_in_roster)
        grouped_html += f"""
      <div style="{dept_hdr}">Supervisors:</div>
      <ul style="{ul_style}">{sup_li_roster}</ul>"""
    elif supervisor_display:
        grouped_html += f"""
      <div style="{dept_hdr}">Supervisors:</div>
      <ul style="{ul_style}"><li><strong>{supervisor_display}</strong></li></ul>"""
    else:
        grouped_html += f"""
      <div style="{dept_hdr}">Supervisors:</div>
      <ul style="{ul_style}"><li>&nbsp;</li></ul>"""

    # ثانياً: باقي الأقسام من roster (تخطّى supervisors — مُعالَج أعلاه)
    for dept, emps in by_dept.items():
        if dept.strip().lower() in MANUAL_DEPTS:
            continue
        items_li = "".join(f"<li>{_fmt_name(e)}</li>\n" for e in emps)
        grouped_html += f"""
      <div style="{dept_hdr}">{dept}:</div>
      <ul style="{ul_style}">{items_li}</ul>"""

    if not grouped_html:
        grouped_html = f'<ul style="{ul_style}"><li style="color:#64748b;">No roster data available.</li></ul>'

    # C) Support Team
    sup_li = "".join(f"<li>{_fmt_name(e)}</li>\n" for e in on_duty if _is_support(e))

    section_b = f'<div style="{hdr_style}">B) CTU Staff On Duty:</div><ul style="{ul_style}">{nil_item}</ul>'
    section_c = f'<div style="{hdr_style}">C) Support Team:</div><ul style="{ul_style}">{sup_li if sup_li else nil_item}</ul>'
    section_d = f'<div style="{hdr_style}">D) Sick Leave / No Show / Others:</div><ul style="{ul_style}">{nil_item}</ul>'
    section_e = f'<div style="{hdr_style}">E) Annual Leave / Course / Off in Lieu:</div><ul style="{ul_style}">{nil_item}</ul>'
    section_f = f'<div style="{hdr_style}">F) Trainee:</div><ul style="{ul_style}">{nil_item}</ul>'
    section_g = f'<div style="{hdr_style}">G) Overtime Justification:</div><ul style="{ul_style}">{nil_item}</ul>'

    return f"""
        <td colspan="2" valign="top" style="{td_style}">
          {grouped_html}
          {section_b}
          {section_c}
          {section_d}
          {section_e}
          {section_f}
          {section_g}
        </td>"""

def build_shift_report(date_dir: str, shift: str) -> None:
    folder = DATA_DIR / date_dir / shift
    if not folder.exists():
        return

    meta         = load_json(folder / "meta.json", {"flights": {}})
    flight_files = sorted(p for p in folder.glob("*.json") if p.name != "meta.json")
    flights      = [json.loads(p.read_text(encoding="utf-8")) for p in flight_files]

    # ── Filter offload: only keep flights whose date matches the report date ──
    from datetime import datetime as _dt_filter
    def _flight_date_matches(flt_date_str: str, report_date: str) -> bool:
        """Check if flight date (e.g. '27FEB', '27FEB25', '2025-02-27') matches report_date (YYYY-MM-DD)."""
        if not flt_date_str or not flt_date_str.strip():
            return True  # no date = don't filter out
        fd = flt_date_str.strip().upper()
        try:
            rd = _dt_filter.strptime(report_date, "%Y-%m-%d")
        except Exception:
            return True
        # Try various formats
        for fmt in ("%d%b%y", "%d%b%Y", "%d%b", "%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y"):
            try:
                parsed = _dt_filter.strptime(fd, fmt)
                # For formats without year, assume report year
                if fmt == "%d%b":
                    parsed = parsed.replace(year=rd.year)
                if parsed.day == rd.day and parsed.month == rd.month and parsed.year == rd.year:
                    return True
                else:
                    return False
            except ValueError:
                continue
        return True  # unknown format = don't filter out

    flights = [f for f in flights if _flight_date_matches(f.get("date", ""), date_dir)]

    shift_labels = {
        "shift1": {"ar": "صباح",     "en": "Morning",   "time": "06:00 – 15:00"},
        "shift2": {"ar": "ظهر/مساء", "en": "Afternoon", "time": "13:00 – 22:00"},
        "shift3": {"ar": "ليل",      "en": "Night",      "time": "21:00 – 06:00"},
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
    roster = fetch_roster_staff(date_dir, shift)

    # ── Supervisor name resolution (on-duty only) ──
    # Only use actual supervisors from roster — no acting/deputy logic.
    # If no supervisor found, leave both display and signature blank.
    on_duty_list = roster.get("on_duty", [])

    supervisor_name = ""

    for emp in on_duty_list:
        if "supervisor" in emp.get("dept", "").lower():
            supervisor_name = emp.get("name", "").strip()
            if supervisor_name:
                break

    # Empty = leave blank (no acting supervisor)
    supervisor_display = supervisor_name if supervisor_name else ""
    signature_display = supervisor_display

    manpower_cols = _render_manpower_section(roster, supervisor_display)

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

<table width="760" cellpadding="0" cellspacing="0" border="0" id="report-content"
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
    <!-- عنوان جدول الأوفلود -->
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:14px;">
      <tr>
        <td width="4" style="background-color:#c2410c;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#fff7ed;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#c2410c; letter-spacing:1px;">OFFLOADING CARGO #</span>
        </td>
      </tr>
    </table>
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

  <!-- ═══ FOOTER / SIGNATURE ═══ -->
  <tr>
    <td style="padding:22px 24px 24px 24px; background-color:#f8faff; border-top:2px solid #0b3a78;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td valign="top">

            <!-- Best Regards line -->
            <div style="font-family:Calibri,Arial,sans-serif; font-size:13px; color:#555; margin-bottom:6px;">
              Best Regards,
            </div>

            <!-- Name + Title -->
            <div style="font-family:Calibri,Arial,sans-serif; margin-bottom:10px;">
              <div style="font-size:15px; font-weight:700; color:#0b3a78; line-height:1.3;">
                {signature_display}
              </div>
              <div style="font-size:12.5px; font-weight:600; color:#444; margin-top:2px;">
                Duty Supervisor – Export Operation
              </div>
            </div>

            <!-- Divider -->
            <div style="border-top:1px solid #d0d8ea; margin-bottom:10px;"></div>

            <!-- Company block -->
            <table cellpadding="0" cellspacing="0" border="0">
              <tr>
                <!-- TRANSOM logo text -->
                <td valign="top" style="padding-right:18px; border-right:2px solid #cc0000;">
                  <div style="font-family:Arial Black,Arial,sans-serif; font-size:22px; font-weight:900;
                              color:#aa0000; letter-spacing:1.5px; line-height:1;">TRANSOM</div>
                  <div style="font-family:Arial,sans-serif; font-size:10px; font-weight:600;
                              color:#888; letter-spacing:3px; margin-top:1px; text-transform:uppercase;">Cargo</div>
                </td>
                <!-- Contact info -->
                <td valign="top" style="padding-left:16px;">
                  <div style="font-family:Calibri,Arial,sans-serif; font-size:13px; font-weight:700;
                              color:#1b1f2a; line-height:1.3; margin-bottom:5px;">
                    Transom Cargo LLC.
                  </div>
                  <div style="font-family:Calibri,Arial,sans-serif; font-size:12px; color:#444; line-height:1.8;">
                    P.O. Box: 618, P.C: 111<br>
                    Sultanate of Oman<br>
                    Phone No. 97297474<br>
                    <a href="http://www.transomcargo.om"
                       style="color:#0b3a78; text-decoration:none; font-weight:600;">
                      www.transomcargo.om
                    </a>
                  </div>
                </td>
              </tr>
            </table>

            <!-- Certifications image -->
            <div style="margin-top:12px;">
              <img src="data:image/png;base64,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ShJ/Jal1sdoAZOHcT7LUWIn+094/KqBPS/DcYPjdQ+o1saW2fit1mJkPn6FWtp/NGvrUxw+XkV2hm19FMETjFreAS9kNzH7ZxPzCzy3/AIfE18k1bmJdHOneNKbdhY5HkSWztL80l9YPzHLsD+AFT5KQlISkAAdgB6VTltwd/Lnl3D1KtRWMt4cs+8+irjph0jsuIylXu4yHb/kjxK3rlL2pSVHz4Ak8f7xJUfnrtVkUpVCSR0hvONVcYxrBRoSlKVot0pSlESsO82u3Xm2vW26wmJsN9PF1l5AUlQ/A1mUoio+/dH77YI0hvp/c4s2zOkrexm/J8eIr1/ZqPdB+Xl37lVVplLFkQsRc7x69Yk8kcOVzhqulvT8ktSW1JkJH9lK1JHyrruvy62282W3W0uIUNFKhsH8qkEhGajMY0XEq+m2NXRCHrHfLTJ8T7LcC/Mrc/wDAlJZWn8C4fxr6roZfVAKZi5CtB8iIUJe/zTNIrqi+dKunN6Wtyfh9pU4v7TjTAZUT89o0ai8n2dul7q+TVsnRh+61Oc1/MmpRNxWhiXPzvRaXCTzuRukVHqqWq3REfmozFkf5TXq3jXTmwPpTcMgsb0nW22ozjt6ecV+6kNpZYSr6L5j8a6ItvQPpZCUlRxv3pQ9ZMp1e/wAuWv5VNsfxTGcfTxsdgtlu/tR4yEKP4kDZrBmQRLn/ABnH8wvi4xxHDHLQwz/UXvKwnxI/zMeGhIbZPyKUEH1NWrgXSay4/dDkN5mScmyVfddzuB5KQf8Au0HYQP4keQOu1WJSojISpAwBKUpWi3SlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURKUpREpSlESlKURf/Z"
                   style="max-width:430px; width:100%; height:auto; display:block;"
                   alt="Certifications">
            </div>

          </td>
          <td align="right" valign="bottom" style="padding-left:10px;">
            <div style="font-family:Calibri,Arial,sans-serif; font-size:10.5px; color:#9aabca; font-style:italic; white-space:nowrap;">
              Generated automatically<br>{date_dir}
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

<!-- ═══ BUTTONS BAR (خارج التقرير) ═══ -->
<div style="max-width:760px; margin:12px 10px 30px; display:flex; gap:10px; justify-content:flex-end; flex-wrap:wrap;">

  <button id="btn-copy"
    type="button"
    style="font-family:Calibri,Arial,sans-serif; font-size:13px; font-weight:700; color:#fff; background:#0b3a78; border:none; border-radius:8px; padding:10px 20px; cursor:pointer; box-shadow:0 2px 8px rgba(11,58,120,.25);">
    📋 Copy Report
  </button>

  <button id="btn-manage-emails"
    type="button"
    style="font-family:Calibri,Arial,sans-serif; font-size:13px; font-weight:700; color:#fff; background:#475569; border:none; border-radius:8px; padding:10px 20px; cursor:pointer; box-shadow:0 2px 8px rgba(71,85,105,.25);">
    📝 Edit Email List
  </button>

  <button id="btn-email"
    type="button"
    style="font-family:Calibri,Arial,sans-serif; font-size:13px; font-weight:700; color:#fff; background:#c2410c; border:none; border-radius:8px; padding:10px 20px; cursor:pointer; box-shadow:0 2px 8px rgba(194,65,12,.25);">
    ✉️ Send Email Now
  </button>

</div>

<script>
(function(){{
  var REPORT_SECRET = '887155';
  var COPY_SUCCESS_MS = 2500;
  var REPO_OWNER = 'khalidsaif912';
  var REPO_NAME = 'offload-monitor';
  var RECIPIENTS_PATH = 'docs/data/email_recipients.json';
  var RECIPIENTS_RAW_URL = 'https://raw.githubusercontent.com/' + REPO_OWNER + '/' + REPO_NAME + '/main/' + RECIPIENTS_PATH;
  var RECIPIENTS_EDIT_URL = 'https://github.com/' + REPO_OWNER + '/' + REPO_NAME + '/edit/main/' + RECIPIENTS_PATH;
  var RECIPIENTS_NEW_URL = 'https://github.com/' + REPO_OWNER + '/' + REPO_NAME + '/new/main?filename=' + encodeURIComponent(RECIPIENTS_PATH);
  var DEFAULT_RECIPIENTS = [];

  function askSendSecret(){{
    var entered = prompt('Enter send password:');
    if(entered === null) return false;
    if((entered || '').trim() !== REPORT_SECRET){{
      alert('Wrong password');
      return false;
    }}
    return true;
  }}

  function resetButton(btn, text, color, delay){{
    setTimeout(function(){{
      btn.innerText = text;
      btn.style.background = color;
      btn.disabled = false;
    }}, delay || 3000);
  }}

  function plainCopyFallback(text, html){{
    if(html){{
      var holder = document.createElement('div');
      holder.innerHTML = html;
      holder.style.position = 'fixed';
      holder.style.left = '-99999px';
      holder.style.top = '0';
      document.body.appendChild(holder);
      var range = document.createRange();
      range.selectNodeContents(holder);
      var sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      try {{
        if(document.execCommand('copy')){{
          sel.removeAllRanges();
          document.body.removeChild(holder);
          return;
        }}
      }} catch (e) {{}}
      sel.removeAllRanges();
      document.body.removeChild(holder);
    }}

    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    try {{ document.execCommand('copy'); }} catch (e) {{}}
    document.body.removeChild(ta);
  }}

  function forceReadableContrast(node){{
    if(!node || node.nodeType !== 1) return;
    var bg = (node.style.backgroundColor || '').toLowerCase().replace(/\s+/g, '');
    var darkBgs = ['rgb(11,58,120)','#0b3a78','rgb(30,64,175)','#1e40af','rgb(37,99,235)','#2563eb','rgb(10,31,82)','#0a1f52'];
    if(darkBgs.indexOf(bg) !== -1){{
      node.style.backgroundColor = '#eef3fc';
      node.style.color = '#0b3a78';
    }}
    var clr = (node.style.color || '').toLowerCase().replace(/\s+/g, '');
    if(clr === '#ffffff' || clr === '#fff' || clr === 'rgb(255,255,255)'){{
      node.style.color = '#0b3a78';
    }}
  }}

  function inlineComputedStyles(source, target){{
    if(!source || !target || source.nodeType !== 1 || target.nodeType !== 1) return;
    var cs = window.getComputedStyle(source);
    var props = [
      'display','width','min-width','max-width','height','min-height','max-height',
      'box-sizing','margin','margin-top','margin-right','margin-bottom','margin-left',
      'padding','padding-top','padding-right','padding-bottom','padding-left',
      'border','border-top','border-right','border-bottom','border-left',
      'border-collapse','border-spacing','table-layout','vertical-align',
      'background','background-color','color','font','font-family','font-size','font-weight',
      'font-style','line-height','letter-spacing','text-align','text-transform','text-decoration',
      'white-space','word-break','overflow-wrap','opacity','border-radius','box-shadow'
    ];
    props.forEach(function(prop){{
      var val = cs.getPropertyValue(prop);
      if(val) target.style.setProperty(prop, val);
    }});
    forceReadableContrast(target);

    if(target.tagName === 'TABLE'){{
      target.setAttribute('cellpadding', target.getAttribute('cellpadding') || '0');
      target.setAttribute('cellspacing', target.getAttribute('cellspacing') || '0');
      target.setAttribute('border', target.getAttribute('border') || '0');
      target.setAttribute('role', 'presentation');
    }}

    var sChildren = source.children || [];
    var tChildren = target.children || [];
    for(var i = 0; i < sChildren.length; i++){{
      if(tChildren[i]) inlineComputedStyles(sChildren[i], tChildren[i]);
    }}
  }}

  function buildReportHtml(){{
    var el = document.getElementById('report-content');
    if(!el) return null;
    var clone = el.cloneNode(true);
    inlineComputedStyles(el, clone);

    /* ── Outlook fix: convert dark backgrounds to light ── */
    var darkBgs = ['rgb(11, 58, 120)','rgb(11,58,120)','#0b3a78',
                   'rgb(30, 64, 175)','rgb(30,64,175)','#1e40af',
                   'rgb(37, 99, 235)','rgb(37,99,235)','#2563eb',
                   'rgb(10, 31, 82)','rgb(10,31,82)','#0a1f52'];
    function isDarkBg(bg){{
      bg = bg.replace(/\s+/g,'');
      for(var i=0;i<darkBgs.length;i++){{ if(bg===darkBgs[i].replace(/\s+/g,'')) return true; }}
      return false;
    }}
    clone.querySelectorAll('td,th,tr,div,span,strong').forEach(function(el2){{
      var bg2 = (el2.style.backgroundColor || '').toLowerCase();
      if(isDarkBg(bg2)){{
        el2.style.backgroundColor = '#eef3fc';
        el2.style.color = '#0b3a78';
        el2.querySelectorAll('*').forEach(function(child){{
          var c = (child.style.color || '').toLowerCase().replace(/\s+/g,'');
          if(c === '#ffffff' || c === '#fff' || c === 'rgb(255,255,255)'){{
            child.style.color = '#0b3a78';
          }}
        }});
      }}
      /* Remove linear-gradient from background */
      var bgFull = (el2.style.background || '').toLowerCase();
      if(bgFull.indexOf('linear-gradient') !== -1){{
        el2.style.background = '';
        el2.style.backgroundColor = '#eef3fc';
        el2.style.color = '#0b3a78';
      }}
      /* Fix white text globally */
      var clr2 = (el2.style.color || '').toLowerCase().replace(/\s+/g,'');
      if(clr2 === '#ffffff' || clr2 === '#fff' || clr2 === 'rgb(255,255,255)'){{
        el2.style.color = '#0b3a78';
      }}
    }});

    clone.querySelectorAll('table').forEach(function(tbl){{
      tbl.style.borderCollapse = 'collapse';
      if(!tbl.style.tableLayout) tbl.style.tableLayout = 'fixed';
      if(!tbl.style.width) tbl.style.width = '100%';
    }});
    clone.querySelectorAll('td,th').forEach(function(cell){{
      if(!cell.style.verticalAlign) cell.style.verticalAlign = 'top';
    }});
    return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><style>body{{background:#eef1f7;font-family:Calibri,Arial,sans-serif;margin:0;padding:10px 0;-webkit-text-size-adjust:100%;}}table{{border-collapse:collapse;table-layout:fixed;}}img{{max-width:100%;height:auto;display:block;}}</style></head><body>' + clone.outerHTML + '</body></html>';
  }}

  function onCopyReport(){{
    var btn = document.getElementById('btn-copy');
    var el = document.getElementById('report-content');
    if(!btn || !el) return;

    var full = buildReportHtml();
    if(!full) return;

    function markCopied(){{
      btn.innerText = '✅ Copied!';
      btn.style.background = '#059669';
      resetButton(btn, '📋 Copy Report', '#0b3a78', COPY_SUCCESS_MS);
    }}

    if(navigator.clipboard && window.ClipboardItem){{
      var item = new ClipboardItem({{
        'text/html': new Blob([full], {{type:'text/html'}}),
        'text/plain': new Blob([el.innerText], {{type:'text/plain'}})
      }});
      navigator.clipboard.write([item]).then(markCopied).catch(function(){{
        if(navigator.clipboard && navigator.clipboard.writeText){{
          navigator.clipboard.writeText(el.innerText).then(markCopied).catch(function(){{
            plainCopyFallback(el.innerText, full);
            markCopied();
          }});
        }} else {{
          plainCopyFallback(el.innerText, full);
          markCopied();
        }}
      }});
    }} else {{
      plainCopyFallback(el.innerText, full);
      markCopied();
    }}
  }}

  function normalizeRecipients(items){{
    var seen = {{}};
    var out = [];
    (items || []).forEach(function(v){{
      var s = String(v || '').trim().toLowerCase();
      if(!s) return;
      if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) return;
      if(seen[s]) return;
      seen[s] = true;
      out.push(s);
    }});
    return out;
  }}

  function getLocalRecipients(){{
    try {{
      return normalizeRecipients(JSON.parse(localStorage.getItem('saved_recipients') || '[]'));
    }} catch(e) {{
      return [];
    }}
  }}

  function setLocalRecipients(items){{
    localStorage.setItem('saved_recipients', JSON.stringify(normalizeRecipients(items)));
  }}

  function loadRecipients(){{
    var fallback = getLocalRecipients();
    if(fallback.length) return Promise.resolve(fallback);
    return fetch(RECIPIENTS_RAW_URL + '?t=' + Date.now(), {{cache:'no-store'}})
      .then(function(r){{
        if(!r.ok) throw new Error('missing');
        return r.json();
      }})
      .then(function(data){{
        var arr = normalizeRecipients((data && data.recipients) || DEFAULT_RECIPIENTS);
        if(arr.length) setLocalRecipients(arr);
        return arr;
      }})
      .catch(function(){{
        return normalizeRecipients(DEFAULT_RECIPIENTS);
      }});
  }}

  function renderRecipientList(container, recipients){{
    if(!container) return;
    if(!recipients.length){{
      container.innerHTML = '<div style="padding:10px;border:1px dashed #cbd5e1;border-radius:8px;color:#64748b;">No saved recipients yet.</div>';
      return;
    }}
    container.innerHTML = recipients.map(function(email, idx){{
      return '<label style="display:flex;align-items:center;justify-content:space-between;gap:10px;padding:8px 10px;border:1px solid #e2e8f0;border-radius:8px;margin-bottom:8px;background:#fff;">'
        + '<span style="display:flex;align-items:center;gap:10px;flex:1;min-width:0;">'
        + '<input type="checkbox" data-email="' + email.replace(/"/g,'&quot;') + '" checked>'
        + '<span style="word-break:break-all;">' + email + '</span>'
        + '</span>'
        + '<button type="button" data-del="' + idx + '" style="border:none;background:#dc2626;color:#fff;border-radius:6px;padding:6px 10px;cursor:pointer;font-weight:700;">Delete</button>'
        + '</label>';
    }}).join('');
  }}

  function openRecipientsManager(mode){{
    return loadRecipients().then(function(initial){{
      var recipients = initial.slice();
      return new Promise(function(resolve){{
        var backdrop = document.createElement('div');
        backdrop.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:9999;display:flex;align-items:center;justify-content:center;padding:16px;';
        var box = document.createElement('div');
        box.style.cssText = 'width:min(560px,100%);max-height:90vh;overflow:auto;background:#f8fafc;border-radius:16px;box-shadow:0 20px 60px rgba(15,23,42,.35);padding:18px;font-family:Calibri,Arial,sans-serif;';
        backdrop.appendChild(box);
        box.innerHTML = ''
          + '<div style="display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:12px;">'
          + '<div><div style="font-size:22px;font-weight:800;color:#0f172a;">Email Recipients</div><div style="font-size:13px;color:#64748b;">Choose all, some, or add a new email.</div></div>'
          + '<button type="button" id="picker-close" style="border:none;background:#e2e8f0;border-radius:8px;padding:8px 12px;cursor:pointer;font-weight:700;">Close</button>'
          + '</div>'
          + '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;">'
          + '<button type="button" id="picker-all" style="border:none;background:#0b3a78;color:#fff;border-radius:8px;padding:8px 12px;cursor:pointer;font-weight:700;">Select All</button>'
          + '<button type="button" id="picker-none" style="border:none;background:#64748b;color:#fff;border-radius:8px;padding:8px 12px;cursor:pointer;font-weight:700;">Unselect All</button>'
          + '<button type="button" id="picker-open-gh" style="border:none;background:#475569;color:#fff;border-radius:8px;padding:8px 12px;cursor:pointer;font-weight:700;">Open on GitHub</button>'
          + '</div>'
          + '<div style="display:flex;gap:8px;margin-bottom:12px;">'
          + '<input id="picker-new-email" type="email" placeholder="new@example.com" style="flex:1;min-width:0;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;">'
          + '<button type="button" id="picker-add" style="border:none;background:#059669;color:#fff;border-radius:8px;padding:10px 14px;cursor:pointer;font-weight:700;">Add</button>'
          + '</div>'
          + '<div id="picker-list"></div>'
          + '<div style="display:flex;justify-content:flex-end;gap:8px;margin-top:14px;">'
          + (mode === 'send' ? '<button type="button" id="picker-send" style="border:none;background:#c2410c;color:#fff;border-radius:8px;padding:10px 16px;cursor:pointer;font-weight:800;">Send Now</button>' : '')
          + '<button type="button" id="picker-save-only" style="border:none;background:#0b3a78;color:#fff;border-radius:8px;padding:10px 16px;cursor:pointer;font-weight:800;">Save</button>'
          + '</div>';
        document.body.appendChild(backdrop);
        var list = box.querySelector('#picker-list');
        function rerender(){{ renderRecipientList(list, recipients); }}
        function close(result){{ document.body.removeChild(backdrop); resolve(result || null); }}
        rerender();
        box.querySelector('#picker-close').onclick = function(){{ close(null); }};
        box.querySelector('#picker-all').onclick = function(){{ list.querySelectorAll('input[type="checkbox"]').forEach(function(cb){{ cb.checked = true; }}); }};
        box.querySelector('#picker-none').onclick = function(){{ list.querySelectorAll('input[type="checkbox"]').forEach(function(cb){{ cb.checked = false; }}); }};
        box.querySelector('#picker-open-gh').onclick = function(){{ window.open(RECIPIENTS_EDIT_URL, '_blank'); }};
        box.querySelector('#picker-add').onclick = function(){{
          var input = box.querySelector('#picker-new-email');
          var val = String(input.value || '').trim().toLowerCase();
          if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(val)){{ alert('Enter a valid email address'); return; }}
          recipients = normalizeRecipients(recipients.concat([val]));
          input.value = '';
          rerender();
        }};
        list.addEventListener('click', function(ev){{
          var idx = ev.target && ev.target.getAttribute('data-del');
          if(idx === null) return;
          recipients.splice(Number(idx), 1);
          rerender();
        }});
        box.querySelector('#picker-save-only').onclick = function(){{
          var selected = normalizeRecipients(recipients);
          setLocalRecipients(selected);
          alert('Saved locally in this browser. To save on GitHub, use Open on GitHub.');
          close({{selected:selected, saved:true}});
        }};
        if(mode === 'send'){{
          box.querySelector('#picker-send').onclick = function(){{
            var selected = Array.from(list.querySelectorAll('input[type="checkbox"]:checked')).map(function(cb){{ return cb.getAttribute('data-email'); }});
            selected = normalizeRecipients(selected);
            if(!selected.length){{ alert('Choose at least one recipient'); return; }}
            setLocalRecipients(normalizeRecipients(recipients));
            close({{selected:selected, saved:true}});
          }};
        }}
      }});
    }});
  }}

  function openRecipientsFile(){{
    if(!askSendSecret()) return;
    loadRecipients().then(function(items){{
      if(items.length){{
        window.open(RECIPIENTS_EDIT_URL, '_blank');
      }} else {{
        var template = `{{
  "recipients": [
    "example1@company.com",
    "example2@company.com"
  ],
  "updated_at": "",
  "source": "offload-monitor-ui"
}}`;
        var url = RECIPIENTS_NEW_URL + '&value=' + encodeURIComponent(template);
        window.open(url, '_blank');
      }}
    }});
  }}

  function sendEmailNow(){{
    if(!askSendSecret()) return;
    openRecipientsManager('send').then(function(result){{
      if(!result || !result.selected || !result.selected.length) return;
      var ok = confirm('إرسال تقرير هذه المناوبة بالإيميل الآن؟\\n\\nShift: {shift}\\nDate: {date_dir}\\nRecipients: ' + result.selected.join(', '));
      if(!ok) return;

      var btn = document.getElementById('btn-email');
      if(!btn) return;
      btn.innerText = '⏳ Sending…';
      btn.disabled = true;
      btn.style.background = '#64748b';

      var pat = localStorage.getItem('gh_pat');
      if(!pat){{
        pat = prompt('أدخل GitHub Personal Access Token (repo scope):');
        if(!pat){{
          btn.innerText = '✉️ Send Email Now';
          btn.style.background = '#c2410c';
          btn.disabled = false;
          return;
        }}
        localStorage.setItem('gh_pat', pat);
      }}

      fetch('https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/dispatches', {{
        method:'POST',
        headers:{{
          'Accept':'application/vnd.github+json',
          'Authorization':'Bearer ' + pat,
          'Content-Type':'application/json'
        }},
        body:JSON.stringify({{
          event_type:'send_report_now',
          client_payload:{{date_dir:'{date_dir}', shift:'{shift}', recipients: result.selected}}
        }})
      }}).then(function(r){{
        if(r.ok || r.status === 204){{
          btn.innerText = '✅ Email Sent!';
          btn.style.background = '#059669';
          resetButton(btn, '✉️ Send Email Now', '#c2410c', 4000);
        }} else {{
          if(r.status === 401){{ localStorage.removeItem('gh_pat'); }}
          btn.innerText = '❌ Failed (' + r.status + ')';
          btn.style.background = '#dc2626';
          resetButton(btn, '✉️ Send Email Now', '#c2410c', 3000);
        }}
      }}).catch(function(){{
        btn.innerText = '❌ Error';
        btn.style.background = '#dc2626';
        resetButton(btn, '✉️ Send Email Now', '#c2410c', 3000);
      }});
    }});
  }}

  var copyBtn = document.getElementById('btn-copy');
  var manageBtn = document.getElementById('btn-manage-emails');
  var emailBtn = document.getElementById('btn-email');
  if(copyBtn) copyBtn.addEventListener('click', onCopyReport);
  if(manageBtn) manageBtn.addEventListener('click', openRecipientsFile);
  if(emailBtn) emailBtn.addEventListener('click', sendEmailNow);
}})();
</script>

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


def build_nil_shift_report(date_dir: str, shift: str, now) -> None:
    """Build NIL shift report using build_shift_report with empty data.
    Always rebuilds to pick up latest template and roster changes."""
    # إذا كان لهذه المناوبة بيانات حقيقية — لا تستبدلها
    data_folder = DATA_DIR / date_dir / shift
    data_folder.mkdir(parents=True, exist_ok=True)
    meta_file = data_folder / "meta.json"
    if not meta_file.exists():
        meta_file.write_text("{}", encoding="utf-8")
    # إذا يوجد رحلات حقيقية — اترك build_shift_report يتعامل معها
    real_flights = [p for p in data_folder.glob("*.json") if p.name != "meta.json"]
    if real_flights:
        return
    # ابنِ تقرير NIL باستخدام نفس build_shift_report
    build_shift_report(date_dir, shift)
    print(f"  [NIL report] Built: {date_dir}/{shift}")

def build_root_index(now: datetime) -> None:
    """Modern home page with accordion days; current day opened by default.
    Always shows all days of current month even with no offload data."""
    import calendar as _cal
    DOCS_DIR.mkdir(parents=True, exist_ok=True)

    today        = now.strftime("%Y-%m-%d")
    last_updated = now.strftime("%Y-%m-%d %H:%M")

    # أيام الشهر الحالي حتى اليوم فقط (لا أيام مستقبلية)
    _days_in_month = _cal.monthrange(now.year, now.month)[1]
    _month_days = {
        f"{now.year:04d}-{now.month:02d}-{d:02d}"
        for d in range(1, now.day + 1)  # من أول الشهر حتى اليوم فقط
    }
    # أيام محفوظة من أشهر أخرى (أقدم من الشهر الحالي)
    _saved_days: set = set()
    if DATA_DIR.exists():
        for _p in DATA_DIR.iterdir():
            if _p.is_dir() and re.match(r"\d{4}-\d{2}-\d{2}", _p.name):
                if _p.name < f"{now.year:04d}-{now.month:02d}-01":
                    _saved_days.add(_p.name)

    day_dirs = sorted(_month_days | _saved_days, reverse=True)

    shift_meta = {
        "shift1": {"label": "Morning",   "ar": "صباح", "time": "06:00 – 15:00", "icon": "🌅"},
        "shift2": {"label": "Afternoon", "ar": "ظهر",  "time": "15:00 – 22:00", "icon": "☀️"},
        "shift3": {"label": "Night",     "ar": "ليل",  "time": "22:00 – 06:00", "icon": "🌙"},
    }

    total_days = len(day_dirs)

    # بناء/تحديث تقارير لكل الأيام والمناوبات
    for day in day_dirs:
        for shift in ("shift1", "shift2", "shift3"):
            # احذف التقرير القديم إذا لم تكن هناك رحلات حقيقية (لإجبار إعادة البناء)
            real = list((DATA_DIR / day / shift).glob("*.json")) if (DATA_DIR / day / shift).exists() else []
            real = [p for p in real if p.name != "meta.json"]
            if not real:
                old_report = DOCS_DIR / day / shift / "index.html"
                if old_report.exists():
                    old_report.unlink()
            build_nil_shift_report(day, shift, now)

    days_html = ""
    for day in day_dirs:
        is_today   = day == today
        is_future  = day > today
        open_attr  = " open" if is_today else ""

        # عد الرحلات
        day_flights = 0
        for shift in ("shift1", "shift2", "shift3"):
            shift_folder = DATA_DIR / day / shift
            if shift_folder.exists():
                day_flights += sum(1 for f in shift_folder.glob("*.json") if f.name != "meta.json")

        badge        = '<span class="today-badge">TODAY</span>' if is_today else ('<span class="today-badge" style="background:#64748b;">UPCOMING</span>' if is_future else "")
        flights_pill = f'<span class="day-pill">{day_flights} flights</span>' if day_flights else ""

        rows = ""
        for shift in ("shift1", "shift2", "shift3"):
            shift_report = DOCS_DIR / day / shift / "index.html"
            meta_s = shift_meta.get(shift, {"label": shift, "ar": shift, "time": "", "icon": "✈"})
            shift_flt_count = 0
            shift_folder = DATA_DIR / day / shift
            if shift_folder.exists():
                shift_flt_count = sum(1 for f in shift_folder.glob("*.json") if f.name != "meta.json")
            flt_txt = f"{shift_flt_count} flight{'s' if shift_flt_count != 1 else ''}" if shift_flt_count else ""

            if shift_report.exists():
                # مناوبة فيها تقرير — رابط
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
            else:
                # مناوبة NIL — رابط قابل للضغط
                rows += f"""
            <a class="shift-card" href="{day}/{shift}/">
                <div class="sc-icon">{meta_s['icon']}</div>
                <div class="sc-body">
                    <div class="sc-title">{meta_s['ar']} <span class="sc-en">/ {meta_s['label']}</span></div>
                    <div class="sc-time">{meta_s['time']}</div>
                </div>
                <div class="sc-right">
                    <span style="font-size:11px;color:#94a3b8;font-weight:600;">NIL</span>
                    <span class="sc-arrow">›</span>
                </div>
            </a>"""

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
    <title>Export Warehouse Activity Report</title>
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
                <h1>✈ Export Warehouse Activity Report</h1>
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
        info = fetch_flight_info_airlabs(flt, flight_date=flight_date_iso, dep_iata="MCT")

        if not info:
            skipped_count += 1
            continue

        changed = False

        # Update destination — always overwrite with AirLabs data
        new_dest = (info.get("dest") or "").strip()
        old_dest = (flight.get("destination") or "").strip()
        if new_dest and new_dest != old_dest:
            print(f"  [retro] {flt}/{date}: dest {old_dest!r} → {new_dest!r}")
            flight["destination"] = new_dest
            changed = True
        elif not new_dest:
            print(f"  [retro] {flt}/{date}: AirLabs returned no dest")

        # Update STD/ETD only if missing
        std = (info.get("std") or "").strip()
        etd = (info.get("etd") or "").strip()
        new_std_etd = ""
        if std and etd and std != etd:
            new_std_etd = f"{std}|{etd}"
        elif std:
            new_std_etd = std

        if new_std_etd and new_std_etd != (flight.get("std_etd") or "").strip():
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
#  إرسال التقرير بالبريد الإلكتروني
# ══════════════════════════════════════════════════════════════════

def _extract_report_content_html(page_html: str) -> str:
    """Return only the main report container without action buttons/scripts."""
    soup = BeautifulSoup(page_html, "html.parser")
    report = soup.find(id="report-content")
    if report:
        return str(report)
    body = soup.body
    return str(body) if body else page_html


def _build_email_html(page_html: str) -> str:
    """Build a mobile-friendly HTML email without page buttons or scripts."""
    report_html = _extract_report_content_html(page_html)
    report_html = report_html.replace('width="760"', 'width="100%"')
    report_html = report_html.replace('style="width:760px; max-width:760px; background-color:#ffffff; border:1px solid #d0d5e8;"',
                                      'style="width:100%; max-width:760px; background-color:#ffffff; border:1px solid #d0d5e8; margin:0 auto;"')

    return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    html, body {{
      margin: 0 !important;
      padding: 0 !important;
      width: 100% !important;
      background: #eef1f7 !important;
      font-family: Calibri, Arial, sans-serif !important;
      -webkit-text-size-adjust: 100%;
      -ms-text-size-adjust: 100%;
    }}
    body, table, td, div, p, a, li {{
      font-size: 15px !important;
      line-height: 1.55 !important;
    }}
    table {{ border-collapse: collapse; }}
    img {{ border: 0; display: block; max-width: 100%; height: auto; }}
    .mobile-wrap {{ width: 100%; padding: 10px 8px 18px; box-sizing: border-box; }}
    @media only screen and (max-width: 640px) {{
      body, table, td, div, p, a, li {{
        font-size: 16px !important;
        line-height: 1.65 !important;
      }}
      .mobile-wrap {{ padding: 6px 4px 14px !important; }}
      table[width="100%"] {{ width: 100% !important; }}
      td[style*="font-size:20px"] div {{ font-size: 18px !important; }}
      td[style*="font-size:13.5px"], div[style*="font-size:13.5px"], span[style*="font-size:12px"], td[style*="font-size:12px"] {{
        font-size: 15px !important;
      }}
    }}
  </style>
</head>
<body>
  <div class="mobile-wrap">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="width:100%; background:#eef1f7;">
      <tr>
        <td align="center">
          {report_html}
        </td>
      </tr>
    </table>
  </div>
</body>
</html>"""


def send_shift_report_email(date_dir: str, shift: str) -> None:
    """إرسال تقرير المناوبة بالبريد الإلكتروني كـ HTML كامل."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    smtp_user      = os.environ.get("EMAIL_SENDER", "").strip()
    smtp_password  = os.environ.get("EMAIL_APP_PASSWORD", "").strip()
    recipients_raw = os.environ.get("EMAIL_RECIPIENTS", "").strip()
    recipients = []

    event_path = os.environ.get("GITHUB_EVENT_PATH", "").strip()
    if event_path and Path(event_path).exists():
        try:
            event_data = json.loads(Path(event_path).read_text(encoding="utf-8"))
            payload = event_data.get("client_payload") or {}
            payload_recipients = payload.get("recipients") or []
            if isinstance(payload_recipients, list):
                recipients = [str(r).strip() for r in payload_recipients if str(r).strip()]
        except Exception as exc:
            print(f"  [email] Could not read recipients from event payload: {exc}")

    if not recipients:
        recipients = [r.strip() for r in recipients_raw.split(",") if r.strip()]

    if not smtp_user or not smtp_password or not recipients:
        print("  [email] Skipping — EMAIL_SENDER / EMAIL_APP_PASSWORD / recipients not set.")
        return

    report_file = DOCS_DIR / date_dir / shift / "index.html"
    if not report_file.exists():
        print(f"  [email] Report not found: {report_file}")
        return

    page_html = report_file.read_text(encoding="utf-8")
    html_content = _build_email_html(page_html)

    shift_names = {
        "shift1": "Morning Shift (06:00–15:00)",
        "shift2": "Afternoon Shift (15:00–22:00)",
        "shift3": "Night Shift (22:00–06:00)",
    }
    subject = f"Export Warehouse Activity Report — {date_dir} | {shift_names.get(shift, shift)}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = smtp_user
    msg["To"]      = ", ".join(recipients)
    plain_text = f"""Export Warehouse Activity Report
Date: {date_dir}
Shift: {shift_names.get(shift, shift)}
"""
    msg.attach(MIMEText(plain_text, "plain", "utf-8"))
    msg.attach(MIMEText(html_content, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(smtp_user, smtp_password)
            server.sendmail(smtp_user, recipients, msg.as_string())
        print(f"  [email] Sent: {subject} → {recipients}")
    except Exception as exc:
        print(f"  [email] ERROR: {exc}")


EMAIL_SENT_DIR: Path = Path("data/.email_sent")

def _email_sent_key(date_dir: str, shift: str) -> Path:
    """مسار ملف الحالة لمعرفة إذا تم إرسال الإيميل لهذه المناوبة."""
    EMAIL_SENT_DIR.mkdir(parents=True, exist_ok=True)
    return EMAIL_SENT_DIR / f"{date_dir}_{shift}.sent"


def should_send_email(now, shift: str) -> bool:
    """نافذة الإرسال: من ساعة قبل نهاية المناوبة حتى نهايتها.

    shift1: نافذة 13:00 – 15:00 (المناوبة تنتهي 15:00)
    shift2: نافذة 20:00 – 22:00 (المناوبة تنتهي 22:00)
    shift3: نافذة 04:00 – 06:00 (المناوبة تنتهي 06:00)
    """
    # (start_h, start_m, end_h, end_m)
    windows = {
        "shift1": (13, 0, 15, 0),
        "shift2": (20, 0, 22, 0),
        "shift3": (4,  0,  6, 0),
    }
    w = windows.get(shift)
    if not w:
        return False

    current = now.hour * 60 + now.minute
    start_m = w[0] * 60 + w[1]
    end_m   = w[2] * 60 + w[3]

    # shift3 يعبر منتصف الليل — لا يعبر هنا لأن النافذة 04:00–06:00
    return start_m <= current <= end_m


def maybe_send_email(now, date_dir: str, shift: str) -> None:
    """أرسل الإيميل مرة واحدة فقط خلال نافذة الإرسال (يمنع التكرار)."""
    if not os.environ.get("EMAIL_SENDER", "").strip():
        return
    if not should_send_email(now, shift):
        return
    sent_file = _email_sent_key(date_dir, shift)
    if sent_file.exists():
        print(f"  [email] Already sent for {date_dir}/{shift} — skipping.")
        return
    print(f"  [email] Sending {shift} report for {date_dir}…")
    send_shift_report_email(date_dir, shift)
    sent_file.write_text(now.isoformat(), encoding="utf-8")


# ══════════════════════════════════════════════════════════════════
#  فلترة الرحلات القديمة بناءً على تاريخ الإيميل والمناوبة
# ══════════════════════════════════════════════════════════════════

def _shift_window_for(ref_dt: datetime) -> tuple[datetime, datetime]:
    """إرجاع (start, end) للمناوبة التي تنتمي إليها ref_dt بالتوقيت المحلي.

    المناوبات:
      shift1 : 06:00 – 14:30
      shift2 : 14:30 – 21:30
      shift3 : 21:30 – 06:00 (اليوم التالي)
    """
    tz  = ZoneInfo(TIMEZONE)
    loc = ref_dt.astimezone(tz)
    mins = loc.hour * 60 + loc.minute

    base = loc.replace(hour=0, minute=0, second=0, microsecond=0)

    if 6 * 60 <= mins < 14 * 60 + 30:          # shift1
        start = base.replace(hour=6)
        end   = base.replace(hour=14, minute=30)
    elif 14 * 60 + 30 <= mins < 21 * 60 + 30:  # shift2
        start = base.replace(hour=14, minute=30)
        end   = base.replace(hour=21, minute=30)
    else:                                        # shift3 يعبر منتصف الليل
        if loc.hour < 6:
            # بعد منتصف الليل — المناوبة بدأت أمس
            start = (base - timedelta(days=1)).replace(hour=21, minute=30)
        else:
            start = base.replace(hour=21, minute=30)
        end = (start + timedelta(hours=8, minutes=30))

    return start, end


def filter_flights_by_shift(flights: list[dict], now: datetime) -> list[dict]:
    """احتفظ فقط بالرحلات التي تاريخها يقع داخل نافذة مناوبة وقت الإيميل (now).

    المنطق:
    - نحوّل تاريخ الرحلة (مثل '27FEB' أو '2025-02-27') إلى ISO.
    - نقارنه بتاريخ نافذة المناوبة الحالية.
    - إذا كان تاريخ الرحلة **قبل** تاريخ بداية المناوبة → قديم → نتجاهله.
    - إذا كان بدون تاريخ أو لم يُوزَّع → نُبقيه (لا نحذفه).
    """
    shift_start, shift_end = _shift_window_for(now)
    tz = ZoneInfo(TIMEZONE)

    kept    = []
    skipped = []

    for f in flights:
        raw_date = (f.get("date") or "").strip()
        if not raw_date:
            kept.append(f)
            continue

        iso = normalize_flight_date(raw_date, now)
        if not iso:
            kept.append(f)
            continue

        try:
            flight_date = datetime.fromisoformat(iso).replace(tzinfo=tz)
        except ValueError:
            kept.append(f)
            continue

        # الرحلة تنتمي لنفس يوم المناوبة أو أحدث → احتفظ بها
        # الرحلة أقدم من بداية المناوبة بأكثر من يوم كامل → قديمة
        age_hours = (shift_start - flight_date).total_seconds() / 3600

        if age_hours > 24:
            skipped.append(f.get("flight","?") + "/" + raw_date)
        else:
            kept.append(f)

    if skipped:
        print(f"  [filter] Skipped {len(skipped)} old flight(s): {', '.join(skipped)}")
    print(f"  [filter] Kept {len(kept)} flight(s) for this shift window.")
    return kept

def main() -> None:
    now = datetime.now(ZoneInfo(TIMEZONE))
    print(f"[{now.isoformat()}] Starting…")

    # ── وضع الإرسال الفوري (triggered من زر في الصفحة) ──
    force_date  = os.getenv("FORCE_SEND_DATE",  "").strip()
    force_shift = os.getenv("FORCE_SEND_SHIFT", "").strip()
    print(f"[debug] FORCE_SEND_DATE={force_date!r}  FORCE_SEND_SHIFT={force_shift!r}")
    if force_date and force_shift:
        print(f"[force-send] Sending {force_shift} report for {force_date}…")
        report_file = DOCS_DIR / force_date / force_shift / "index.html"
        if not report_file.exists():
            print(f"[force-send] Report not found: {report_file} — rebuilding first…")
            build_shift_report(force_date, force_shift)
        # احذف ملف .sent حتى يُرسل حتى لو أُرسل مسبقاً
        sent_file = _email_sent_key(force_date, force_shift)
        if sent_file.exists():
            sent_file.unlink()
        send_shift_report_email(force_date, force_shift)
        sent_file.write_text(now.isoformat(), encoding="utf-8")
        print("[force-send] Done. ✓")
        return

    # ── وضع الإثراء الرجعي ──
    if os.getenv("RETRO_ENRICH", "").strip().lower() in ("1", "true", "yes", "y"):
        print("RETRO_ENRICH=1 detected. Running retroactive enrichment for all saved flights…")
        retroactive_enrich_all(now)
        print("RETRO_ENRICH done — rebuilding root index and NIL reports…")
        build_root_index(now)
        return

    # ── وضع إعادة بناء جميع التقارير (بدون إثراء) ──
    if os.getenv("REBUILD_ALL", "").strip().lower() in ("1", "true", "yes", "y"):
        print("REBUILD_ALL=1 detected. Rebuilding ALL shift reports with latest template…")
        rebuilt = 0
        for date_dir_p in sorted(p for p in DATA_DIR.iterdir() if p.is_dir()):
            for _s in ("shift1", "shift2", "shift3"):
                if (date_dir_p / _s).exists():
                    build_shift_report(date_dir_p.name, _s)
                    print(f"  rebuilt: {date_dir_p.name}/{_s}")
                    rebuilt += 1
        build_root_index(now)
        print(f"REBUILD_ALL done. {rebuilt} reports rebuilt. ✓")
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
            print("No change detected — building NIL reports and root index…")
            build_root_index(now)
            STATE_FILE.write_text(new_hash, encoding="utf-8")
            # ── إرسال البريد قبل نهاية المناوبة (حتى لو لا يوجد تغيير) ──
            today_str = now.strftime("%Y-%m-%d")
            for _shift in ("shift1", "shift2", "shift3"):
                maybe_send_email(now, today_str, _shift)
            return
        if old_hash == new_hash and FORCE_REBUILD:
            print("No change detected, but FORCE_REBUILD=1 → continuing to rebuild.")

    print("Change detected. Parsing…")
    flights = extract_flights(html)

    if not flights:
        print("WARNING: No flights extracted. Check HTML structure.")
        STATE_FILE.write_text(new_hash, encoding="utf-8")
        return

    # ── فلترة الرحلات القديمة بناءً على تاريخ الإيميل والمناوبة ──
    flights = filter_flights_by_shift(flights, now)

    if not flights:
        print("WARNING: All flights filtered out as old/stale — no data for this shift.")
        STATE_FILE.write_text(new_hash, encoding="utf-8")
        build_root_index(now)
        return

    # ── Enrich with AirLabs ──
    if os.environ.get("AIRLABS_API_KEY", "").strip():
        enriched = 0
        for f in flights:
            flt = (f.get("flight") or "").strip()
            if not flt:
                continue
            info = fetch_flight_info_airlabs(flt, flight_date=normalize_flight_date(f.get('date',''), now), dep_iata="MCT")
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
    ensure_email_recipients_file()
    build_shift_report(date_dir, shift)
    build_root_index(now)

    STATE_FILE.write_text(new_hash, encoding="utf-8")
    print(f"Done. ✓  ({len(flights)} flights saved)")

    # ── إرسال البريد قبل نهاية المناوبة ──
    today_str = now.strftime("%Y-%m-%d")
    for _shift in ("shift1", "shift2", "shift3"):
        maybe_send_email(now, today_str, _shift)


if __name__ == "__main__":
    main()

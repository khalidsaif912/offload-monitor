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
import calendar as _cal
from collections import OrderedDict
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
import time
import random
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup


# ══════════════════════════════════════════════════════════════════
#  الإعدادات العامة
# ══════════════════════════════════════════════════════════════════

# ── Confidence scores per source (0-100) ──
SOURCE_CONFIDENCE = {
    "airlabs_schedules": 95,
    "airlabs_flights": 85,
    "local_db": 75,          # ← mct_flights.json: أعلى من Flightradar وأقل من AirLabs
    "flightradar24": 60,
    "muscat_airport": 20,
    "manual_override": 100,
}



# ══════════════════════════════════════════════════════════════════
#  HTTP Session مع headers واقعية وRetry تلقائي
# ══════════════════════════════════════════════════════════════════

_REALISTIC_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "DNT": "1",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Cache-Control": "max-age=0",
}

# ── Rate-limit tracker ──
_last_request_time: dict[str, float] = {}
_MIN_DELAY_SECONDS = 3.0  # الحد الأدنى بين طلبين لنفس الموقع

def _rate_limited_get(url: str, **kwargs) -> requests.Response:
    """GET مع rate limiting تلقائي حسب الدومين — يستخدم الـ Session المشتركة."""
    from urllib.parse import urlparse
    domain = urlparse(url).netloc
    now_ts = time.time()
    last = _last_request_time.get(domain, 0)
    wait = _MIN_DELAY_SECONDS - (now_ts - last) + random.uniform(0.5, 1.5)
    if wait > 0:
        time.sleep(wait)

    # دمج الheaders
    headers = dict(_REALISTIC_HEADERS)
    headers.update(kwargs.pop("headers", {}))
    kwargs["headers"] = headers

    resp = _SESSION.get(url, **kwargs)
    _last_request_time[domain] = time.time()
    return resp

# ══════════════════════════════════════════════════════════════════
#  Session مشتركة (تُنشأ مرة واحدة فقط طوال عمر السكربت)
# ══════════════════════════════════════════════════════════════════

_SESSION = requests.Session()
_SESSION.headers.update(_REALISTIC_HEADERS)
_SESSION.mount("https://", HTTPAdapter(max_retries=Retry(
    total=3,
    backoff_factor=2,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["GET"],
)))
_SESSION.mount("http://", HTTPAdapter(max_retries=Retry(
    total=3,
    backoff_factor=2,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["GET"],
)))


# ══════════════════════════════════════════════════════════════════
#  Cache للرحلات المُجلَبة من الشبكة (يمنع الطلبات المكررة)
# ══════════════════════════════════════════════════════════════════

_flight_info_cache: dict[str, tuple[dict | None, str | None]] = {}


# ══════════════════════════════════════════════════════════════════
#  قاعدة البيانات المحلية  mct_flights.json
# ══════════════════════════════════════════════════════════════════

LOCAL_DB_PATH: Path = Path(os.getenv("MCT_FLIGHTS_DB", "mct_flights.json"))
_local_db: dict[str, dict] | None = None   # None = لم يُحمَّل بعد


def _load_local_db() -> dict[str, dict]:
    """تحميل mct_flights.json مرة واحدة وتخزينه في الذاكرة."""
    global _local_db
    if _local_db is not None:
        return _local_db

    if not LOCAL_DB_PATH.exists():
        print(f"  [local_db] {LOCAL_DB_PATH} not found — skipping local lookup.")
        _local_db = {}
        return _local_db

    try:
        raw = json.loads(LOCAL_DB_PATH.read_text(encoding="utf-8"))
        # نوحّد المفاتيح: إزالة المسافات + أحرف كبيرة
        _local_db = {
            normalize_flight_number(k): v
            for k, v in raw.items()
            if isinstance(v, dict)
        }
        print(f"  [local_db] Loaded {len(_local_db)} flights from {LOCAL_DB_PATH}")
    except Exception as exc:
        print(f"  [local_db] Failed to load {LOCAL_DB_PATH}: {exc}")
        _local_db = {}

    return _local_db


def fetch_flight_info_local_db(
    flight_iata: str,
    *,
    flight_date: str | None = None,
    dep_iata: str | None = None,
    arr_iata: str | None = None,
) -> dict | None:
    """البحث في mct_flights.json كمصدر محلي فوري (بدون شبكة).

    يُعيد dict بنفس شكل باقي المصادر، أو None إذا لم تُوجَد الرحلة.
    ملاحظة: الـ STD في الملف هو الجدول الثابت فقط — لا يوجد ETD/ATD.
    """
    db = _load_local_db()
    entry = db.get(normalize_flight_number(flight_iata))
    if not entry:
        return None

    std  = (entry.get("std")  or "").strip()
    etd  = (entry.get("etd")  or "").strip()   # اختياري في الملف
    dest = (entry.get("dest") or "").strip().upper()

    if not any([std, etd, dest]):
        return None

    print(f"  [local_db] {flight_iata} → dest={dest!r}, std={std!r}, etd={etd!r}")
    return {
        "std":        std,
        "etd":        etd,
        "atd":        "",
        "dest":       dest,
        "source":     "local_db",
        "confidence": SOURCE_CONFIDENCE["local_db"],
    }


ONEDRIVE_URL: str = os.getenv("ONEDRIVE_FILE_URL", "")
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
ROSTER_IMPORT_RAW_BASE: str = "https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/import"
ROSTER_JSON_RAW: str = "https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/data/roster.json"
ROSTER_JSON_GH:  str = "https://github.com/khalidsaif912/roster-site/blob/main/docs/data/roster.json"

MANPOWER_JSON_PATH: Path = Path(os.getenv("MANPOWER_JSON_PATH", "manpower.json"))


# ══════════════════════════════════════════════════════════════════
#  الدوال المساعدة
# ══════════════════════════════════════════════════════════════════

def download_file() -> tuple[str, str]:
    """Download the OneDrive file and return (html_text, last_modified_local_str).

    last_modified_local_str is HH:MM in TIMEZONE, derived from the HTTP
    Last-Modified header.  Falls back to '' if the header is missing.
    """
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

    # استخراج وقت آخر تعديل للملف (يقارب وقت إرسال/استلام الإيميل)
    lm_str = ""
    lm_header = response.headers.get("Last-Modified", "")
    if lm_header:
        try:
            from email.utils import parsedate_to_datetime
            lm_dt_utc = parsedate_to_datetime(lm_header)
            lm_local  = lm_dt_utc.astimezone(ZoneInfo(TIMEZONE))
            lm_str    = lm_local.strftime("%H:%M")
            print(f"  [OneDrive] Last-Modified: {lm_header} → local: {lm_str}")
        except Exception as exc:
            print(f"  [OneDrive] Failed to parse Last-Modified: {exc}")

    return response.text, lm_str


def compute_sha256(content: str) -> str:
    return hashlib.sha256(content.encode("utf-8", errors="ignore")).hexdigest()


def normalize_flight_date(date_str: str, now: datetime) -> str:
    """Convert common email-style dates into 'YYYY-MM-DD'.

    Supported examples:
      27FEB, 27FEB26, 27 FEB, 27-FEB, 27.FEB, 27-FEB-26, 27 FEB 2026, 2026-02-27
    Returns '' if parsing fails.
    """
    s = (date_str or "").strip().upper()
    if not s:
        return ""

    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return s

    compact = re.sub(r"[^A-Z0-9]", "", s)
    spaced = s.replace("/", " ").replace("-", " ").replace(".", " ")
    spaced = re.sub(r"\s+", " ", spaced).strip()

    months = {
        "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
        "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
    }

    day = None
    mon = None
    year = None

    patterns = [
        re.fullmatch(r"(\d{1,2})([A-Z]{3})(\d{4})", compact),
        re.fullmatch(r"(\d{1,2})([A-Z]{3})(\d{2})", compact),
        re.fullmatch(r"(\d{1,2})([A-Z]{3})", compact),
        re.fullmatch(r"(\d{1,2})\s*([A-Z]{3})\s*(\d{4})", spaced),
        re.fullmatch(r"(\d{1,2})\s*([A-Z]{3})\s*(\d{2})", spaced),
        re.fullmatch(r"(\d{1,2})\s*([A-Z]{3})", spaced),
    ]

    for m in patterns:
        if not m:
            continue
        day = int(m.group(1))
        mon = m.group(2)[:3]
        if len(m.groups()) >= 3 and m.group(3):
            y = m.group(3)
            year = int(y) if len(y) == 4 else 2000 + int(y)
        else:
            year = now.year
        break

    if day is None or mon not in months or year is None:
        return ""

    try:
        d = datetime(year, months[mon], day).date()
    except ValueError:
        return ""

    if (d - now.date()).days > 180:
        try:
            d = datetime(year - 1, months[mon], day).date()
        except ValueError:
            pass

    return d.isoformat()


def normalize_flight_number(flight_iata: str) -> str:
    """Normalize flight numbers like 'WY 251' -> 'WY251'."""
    return re.sub(r"\s+", "", (flight_iata or "").strip().upper())


def _time_only(val: str, tz: str = TIMEZONE) -> str:
    """Return HH:MM from a datetime-like string, converting to local timezone.

    Handles:
      - ISO 8601 with UTC offset  : '2026-03-01T11:00:00+00:00' → '15:00' (MCT)
      - ISO 8601 with Z suffix    : '2026-03-01T11:00:00Z'       → '15:00' (MCT)
      - Space-separated datetime  : '2026-03-01 11:00:00'        → converted if TZ present
      - Time-only string          : '15:00'                       → returned as-is
    """
    s = (val or "").strip()
    if not s:
        return ""

    # Try full ISO 8601 parsing with timezone info
    # Normalise Z → +00:00 for Python < 3.11 compatibility
    s_norm = s.replace("Z", "+00:00").replace("z", "+00:00")

    for fmt in (
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S.%f%z",
        "%Y-%m-%d %H:%M:%S%z",
        "%Y-%m-%d %H:%M%z",
    ):
        try:
            dt_parsed = datetime.strptime(s_norm, fmt)
            dt_local  = dt_parsed.astimezone(ZoneInfo(tz))
            result    = dt_local.strftime("%H:%M")
            if result != s_norm[:5]:          # only log when conversion actually changed the value
                print(f"  [tz-convert] {s!r} → {result!r} ({tz})")
            return result
        except ValueError:
            continue

    # Fallback: extract HH:MM as-is (no timezone info present)
    # NOTE: Bare HH:MM from AirLabs dep_time is already in local airport time.
    #       Only ISO datetimes with explicit UTC offset (handled above) need conversion.
    #       We do NOT blindly assume bare HH:MM is UTC — that would break dep_time fields.
    m = re.search(r"(\d{2}:\d{2})", s)
    return m.group(1) if m else ""



def _airlabs_best_row(rows: list, flight_iata: str, flight_date: str | None) -> dict | None:
    """Pick the best matching row using weighted scoring.

    Scoring:
      +50  exact flight_iata match
      +40  date match in dep_scheduled
      +30  date match in dep_time (less reliable)
      +20  dep_iata == MCT or arr_iata matches
      +10  has dep_scheduled field
      +5   has dep_estimated field

    Minimum threshold: 50 (must match flight OR date).
    Never returns first row blindly.
    """
    if not rows:
        return None

    scored = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        score = 0

        # 1. Flight IATA match
        row_iata = (row.get("flight_iata") or "").strip().upper()
        if row_iata == flight_iata:
            score += 50

        # 2. Date match in dep_scheduled (most reliable time field)
        if flight_date:
            dep_sched = str(row.get("dep_scheduled") or "")
            dep_time = str(row.get("dep_time") or "")
            if flight_date in dep_sched:
                score += 40
            elif flight_date in dep_time:
                score += 30  # less reliable field

        # 3. Route relevance (MCT connection)
        if (row.get("dep_iata") or "").strip().upper() == "MCT":
            score += 20
        elif (row.get("arr_iata") or "").strip().upper() == "MCT":
            score += 10

        # 4. Data richness bonus
        if (row.get("dep_scheduled") or "").strip():
            score += 10
        if (row.get("dep_estimated") or "").strip():
            score += 5

        scored.append((score, row))

    if not scored:
        return None

    scored.sort(key=lambda x: x[0], reverse=True)

    # Minimum threshold to prevent wrong-day/wrong-route selection
    if scored[0][0] >= 50:
        best = scored[0][1]
        if len(scored) > 1:
            print(
                f"  [AirLabs] {len(scored)} candidates for {flight_iata} "
                f"date={flight_date!r}; selected score={scored[0][0]} "
                f"arr_iata={best.get('arr_iata')!r}"
            )
        return best

    print(f"  [AirLabs] No candidate scored >=50 for {flight_iata} date={flight_date!r} (best={scored[0][0]})")
    return None


def fetch_flight_info_airlabs(
    flight_iata: str,
    *,
    flight_date: str | None = None,
    dep_iata: str | None = None,
    arr_iata: str | None = None,
) -> dict | None:
    """Fetch flight info (STD/ETD + DEST) from AirLabs if AIRLABS_API_KEY is set.

    Strategy:
      1. Try /schedules endpoint first (more reliable for scheduled flights).
      2. Fall back to /flights (real-time) if /schedules returns nothing.

    Returns a dict like:
      {"std":"HH:MM", "etd":"HH:MM", "dest":"ADD"}
    Times are returned as *time only* (HH:MM), because the report already shows the date.
    """
    api_key = os.environ.get("AIRLABS_API_KEY", "").strip()
    if not api_key:
        return None

    flight_iata = normalize_flight_number(flight_iata)
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

    # ── 1) Try /schedules first (more reliable for scheduled flights) ──
    sched_extra: dict[str, str] = {}
    if flight_date:
        sched_extra["flight_date"] = flight_date
    if dep_iata:
        sched_extra["dep_iata"] = dep_iata.strip().upper()
    if arr_iata:
        sched_extra["arr_iata"] = arr_iata.strip().upper()
    rows = _fetch("schedules", sched_extra)
    best = _airlabs_best_row(rows, flight_iata, flight_date)

    # ── 2) Fall back to /flights (real-time) ──────────────────────
    if not best:
        extra: dict[str, str] = {}
        if dep_iata:
            extra["dep_iata"] = dep_iata.strip().upper()
        if arr_iata:
            extra["arr_iata"] = arr_iata.strip().upper()
        rows = _fetch("flights", extra)
        best = _airlabs_best_row(rows, flight_iata, flight_date)

    if not best:
        return None

    # ═══ CRITICAL FIX: Correct field mapping ═══
    # STD = dep_scheduled ONLY (the published timetable)
    # ETD = dep_estimated (airline's updated prediction)
    # ATD = dep_time / dep_actual (what actually happened)
    # NEVER use dep_time as STD fallback!
    std  = _time_only(str(best.get("dep_scheduled") or ""))
    etd  = _time_only(str(best.get("dep_estimated") or best.get("dep_estimated_utc") or ""))
    atd  = _time_only(str(best.get("dep_time") or best.get("dep_actual") or best.get("dep_time_utc") or ""))
    dest = str(best.get("arr_iata") or "").strip().upper()

    # Determine source based on which endpoint was used
    source = "airlabs_schedules" if best.get("dep_scheduled") else "airlabs_flights"

    # Validation: if ATD is >12h before STD, likely wrong date match
    if std and atd:
        try:
            std_mins = int(std.split(":")[0]) * 60 + int(std.split(":")[1])
            atd_mins = int(atd.split(":")[0]) * 60 + int(atd.split(":")[1])
            if abs(std_mins - atd_mins) > 720:  # >12 hours difference
                print(f"  [AirLabs] WARNING: ATD={atd} vs STD={std} differ by >12h — possible date mismatch")
        except (ValueError, IndexError):
            pass

    print(f"  [AirLabs] {flight_iata} → dest={dest!r}, std={std!r}, etd={etd!r}, atd={atd!r} [{source}]")
    return {
        "std": std,
        "etd": etd,
        "atd": atd,
        "dest": dest,
        "source": source,
        "confidence": SOURCE_CONFIDENCE[source],
    }





def _pick_dest_from_text(text: str, flight_iata: str) -> str:
    flight_iata = normalize_flight_number(flight_iata)
    if not text or not flight_iata:
        return ""
    patterns = [
        rf"{re.escape(flight_iata)}[^\n]{{0,120}}?TO\s+([A-Z][A-Z .'-]{{2,40}})\s*\(([A-Z]{{3}})\)",
        rf"{re.escape(flight_iata)}[^\n]{{0,120}}?([A-Z][A-Z .'-]{{2,40}})\s*\(([A-Z]{{3}})\)",
        rf"FROM\s+MUSCAT\s*\(MCT\)\s+TO\s+([A-Z][A-Z .'-]{{2,40}})\s*\(([A-Z]{{3}})\)",
    ]
    up = text.upper()
    for pat in patterns:
        m = re.search(pat, up, flags=re.IGNORECASE)
        if m:
            return (m.group(2) or "").strip().upper()
    return ""


def fetch_flight_info_flightradar(
    flight_iata: str,
    *,
    flight_date: str | None = None,
    dep_iata: str | None = None,
    arr_iata: str | None = None,
) -> dict | None:
    """Scrape a public Flightradar24 flight page as a fallback source."""
    flight_iata = normalize_flight_number(flight_iata)
    if not flight_iata:
        return None

    url = f"https://www.flightradar24.com/data/flights/{flight_iata.lower()}"
    try:
        resp = _rate_limited_get(url, timeout=30)
        resp.raise_for_status()
    except requests.exceptions.HTTPError as exc:
        if exc.response is not None and exc.response.status_code == 403:
            print(f"  [Flightradar] 403 Forbidden for {flight_iata} — waiting 10s before next request")
            time.sleep(10)
        else:
            print(f"  [Flightradar] request error for {flight_iata}: {exc}")
        return None
    except Exception as exc:
        print(f"  [Flightradar] request error for {flight_iata}: {exc}")
        return None

    page_text = BeautifulSoup(resp.text, "html.parser").get_text(" ", strip=True)
    page_text = re.sub(r"\s+", " ", page_text)
    up = page_text.upper()

    date_pat = ""
    if flight_date:
        try:
            dt = datetime.strptime(flight_date, "%Y-%m-%d")
            date_pat = dt.strftime("%d %b %Y").upper()
        except Exception:
            date_pat = ""

    segment = up
    if date_pat and date_pat in up:
        pos = up.find(date_pat)
        segment = up[pos: pos + 1200]

    std = ""
    etd = ""
    atd = ""
    dest = arr_iata.strip().upper() if arr_iata else ""

    m_std = re.search(r"STD\s*(\d{2}:\d{2})", segment)
    if m_std:
        std = m_std.group(1)

    m_atd = re.search(r"ATD\s*(\d{2}:\d{2})", segment)
    if m_atd:
        atd = m_atd.group(1)

    m_est = re.search(r"ESTIMATED(?: DEPARTURE)?\s*(\d{2}:\d{2})", segment)
    if m_est:
        etd = m_est.group(1)

    dep_iata = (dep_iata or "").strip().upper()
    m_route = re.search(r"FROM\s+([A-Z .'-]+)\s*\(([A-Z]{3})\)\s+TO\s+([A-Z .'-]+)\s*\(([A-Z]{3})\)", segment)
    if m_route:
        dep_code = m_route.group(2).strip().upper()
        arr_code = m_route.group(4).strip().upper()
        if not dep_iata or dep_code == dep_iata:
            dest = arr_code

    if not dest:
        dest = _pick_dest_from_text(up, flight_iata)

    if not any([std, etd, atd, dest]):
        print(f"  [Flightradar] No useful match for {flight_iata}")
        return None

    print(f"  [Flightradar] {flight_iata} → dest={dest!r}, std={std!r}, etd={etd!r}, atd={atd!r}")
    return {
        "std": std,
        "etd": etd,
        "atd": atd,
        "dest": dest,
        "source": "flightradar24",
        "confidence": SOURCE_CONFIDENCE["flightradar24"],
    }


def fetch_flight_info_muscatairport(
    flight_iata: str,
    *,
    flight_date: str | None = None,
    dep_iata: str | None = None,
    arr_iata: str | None = None,
) -> dict | None:
    """Scrape Muscat Airport departures page as a last fallback source."""
    flight_iata = normalize_flight_number(flight_iata)
    if not flight_iata:
        return None

    url = "https://www.muscatairport.co.om/flight-status?date_type=0&type=2"
    try:
        resp = _rate_limited_get(url, timeout=30)
        resp.raise_for_status()
    except requests.exceptions.HTTPError as exc:
        if exc.response is not None and exc.response.status_code == 403:
            print(f"  [MuscatAirport] 403 Forbidden for {flight_iata} — waiting 10s before next request")
            time.sleep(10)
        else:
            print(f"  [MuscatAirport] request error for {flight_iata}: {exc}")
        return None
    except Exception as exc:
        print(f"  [MuscatAirport] request error for {flight_iata}: {exc}")
        return None

    page_text = BeautifulSoup(resp.text, "html.parser").get_text(" ", strip=True)
    page_text = re.sub(r"\s+", " ", page_text)
    up = page_text.upper()

    idx = up.find(flight_iata)
    if idx == -1:
        print(f"  [MuscatAirport] No listing found for {flight_iata}")
        return None

    segment = up[max(0, idx - 120): idx + 700]
    dest = arr_iata.strip().upper() if arr_iata else ""

    times = re.findall(r"\b(\d{2}:\d{2})\b", segment)

    # ═══ FIX: Only extract STD (first time). Do NOT guess ETD from second time.
    # Muscat Airport page has no labeled fields — assigning times blindly is dangerous.
    std = times[0] if times else ""
    # ETD is explicitly NOT set — confidence is too low to guess
    etd = ""

    if not dest:
        dest = _pick_dest_from_text(segment, flight_iata)

    if not any([std, dest]):
        print(f"  [MuscatAirport] No useful match for {flight_iata}")
        return None

    print(f"  [MuscatAirport] {flight_iata} → dest={dest!r}, std={std!r} (conf=20, etd intentionally blank)")
    return {
        "std": std,
        "etd": "",  # intentionally empty — unreliable source
        "atd": "",
        "dest": dest,
        "source": "muscat_airport",
        "confidence": SOURCE_CONFIDENCE["muscat_airport"],
    }


def fetch_flight_info_with_fallbacks(
    flight_iata: str,
    *,
    flight_date: str | None = None,
    dep_iata: str | None = None,
    arr_iata: str | None = None,
) -> tuple[dict | None, str | None]:
    """Try: LocalDB → AirLabs → Flightradar24.

    الترتيب الجديد:
      1. local_db   (mct_flights.json) — فوري، بدون شبكة، conf=75
      2. AirLabs    — شبكة، conf=85/95
      3. Flightradar — شبكة، conf=60، فقط إذا بقيت حقول فارغة
      ✗ MuscatAirport — مُعطَّل (يُعيد 403 دائمًا)

    التوقف المبكر: إذا اكتملت (std + dest) من مصدر واحد → لا نكمل.
    Cache: نتيجة كل رحلة تُخزَّن ولا تُعاد جلبها مرة ثانية.
    """
    flight_iata = normalize_flight_number(flight_iata)

    # ── Cache check ──
    cache_key = f"{flight_iata}:{flight_date or ''}"
    if cache_key in _flight_info_cache:
        print(f"  [cache] {flight_iata} → served from cache")
        return _flight_info_cache[cache_key]

    sources = [
        ("local_db",   fetch_flight_info_local_db),
        ("AirLabs",    fetch_flight_info_airlabs),
        ("Flightradar", fetch_flight_info_flightradar),
        # MuscatAirport مُعطَّل — يُعيد 403 دائمًا ويُضيف 10 ثوانٍ تأخير
    ]

    # ═══ Per-field confidence tracking ═══
    final: dict[str, dict] = {
        "std":  {"val": "", "conf": 0},
        "etd":  {"val": "", "conf": 0},
        "atd":  {"val": "", "conf": 0},
        "dest": {"val": "", "conf": 0},
    }
    used_sources: list[str] = []

    for source_name, fn in sources:
        # ── توقف مبكر: إذا عندنا STD + DEST لا نحتاج المزيد ──
        if final["std"]["val"] and final["dest"]["val"]:
            print(f"  [fallback] {flight_iata}: STD+DEST complete after {used_sources} — skipping remaining sources")
            break

        try:
            info = fn(
                flight_iata,
                flight_date=flight_date,
                dep_iata=dep_iata,
                arr_iata=arr_iata,
            )
        except Exception as exc:
            print(f"  [fallback] {source_name} failed for {flight_iata}: {exc}")
            time.sleep(3)
            try:
                info = fn(flight_iata, flight_date=flight_date, dep_iata=dep_iata, arr_iata=arr_iata)
            except Exception as exc2:
                print(f"  [fallback] {source_name} retry also failed: {exc2}")
                continue

        if not info:
            continue

        conf = info.get("confidence", 0)
        filled_something = False

        for field in ("std", "etd", "atd", "dest"):
            new_val = (info.get(field) or "").strip()
            if new_val and conf > final[field]["conf"]:
                old_val = final[field]["val"]
                if old_val and old_val != new_val:
                    print(f"  [fallback] {flight_iata}.{field}: {old_val!r} -> {new_val!r} (conf {conf} > {final[field]['conf']})")
                final[field] = {"val": new_val, "conf": conf}
                filled_something = True

        if filled_something:
            used_sources.append(source_name)

    # Flatten to simple dict
    merged = {k: v["val"] for k, v in final.items()}

    # تحذير إذا بقيت حقول فارغة بعد استنفاد جميع المصادر
    missing = [k for k in ("std", "dest") if not merged[k].strip()]
    if missing:
        print(f"  [⚠ fallback] {flight_iata}: still missing after all sources: {', '.join(missing)}")

    if not any(merged[k].strip() for k in ("std", "etd", "dest")):
        result = (None, None)
    else:
        combined_source = "+".join(used_sources) if used_sources else None
        result = (merged, combined_source)

    # ── تخزين في الـ cache ──
    _flight_info_cache[cache_key] = result
    return result


def get_shift(now: datetime) -> str:
    """تحديد المناوبة الحالية بناءً على الوقت.

    المناوبات:
      shift1 : 06:00 – 15:00
      shift2 : 13:00 – 22:00 (تداخل مع shift1 بين 13:00-15:00)
      shift3 : 21:00 – 06:00 (تداخل مع shift2 بين 21:00-22:00)

    قواعد الأوفلود (القطع):
      بعد 14:30 → shift2
      بعد 21:30 → shift3
      بعد 05:30 → shift1
    """
    mins = now.hour * 60 + now.minute
    # نستخدم أوقات القطع للأوفلود لتحديد المناوبة الفعلية
    if 5 * 60 + 30 <= mins < 14 * 60 + 30:
        return "shift1"
    if 14 * 60 + 30 <= mins < 21 * 60 + 30:
        return "shift2"
    return "shift3"


def get_shift_date(now: datetime, shift: str | None = None) -> str:
    """Return the correct date_dir for the current shift.
    For shift3 after midnight (hour < 6), return yesterday's date.
    """
    if shift is None:
        shift = get_shift(now)
    if shift == "shift3" and now.hour < 6:
        return (now - timedelta(days=1)).strftime("%Y-%m-%d")
    return now.strftime("%Y-%m-%d")


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
                if re.match(r"^(CBT|BT|AKE|PMC|PAG|ULD|AKH|RKN|QKE|PKC|AAK|AKN|DQF|DQN|FQA|FQN|PGA|PLA|PLB|RKN|SAA)\w*", awb_clean) and not any([pcs, kgs, desc, rsn]):
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
            dest_c     = ""

            # ابحث في الأسطر التالية عن البيانات والسبب
            j = i + 1
            while j < len(lines) and j < i + 30:
                m_data = data_pat.match(lines[j])
                if m_data:
                    awb   = m_data.group(1).strip()
                    pcs   = m_data.group(2).strip()
                    desc  = m_data.group(3).strip()
                    cls_  = m_data.group(4).strip()
                    kgs   = m_data.group(5).strip()
                    dest  = m_data.group(6).strip()
                    dest_c = dest
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
                    "destination": dest_c,
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

def save_flights(flights: list[dict], now: datetime) -> tuple[str, str, dict, list[str]]:
    """
    Save flights under the folder of the actual flight date, not merely the email/runtime date.

    Returns:
        operational_date_dir: shift date derived from runtime (kept for compatibility/logging)
        shift: current shift key
        operational_meta: meta.json for the runtime-derived folder (or empty fallback)
        affected_date_dirs: sorted list of date folders that were actually written
    """
    shift = get_shift(now)
    operational_date_dir = get_shift_date(now, shift)

    metas_by_folder: dict[Path, dict] = {}
    affected_date_dirs: set[str] = set()

    for flight in flights:
        flight_date_dir = normalize_flight_date(flight.get("date", ""), now) or operational_date_dir
        affected_date_dirs.add(flight_date_dir)

        folder = DATA_DIR / flight_date_dir / shift
        folder.mkdir(parents=True, exist_ok=True)

        meta_path = folder / "meta.json"
        meta = metas_by_folder.get(meta_path)
        if meta is None:
            meta = load_json(meta_path, {"flights": {}})
            if not isinstance(meta, dict) or "flights" not in meta:
                meta = {"flights": {}}
            metas_by_folder[meta_path] = meta

        filename = slugify(
            f"{flight['flight']}_{flight.get('date','')}_{flight.get('destination','')}"
        ) + ".json"
        file_path = folder / filename
        existed = file_path.exists()

        payload = {
            **flight,
            "saved_at": now.isoformat(),
            "storage_date_dir": flight_date_dir,
            "storage_shift": shift,
        }
        file_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        entry = meta["flights"].get(filename, {
            "flight": flight["flight"],
            "date": flight.get("date", ""),
            "dest": flight.get("destination", ""),
            "created_at": now.isoformat(),
            "updated_at": now.isoformat(),
            "updates": 0,
        })
        if existed:
            entry["updates"] = int(entry.get("updates", 0)) + 1
        entry["updated_at"] = now.isoformat()
        entry["storage_date_dir"] = flight_date_dir
        entry["storage_shift"] = shift
        meta["flights"][filename] = entry

    for meta_path, meta in metas_by_folder.items():
        write_json(meta_path, meta)

    operational_meta = metas_by_folder.get(
        DATA_DIR / operational_date_dir / shift / "meta.json",
        {"flights": {}},
    )
    return operational_date_dir, shift, operational_meta, sorted(affected_date_dirs)










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











# ══════════════════════════════════════════════════════════════════
#  الروستر — جلب الموظفين وعرضهم
# ══════════════════════════════════════════════════════════════════

# Mapping: shift key → label used in daily roster HTML
_SHIFT_TO_ROSTER_LABEL = {
    "shift1": "Morning",
    "shift2": "Afternoon",
    "shift3": "Night",
}

# Leave/off labels to separate from on-duty
_LEAVE_LABELS = {"annual leave", "sick leave", "emergency leave", "off day", "training"}
_ROSTER_EXCLUDED_DEPTS = {"officers"}
_ROSTER_EXCLUDED_SNS = set()  # no global SN exclusions — handled per-section


def _normalize_shift_label(value: str) -> str:
    return (value or "").strip().lower()


def _roster_request_headers() -> dict[str, str]:
    return {
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache",
        "Expires": "0",
        "User-Agent": "Mozilla/5.0",
    }


def _fetch_daily_roster_html(date_dir: str) -> str:
    """Fetch the Export daily roster HTML page for a specific date."""
    day_url = f"{ROSTER_PAGE_URL.rstrip('/')}/date/{date_dir}/"
    response = requests.get(
        day_url,
        timeout=20,
        headers=_roster_request_headers(),
    )
    response.raise_for_status()
    return response.text


def _fetch_import_roster_html(date_dir: str) -> str:
    """Fetch the Import daily roster HTML page for a specific date.

    Tries the published page first, then falls back to raw GitHub HTML.
    """
    candidates = [
        f"{ROSTER_PAGE_URL.rstrip('/')}/import/{date_dir}/",
        f"{ROSTER_IMPORT_RAW_BASE.rstrip('/')}/{date_dir}/index.html",
    ]
    last_exc: Exception | None = None
    for url in candidates:
        try:
            response = requests.get(
                url,
                timeout=20,
                headers=_roster_request_headers(),
            )
            response.raise_for_status()
            return response.text
        except Exception as exc:
            last_exc = exc
    raise last_exc or RuntimeError(f"Import roster fetch failed for {date_dir}")


def _normalize_import_roster_lines(html: str) -> list[str]:
    text = BeautifulSoup(html, "html.parser").get_text("\n")
    lines: list[str] = []
    for raw_line in text.splitlines():
        line = re.sub(r"\s+", " ", (raw_line or "")).strip()
        if line:
            lines.append(line)
    return lines


def _extract_roster_shift_from_line(line: str) -> str:
    low = _normalize_shift_label(line)
    if "morning" in low:
        return "morning"
    if "afternoon" in low:
        return "afternoon"
    if "night" in low:
        return "night"
    if "off day" in low:
        return "off day"
    if "annual leave" in low:
        return "annual leave"
    if "sick leave" in low:
        return "sick leave"
    if "emergency leave" in low:
        return "emergency leave"
    if "training" in low:
        return "training"
    return ""


_IMPORT_DEPT_HEADERS: dict[str, str] = {
    "documentation": "Documentation",
    "flight dispatch (export)": "Flight Dispatch (Export)",
    "flight dispatch (import)": "Flight Dispatch (Import)",
    "import checkers": "Import Checkers",
    "import operators": "Import Operators",
    "release control": "Release Control",
    "supervisors": "Supervisors",
}
_IMPORT_FLIGHT_DISPATCH_KEYS: dict[str, str] = {
    "flight dispatch (export)": "fd_export",
    "flight dispatch (import)": "fd_import",
}


def _parse_import_employee_line(line: str, dept: str) -> dict | None:
    m = re.match(r"^(.+?)\s*[·•\-–]\s*(\d{3,6})\b.*$", line)
    if not m:
        return None
    return {
        "name": m.group(1).strip(),
        "sn": m.group(2).strip(),
        "dept": dept,
    }


def fetch_import_flight_dispatch_staff(date_dir: str, shift: str) -> dict:
    """Fetch Flight Dispatch staff from Import roster for the given date/shift."""
    target_shift = _normalize_shift_label(_SHIFT_TO_ROSTER_LABEL.get(shift, shift))
    result = {"fd_export": [], "fd_import": []}
    if not target_shift:
        return result

    try:
        html = _fetch_import_roster_html(date_dir)
        lines = _normalize_import_roster_lines(html)
    except Exception as exc:
        print(f"  [roster-import] Failed to fetch/parse {date_dir}: {exc}")
        return result

    current_dept_key = ""
    current_shift = ""

    for line in lines:
        low = _normalize_shift_label(line)

        if low in _IMPORT_DEPT_HEADERS:
            current_dept_key = low
            current_shift = ""
            continue

        if not current_dept_key:
            continue

        shift_label = _extract_roster_shift_from_line(line)
        if shift_label:
            current_shift = shift_label
            continue

        if low.startswith("total "):
            continue
        if low.startswith("view full roster") or low.startswith("last updated:"):
            break

        # Once we leave the Flight Dispatch sections, ignore the rest.
        if low in _IMPORT_DEPT_HEADERS and low not in _IMPORT_FLIGHT_DISPATCH_KEYS:
            current_dept_key = ""
            current_shift = ""
            continue

        if current_dept_key not in _IMPORT_FLIGHT_DISPATCH_KEYS:
            continue
        if current_shift != target_shift:
            continue

        item = _parse_import_employee_line(line, _IMPORT_DEPT_HEADERS[current_dept_key])
        if item:
            result[_IMPORT_FLIGHT_DISPATCH_KEYS[current_dept_key]].append(item)

    print(
        f"  [roster-import] {date_dir}/{shift}: "
        f"{len(result['fd_export'])} fd-export, {len(result['fd_import'])} fd-import"
    )
    return result


def fetch_roster_staff(date_dir: str, shift: str) -> dict:
    """Fetch staff on duty from the Export daily roster HTML page for the given date/shift.

    Returns:
        {
          "on_duty":  [{"name": str, "dept": str, "sn": str}, ...],
          "on_leave": [{"name": str, "dept": str, "status": str}, ...],
        }
    """
    target_shift = _normalize_shift_label(_SHIFT_TO_ROSTER_LABEL.get(shift, shift))
    if not target_shift:
        return {"on_duty": [], "on_leave": []}

    try:
        html = _fetch_daily_roster_html(date_dir)
        soup = BeautifulSoup(html, "html.parser")
    except Exception as e:
        print(f"  [roster-html] Failed to fetch/parse {date_dir}: {e}")
        return {"on_duty": [], "on_leave": []}

    on_duty: list[dict] = []
    on_leave: list[dict] = []

    for dept_card in soup.select(".deptCard"):
        dept_el = dept_card.select_one(".deptTitle")
        dept = dept_el.get_text(" ", strip=True) if dept_el else "Unknown"
        dept_norm = dept.strip().lower()

        for shift_card in dept_card.select("details.shiftCard"):
            shift_label = _normalize_shift_label(shift_card.get("data-shift") or "")

            for emp_row in shift_card.select(".empRow"):
                name_el = emp_row.select_one(".empName")
                if not name_el:
                    continue

                raw_name = name_el.get_text(" ", strip=True)
                sn = ""
                name = raw_name

                # Matches both:
                #   Mohamed Al Amri - 81404
                #   Mohamed Al Subhi - 82592 (Inventory)
                m = re.match(r"^(.+?)\s*[-–]\s*(\d+)(?:\s*\(.*?\))?\s*$", raw_name)
                if m:
                    name = m.group(1).strip()
                    sn = m.group(2).strip()

                item = {"name": name, "sn": sn, "dept": dept}

                if shift_label in _LEAVE_LABELS:
                    item["status"] = shift_label.title()
                    on_leave.append(item)
                    continue

                if shift_label != target_shift:
                    continue

                if dept_norm in _ROSTER_EXCLUDED_DEPTS:
                    continue

                on_duty.append(item)

    print(f"  [roster-html] {date_dir}/{shift}: {len(on_duty)} on duty, {len(on_leave)} on leave")
    all_depts = sorted({e.get("dept", "?") for e in on_duty + on_leave})
    print(f"  [roster-debug] all depts seen: {all_depts}")
    return {"on_duty": on_duty, "on_leave": on_leave}



# ══════════════════════════════════════════════════════════════════
#  بناء صفحات HTML
# ══════════════════════════════════════════════════════════════════

def _render_offload_table(flights: list[dict], meta: dict) -> str:
    """Render offload section as a single vertical table (Type B style).
    Columns: ITEM | DATE | FLIGHT | STD/ETD | DEST | Email Received Time |
             Physical Cargo Received from Ramp | Trolley/ULD Number |
             Offloading Process Completed in CMS | Offloading Pieces Verification |
             Offloading Reason | Remarks/Additional Information
    """
    if not flights:
        flights = []

    # ── Styles ──
    hdr_bg      = "#dce6f4"
    hdr_color   = "#1b1f2a"
    hdr_border  = "#a8bcd8"
    row_even    = "#ffffff"
    row_odd     = "#f4f7fc"
    cell_border = "#d0d9ee"
    nil_color   = "#64748b"
    text_dark   = "#1b1f2a"
    totals_bg   = "#eef3fc"
    totals_border = "#0b3a78"
    totals_color = "#0b3a78"

    # ── Column headers ──
    columns = [
        ("ITEM", "40px"),
        ("DATE", "80px"),
        ("FLIGHT", "80px"),
        ("STD/ETD", "80px"),
        ("DEST", "60px"),
        ("Email Received Time", "90px"),
        ("Physical Cargo Received from Ramp", "100px"),
        ("Trolley/ ULD Number", "90px"),
        ("Offloading Process Completed in CMS", "100px"),
        ("Offloading Pieces Verification", "100px"),
        ("Offloading Reason", "100px"),
        ("Remarks/Additional Information", ""),
    ]

    col_headers = "<tr>"
    for label, width in columns:
        w = f"width:{width};" if width else ""
        col_headers += (
            f'<td style="padding:8px 6px; background-color:{hdr_bg}; color:{hdr_color};'
            f'font-weight:700; font-size:11px; font-family:Calibri,Arial,sans-serif;'
            f'border:1px solid {hdr_border}; text-align:center; vertical-align:middle; {w}">'
            f'{label}</td>'
        )
    col_headers += "</tr>"

    # ── Deduplicate flights by flight number (keep first occurrence) ──
    seen_flights: set[str] = set()
    unique_flights: list[dict] = []
    for f in flights:
        fkey = (f.get("flight", "") or "").strip().upper()
        if fkey and fkey in seen_flights:
            continue
        if fkey:
            seen_flights.add(fkey)
        unique_flights.append(f)
    flights = unique_flights

    # ── Data rows ──
    data_rows = ""
    item_num = 0

    # tabindex counter for Tab navigation
    _ti = [1]
    def _next_ti():
        v = _ti[0]; _ti[0] += 1; return v

    def _format_full_date(raw: str) -> str:
        """Force date into DD-MMM-YY format (e.g. 15-MAR-26) no matter what input arrives."""
        raw = (raw or "").strip()
        if not raw or raw == "—":
            # No date provided — use today's date in Muscat timezone
            today = datetime.now(ZoneInfo(TIMEZONE))
            return today.strftime("%d-%b-%y").upper()

        raw_up = raw.upper().replace("/", "-").replace(".", "-")

        # 1) ISO: 2026-03-15 or 2026-03-15T...
        m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", raw_up)
        if m:
            try:
                dt = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                return dt.strftime("%d-%b-%y").upper()
            except ValueError:
                pass

        # 2) Full with year: 15-MAR-26, 15MAR26, 15-MAR-2026, 15MAR2026
        for fmt in ("%d-%b-%y", "%d%b%y", "%d-%b-%Y", "%d%b%Y"):
            try:
                dt = datetime.strptime(raw_up, fmt)
                return dt.strftime("%d-%b-%y").upper()
            except ValueError:
                pass

        # 3) Short without year: 15MAR or 15-MAR — attach current year
        m = re.match(r"(\d{1,2})-?([A-Z]{3})$", raw_up)
        if m:
            try:
                yr = datetime.now(ZoneInfo(TIMEZONE)).year
                dt = datetime.strptime(f"{m.group(1)}{m.group(2)}{yr}", "%d%b%Y")
                return dt.strftime("%d-%b-%y").upper()
            except ValueError:
                pass

        # 4) Try extracting any day+month from the string
        m = re.search(r"(\d{1,2})\s*-?\s*([A-Z]{3})\s*-?\s*(\d{2,4})?", raw_up)
        if m:
            day, mon = m.group(1), m.group(2)
            yr_str = m.group(3)
            if not yr_str:
                yr_str = str(datetime.now(ZoneInfo(TIMEZONE)).year)
            try:
                dt = datetime.strptime(f"{day}{mon}{yr_str}", "%d%b%Y" if len(yr_str) == 4 else "%d%b%y")
                return dt.strftime("%d-%b-%y").upper()
            except ValueError:
                pass

        # 5) Absolute fallback — return today's date
        today = datetime.now(ZoneInfo(TIMEZONE))
        return today.strftime("%d-%b-%y").upper()

    def _to_muscat_time(time_str: str) -> str:
        """Convert time string to Muscat local time (UTC+4).

        - ISO datetime with explicit tz offset -> convert to Muscat.
        - Bare HH:MM from HTML table (STD/ETD column) -> treat as UTC and add +4h.
          Reason: The offload report stores STD/ETD in UTC (e.g. 06:40 UTC = 10:40 MCT).
        """
        s = (time_str or "").strip()
        if not s:
            return ""
        # Full ISO datetime -> convert to Muscat
        try:
            dt = datetime.fromisoformat(s)
            if dt.tzinfo is not None:
                dt = dt.astimezone(ZoneInfo(TIMEZONE))
                return dt.strftime("%H:%M")
            # ISO without tz -> treat as UTC
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
            return dt.astimezone(ZoneInfo(TIMEZONE)).strftime("%H:%M")
        except (ValueError, TypeError):
            pass
        # Bare HH:MM -> treat as UTC and convert to Muscat (UTC+4)
        m_t = re.match(r"^(\d{1,2}):(\d{2})$", s)
        if m_t:
            try:
                today = datetime.now(ZoneInfo(TIMEZONE)).date()
                dt_utc = datetime(today.year, today.month, today.day,
                                  int(m_t.group(1)), int(m_t.group(2)),
                                  tzinfo=ZoneInfo("UTC"))
                converted = dt_utc.astimezone(ZoneInfo(TIMEZONE)).strftime("%H:%M")
                if converted != s:
                    print(f"  [tz-convert] STD/ETD {s!r} (UTC) -> {converted!r} (MCT)")
                return converted
            except Exception:
                pass
        return s

    for flight in flights:
        flt  = (flight.get("flight", "") or "—").upper()
        date = _format_full_date(flight.get("date", "") or "")
        dest = (flight.get("destination", "") or "—").upper()
        std_raw = flight.get("std_etd", "") or ""
        std_val, etd_val = _format_std_etd(std_raw)
        # Convert times to Muscat timezone
        std_val = _to_muscat_time(std_val)
        etd_val = _to_muscat_time(etd_val)
        std_etd_display = f"{std_val}/{etd_val}" if std_val or etd_val else "—"
        if std_val and not etd_val:
            std_etd_display = std_val
        elif etd_val and not std_val:
            std_etd_display = etd_val

        # Ops status
        email = (flight.get("email_time") or "").strip()
        if not email:
            _saved_at = (flight.get("saved_at") or "").strip()
            if _saved_at:
                try:
                    _sa_dt = datetime.fromisoformat(_saved_at)
                    if _sa_dt.tzinfo is not None:
                        _sa_dt = _sa_dt.astimezone(ZoneInfo(TIMEZONE))
                    email = _sa_dt.strftime("%H:%M")
                except Exception:
                    email = ""
        physical = (flight.get("physical")   or "").strip().upper() or ""
        cms      = (flight.get("cms")        or "").strip().upper() or ""
        # Pieces verification: sum PCS from all items
        total_pcs = 0
        for it in flight.get("items", []):
            try:
                total_pcs += int(it.get("pcs", 0) or 0)
            except (ValueError, TypeError):
                pass
        verified = str(total_pcs) if total_pcs > 0 else ""
        remarks  = (flight.get("remarks")    or "").strip().upper() or ""

        items = flight.get("items", [])
        real_items = [i for i in items if (i.get("awb","") or "").strip()]

        # ── Single row per flight ──
        item_num += 1
        bg = row_odd if item_num % 2 == 0 else row_even

        td_s = (f'style="padding:7px 6px;border:1px solid {cell_border};'
                f'font-size:12px;font-family:Calibri,Arial,sans-serif;color:{text_dark};'
                f'background:{bg};text-align:center;vertical-align:middle;"')

        # Offloading reason: combine unique reasons from items
        reasons = []
        for it in real_items:
            r = (it.get("reason", "") or "").strip().upper()
            if r and r not in reasons:
                reasons.append(r)
        reason_display = ", ".join(reasons) if reasons else ""

        # Trolley/ULD: only use trolley field — never fall back to AWB numbers
        uld_parts = []
        for it in real_items:
            u = (it.get("trolley", "") or "").strip().upper()
            if u and u not in uld_parts:
                uld_parts.append(u)
        uld_display = ", ".join(uld_parts) if uld_parts else ""

        data_rows += f"""
      <tr>
        <td {td_s}><strong>{item_num}</strong></td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}" data-col="date">{date}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}" data-col="flight">{flt}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}" data-col="std">{std_etd_display}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}" data-col="dest">{dest}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}" data-col="email">{email}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}">{physical}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}">{uld_display}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}">{cms}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}">{verified}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}">{reason_display}</td>
        <td {td_s} contenteditable="true" tabindex="{_next_ti()}">{remarks}</td>
      </tr>"""

    # ── 3 empty rows for manual entry ──
    _empty_td = (f'style="padding:7px 6px;border:1px solid {cell_border};'
                 f'font-size:12px;font-family:Calibri,Arial,sans-serif;color:{text_dark};'
                 f'background:{row_even};text-align:center;"')
    for _ in range(3):
        item_num += 1
        data_rows += f"""
      <tr>
        <td {_empty_td}><strong>{item_num}</strong></td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}" data-col="date">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}" data-col="flight">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}" data-col="std">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}" data-col="dest">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}" data-col="email">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}">&nbsp;</td>
        <td {_empty_td} contenteditable="true" tabindex="{_next_ti()}">&nbsp;</td>
      </tr>"""

    # ── NIL case ──
    if not flights:
        data_rows = f"""
      <tr id="nil-row">
        <td colspan="12" style="padding:10px 10px; border:1px solid {cell_border};
            color:{nil_color}; text-align:center; font-style:italic; font-size:12px;
            font-family:Calibri,Arial,sans-serif; background:{row_even};">
          <span id="nil-text" contenteditable="true" style="outline:none;display:inline-block;min-width:200px;">NIL \u2014 No offload data recorded for this shift.</span>
          &nbsp;<button onclick="var r=document.getElementById('nil-row');if(r)r.remove();triggerAutosave();"
            style="font-size:10px;padding:1px 7px;cursor:pointer;background:#eef3fc;border:1px solid #0b3a78;color:#0b3a78;border-radius:3px;vertical-align:middle;">\u2715 Remove</button>
        </td>
      </tr>"""
        # Add 3 empty rows even for NIL
        for i in range(1, 4):
            data_rows += f"""
      <tr>
        <td {_empty_td}><strong>{i}</strong></td>
        <td {_empty_td} contenteditable="true" data-col="date">&nbsp;</td>
        <td {_empty_td} contenteditable="true" data-col="flight">&nbsp;</td>
        <td {_empty_td} contenteditable="true" data-col="std">&nbsp;</td>
        <td {_empty_td} contenteditable="true" data-col="dest">&nbsp;</td>
        <td {_empty_td} contenteditable="true" data-col="email">&nbsp;</td>
        <td {_empty_td} contenteditable="true">&nbsp;</td>
        <td {_empty_td} contenteditable="true">&nbsp;</td>
        <td {_empty_td} contenteditable="true">&nbsp;</td>
        <td {_empty_td} contenteditable="true">&nbsp;</td>
        <td {_empty_td} contenteditable="true">&nbsp;</td>
        <td {_empty_td} contenteditable="true">&nbsp;</td>
      </tr>"""

    table_html = f"""
    <table width="100%" cellpadding="0" cellspacing="0" border="0"
           style="border-collapse:collapse; font-family:Calibri,Arial,sans-serif; margin-top:12px; margin-bottom:14px;">
      {col_headers}
      <tbody id="offload-tbody">
      {data_rows}
      </tbody>
    </table>"""

    return f'<div class="offload-scroll" style="margin-top:12px;">{table_html}</div>'


def _render_manpower_section(roster: dict, supervisor_display: str = "", import_roster: dict | None = None) -> str:
    """Render Section 6 MANPOWER — grouped by dept, sections B-G."""
    # re is already imported at module level
    on_duty = roster.get("on_duty", [])
    import_roster = import_roster or {"fd_export": [], "fd_import": []}

    td_style  = "font-family:Calibri,Arial,sans-serif;font-size:13px;color:#1b1f2a;line-height:1.8;"
    hdr_style = "color:#0b3a78;font-weight:700;font-size:13px;"
    dept_hdr  = "color:#0b3a78;font-weight:700;font-size:12px;margin:8px 0 2px 0;"
    ul_style  = "margin:2px 0 10px 20px;padding:0;"
    ul_class  = "mp-list"
    nil_item  = '<li style="color:#64748b;">NIL</li>'

    EXCLUDED_DEPTS   = {"officers"}
    # هؤلاء يُعرضون في أقسامهم الخاصة (Inventory / Support Team) — لا في القسم العام
    INVENTORY_SNS    = {"82592", "990737"}
    SUPPORT_SNS      = {"82653", "82565"}
    ALL_SPECIAL_SNS  = INVENTORY_SNS | SUPPORT_SNS

    def _is_support(e):
        return "support" in (e.get("name","") + e.get("dept","")).lower()

    def _is_excluded(e):
        dept = e.get("dept","").strip().lower()
        sn   = str(e.get("sn","")).strip()
        return dept in EXCLUDED_DEPTS or sn in ALL_SPECIAL_SNS or _is_support(e)

    def _fmt_name(emp):
        raw  = emp.get("name","").strip()
        sn   = str(emp.get("sn") or "").strip()
        # استخراج SN والاسم إذا كانا مدمجَين في raw
        m = re.match(r"^(.+?)\s*-\s*(\d{4,})\s*(?:\((.+?)\))?$", raw)
        if m:
            name_part = m.group(1).strip()
            sn_part   = m.group(2).strip()
        else:
            name_part = raw
            sn_part   = sn
        note_html = ""
        # عرض الرقم فقط بدون حروف قبله
        parts = []
        if sn_part:
            parts.append(f'<strong style="color:#1b1f2a;">{sn_part}</strong>')
        parts.append(f'<span style="color:#1b1f2a;">{name_part}</span>')
        inner_html = '&nbsp;&nbsp;'.join(parts)
        return (
            f'<span data-sn="{sn_part}" data-name="{name_part}" '
            f'style="font-family:Calibri,Arial,sans-serif;">'
            f'{inner_html}'
            f'</span>'
            f'{note_html}'
        )

    # الموظفون الرئيسيون مجمّعون بالقسم
    # OrderedDict is already imported at module level
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
        and str(e.get("sn","")).strip() not in ALL_SPECIAL_SNS
    ]
    if sup_in_roster:
        sup_li_roster = "".join(f'<li contenteditable="true" style="outline:none;">{_fmt_name(e)}</li>\n' for e in sup_in_roster)
        grouped_html += f"""
      <div style="{dept_hdr}">Supervisors:</div>
      <ul id="ul-supervisors" class="{ul_class}" style="{ul_style}">{sup_li_roster}</ul>"""
    elif supervisor_display:
        grouped_html += f"""
      <div style="{dept_hdr}">Supervisors:</div>
      <ul id="ul-supervisors" class="{ul_class}" style="{ul_style}"><li contenteditable="true" style="outline:none;"><strong>{supervisor_display}</strong></li></ul>"""
    else:
        grouped_html += f"""
      <div style="{dept_hdr}">Supervisors:</div>
      <ul id="ul-supervisors" class="{ul_class}" style="{ul_style}"><li contenteditable="true" style="outline:none;">&nbsp;</li></ul>"""

    # ثانياً: باقي الأقسام من roster (تخطّى supervisors — مُعالَج أعلاه)
    for dept, emps in by_dept.items():
        if dept.strip().lower() in MANUAL_DEPTS:
            continue
        dept_id = "ul-dept-" + re.sub(r'[^a-z0-9]', '', dept.lower())
        items_li = "".join(f'<li contenteditable="true" style="outline:none;">{_fmt_name(e)}</li>\n' for e in emps)
        grouped_html += f"""
      <div style="{dept_hdr}">{dept}:</div>
      <ul id="{dept_id}" class="{ul_class}" style="{ul_style}">{items_li}</ul>"""

    if not grouped_html:
        grouped_html = f'<ul class="{ul_class}" style="{ul_style}"><li style="color:#64748b;">No roster data available.</li></ul>'

    # ── Inventory: من الروستر بـ SN ──
    _inventory_from_roster = [e for e in on_duty if str(e.get("sn","")).strip() in INVENTORY_SNS]

    def _fmt_emp_row(name, sn):
        sn_display = sn if sn else ""
        row_html = (
            f'<li contenteditable="true" style="outline:none;">'
            f'<span data-sn="{sn}" data-name="{name}" '
            f'style="font-family:Calibri,Arial,sans-serif;">'
        )
        if sn_display:
            row_html += f'<strong style="color:#1b1f2a;">{sn_display}</strong>&nbsp;&nbsp;'
        row_html += (
            f'<span style="color:#1b1f2a;">{name}</span>'
            f'</span></li>'
        )
        return row_html

    if _inventory_from_roster:
        _inventory_staff_items = "\n".join(_fmt_emp_row(e["name"], e["sn"]) for e in _inventory_from_roster)
    else:
        _inventory_staff_items = '<li style="color:#64748b;">—</li>'

    # ── Support Team: من الروستر بـ SN أو dept/name يحتوي support ──
    _support_by_sn   = [e for e in on_duty if str(e.get("sn","")).strip() in SUPPORT_SNS]
    _support_by_name = [e for e in on_duty if _is_support(e) and str(e.get("sn","")).strip() not in SUPPORT_SNS]
    _seen_sup = set()
    _combined_support = []
    for e in _support_by_sn + _support_by_name:
        k = str(e.get("sn","")).strip() or e.get("name","")
        if k not in _seen_sup:
            _seen_sup.add(k)
            _combined_support.append(e)

    if _combined_support:
        _c_items = "\n".join(
            f'<li contenteditable="true" style="outline:none;">{_fmt_name(e)}</li>'
            for e in _combined_support
        )
    else:
        _c_items = '<li contenteditable="true" style="outline:none;">&nbsp;</li>'

    _btn_style = "font-size:11px;padding:1px 8px;margin:2px 0 8px;cursor:pointer;background:#eef3fc;border:1px solid #0b3a78;color:#0b3a78;border-radius:3px;"
    _li_edit   = '<li contenteditable="true" style="outline:none;">&nbsp;</li>'

    # section_b
    section_b = (
        f'<div style="{hdr_style}">CTU Staff On Duty:</div>'
        f'<ul id="ul-ctu" class="{ul_class}" style="{ul_style}">{_li_edit}</ul>'
    )

    # Inventory section
    section_inventory = (
        f'<div style="{hdr_style}">Inventory:</div>'
        f'<ul id="ul-inventory" class="{ul_class}" style="{ul_style}">{_inventory_staff_items}</ul>'
    )

    # section_c
    section_c = (
        f'<div style="{hdr_style}">Support Team:</div>'
        f'<ul id="ul-support" class="{ul_class}" style="{ul_style}">{_c_items}</ul>'
    )

    fd_export_staff = import_roster.get("fd_export", [])
    fd_import_staff = import_roster.get("fd_import", [])
    fd_combined_staff = []
    fd_seen = set()
    for e in fd_export_staff + fd_import_staff:
        key = (str(e.get("sn", "")).strip(), str(e.get("name", "")).strip().lower())
        if key in fd_seen:
            continue
        fd_seen.add(key)
        fd_combined_staff.append(e)

    if fd_combined_staff:
        _fd_items = "\n".join(_fmt_emp_row(e["name"], e["sn"]) for e in fd_combined_staff)
    else:
        _fd_items = '<li style="color:#64748b;">NIL</li>'

    section_flight_dispatch = (
        f'<div style="{hdr_style}">Flight Dispatch:</div>'
        f'<ul id="ul-flight-dispatch" class="{ul_class}" style="{ul_style}">{_fd_items}</ul>'
    )

    section_d = (
        f'<div style="{hdr_style}">Sick Leave / No Show / Others:</div>'
        f'<ul id="ul-sickleave" class="{ul_class}" style="{ul_style}">{_li_edit}</ul>'
    )
    section_e = (
        f'<div style="{hdr_style}">Annual Leave / Course / Off in Lieu:</div>'
        f'<ul id="ul-annualleave" class="{ul_class}" style="{ul_style}">{_li_edit}</ul>'
    )
    section_f = (
        f'<div style="{hdr_style}">Trainee:</div>'
        f'<ul id="ul-trainee" class="{ul_class}" style="{ul_style}">{_li_edit}</ul>'
    )
    section_g = (
        f'<div style="{hdr_style}">Overtime Justification:</div>'
        f'<ul id="ul-overtime" class="{ul_class}" style="{ul_style}">{_li_edit}</ul>'
    )

    return f"""
        <td colspan="2" valign="top" style="{td_style}">
          {grouped_html}
          {section_b}
          {section_inventory}
          {section_c}
          {section_flight_dispatch}
          {section_d}
          {section_e}
          {section_f}
          {section_g}
        </td>"""



def _load_manpower_json_staff_map() -> dict[str, str]:
    """Load manpower.json and flatten all IDs to numeric SN -> NAME map."""
    path = MANPOWER_JSON_PATH
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"  [manpower.json] Failed to load {path}: {exc}")
        return {}

    out: dict[str, str] = {}

    def walk(node):
        if isinstance(node, dict):
            name = str(node.get("name", "")).strip()
            emp_id = str(node.get("id", "")).strip()
            if name and emp_id:
                m = re.search(r"(\d{3,10})", emp_id)
                if m:
                    sn = m.group(1)
                    out.setdefault(sn, name)
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(data)
    print(f"  [manpower.json] Loaded {len(out)} staff IDs from {path}")
    return out

def _update_flight_json(folder: Path, flight: dict) -> None:
    """Update a saved flight JSON file with new data (e.g. enriched STD/ETD)."""
    filename = slugify(
        f"{flight['flight']}_{flight.get('date','')}_{flight.get('destination','')}"
    ) + ".json"
    file_path = folder / filename
    if file_path.exists():
        try:
            existing = json.loads(file_path.read_text(encoding="utf-8"))
            # السماح بتحديث الحقول بقيم فارغة، مع حماية المفاتيح الأساسية
            _protected = {"flight", "date", "items"}
            existing.update({k: v for k, v in flight.items() if k not in _protected or v})
            file_path.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass


def build_shift_report(date_dir: str, shift: str) -> None:
    folder = DATA_DIR / date_dir / shift
    if not folder.exists():
        return

    meta         = load_json(folder / "meta.json", {"flights": {}})
    flight_files = sorted(p for p in folder.glob("*.json") if p.name != "meta.json")
    flights      = [json.loads(p.read_text(encoding="utf-8")) for p in flight_files]

    # ── Filter offload: only keep flights whose date matches the report date ──
    # datetime is already imported at module level
    def _flight_date_matches(flt_date_str: str, report_date: str) -> bool:
        """Strictly keep only flights whose parsed date equals report_date."""
        if not flt_date_str or not flt_date_str.strip():
            return False
        parsed = normalize_flight_date(flt_date_str, datetime.now(ZoneInfo(TIMEZONE)))
        if not parsed:
            return False
        return parsed == report_date

    flights = [f for f in flights if _flight_date_matches(f.get("date", ""), date_dir)]

    # ── Enrich STD/ETD + DEST — always re-evaluate with confidence scoring ──
    enriched_count = 0
    for f in flights:
        flt = normalize_flight_number(f.get("flight") or "")
        if not flt:
            continue

        try:
            info, source_name = fetch_flight_info_with_fallbacks(
                flt,
                flight_date=normalize_flight_date(f.get("date", ""), datetime.now(ZoneInfo(TIMEZONE))) or date_dir,
                dep_iata="MCT",
                arr_iata=(f.get("destination") or "").strip() or None,
            )
            if not info:
                print(f"  [⚠ enrich] {flt}: no data from any source")
                continue

            current_std_etd = (f.get("std_etd") or "").strip()
            current_dest = (f.get("destination") or "").strip().upper()

            std = (info.get("std") or "").strip()
            etd = (info.get("etd") or "").strip()
            dest = (info.get("dest") or "").strip().upper()

            # ═══ FIX: Always prefer fresh high-confidence data ═══
            # Build new std_etd from fresh data
            if std and etd and std != etd:
                new_std_etd = f"{std}|{etd}"
            elif std:
                new_std_etd = std
            elif etd:
                new_std_etd = etd
            else:
                new_std_etd = current_std_etd  # keep existing only if nothing new

            # ═══ FIX: Overwrite even if current value exists ═══
            changed = False
            if new_std_etd and new_std_etd != current_std_etd:
                old = current_std_etd
                f["std_etd"] = new_std_etd
                f["enrichment_source"] = source_name
                f["last_enriched"] = datetime.now(ZoneInfo(TIMEZONE)).isoformat()
                changed = True
                print(f"  [CORRECTION] {flt} STD/ETD: {old!r} -> {new_std_etd!r} via {source_name}")

            if dest and dest != current_dest:
                f["destination"] = dest
                changed = True
                print(f"  [CORRECTION] {flt} DEST: {current_dest!r} -> {dest!r}")

            if changed:
                enriched_count += 1
                _update_flight_json(folder, f)

        except Exception as exc:
            print(f"  [report-enrich] {flt}: {exc}")

    print(f"Enriched {enriched_count} flight(s) via confidence-scored fallback chain.")

    shift_labels = {
        "shift1": {"ar": "صباح",     "en": "Morning",   "time": "06:00 – 15:00"},
        "shift2": {"ar": "ظهر/مساء", "en": "Afternoon", "time": "13:00 – 22:00"},
        "shift3": {"ar": "ليل",      "en": "Night",      "time": "21:00 – 06:00"},
    }
    sl            = shift_labels.get(shift, {"ar": shift, "en": shift, "time": ""})
    sl_en         = sl["en"]
    sl_time       = sl["time"]
    shift_label   = f"{sl_en} Shift — {sl_time}"
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
    import_roster = fetch_import_flight_dispatch_staff(date_dir, shift)

    # ── Supervisor name resolution (on-duty only) ──
    # Only use actual supervisors from roster — no acting/deputy logic.
    # If no supervisor found, leave both display and signature blank.
    on_duty_list = roster.get("on_duty", [])

    supervisor_name = ""

    EXCLUDED_SUPERVISOR_SNS = {"990737"}  # Said Al Amri — excluded from supervisor display

    for emp in on_duty_list:
        if "supervisor" in emp.get("dept", "").lower():
            if str(emp.get("sn", "")).strip() in EXCLUDED_SUPERVISOR_SNS:
                continue
            supervisor_name = emp.get("name", "").strip()
            if supervisor_name:
                break

    # Empty = leave blank (no acting supervisor)
    supervisor_display = supervisor_name if supervisor_name else ""
    signature_display = supervisor_display

    manpower_cols = _render_manpower_section(roster, supervisor_display, import_roster)

    # Format date for display
    # Night Shift تمتد عبر يومين (مثلاً 22/23 MAR)، لذلك نعرض كلا اليومين في الرأس
    try:
        d = datetime.strptime(date_dir, "%Y-%m-%d")
        if shift == "shift3":
            d_next = d + timedelta(days=1)
            date_display = f"{d.day}/{d_next.strftime('%d %b %Y').upper()}"
        else:
            date_display = d.strftime("%d %b %Y").upper()
    except Exception:
        date_display = date_dir

    # ── تجهيز JSON للرحلات المحلية كمتغير خارج الـ f-string ──
    local_flights_js = json.dumps(_load_local_db(), ensure_ascii=False)

    # ── تجهيز JSON لكل الموظفين لاستخدامه في autocomplete ──
    _staff_map: dict[str, str] = _load_manpower_json_staff_map()
    for _emp in roster.get("on_duty", []) + roster.get("on_leave", []):
        _sn = str(_emp.get("sn", "")).strip()
        _nm = str(_emp.get("name", "")).strip()
        if _sn and _nm and _sn not in _staff_map:
            _staff_map[_sn] = _nm
    for _emp in import_roster.get("fd_export", []) + import_roster.get("fd_import", []):
        _sn = str(_emp.get("sn", "")).strip()
        _nm = str(_emp.get("name", "")).strip()
        if _sn and _nm and _sn not in _staff_map:
            _staff_map[_sn] = _nm
    all_staff_js = json.dumps(_staff_map, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <title>Export Warehouse Activity Report — {date_display}</title>
  <style>
    *,*::before,*::after{{box-sizing:border-box;}}
    body{{margin:0;padding:0;background:#eef1f7;font-family:Calibri,Arial,sans-serif;}}
    .page-wrap{{background:#eef1f7;padding:16px 6px 40px;min-height:100vh;overflow-x:auto;}}
    #report-content{{width:900px;max-width:900px;background:#fff;border:1px solid #d0d5e8;margin:0 auto;table-layout:fixed;}}
    .btn-bar{{max-width:900px;margin:0 auto;display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;padding:10px 4px;position:sticky;bottom:0;z-index:9999;background:#eef1f7;border-top:1px solid #d0d5e8;box-shadow:0 -2px 8px rgba(11,58,120,.10);}}
    .btn-bar button{{font-family:Calibri,Arial,sans-serif;font-size:13px;font-weight:700;color:#fff;border:none;border-radius:8px;padding:10px 18px;cursor:pointer;}}
    /* جدول الأوفلود — يسمح بالتمرير الأفقي على الجوال */
    .offload-scroll{{overflow-x:auto;-webkit-overflow-scrolling:touch;width:100%;}}
    .offload-scroll table{{min-width:900px;width:100%;}}

    /* ══════════════ MOBILE ══════════════ */
    @media(max-width:700px){{
      /* عام */
      .page-wrap{{padding:4px 0 24px;}}
      #report-content{{border-radius:0;border-left:0;border-right:0;}}

      /* الهيدر */
      .hdr-inner{{display:block!important;}}
      .hdr-right{{display:none!important;}}  /* إخفاء "Transom Cargo" على الجوال */
      .hdr-title{{font-size:15px!important;line-height:1.3!important;}}
      .hdr-meta{{font-size:11px!important;margin-top:5px!important;line-height:1.6!important;}}

      /* padding الأقسام */
      .sec-pad{{padding-left:12px!important;padding-right:12px!important;}}

      /* أزرار الأسفل */
      .btn-bar{{justify-content:stretch;margin:0;gap:6px;padding:8px 6px;}}
      .btn-bar button{{flex:1 1 45%;font-size:12px;padding:10px 4px;border-radius:6px;}}

      /* التوقيع */
      .sig-wrap td{{display:block!important;width:100%!important;padding-bottom:12px!important;border-right:none!important;}}

      /* خط الأسماء في MANPOWER */
      .mp-list li{{font-size:12px!important;line-height:1.7!important;}}
      .mp-list span[style*="width:200px"]{{width:150px!important;}}

      /* footer */
      .report-footer-td{{padding:14px 12px!important;font-size:11px!important;}}

      /* section headers — أصغر */
      .sec-label{{font-size:11px!important;}}

      /* النصوص العامة */
      .sec-body{{font-size:12.5px!important;}}
      ul li{{font-size:12.5px!important;}}
    }}

    @media(max-width:400px){{
      .btn-bar button{{flex:1 1 100%;}}
      .hdr-title{{font-size:13px!important;}}
    }}
  </style>
</head>
<body>
<div class="page-wrap">
<div style="width:900px;max-width:900px;margin:0 auto;">

<table width="900" cellpadding="0" cellspacing="0" border="0" id="report-content"
       style="background-color:#ffffff; border:1px solid #d0d5e8;">

  <!-- ═══ HEADER ═══ -->
  <tr>
    <td data-email-keep="1" style="background-color:#1e5799; padding:0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
        <tr>
          <td width="8" data-email-keep="1" style="background-color:#f59e0b; font-size:1px; line-height:1px;">&nbsp;</td>
          <td data-email-keep="1" style="padding:20px 22px 18px 16px; background-color:#1e5799;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" class="hdr-inner" style="display:table;">
              <tr>
                <td data-email-keep="1" class="hdr-left" style="background-color:#1e5799;">
                  <div class="hdr-title" data-email-keep="1" style="font-family:Calibri,Arial,sans-serif; font-size:21px; font-weight:800; color:#ffffff; letter-spacing:0.5px; line-height:1.25;">
                    <span style="color:#ffffff;">&#9992;&nbsp; Export Warehouse Activity Report</span>
                  </div>
                  <div class="hdr-meta" data-email-keep="1" style="font-family:Calibri,Arial,sans-serif; font-size:13px; color:#ffffff; margin-top:7px; letter-spacing:0.2px;">
                    <span style="color:#ffffff;">Shift Date:&nbsp;</span><strong style="color:#fde68a; font-weight:700;">{date_display}</strong>
                    <span style="color:#ffffff;">&nbsp;&nbsp;|&nbsp;&nbsp;Time:&nbsp;</span><strong style="color:#fde68a; font-weight:700;">{sl_time} LT</strong>
                    <span style="color:#ffffff;">&nbsp;&nbsp;|&nbsp;&nbsp;</span><strong style="color:#fde68a; font-weight:700;">{sl_en} Shift</strong>
                  </div>
                </td>
                <td align="right" valign="middle" data-email-keep="1" class="hdr-right" style="padding-right:6px; background-color:#1e5799;">
                  <div data-email-keep="1" style="font-family:Calibri,Arial,sans-serif; font-size:12px; color:#93c5fd; line-height:1.5; text-align:right;">
                    <span style="color:#93c5fd;">Transom Cargo LLC.</span><br>
                    <strong style="color:#ffffff; font-weight:700;">Export Operations</strong>
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ═══ SECTION 1: OPERATIONAL ACTIVITIES ═══ -->
  <tr><td class="sec-pad" style="padding:18px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">1.&nbsp; OPERATIONAL ACTIVITIES</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <strong style="color:#0b3a78;">Load Plan:</strong>
      <ul id="ul-loadplan" data-flight-list="1" style="margin:4px 0 6px 22px; padding:0; color:#1b1f2a;"><li contenteditable="true" tabindex="50" style="outline:none; min-width:40px;">&nbsp;</li></ul>
      <br>
      <strong style="color:#0b3a78;">Advance Loading:</strong>
      <ul id="ul-advloading" data-flight-list="1" style="margin:4px 0 6px 22px; padding:0; color:#1b1f2a;"><li contenteditable="true" tabindex="51" style="outline:none; min-width:40px;">&nbsp;</li></ul>
      <br>
      <strong style="color:#0b3a78;">CSD Rescreening:</strong>
      <ul id="ul-csdrescreening" style="margin:4px 0 6px 22px; padding:0; color:#1b1f2a;"><li contenteditable="true" tabindex="52" style="outline:none; min-width:40px;">&nbsp;</li></ul>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 2: BRIEFINGS ═══ -->
  <tr><td class="sec-pad" style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">2.&nbsp; BRIEFINGS CONDUCTED</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <ul id="ul-briefings" style="margin:0 0 0 22px; padding:0; color:#1b1f2a;">
        <li contenteditable="true" tabindex="60" style="outline:none;">Safety toolbox conducted.</li>
        <li contenteditable="true" tabindex="61" style="outline:none;">ULD and net serviceability checked.</li>
        <li contenteditable="true" tabindex="62" style="outline:none;">Staff reminded about punctuality, proper cargo loading/counting, and no mobile phone use while driving.</li>
        <li contenteditable="true" tabindex="63" style="outline:none;">Briefing on <strong>EY CCS 25-011</strong> (correct pallet stack build-up) – read &amp; sign completed.</li>
        <li contenteditable="true" tabindex="64" style="outline:none;">Briefing on <strong>QR CGO CSA 09-25</strong> (weight scale discrepancies) – read &amp; sign completed.</li>
        <li contenteditable="true" tabindex="65" style="outline:none;">Process briefing for shipments UWS discrepancies.</li>
        <li contenteditable="true" tabindex="66" style="outline:none;"><strong>WY instruction:</strong> No shipment to CAI with handwritten labels – read &amp; sign completed.</li>
        <li contenteditable="true" tabindex="67" style="outline:none;">EY safety notification related to cargo handling discussed.</li>
        <li contenteditable="true" tabindex="68" style="outline:none;">Staff reminded to complete LMS training and informed about new roster.</li>
        <li contenteditable="true" tabindex="69" style="outline:none;">Instructions on attaching printed ULD tags on trolleys.</li>
        <li contenteditable="true" tabindex="70" style="outline:none;">Stationery logbook kept at supervisor desk.</li>
      </ul>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 3: FLIGHT PERFORMANCE ═══ -->
  <tr><td class="sec-pad" style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">3.&nbsp; FLIGHT PERFORMANCE</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      ✅&nbsp; All flights departed on time; no delay related to Cargo.
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 3B: OPERATIONAL NOTES ═══ -->
  <tr><td style="padding:0 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#5b6a8a;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#f4f5f9;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#3d4a63; letter-spacing:1px;">OPERATIONAL NOTES</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <ul id="ul-opnotes" style="margin:0 0 0 22px; padding:0; color:#1b1f2a;">
        <li contenteditable="true" tabindex="80" style="outline:none;">All flights departed on time as per RDM Mr. Saleh.</li>
        <li contenteditable="true" tabindex="81" style="outline:none;">DG embargo station check completed.</li>
        <li contenteditable="true" tabindex="82" style="outline:none;">Pigeonhole check done for any pending documents.</li>
      </ul>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 4: CHECKS & COMPLIANCE + OFFLOAD ═══ -->
  <tr><td class="sec-pad" style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">4.&nbsp; CHECKS &amp; COMPLIANCE</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <strong style="color:#0b3a78;">DG Check:</strong> DG Embargo station check done. AWB left behind: NIL.
    </div>
    <!-- عنوان جدول الأوفلود -->
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:14px;">
      <tr>
        <td width="4" style="background-color:#c2410c;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#fff7ed;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#c2410c; letter-spacing:1px;">OFFLOADING CARGO #</span>
        </td>
      </tr>
    </table>
    {offload_table_html}
    <div style="border-top:1px solid #e4e9f5; margin-top:16px;"></div>
  </td></tr>

  <!-- ═══ SECTION 5: SAFETY ═══ -->
  <tr><td class="sec-pad" style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#1a7a3c;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#edf7f1;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#1a7a3c; letter-spacing:1px;">5.&nbsp; SAFETY</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      Safety briefing conducted to all staff, drivers and porters. All checkers reminded to verify net expiration &amp; ULD serviceability.
      <br><br>
      <strong style="color:#1a7a3c;">✅ Note:</strong> All staff and drivers are wearing proper PPEs.
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 6: MANPOWER ═══ -->
  <tr><td class="sec-pad" style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">6.&nbsp; MANPOWER</span>
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
  <tr><td class="sec-pad" style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#1a7a3c;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#edf7f1;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#1a7a3c; letter-spacing:1px;">7.&nbsp; EQUIPMENT STATUS</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      ✅&nbsp; <strong>ALL EQUIPMENT ARE OK.</strong>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 8: HANDOVER ═══ -->
  <tr><td class="sec-pad" style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">8.&nbsp; HANDOVER DETAILS</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <ul id="ul-handover" style="margin:0 0 10px 22px; padding:0;">
        <li contenteditable="true" tabindex="90" style="outline:none;">READ AND SIGN.</li>
        <li contenteditable="true" tabindex="91" style="outline:none;">Shell &amp; Al-Maha Card Fuel.</li>
        <li contenteditable="true" tabindex="92" style="outline:none;">DIP MAIL Cage Keys.</li>
        <li contenteditable="true" tabindex="93" style="outline:none;">Supervisor mobile phone H/O in good condition.</li>
        <li contenteditable="true" tabindex="94" style="outline:none;">DSE RADIO.</li>
        <li contenteditable="true" tabindex="95" style="outline:none;">All trolleys arranged for early flight.</li>
      </ul>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 9: SPECIAL H/O ═══ -->
  <tr><td class="sec-pad" style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#0b3a78;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#eef3fc;">
          <span class="sec-label" style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#0b3a78; letter-spacing:1px;">9-&nbsp; Special H/O:</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
      <ul id="ul-special-handover" style="margin:0 0 10px 22px; padding:0;">
        <li contenteditable="true" tabindex="96" style="outline:none;">&nbsp;</li>
      </ul>
    </div>
    <div style="border-top:1px solid #e4e9f5; margin-top:14px;"></div>
  </td></tr>

  <!-- ═══ SECTION 10: OTHER ═══ -->
  <tr><td class="sec-pad" style="padding:14px 24px 0 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td width="4" style="background-color:#7a5200;">&nbsp;</td>
        <td style="padding:6px 10px; background-color:#fdf6ec;">
          <span style="font-family:Calibri,Arial,sans-serif; font-size:12px; font-weight:700; color:#7a5200; letter-spacing:1px;">10.&nbsp; OTHER</span>
        </td>
      </tr>
    </table>
    <div class="sec-body" style="font-family:Calibri,Arial,sans-serif; font-size:13.5px; color:#1b1f2a; line-height:1.7; margin-top:10px; padding:0 4px;">
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
                <td valign="top" style="padding-right:18px; border-right:2px solid #8b0000;">
                  <table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
                    <tr>
                      <td width="5" bgcolor="#8b0000" style="background-color:#8b0000; font-size:1px; line-height:1px;">&nbsp;</td>
                      <td bgcolor="#fdf0f0" style="background-color:#fdf0f0; padding:5px 10px;">
                        <font face="Arial Black,Arial,sans-serif" color="#8b0000">
                          <b style="font-size:22px; font-weight:900; letter-spacing:2px; color:#8b0000; font-family:Arial Black,Arial,sans-serif;">TRANSOM</b>
                        </font><br>
                        <font face="Arial,sans-serif" color="#888888">
                          <span style="font-size:9px; font-weight:600; letter-spacing:3px; color:#888888;">CARGO</span>
                        </font>
                      </td>
                    </tr>
                  </table>
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

        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style="background-color:#0b3a78; height:5px; font-size:1px; line-height:1px;">&nbsp;</td>
  </tr>

</table>

<!-- ═══ BUTTONS BAR ═══ -->
<div class="btn-bar" data-no-copy="1">
  <button id="btn-copy" type="button" data-no-copy="1" style="background:#0b3a78;box-shadow:0 2px 8px rgba(11,58,120,.25);">📋 Copy Report</button>
  <button id="btn-cloud-sync" type="button" data-no-copy="1" style="background:#0f766e;box-shadow:0 2px 8px rgba(15,118,110,.25);">☁️ Cloud Sync</button>
  <button id="btn-manage-emails" type="button" data-no-copy="1" style="background:#475569;box-shadow:0 2px 8px rgba(71,85,105,.25);">📝 Edit Email List</button>
  <button id="btn-email" type="button" data-no-copy="1" style="background:#c2410c;box-shadow:0 2px 8px rgba(194,65,12,.25);">✉️ Send Email Now</button>
</div>

</div><!-- /max-width wrapper -->
</div><!-- /page-wrap -->

<script>
/* ── Config injected by Python build ── */
window._AIRLABS_KEY        = '';          /* ضع مفتاح AirLabs هنا إذا توفّر */
window._LOCAL_MCT_FLIGHTS  = {local_flights_js};
window._ALL_STAFF          = {all_staff_js};
window._REPORT_CLOUD_PATH  = 'docs/data/report_edits/{date_dir}/{shift}.json';
</script>

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

  /* ── قوائم قابلة للتعديل: إضافة وحذف عناصر ── */
  function addListItem(ulId){{
    var ul = document.getElementById(ulId);
    if(!ul) return;
    var li = createEditableListItem(ul);
    ul.appendChild(li);
    _attachLiEvents(li);
    focusEditableListItem(li, isManpowerListUl(ul));
    if(typeof triggerAutosave === 'function') triggerAutosave();
  }}

  function _attachLiEvents(li){{
    if(!li || li._baseLiSetup) return;
    li._baseLiSetup = true;
    var ul = li.parentElement;
    if(!ul) return;

    if(isManpowerListUl(ul)){{
      if(typeof window.setupManpowerLi === 'function') window.setupManpowerLi(li);
      return;
    }}

    li.addEventListener('keydown', function(e){{
      if(e.key === 'Enter'){{
        e.preventDefault();
        var newLi = createEditableListItem(li.parentElement);
        li.parentElement.insertBefore(newLi, li.nextSibling);
        _attachLiEvents(newLi);
        focusEditableListItem(newLi, false);
        return;
      }}
      if((e.key === 'Backspace'||e.key==='Delete') && (li.innerText.trim()===''||li.innerText.trim()===' ')){{
        if(li.parentElement && li.parentElement.children.length > 1){{
          e.preventDefault();
          var prev = li.previousElementSibling || li.nextElementSibling;
          li.parentElement.removeChild(li);
          if(prev) prev.focus();
        }}
      }}
    }});

    if(ul && (ul.dataset.flightList || ul.getAttribute('data-flight-list'))){{
      var isLP = !!(ul && ul.id === 'ul-loadplan');
      if(typeof window.setupFlightListItem === 'function') window.setupFlightListItem(li, isLP);
    }}
  }}

  /* ربط الـ li الموجودة مسبقاً بالأحداث عند تحميل الصفحة */
  document.addEventListener('DOMContentLoaded', function(){{
    var editableLists = [
      'ul-loadplan','ul-advloading','ul-csdrescreening','ul-handover','ul-special-handover',
      'ul-supervisors','ul-ctu','ul-inventory','ul-support','ul-fd-export','ul-fd-import',
      'ul-sickleave','ul-annualleave','ul-trainee','ul-overtime'
    ];
    editableLists.forEach(function(id){{
      var ul = document.getElementById(id);
      if(!ul) return;
      Array.from(ul.querySelectorAll('li[contenteditable]')).forEach(function(li){{
        _attachLiEvents(li);
      }});
    }});
    /* أيضاً أقسام الأقسام الديناميكية */
    document.querySelectorAll('ul[id^="ul-dept-"]').forEach(function(ul){{
      Array.from(ul.querySelectorAll('li[contenteditable]')).forEach(function(li){{
        _attachLiEvents(li);
      }});
    }});
  }});
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

  /* ══════════════════════════════════════════════════════════
     buildReportHtml — HTML نظيف للإيميل/النسخ
     القاعدة: البنر (data-email-keep) يحتفظ بألوانه دائماً
     باقي العناصر: خلفيات داكنة → فاتحة، نص أبيض → داكن
     ══════════════════════════════════════════════════════════ */
  function buildReportHtml(){{
    var el = document.getElementById('report-content');
    if(!el) return null;

    /* خطوة 1: نسخ بسيط — inline styles موجودة في HTML مباشرةً */
    var clone = el.cloneNode(true);

    function normC(s){{ return (s||'').toLowerCase().replace(/\\s/g,''); }}
    var DARK=[
      'rgb(11,58,120)','#0b3a78','rgb(30,64,175)','#1e40af',
      'rgb(37,99,235)','#2563eb','rgb(10,31,82)','#0a1f52',
      'rgb(26,110,207)','#1a6ecf'
    ];
    function isDark(s){{
      var n=normC(s);
      for(var i=0;i<DARK.length;i++) if(n===normC(DARK[i])) return true;
      return false;
    }}
    function isWhite(s){{
      var n=normC(s);
      return n==='#ffffff'||n==='#fff'||n==='rgb(255,255,255)';
    }}
    function inBanner(node){{
      var p=node;
      while(p&&p!==clone){{
        if(p.getAttribute&&p.getAttribute('data-email-keep')==='1') return true;
        p=p.parentElement;
      }}
      return false;
    }}

    /* خطوة 2: تحويل العناصر خارج البنر فقط */
    clone.querySelectorAll('*').forEach(function(e){{
      if(inBanner(e)) return;
      if(isDark(e.style.backgroundColor)){{
        e.style.backgroundColor='#eef3fc';
        e.style.color='#0b3a78';
      }}
      var bg=e.style.background||'';
      if(bg.toLowerCase().indexOf('linear-gradient')!==-1){{
        e.style.background='';
        e.style.backgroundColor='#eef3fc';
        e.style.color='#0b3a78';
      }}
      if(isWhite(e.style.color)) e.style.color='#1b1f2a';
    }});

    /* خطوة 3: البنر — تثبيت ألوانه بشكل قاطع بعد خطوة 2 */
    clone.querySelectorAll('[data-email-keep]').forEach(function(e){{
      var bg=normC(e.style.backgroundColor);
      /* الشريط الذهبي يبقى */
      if(bg&&bg!==normC('#f59e0b')&&bg!=='rgb(245,158,11)')
        e.style.backgroundColor='#1e5799';
    }});
    var bRoot=clone.querySelector('[data-email-keep="1"]');
    if(bRoot){{
      /* تطبيق اللون الأبيض على الجذر حتى يرثه كل النص غير المُصنَّف */
      bRoot.style.color='#ffffff';
      bRoot.querySelectorAll('*').forEach(function(e){{
        var c=normC(e.style.color);
        /* ذهبي → يبقى */
        if(c===normC('#fde68a')||c==='rgb(253,230,138)') return;
        /* أزرق فاتح → يبقى */
        if(c===normC('#93c5fd')||c==='rgb(147,197,253)') return;
        /* فرض الأبيض على الجميع بغض النظر عن القيمة الحالية */
        e.style.color='#ffffff';
      }});
    }}

    /* خطوة 4: إصلاح الجداول */
    clone.querySelectorAll('table').forEach(function(t){{
      t.setAttribute('cellpadding','0');
      t.setAttribute('cellspacing','0');
      t.setAttribute('border','0');
      t.setAttribute('role','presentation');
      t.style.borderCollapse='collapse';
      if(t.id==='report-content'){{ t.setAttribute('width','900'); t.style.width='900px'; t.style.maxWidth='900px'; t.style.tableLayout='fixed'; }} else if(!t.style.width) {{ t.style.width='100%'; }}
    }});
    clone.querySelectorAll('td,th').forEach(function(c){{
      if(!c.style.verticalAlign) c.style.verticalAlign='top';
    }});

    /* خطوة 5: تنظيف */
    clone.querySelectorAll('[contenteditable]').forEach(function(e){{
      e.removeAttribute('contenteditable');
      e.style.outline='none';
    }});
    clone.querySelectorAll('[tabindex]').forEach(function(e){{
      e.removeAttribute('tabindex');
    }});
    /* إزالة أزرار + Add من النسخة المُرسَلة */
    clone.querySelectorAll('button').forEach(function(e){{
      if(e.textContent.trim().indexOf('Add')!==-1 || e.textContent.trim().indexOf('+')!==-1)
        e.parentNode && e.parentNode.removeChild(e);
    }});
    /* إزالة عناصر data-no-copy (أزرار النسخ/الإرسال) من نسخة الإيميل */
    clone.querySelectorAll('[data-no-copy]').forEach(function(e){{
      e.parentNode && e.parentNode.removeChild(e);
    }});
    var br=clone.querySelector('#back-link-row');
    if(br && br.parentNode) br.parentNode.removeChild(br);

    return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><style>body{{background:#eef1f7;font-family:Calibri,Arial,sans-serif;margin:0;padding:10px 0;-webkit-text-size-adjust:100%;}}table{{border-collapse:collapse;}}img{{max-width:100%;height:auto;display:block;}}</style></head><body>'+clone.outerHTML+'</body></html>';
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
      if(!/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(s)) return;
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
          if(!/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(val)){{ alert('Enter a valid email address'); return; }}
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
    openRecipientsManager('send').then(function(result){{
      if(!result || !result.selected || !result.selected.length) return;
      var ok = confirm('إرسال تقرير هذه المناوبة بالإيميل الآن؟\\n\\nShift: {shift}\\nDate: {date_dir}\\nRecipients: ' + result.selected.join(', '));
      if(!ok) return;

      var btn = document.getElementById('btn-email');
      if(!btn) return;
      btn.innerText = '⏳ Saving edits…';
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

      /* ── Step 1: Capture DOM edits (inline styles, clean email-safe HTML) ── */
      var editedHtml = buildReportHtml();
      if(!editedHtml){{
        btn.innerText = '❌ No report found';
        btn.style.background = '#dc2626';
        resetButton(btn, '✉️ Send Email Now', '#c2410c', 3000);
        return;
      }}

      /* ── Step 2: Get current file SHA then upload the edited HTML to repo ── */
      var filePath = 'docs/{date_dir}/{shift}/index.html';
      var apiBase = 'https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/contents/';
      var headers = {{
        'Accept':'application/vnd.github+json',
        'Authorization':'Bearer ' + pat,
        'Content-Type':'application/json'
      }};

      btn.innerText = '⏳ Uploading edits…';

      fetch(apiBase + filePath + '?t=' + Date.now(), {{headers:headers}})
      .then(function(r){{
        if(r.status === 401){{ localStorage.removeItem('gh_pat'); throw new Error('AUTH'); }}
        if(!r.ok && r.status !== 404) throw new Error('GET_SHA_' + r.status);
        return r.ok ? r.json() : null;
      }})
      .then(function(fileInfo){{
        var sha = fileInfo ? fileInfo.sha : undefined;
        /* Use editedHtml — includes user edits + inline styles + email-safe colors */
        var encoded = btoa(unescape(encodeURIComponent(editedHtml)));
        var putBody = {{
          message: 'Update report with manual edits ({date_dir}/{shift})',
          content: encoded,
          branch: 'main'
        }};
        if(sha) putBody.sha = sha;

        return fetch(apiBase + filePath, {{
          method: 'PUT',
          headers: headers,
          body: JSON.stringify(putBody)
        }});
      }})
      .then(function(r){{
        if(r.status === 401){{ localStorage.removeItem('gh_pat'); throw new Error('AUTH'); }}
        if(!r.ok) throw new Error('PUT_' + r.status);
        btn.innerText = '⏳ Sending email…';

        /* ── Step 3: Trigger the dispatch to send email ── */
        return fetch('https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/dispatches', {{
          method:'POST',
          headers: headers,
          body:JSON.stringify({{
            event_type:'send_report_now',
            client_payload:{{date_dir:'{date_dir}', shift:'{shift}', recipients: result.selected}}
          }})
        }});
      }})
      .then(function(r){{
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
      }})
      .catch(function(err){{
        var msg = (err && err.message) || 'Error';
        if(msg === 'AUTH'){{
          btn.innerText = '❌ Token expired — retry';
        }} else {{
          btn.innerText = '❌ ' + msg;
        }}
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

/* ══════════════════════════════════════════════════════
   SMART AUTOCOMPLETE — Load Plan / Advance Loading / Offload Table / MANPOWER
   ══════════════════════════════════════════════════════ */
(function(){{

  /* ── 0) Config & Helpers ── */
  var AIRLABS_KEY = (typeof window._AIRLABS_KEY !== 'undefined') ? window._AIRLABS_KEY : '';
  var FLT_RE = /^([A-Z]{{2,3}})(\d{{1,5}})$/i;

  var LOCAL_FLIGHTS = {{}};
  try {{ var lf = window._LOCAL_MCT_FLIGHTS; if(lf) LOCAL_FLIGHTS = lf; }} catch(ex) {{}}

  function todayFull() {{
    var now = new Date(Date.now() + 4*3600*1000);
    var d   = now.getUTCDate();
    var mon = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'][now.getUTCMonth()];
    var yr  = now.getUTCFullYear();
    return (d<10?'0'+d:d)+mon+yr;
  }}
  function todayISO() {{
    return new Date(Date.now() + 4*3600*1000).toISOString().slice(0,10);
  }}

  function forceUpper(el) {{
    var old = cleanEditableText(el);
    var up  = old.toUpperCase();
    if(old !== up) {{
      el.textContent = up;
      setCursorEnd(el);
    }}
    return up;
  }}

  function cleanEditableText(el) {{
    return String((el && el.innerText) || '')
      .replace(/[•·▪◦●]/g,' ')
      .replace(/\u00a0/g,' ')
      .replace(/^[\s\-–—]+/, '')
      .replace(/\s+/g,' ')
      .trim();
  }}

  function isEffectivelyEmptyManpowerLi(li) {{
    if(!li) return true;
    var txt = getManpowerText ? getManpowerText(li) : cleanEditableText(li);
    txt = String(txt || '').replace(/\u200b/g, '').replace(/\u00a0/g, ' ').trim();
    return !txt;
  }}

  function removeManpowerLiIfEmpty(li) {{
    if(!li || !li.parentElement) return false;
    if(!isEffectivelyEmptyManpowerLi(li)) return false;
    var ul = li.parentElement;
    if(ul.children.length <= 1) return false;
    var prev = li.previousElementSibling;
    var next = li.nextElementSibling;
    ul.removeChild(li);
    var target = prev || next;
    if(target) focusEditableListItem(target, true);
    if(typeof triggerAutosave === 'function') triggerAutosave();
    return true;
  }}

  function normalizeFlightTyped(text) {{
    var raw = String(text || '').toUpperCase().replace(/[^A-Z0-9]/g,'');
    if(!raw) return '';
    if(/^\d{1,5}$/.test(raw)) return raw;
    return raw;
  }}

  function flightCandidatesFromTyped(typed) {{
    var out = [];
    var seen = {{}};
    var digitsOnly = /^\d{{1,5}}$/.test(typed || '');

    Object.keys(LOCAL_FLIGHTS || {{}}).forEach(function(code) {{
      var key = String(code || '').toUpperCase().replace(/\s+/g,'');
      if(!key || seen[key]) return;
      var m = key.match(/^([A-Z]{{2,3}})(\d{{1,5}})$/);
      if(!m) return;
      var num = m[2];
      if((digitsOnly && num.indexOf(typed) === 0) || (!digitsOnly && key.indexOf(typed) === 0)) {{
        seen[key] = true;
        out.push(key);
      }}
    }});

    out.sort(function(a,b) {{
      var ad = a.replace(/^[A-Z]+/, '');
      var bd = b.replace(/^[A-Z]+/, '');
      if(ad === typed && bd !== typed) return -1;
      if(bd === typed && ad !== typed) return 1;
      if(a === typed && b !== typed) return -1;
      if(b === typed && a !== typed) return 1;
      return a.localeCompare(b);
    }});
    return out.slice(0, 8);
  }}

  function showGhost(el, txt) {{
    var g = el.parentElement && el.parentElement.querySelector(':scope > .ghost-tip');
    if(!g && el.tagName === 'TD') g = el.querySelector('.ghost-tip');
    if(!g) {{
      g = document.createElement('span');
      g.className = 'ghost-tip';
      g.style.cssText = 'color:#94a3b8;font-style:italic;font-size:11px;display:block;pointer-events:none;user-select:none;white-space:nowrap;';
      if(el.tagName === 'TD') {{ el.appendChild(g); }}
      else if(el.parentElement) {{ el.parentElement.insertBefore(g, el.nextSibling); }}
    }}
    g.textContent = '\u2192 '+txt+' (Tab)';
  }}
  function removeGhost(el) {{
    [el, el.parentElement].forEach(function(p) {{
      if(!p) return;
      p.querySelectorAll('.ghost-tip').forEach(function(g) {{ g.remove(); }});
    }});
  }}

  /* ── Flight info fetch (local DB first, then AirLabs) ── */
  function parseTime(val) {{
    if(!val) return '';
    val = String(val).trim();
    var iso = val.match(/T(\d{{2}}:\d{{2}})/);
    if(iso) return iso[1];
    var hm = val.match(/^(\d{{1,2}}:\d{{2}})$/);
    if(hm) return hm[1];
    var d4 = val.match(/^(\d{{4}})$/);
    if(d4) return d4[1].slice(0,2)+':'+d4[1].slice(2);
    return '';
  }}

  function fetchFlightInfo(flt, isoDate, cb) {{
    var local = LOCAL_FLIGHTS[flt.toUpperCase()];
    if(local && (local.dest || local.std)) {{
      cb({{
        std:  parseTime(local.std  || ''),
        etd:  parseTime(local.etd  || ''),
        dest: (local.dest || '').toUpperCase()
      }});
      return;
    }}
    if(!AIRLABS_KEY) {{ cb(null); return; }}
    var url = 'https://airlabs.co/api/v9/schedules?api_key='+AIRLABS_KEY
            + '&flight_iata='+flt.toUpperCase()+'&dep_iata=MCT';
    if(isoDate) url += '&flight_date='+isoDate;
    fetch(url, {{cache:'no-cache'}})
      .then(function(r) {{ return r.json(); }})
      .then(function(j) {{
        var rows = j.response || j.data || [];
        if(!rows.length) {{
          return fetch('https://airlabs.co/api/v9/flights?api_key='+AIRLABS_KEY
                      +'&flight_iata='+flt.toUpperCase()+'&dep_iata=MCT', {{cache:'no-cache'}})
            .then(function(r2) {{ return r2.json(); }})
            .then(function(j2) {{
              var r2rows = j2.response || j2.data || [];
              if(!r2rows.length) {{ cb(null); return; }}
              doRow(r2rows[0]);
            }});
        }}
        doRow(rows[0]);
      }})
      .catch(function() {{ cb(null); }});
    function doRow(row) {{
      cb({{
        std:  parseTime(row.dep_scheduled || ''),
        etd:  parseTime(row.dep_estimated || ''),
        dest: (row.arr_iata||'').toUpperCase()
      }});
    }}
  }}

  /* ══════════════════════════════════════
     A) Load Plan & Advance Loading
        • كتابة رقم رحلة → ghost suggestion
        • Tab → تعبئة النص كاملاً
        • Enter → سطر جديد بنفس الإعداد
     ══════════════════════════════════════ */
  function setupFlightListItem(li, isLoadPlan) {{
    if(li._flightSetup) return;   /* منع التكرار */
    li._flightSetup = true;

    var _ghost = '';
    var _reqId = 0;

    function applyFlightGhost(flt, info) {{
      var dest = (info && info.dest) ? info.dest : '???';
      _ghost = isLoadPlan
        ? (flt + '/' + dest + '/' + todayFull() + '.')
        : (flt + '/' + dest + '/' + todayFull() + ' \u2014 Completed as per plan');
      li.dataset.ghostVal = _ghost;
      showGhost(li, _ghost);
    }}

    function updateFlightSuggestion() {{
      var typed = normalizeFlightTyped(cleanEditableText(li));
      if(!typed) {{ removeGhost(li); li.dataset.ghostVal=''; return; }}

      var candidates = flightCandidatesFromTyped(typed);
      var flt = '';
      if(FLT_RE.test(typed)) flt = typed;
      else if(candidates.length === 1) flt = candidates[0];
      else if(candidates.length > 1) {{
        li.dataset.ghostVal = candidates[0];
        showGhost(li, candidates.slice(0, 4).join('   '));
        return;
      }} else {{
        removeGhost(li); li.dataset.ghostVal=''; return;
      }}

      var reqId = ++_reqId;
      showGhost(li, flt + '/???/' + todayFull() + '\u2026');
      fetchFlightInfo(flt, todayISO(), function(info) {{
        if(reqId !== _reqId) return;
        applyFlightGhost(flt, info);
      }});
    }}

    li.addEventListener('focus', updateFlightSuggestion);
    li.addEventListener('input', updateFlightSuggestion);
    li.addEventListener('keyup', updateFlightSuggestion);
    li.addEventListener('paste', function() {{ setTimeout(updateFlightSuggestion, 30); }});

    /* keydown: Tab → قبول Ghost | Enter → سطر جديد */
    li.addEventListener('keydown', function(e) {{
      if(e.key === 'Tab' && li.dataset.ghostVal) {{
        e.preventDefault();
        var acceptVal = li.dataset.ghostVal;
        var candidates = flightCandidatesFromTyped(normalizeFlightTyped(cleanEditableText(li)));
        if(!/\//.test(acceptVal) && candidates.length) {{
          acceptVal = candidates[0];
          li.innerText = acceptVal;
          setCursorEnd(li);
          li.dataset.ghostVal = '';
          updateFlightSuggestion();
          return;
        }}
        li.innerText = acceptVal;
        li.dataset.ghostVal = '';
        removeGhost(li);
        setCursorEnd(li);
        if(typeof triggerAutosave==='function') triggerAutosave();

      }} else if(e.key === 'Enter') {{
        e.preventDefault();
        removeGhost(li);
        li.dataset.ghostVal = '';
        var ul = li.parentElement;
        if(!ul) return;
        var newLi = document.createElement('li');
        newLi.contentEditable = 'true';
        newLi.style.outline = 'none';
        newLi.innerHTML = '&nbsp;';
        ul.insertBefore(newLi, li.nextSibling);
        setupFlightListItem(newLi, isLoadPlan);   /* ← يربط على الـ li الجديدة مباشرة */
        newLi.focus();
        setCursorEnd(newLi);
      }}
    }});

    li.addEventListener('blur', function() {{ setTimeout(function() {{ if(document.activeElement !== li) removeGhost(li); }}, 120); }});
  }}

  function setCursorEnd(el) {{
    try {{
      var r = document.createRange();
      r.selectNodeContents(el);
      r.collapse(false);
      window.getSelection().removeAllRanges();
      window.getSelection().addRange(r);
    }} catch(ex) {{}}
  }}

  function setCursorStart(el) {{
    try {{
      var r = document.createRange();
      r.selectNodeContents(el);
      r.collapse(true);
      window.getSelection().removeAllRanges();
      window.getSelection().addRange(r);
    }} catch(ex) {{}}
  }}

  function getTextSelectionOffsets(el) {{
    try {{
      var sel = window.getSelection && window.getSelection();
      if(!sel || !sel.rangeCount) return {{start:null, end:null}};
      var range = sel.getRangeAt(0);
      if(!el || !el.contains(range.startContainer) || !el.contains(range.endContainer)) return {{start:null, end:null}};
      var preStart = range.cloneRange();
      preStart.selectNodeContents(el);
      preStart.setEnd(range.startContainer, range.startOffset);
      var preEnd = range.cloneRange();
      preEnd.selectNodeContents(el);
      preEnd.setEnd(range.endContainer, range.endOffset);
      return {{start: preStart.toString().length, end: preEnd.toString().length}};
    }} catch(ex) {{
      return {{start:null, end:null}};
    }}
  }}

  function setCursorByTextOffset(el, offset) {{
    try {{
      if(!el) return;
      var walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null);
      var remaining = Math.max(0, Number(offset || 0));
      var node = null;
      while((node = walker.nextNode())) {{
        var len = (node.nodeValue || '').length;
        if(remaining <= len) {{
          var r = document.createRange();
          r.setStart(node, remaining);
          r.collapse(true);
          var sel = window.getSelection();
          sel.removeAllRanges();
          sel.addRange(r);
          return;
        }}
        remaining -= len;
      }}
      setCursorEnd(el);
    }} catch(ex) {{
      setCursorEnd(el);
    }}
  }}

  function isManpowerListUl(ul) {{
    if(!ul || !ul.id) return false;
    return ul.id.indexOf('ul-dept-') === 0 || MP_IDS.indexOf(ul.id) !== -1;
  }}

  function createEditableListItem(ul) {{
    var li = document.createElement('li');
    li.contentEditable = 'true';
    li.style.outline = 'none';
    li.style.minWidth = '40px';
    if(isManpowerListUl(ul)) {{
      li.textContent = '​';
    }} else {{
      li.innerHTML = '&nbsp;';
    }}
    return li;
  }}

  function focusEditableListItem(li, atStart) {{
    if(!li) return;
    li.focus();
    if(atStart) setCursorStart(li); else setCursorEnd(li);
  }}

  /* ربط القوائم الموجودة عند تحميل الصفحة */
  function initFlightLists() {{
    var ulLP  = document.getElementById('ul-loadplan');
    var ulADV = document.getElementById('ul-advloading');
    if(ulLP)  ulLP.querySelectorAll('li').forEach(function(li)  {{ setupFlightListItem(li, true);  }});
    if(ulADV) ulADV.querySelectorAll('li').forEach(function(li) {{ setupFlightListItem(li, false); }});
  }}
  window.setupFlightListItem = setupFlightListItem;
  window.initFlightLists = initFlightLists;
  if(document.readyState === 'loading') {{
    document.addEventListener('DOMContentLoaded', initFlightLists);
  }} else {{
    initFlightLists();
  }}

  /* "+ Add" button — نُعيد تعريف addListItem ليُضيف autocomplete */
  var _orig = window.addListItem;
  window.addListItem = function(ulId) {{
    _orig(ulId);
    var ul = document.getElementById(ulId);
    if(!ul) return;
    var last = ul.lastElementChild;
    if(!last) return;
    var isLP = (ulId === 'ul-loadplan');
    var isADV= (ulId === 'ul-advloading');
    if(isLP || isADV) {{
      setupFlightListItem(last, isLP);
    }} else if(typeof setupManpowerLi === 'function') {{
      setupManpowerLi(last);
    }}
  }};

  /* ══════════════════════════════════════
     B) جدول الأوفلود
        • click/focus على date → يكمل التاريخ
        • input على flight → يجلب STD+DEST
     ══════════════════════════════════════ */
  function setupTableCell(td) {{
    if(td._cellSetup) return;   /* منع التكرار */
    td._cellSetup = true;

    td.addEventListener('keydown', function(e) {{
      if(e.key !== 'Tab' || e.shiftKey) return;
      var row = td.closest('tr');
      var tbody = row && row.parentElement;
      if(!row || !tbody || tbody.id !== 'offload-tbody') return;
      var editable = Array.from(row.querySelectorAll('td[contenteditable="true"]'));
      if(!editable.length || editable[editable.length - 1] !== td) return;
      var rows = Array.from(tbody.querySelectorAll('tr')).filter(function(r) {{
        return !!r.querySelector('td[contenteditable="true"]');
      }});
      if(!rows.length || rows[rows.length - 1] !== row) return;
      e.preventDefault();
      var tr = appendOffloadRow(false);
      if(tr) {{
        setTimeout(function() {{
          var firstCell = tr.querySelector('td[contenteditable="true"]');
          if(firstCell) {{
            firstCell.focus();
            setCursorEnd(firstCell);
          }}
        }}, 0);
      }}
    }});

    var col = td.dataset.col;
    if(!col) return;

    if(col === 'date') {{
      function autoDate() {{
        var txt = td.innerText.replace(/ /g,'').trim();
        if(!txt) {{
          td.innerText = todayFull();
          if(typeof triggerAutosave==='function') triggerAutosave();
        }}
      }}
      td.addEventListener('click',  autoDate);
      td.addEventListener('focus',  autoDate);
      td.addEventListener('input',  function() {{ forceUpper(td); }});
    }}

    if(col === 'flight') {{
      var _lastFlt = '';
      function normalizeFlightCell() {{
        var raw = cleanEditableText(td).toUpperCase().replace(/[^A-Z0-9]/g,'');
        if((td.textContent || '') !== raw) {{
          td.textContent = raw;
          setCursorEnd(td);
        }}
        return raw;
      }}
      function onFlightInput() {{
        var flt = normalizeFlightCell();
        if(flt === _lastFlt) return;   /* لا تغيير حقيقي */
        _lastFlt = flt;
        if(!FLT_RE.test(flt)) {{ removeGhost(td); return; }}
        var row      = td.closest('tr');
        var dateCell = row && row.querySelector('[data-col="date"]');
        var stdCell  = row && row.querySelector('[data-col="std"]');
        var destCell = row && row.querySelector('[data-col="dest"]');
        var emailCell= row && row.querySelector('[data-col="email"]');
        /* تعبئة التاريخ تلقائياً إذا كان فارغاً */
        if(dateCell) {{
          var dv = dateCell.innerText.replace(/ /g,'').trim();
          if(!dv) {{ dateCell.innerText = todayFull(); dv = todayFull(); }}
        }}
        var isoDate = todayISO();
        if(dateCell) {{
          var dv2 = (dateCell.innerText||'').replace(/ /g,'').trim().toUpperCase();
          var mons = {{JAN:1,FEB:2,MAR:3,APR:4,MAY:5,JUN:6,JUL:7,AUG:8,SEP:9,OCT:10,NOV:11,DEC:12}};
          var dm = dv2.match(/^(\d{{1,2}})([A-Z]{{3}})(\d{{2,4}})?$/);
          if(dm) {{
            var yr = dm[3] ? (dm[3].length===2 ? 2000+parseInt(dm[3]) : parseInt(dm[3])) : new Date().getFullYear();
            isoDate = yr+'-'+(mons[dm[2]]||1).toString().padStart(2,'0')+'-'+dm[1].padStart(2,'0');
          }}
        }}
        showGhost(td, 'fetching '+flt+'…');
        fetchFlightInfo(flt, isoDate, function(info) {{
          removeGhost(td);
          if(!info) return;
          if(stdCell) {{
            var cur = stdCell.innerText.replace(/ /g,'').trim();
            if(!cur || cur === ' ') {{
              var std = info.std||'', etd = info.etd||'';
              stdCell.innerText = (std && etd && std!==etd) ? std+' | '+etd : (std||etd||'');
            }}
          }}
          if(destCell) {{
            var dc = destCell.innerText.replace(/ /g,'').trim();
            if((!dc) && info.dest) destCell.innerText = info.dest;
          }}
          if(typeof triggerAutosave==='function') triggerAutosave();
        }});
      }}
      td.addEventListener('focus',  normalizeFlightCell);
      td.addEventListener('input',  onFlightInput);
      td.addEventListener('keyup',  onFlightInput);  /* يشتغل عند paste أيضاً */
      td.addEventListener('paste', function() {{ setTimeout(onFlightInput, 50); }});
    }}
  }}


  function reindexOffloadRows() {{
    var tbody = document.getElementById('offload-tbody');
    if(!tbody) return;
    var rows = Array.from(tbody.querySelectorAll('tr')).filter(function(row) {{
      return !!row.querySelector('td[contenteditable="true"]');
    }});
    rows.forEach(function(row, idx) {{
      var first = row.querySelector('td');
      if(first) first.innerHTML = '<strong>' + (idx + 1) + '</strong>';
    }});
    var ti = 1;
    rows.forEach(function(row) {{
      row.querySelectorAll('td[contenteditable="true"]').forEach(function(cell) {{
        cell.setAttribute('tabindex', String(ti++));
      }});
    }});
  }}

  function createOffloadRow() {{
    var tr = document.createElement('tr');
    var tdStyle = 'padding:7px 6px;border:1px solid #d0d9ee;font-size:12px;font-family:Calibri,Arial,sans-serif;color:#1b1f2a;background:#ffffff;text-align:center;vertical-align:middle;';
    var cols = [
      {{ key: 'date' }},
      {{ key: 'flight' }},
      {{ key: 'std' }},
      {{ key: 'dest' }},
      {{ key: 'email' }},
      {{}}, {{}}, {{}}, {{}}, {{}}, {{}}
    ];
    var idxCell = document.createElement('td');
    idxCell.setAttribute('style', tdStyle);
    idxCell.innerHTML = '<strong>0</strong>';
    tr.appendChild(idxCell);
    cols.forEach(function(col) {{
      var td = document.createElement('td');
      td.setAttribute('style', tdStyle);
      td.setAttribute('contenteditable', 'true');
      if(col.key) td.dataset.col = col.key;
      td.innerHTML = '&nbsp;';
      tr.appendChild(td);
    }});
    return tr;
  }}

  function appendOffloadRow(focusFirst) {{
    var tbody = document.getElementById('offload-tbody');
    if(!tbody) return null;
    var tr = createOffloadRow();
    tbody.appendChild(tr);
    tr.querySelectorAll('td[contenteditable="true"]').forEach(setupTableCell);
    reindexOffloadRows();
    if(focusFirst) {{
      var firstCell = tr.querySelector('td[contenteditable="true"]');
      if(firstCell) {{
        firstCell.focus();
        setCursorEnd(firstCell);
      }}
    }}
    if(typeof triggerAutosave === 'function') triggerAutosave();
    return tr;
  }}

  function initTableCells() {{
    document.querySelectorAll('#offload-tbody td[contenteditable="true"]').forEach(setupTableCell);
    reindexOffloadRows();
    /* MutationObserver — يُفعّل autocomplete على الصفوف الجديدة في الجدول */
    var tbody = document.getElementById('offload-tbody');
    if(tbody) {{
      new MutationObserver(function(muts) {{
        muts.forEach(function(m) {{
          m.addedNodes.forEach(function(n) {{
            if(n.nodeType===1) {{
              n.querySelectorAll('td[contenteditable="true"]').forEach(setupTableCell);
            }}
          }});
        }});
      }}).observe(tbody, {{childList:true}});
    }}
  }}
  if(document.readyState === 'loading') {{
    document.addEventListener('DOMContentLoaded', initTableCells);
  }} else {{
    initTableCells();
  }}

  /* ══════════════════════════════════════
     C) MANPOWER — SN Autocomplete
     ══════════════════════════════════════ */
  function extractSnNamePair(text) {{
    var txt = String(text || '').replace(/ /g,' ').replace(/\s+/g,' ').trim();
    if(!txt) return null;
    var m = txt.match(/^(?:SN\s*)?(\d{3,10})\s+(.+)$/i);
    if(!m) return null;
    var sn = String(m[1] || '').replace(/[^0-9]/g,'');
    var name = String(m[2] || '').replace(/\s+/g,' ').trim();
    if(!sn || !name) return null;
    return {{ sn: sn, name: name }};
  }}

  function manpowerListSelector() {{
    return MP_IDS.map(function(id) {{ return '#' + id; }}).concat(['ul[id^="ul-dept-"]']).join(',');
  }}

  function buildSnMap() {{
    var map = {{}};
    /* 1) كل موظفي الروستر المحقونين من Python */
    try {{
      var allStaff = window._ALL_STAFF || {{}};
      Object.keys(allStaff).forEach(function(sn) {{
        var snClean = String(sn).replace(/[^0-9]/g,'');
        var name    = String(allStaff[sn] || '').replace(/\s+/g,' ').trim();
        if(snClean && name && !map[snClean]) map[snClean] = name;
      }});
    }} catch(ex) {{}}
    /* 2) الموظفون الظاهرون في DOM */
    document.querySelectorAll('[data-sn][data-name]').forEach(function(el) {{
      var sn   = String(el.dataset.sn || '').replace(/[^0-9]/g,'');
      var name = String(el.dataset.name || '').replace(/\s+/g,' ').trim();
      if(sn && name && !map[sn]) map[sn] = name;
    }});
    /* 3) نص مكتوب يدوياً */
    document.querySelectorAll(manpowerListSelector() + ' li').forEach(function(li) {{
      var pair = extractSnNamePair(cleanEditableText(li));
      if(pair && !map[pair.sn]) map[pair.sn] = pair.name;
    }});
    return map;
  }}

  var snMap = {{}};
  var snDropdownHost = null;
  var snDropdownTarget = null;

  function refreshSnMap() {{
    snMap = buildSnMap();
    return snMap;
  }}

  function closeSnDropdown() {{
    if(snDropdownHost && snDropdownHost.parentNode) snDropdownHost.parentNode.removeChild(snDropdownHost);
    snDropdownHost = null;
    snDropdownTarget = null;
  }}

  function escapeHtml(text) {{
    return String(text || '').replace(/[&<>\"]/g, function(ch) {{
      return ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : '&quot;';
    }});
  }}

  function normalizeManpowerEditable(li) {{
    /* لا نُبدّل شيئاً — التنسيق (span data-sn) يجب أن يبقى كما هو.
       نُنظّف فقط إذا كان محتوى الـ li نصاً خاماً بدون span منسق. */
    if(!li) return;
    var hasFormatted = li.querySelector && li.querySelector('[data-sn][data-name]');
    if(hasFormatted) return;   /* ← منسق → لا تلمسه */
    /* نص خام فقط — لا نحتاج تنسيقاً هنا */
  }}

  function getManpowerText(li) {{
    if(!li) return '';
    var selected = li.querySelector && li.querySelector('[data-sn][data-name]');
    if(selected) {{
      var snSel = String(selected.dataset.sn || '').replace(/[^0-9]/g,'');
      var nameSel = String(selected.dataset.name || '').replace(/\s+/g,' ').trim();
      if(snSel && nameSel) return 'SN' + snSel + ' ' + nameSel;
    }}
    return cleanEditableText(li);
  }}

  function isFormattedManpowerEntry(li) {{
    return !!(li && li.querySelector && li.querySelector('[data-sn][data-name]'));
  }}

  function replaceManpowerText(li, text) {{
    if(!li) return;
    li.textContent = (text || '').replace(/​/g, '');
    setCursorEnd(li);
    if(typeof triggerAutosave === 'function') triggerAutosave();
  }}

  function convertFormattedManpowerToPlain(li, action) {{
    if(!li || !isFormattedManpowerEntry(li)) return false;
    var plain = getManpowerText(li).replace(/​/g, '');
    var sel = getTextSelectionOffsets(li);
    var start = sel.start;
    var end = sel.end;
    if(start == null || end == null) {{
      start = plain.length;
      end = plain.length;
    }}

    if(action === 'backspace') {{
      if(start !== end) {{
        plain = plain.slice(0, start) + plain.slice(end);
      }} else if(start > 0) {{
        plain = plain.slice(0, start - 1) + plain.slice(end);
        start -= 1;
      }}
      end = start;
    }} else if(action === 'delete') {{
      if(start !== end) {{
        plain = plain.slice(0, start) + plain.slice(end);
      }} else if(end < plain.length) {{
        plain = plain.slice(0, start) + plain.slice(end + 1);
      }}
      end = start;
    }} else if(action && action.type === 'insertText') {{
      var insertText = String(action.text || '');
      plain = plain.slice(0, start) + insertText + plain.slice(end);
      start = start + insertText.length;
      end = start;
    }}

    li.textContent = plain || '​';
    setCursorByTextOffset(li, start);
    if(typeof triggerAutosave === 'function') triggerAutosave();
    return true;
  }}

  function extractTypedSn(text) {{
    var txt = String(text || '')
      .replace(/[•·▪◦●]/g,' ')
      .replace(/\u00a0/g,' ')
      .replace(/^[\s\-–—]+/, '')
      .trim();
    if(!txt) return '';
    var m = txt.match(/(?:SN\s*)?(\d{{1,10}})/i);
    return m ? m[1] : '';
  }}

  function collectSnMatches(typed) {{
    typed = String(typed || '').replace(/[^0-9]/g,'');
    if(!typed) return [];
    if(!Object.keys(snMap).length) refreshSnMap();

    var seen = {{}};
    var entries = [];

    function pushEntry(snVal, nameVal) {{
      var sn = String(snVal || '').replace(/[^0-9]/g,'');
      var name = String(nameVal || '').replace(/\s+/g,' ').trim();
      if(!sn || !name || seen[sn]) return;
      seen[sn] = true;
      entries.push({{ sn: sn, name: name }});
    }}

    Object.keys(snMap).forEach(function(snKey) {{
      pushEntry(snKey, snMap[snKey]);
    }});

    document.querySelectorAll('[data-sn][data-name]').forEach(function(el) {{
      pushEntry(el.dataset.sn, el.dataset.name);
    }});

    return entries
      .map(function(item) {{
        var idx = item.sn.indexOf(typed);
        var rank = -1;
        if(item.sn === typed) rank = 0;
        else if(idx === 0) rank = 1;
        else if(typed.length >= 2 && idx > 0) rank = 2;
        else return null;
        return {{
          sn: item.sn,
          name: item.name,
          _rank: rank,
          _idx: idx < 0 ? 999 : idx,
          _delta: Math.abs(item.sn.length - typed.length)
        }};
      }})
      .filter(Boolean)
      .sort(function(a, b) {{
        if(a._rank !== b._rank) return a._rank - b._rank;
        if(a._idx !== b._idx) return a._idx - b._idx;
        if(a._delta !== b._delta) return a._delta - b._delta;
        return a.sn.localeCompare(b.sn);
      }})
      .slice(0, 10)
      .map(function(item) {{
        return {{ sn: item.sn, name: item.name }};
      }});
  }}

  function renderSnSelection(li, item) {{
    if(!li || !item) return;
    li.innerHTML = '<span data-sn="' + escapeHtml(item.sn) + '" data-name="' + escapeHtml(item.name) + '" '
                 + 'style="font-family:Calibri,Arial,sans-serif;">'
                 + '<strong style="color:#1b1f2a;">SN' + escapeHtml(item.sn) + '</strong>'
                 + '&nbsp;&nbsp;'
                 + '<span style="color:#1b1f2a;">' + escapeHtml(item.name) + '</span>'
                 + '</span>';
  }}

  function applySnSelection(li, item) {{
    if(!li || !item) return;
    renderSnSelection(li, item);
    refreshSnMap();
    closeSnDropdown();
    setCursorEnd(li);
    if(typeof triggerAutosave === 'function') triggerAutosave();
  }}

  function positionSnDropdown(li, dd) {{
    if(!li || !dd) return;
    var rect = li.getBoundingClientRect();
    dd.style.left = (window.scrollX + rect.left) + 'px';
    dd.style.top = (window.scrollY + rect.bottom + 4) + 'px';
    dd.style.minWidth = Math.max(rect.width, 240) + 'px';
  }}

  function showSnDropdown(li, matches) {{
    closeSnDropdown();
    if(!matches.length) return;

    var dd = document.createElement('div');
    dd.className = 'sn-drop';
    dd.setAttribute('contenteditable', 'false');
    dd.style.cssText = 'position:absolute;background:#fff;border:1px solid #0b3a78;'
                     + 'border-radius:6px;box-shadow:0 4px 16px rgba(11,58,120,.18);z-index:99999;'
                     + 'max-height:220px;overflow-y:auto;font-size:13px;';

    matches.forEach(function(item) {{
      var opt = document.createElement('div');
      opt.setAttribute('contenteditable', 'false');
      opt.style.cssText = 'padding:7px 12px;cursor:pointer;display:flex;gap:10px;align-items:center;';
      opt.innerHTML = '<span style="font-weight:700;color:#0b3a78;min-width:70px;">SN' + escapeHtml(item.sn) + '</span>'
                    + '<span style="color:#1b1f2a;">' + escapeHtml(item.name) + '</span>';
      opt.onmouseenter = function() {{ opt.style.background = '#eef3fc'; }};
      opt.onmouseleave = function() {{ opt.style.background = ''; }};
      opt.onmousedown = function(ev) {{
        ev.preventDefault();
        applySnSelection(li, item);
      }};
      dd.appendChild(opt);
    }});

    document.body.appendChild(dd);
    positionSnDropdown(li, dd);
    snDropdownHost = dd;
    snDropdownTarget = li;
  }}

  window.addEventListener('load', refreshSnMap);
  document.addEventListener('DOMContentLoaded', refreshSnMap);
  setTimeout(refreshSnMap, 300);
  setTimeout(refreshSnMap, 1500);
  window.addEventListener('resize', function() {{
    if(snDropdownHost && snDropdownTarget) positionSnDropdown(snDropdownTarget, snDropdownHost);
  }});
  window.addEventListener('scroll', function() {{
    if(snDropdownHost && snDropdownTarget) positionSnDropdown(snDropdownTarget, snDropdownHost);
  }}, true);
  document.addEventListener('click', function(ev) {{
    if(!snDropdownHost) return;
    if(snDropdownHost.contains(ev.target)) return;
    if(snDropdownTarget && snDropdownTarget.contains(ev.target)) return;
    closeSnDropdown();
  }});

  function setupManpowerLi(li) {{
    if(!li || li._mpSetup) return;
    li._mpSetup = true;

    function updateManpowerSuggestion() {{
      refreshSnMap();
      closeSnDropdown();
      if(isFormattedManpowerEntry(li)) return;
      var typed = extractTypedSn(getManpowerText(li));
      if(!typed) return;
      var matches = collectSnMatches(typed);
      if(matches.length === 1 && matches[0].sn === typed) {{
        applySnSelection(li, matches[0]);
        return;
      }}
      if(matches.length) showSnDropdown(li, matches);
    }}

    li.addEventListener('focus', function() {{
      refreshSnMap();
      if(isFormattedManpowerEntry(li)) closeSnDropdown();
      else updateManpowerSuggestion();
    }});
    li.addEventListener('click', function() {{
      refreshSnMap();
      if(isFormattedManpowerEntry(li)) closeSnDropdown();
      else updateManpowerSuggestion();
    }});
    li.addEventListener('input', function() {{
      updateManpowerSuggestion();
    }});
    li.addEventListener('keyup', function(ev) {{
      if(!['Enter','Tab','Escape'].includes(ev.key)) updateManpowerSuggestion();
    }});
    li.addEventListener('paste', function(ev) {{
      if(isFormattedManpowerEntry(li)) {{
        var pasted = ((ev.clipboardData || window.clipboardData) && (ev.clipboardData || window.clipboardData).getData('text')) || '';
        pasted = String(pasted || '').replace(/\s+/g, ' ').trim();
        if(pasted) {{
          ev.preventDefault();
          convertFormattedManpowerToPlain(li, {{type:'insertText', text:pasted}});
          updateManpowerSuggestion();
          return;
        }}
      }}
      setTimeout(function() {{ updateManpowerSuggestion(); }}, 30);
    }});

    li.addEventListener('keydown', function(ev) {{
      if(ev.key === 'Enter') {{
        ev.preventDefault();
        ev.stopPropagation();
        closeSnDropdown();
        var ul = li.parentElement;
        if(!ul) return;
        var newLi = createEditableListItem(ul);
        ul.insertBefore(newLi, li.nextSibling);
        _attachLiEvents(newLi);
        focusEditableListItem(newLi, true);
        if(typeof triggerAutosave === 'function') triggerAutosave();
        return;
      }}

      if(ev.key === 'Escape') {{
        closeSnDropdown();
        return;
      }}

      if(ev.key === 'Tab') {{
        if(isFormattedManpowerEntry(li)) return;
        var typedTab = extractTypedSn(getManpowerText(li));
        var matchesTab = collectSnMatches(typedTab);
        if(matchesTab.length) {{
          ev.preventDefault();
          applySnSelection(li, matchesTab[0]);
        }}
        return;
      }}

      if((ev.key === 'Backspace' || ev.key === 'Delete')) {{
        if(isFormattedManpowerEntry(li)) {{
          ev.preventDefault();
          convertFormattedManpowerToPlain(li, ev.key === 'Delete' ? 'delete' : 'backspace');
          closeSnDropdown();
          return;
        }}
        if(isEffectivelyEmptyManpowerLi(li)) {{
          ev.preventDefault();
          removeManpowerLiIfEmpty(li);
          return;
        }}
      }}

      if(isFormattedManpowerEntry(li) && ev.key && ev.key.length === 1 && !ev.ctrlKey && !ev.metaKey && !ev.altKey) {{
        ev.preventDefault();
        convertFormattedManpowerToPlain(li, {{type:'insertText', text:ev.key}});
        updateManpowerSuggestion();
        return;
      }}
    }});

    li.addEventListener('blur', function() {{
      if(isFormattedManpowerEntry(li)) {{
        setTimeout(closeSnDropdown, 200);
        return;
      }}
      var typed = extractTypedSn(getManpowerText(li));
      var matches = collectSnMatches(typed);
      if(typed && matches.length === 1 && matches[0].sn === typed) {{
        applySnSelection(li, matches[0]);
      }} else {{
        var pair = extractSnNamePair(getManpowerText(li));
        if(pair) {{
          renderSnSelection(li, pair);
          refreshSnMap();
        }}
      }}
      setTimeout(closeSnDropdown, 200);
    }});
  }}

    var MP_IDS = ['ul-supervisors','ul-ctu','ul-inventory','ul-support','ul-fd-export','ul-fd-import',
                'ul-sickleave','ul-annualleave','ul-trainee','ul-overtime'];
  function initManpower() {{
    refreshSnMap();
    MP_IDS.forEach(function(id) {{
      var ul = document.getElementById(id);
      if(!ul) return;
      ul.querySelectorAll('li').forEach(setupManpowerLi);
      new MutationObserver(function(muts) {{
        muts.forEach(function(m) {{
          m.addedNodes.forEach(function(n) {{ if(n.tagName==='LI') setupManpowerLi(n); }});
        }});
      }}).observe(ul, {{childList:true}});
    }});
    document.querySelectorAll('[id^="ul-dept-"]').forEach(function(ul) {{
      ul.querySelectorAll('li').forEach(setupManpowerLi);
      new MutationObserver(function(muts) {{
        muts.forEach(function(m) {{
          m.addedNodes.forEach(function(n) {{ if(n.tagName==='LI') setupManpowerLi(n); }});
        }});
      }}).observe(ul, {{childList:true}});
    }});
  }}

  function isManpowerLiTarget(node) {{
    var el = node;
    if(el && el.nodeType === 3) el = el.parentElement;
    var li = el && el.closest ? el.closest('li[contenteditable]') : null;
    if(!li) {{
      try {{
        var sel = window.getSelection && window.getSelection();
        var anchor = sel && sel.anchorNode;
        if(anchor && anchor.nodeType === 3) anchor = anchor.parentElement;
        li = anchor && anchor.closest ? anchor.closest('li[contenteditable]') : null;
      }} catch(ex) {{}}
    }}
    if(!li || !li.parentElement || !li.parentElement.id) return null;
    var ul = li.parentElement;
    var id = ul.id || '';
    if(id.indexOf('ul-dept-') === 0) return li;
    if(MP_IDS.indexOf(id) !== -1) return li;
    return null;
  }}

  function handleManpowerInteractive(node, forceOpen) {{
    var li = isManpowerLiTarget(node);
    if(!li) return;
    refreshSnMap();
    if(isFormattedManpowerEntry(li) && !forceOpen) {{ closeSnDropdown(); return; }}
    var typed = extractTypedSn(getManpowerText(li));
    if(!typed) {{ closeSnDropdown(); return; }}
    var matches = collectSnMatches(typed);
    if(matches.length === 1 && matches[0].sn === typed) {{
      applySnSelection(li, matches[0]);
      return;
    }}
    if(matches.length) showSnDropdown(li, matches); else closeSnDropdown();
  }}

  document.addEventListener('focusin', function(ev) {{
    var li = isManpowerLiTarget(ev.target);
    if(li) handleManpowerInteractive(li, false);
  }});
  document.addEventListener('click', function(ev) {{
    var li = isManpowerLiTarget(ev.target);
    if(li) handleManpowerInteractive(li, false);
  }});
  document.addEventListener('input', function(ev) {{
    var li = isManpowerLiTarget(ev.target);
    if(li) handleManpowerInteractive(li, false);
  }});
  document.addEventListener('keyup', function(ev) {{
    var li = isManpowerLiTarget(ev.target);
    if(li && !['Enter','Tab','Escape'].includes(ev.key)) handleManpowerInteractive(li, false);
  }});
  document.addEventListener('paste', function(ev) {{
    var li = isManpowerLiTarget(ev.target);
    if(li) setTimeout(function(){{ handleManpowerInteractive(li, false); }}, 30);
  }});
  function insertManpowerLineAfter(li) {{
    var ul = li && li.parentElement;
    if(!ul) return null;
    var newLi = createEditableListItem(ul);
    ul.insertBefore(newLi, li.nextSibling);
    _attachLiEvents(newLi);
    focusEditableListItem(newLi, true);
    if(typeof triggerAutosave === 'function') triggerAutosave();
    return newLi;
  }}

  document.addEventListener('keydown', function(ev) {{
    var li = isManpowerLiTarget(ev.target);
    if(!li) return;

    if(ev.key === 'Enter') {{
      ev.preventDefault();
      ev.stopPropagation();
      if(typeof ev.stopImmediatePropagation === 'function') ev.stopImmediatePropagation();
      closeSnDropdown();
      insertManpowerLineAfter(li);
      return;
    }}

    if((ev.key === 'Backspace' || ev.key === 'Delete') && isEffectivelyEmptyManpowerLi(li)) {{
      ev.preventDefault();
      ev.stopPropagation();
      if(typeof ev.stopImmediatePropagation === 'function') ev.stopImmediatePropagation();
      removeManpowerLiIfEmpty(li);
      return;
    }}

    if(ev.key === 'Escape') {{ closeSnDropdown(); return; }}
  }}, true);
  document.addEventListener('focusout', function(ev) {{
    var li = isManpowerLiTarget(ev.target);
    if(!li) return;
    if(isFormattedManpowerEntry(li)) {{
      setTimeout(closeSnDropdown, 200);
      return;
    }}
    var typed = extractTypedSn(getManpowerText(li));
    var matches = collectSnMatches(typed);
    if(typed && matches.length === 1 && matches[0].sn === typed) {{
      applySnSelection(li, matches[0]);
    }} else {{
      var pair = extractSnNamePair(getManpowerText(li));
      if(pair) {{
        renderSnSelection(li, pair);
        refreshSnMap();
      }}
    }}
    setTimeout(closeSnDropdown, 200);
  }});
  window.setupManpowerLi = setupManpowerLi;
  window.initManpower = initManpower;
  window.rebindSmartAutocomplete = function() {{
    try {{ if(typeof window.initFlightLists === 'function') window.initFlightLists(); }} catch(ex) {{}}
    try {{ if(typeof initTableCells === 'function') initTableCells(); }} catch(ex) {{}}
    try {{ if(typeof window.initManpower === 'function') window.initManpower(); }} catch(ex) {{}}
    try {{ if(typeof refreshSnMap === 'function') refreshSnMap(); }} catch(ex) {{}}
  }};
  if(document.readyState === 'loading') {{
    document.addEventListener('DOMContentLoaded', initManpower);
  }} else {{
    initManpower();
  }}

}})();





/* ══════════════════════════════════════════════════════
   AUTOSAVE — IndexedDB + GitHub Cloud Sync
   محلياً: يحفظ كل تعديل فوري داخل المتصفح
   سحابياً: عند تفعيل Cloud Sync يحفظ نفس التقرير في GitHub
   لكي تظهر التعديلات من الهاتف على الكمبيوتر والعكس
   ══════════════════════════════════════════════════════ */
(function(){{
  var DB_NAME = 'offload_autosave';
  var STORE   = 'reports';
  var PAGE_KEY = location.pathname;
  var REPO_OWNER = 'khalidsaif912';
  var REPO_NAME  = 'offload-monitor';
  var CLOUD_PATH = window._REPORT_CLOUD_PATH || '';
  var CLOUD_RAW_URL = CLOUD_PATH ? ('https://raw.githubusercontent.com/' + REPO_OWNER + '/' + REPO_NAME + '/main/' + CLOUD_PATH) : '';
  var CLOUD_API_URL = CLOUD_PATH ? ('https://api.github.com/repos/' + REPO_OWNER + '/' + REPO_NAME + '/contents/' + CLOUD_PATH) : '';
  var _db = null;
  var LS_KEY = 'offload_autosave_cache::' + PAGE_KEY;
  var _saveTimer = null;
  var _cloudTimer = null;
  var _cloudBusy = false;
  var _cloudPending = false;
  var _lastAppliedTs = 0;
  var _autoPromptShown = false;
  var _skipAutoPrompt = sessionStorage.getItem('gh_pat_skip_' + PAGE_KEY) === '1';

  function cloudBtn(){{
    return document.getElementById('btn-cloud-sync');
  }}

  function updateCloudBtnIdle(){{
    var btn = cloudBtn();
    if(!btn) return;
    var hasPat = !!(localStorage.getItem('gh_pat') || '').trim();
    btn.disabled = false;
    btn.innerText = hasPat ? '☁️ Auto Sync On' : '☁️ Set Auto Sync';
    btn.style.background = hasPat ? '#0f766e' : '#475569';
  }}

  function setCloudBtn(text, bg, delay){{
    var btn = cloudBtn();
    if(!btn) return;
    btn.innerText = text;
    if(bg) btn.style.background = bg;
    if(delay){{
      setTimeout(function(){{
        updateCloudBtnIdle();
      }}, delay);
    }}
  }}

  function openDB(cb){{
    if(_db){{ cb(_db); return; }}
    var req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = function(e){{
      e.target.result.createObjectStore(STORE, {{keyPath:'key'}});
    }};
    req.onsuccess = function(e){{ _db = e.target.result; cb(_db); }};
    req.onerror   = function(){{ console.warn('[autosave] IndexedDB open failed'); }};
  }}

  function readLocal(cb){{
    openDB(function(db){{
      var tx = db.transaction(STORE,'readonly');
      var req = tx.objectStore(STORE).get(PAGE_KEY);
      req.onsuccess = function(e){{ cb((e.target && e.target.result) || null); }};
      req.onerror = function(){{ cb(null); }};
    }});
  }}

  function writeLocal(record){{
    openDB(function(db){{
      var tx = db.transaction(STORE,'readwrite');
      tx.objectStore(STORE).put(record);
    }});
  }}

  function collectData(){{
    var data = {{}};

    document.querySelectorAll('[contenteditable="true"]').forEach(function(el){{
      var id = el.id || el.dataset.saveKey;
      if(!id){{
        var path = [];
        var node = el;
        while(node && node !== document.body){{
          var idx = Array.prototype.indexOf.call((node.parentElement||{{}}).children||[], node);
          path.unshift((node.tagName||'')+'['+idx+']');
          node = node.parentElement;
        }}
        id = path.join('>');
        el.dataset.saveKey = id;
      }}
      data[id] = el.innerHTML;
    }});

    document.querySelectorAll('ul[id]').forEach(function(ul){{
      data['__ul__'+ul.id] = ul.innerHTML;
    }});

    data['__nil_row_removed__'] = !document.getElementById('nil-row');

    return {{
      key: PAGE_KEY,
      ts: Date.now(),
      updated_at: new Date().toISOString(),
      cloud_path: CLOUD_PATH,
      data: data
    }};
  }}

  function applyRecord(rec){{
    if(!rec || !rec.data) return;
    var data = rec.data;

    Object.keys(data).forEach(function(k){{
      if(k.indexOf('__ul__') === 0){{
        var ul = document.getElementById(k.replace('__ul__',''));
        if(ul) ul.innerHTML = data[k];
      }}
    }});

    Object.keys(data).forEach(function(k){{
      if(k.indexOf('__') === 0) return;
      var el = document.getElementById(k) || document.querySelector('[data-save-key="'+k+'"]');
      if(el && el.isContentEditable) el.innerHTML = data[k];
    }});

    if(data['__nil_row_removed__']){{
      var nr = document.getElementById('nil-row');
      if(nr) nr.remove();
    }}

    _lastAppliedTs = Math.max(_lastAppliedTs, Number(rec.ts || 0));

    setTimeout(function(){{
      try {{ if(typeof window.rebindSmartAutocomplete === 'function') window.rebindSmartAutocomplete(); }} catch(ex) {{}}
      try {{ if(typeof window.initFlightLists === 'function') window.initFlightLists(); }} catch(ex) {{}}
      try {{ if(typeof initTableCells === 'function') initTableCells(); }} catch(ex) {{}}
      try {{ if(typeof window.initManpower === 'function') window.initManpower(); }} catch(ex) {{}}
      try {{ if(typeof refreshSnMap === 'function') refreshSnMap(); }} catch(ex) {{}}
      try {{ if(typeof window.renumberOffloadRows === 'function') window.renumberOffloadRows(); }} catch(ex) {{}}
    }}, 0);
  }}

  function getGitHubPat(silent){{
    var pat = (localStorage.getItem('gh_pat') || '').trim();
    if(pat || silent) return pat;
    pat = prompt('Enter GitHub token once to enable automatic report sync between phone and computer.\\n\\nPress Cancel to keep local-only saving on this device for now.');
    if(pat){{
      pat = pat.trim();
      localStorage.setItem('gh_pat', pat);
      sessionStorage.removeItem('gh_pat_skip_' + PAGE_KEY);
      _skipAutoPrompt = false;
      updateCloudBtnIdle();
    }} else {{
      sessionStorage.setItem('gh_pat_skip_' + PAGE_KEY, '1');
      _skipAutoPrompt = true;
      setCloudBtn('☁️ Local only', '#64748b', 2600);
    }}
    return (pat || '').trim();
  }}

  function ensureAutoGitHubPat(){{
    var pat = (localStorage.getItem('gh_pat') || '').trim();
    if(pat) return pat;
    if(_skipAutoPrompt || _autoPromptShown) return '';
    _autoPromptShown = true;
    return getGitHubPat(false);
  }}

  function fetchCloudRecord(){{
    if(!CLOUD_RAW_URL) return Promise.resolve(null);
    return fetch(CLOUD_RAW_URL + '?t=' + Date.now(), {{cache:'no-store'}})
      .then(function(r){{
        if(r.status === 404) return null;
        if(!r.ok) throw new Error('RAW_' + r.status);
        return r.json();
      }})
      .catch(function(){{ return null; }});
  }}

  function pushCloud(record, forcePrompt){{
    if(!CLOUD_API_URL) return Promise.resolve(false);
    var pat = getGitHubPat(!forcePrompt ? true : false);
    if(!pat){{
      if(forcePrompt){{
        setCloudBtn('☁️ Token needed', '#b45309', 2600);
      }}
      return Promise.resolve(false);
    }}
    if(_cloudBusy){{
      _cloudPending = true;
      return Promise.resolve(false);
    }}
    _cloudBusy = true;
    setCloudBtn('☁️ Syncing…', '#0284c7');

    var headers = {{
      'Accept':'application/vnd.github+json',
      'Authorization':'Bearer ' + pat,
      'X-GitHub-Api-Version':'2022-11-28'
    }};
    var content = btoa(unescape(encodeURIComponent(JSON.stringify(record, null, 2))));

    return fetch(CLOUD_API_URL + '?t=' + Date.now(), {{headers:headers}})
      .then(function(r){{
        if(r.status === 404) return null;
        if(r.status === 401 || r.status === 403){{
          localStorage.removeItem('gh_pat');
          throw new Error('AUTH');
        }}
        if(!r.ok) throw new Error('GET_' + r.status);
        return r.json();
      }})
      .then(function(existing){{
        var body = {{
          message: 'Sync report edits for ' + PAGE_KEY,
          content: content,
          branch: 'main'
        }};
        if(existing && existing.sha) body.sha = existing.sha;
        return fetch(CLOUD_API_URL, {{
          method: 'PUT',
          headers: Object.assign({{'Content-Type':'application/json'}}, headers),
          body: JSON.stringify(body)
        }});
      }})
      .then(function(r){{
        if(r.status === 401 || r.status === 403){{
          localStorage.removeItem('gh_pat');
          throw new Error('AUTH');
        }}
        if(!r.ok) throw new Error('PUT_' + r.status);
        return r.json();
      }})
      .then(function(){{
        _cloudBusy = false;
        _cloudPending = false;
        _lastAppliedTs = Math.max(_lastAppliedTs, Number(record.ts || 0));
        setCloudBtn('☁️ Synced', '#16a34a', 2500);
        return true;
      }})
      .catch(function(err){{
        _cloudBusy = false;
        _cloudPending = false;
        if((err && err.message) === 'AUTH'){{
          setCloudBtn('☁️ Re-enter token', '#dc2626', 3200);
        }} else {{
          setCloudBtn('☁️ Sync failed', '#dc2626', 3200);
        }}
        return false;
      }});
  }}

  function scheduleCloudPush(record){{
    if(!CLOUD_API_URL) return;
    var pat = ensureAutoGitHubPat();
    if(!pat) return;
    clearTimeout(_cloudTimer);
    _cloudTimer = setTimeout(function(){{
      pushCloud(record, false);
    }}, 1400);
  }}

  function saveNow(forceCloudNow){{
    var record = collectData();
    writeLocal(record);
    if(forceCloudNow) pushCloud(record, false); else scheduleCloudPush(record);
    return record;
  }}

  window.triggerAutosave = function(){{ saveNow(); }};
  window.forceCloudSync = function(){{
    var record = saveNow();
    return pushCloud(record, true).then(function(ok){{
      if(ok){{
        alert('Cloud sync completed. When you open the same report on another device, the latest edits will appear there.');
      }}
      return ok;
    }});
  }};

  function restoreSaved(){{
    Promise.all([
      new Promise(function(resolve){{ readLocal(resolve); }}),
      fetchCloudRecord()
    ]).then(function(results){{
      var localRec = results[0];
      var cloudRec = results[1];
      var chosen = null;

      if(localRec && cloudRec){{
        chosen = Number(cloudRec.ts || 0) > Number(localRec.ts || 0) ? cloudRec : localRec;
      }} else {{
        chosen = cloudRec || localRec;
      }}

      if(chosen) applyRecord(chosen);
      if(cloudRec && (!localRec || Number(cloudRec.ts || 0) > Number(localRec.ts || 0))){{
        writeLocal(cloudRec);
        setCloudBtn('☁️ Cloud loaded', '#16a34a', 2200);
      }} else if(localRec && cloudRec && Number(localRec.ts || 0) > Number(cloudRec.ts || 0)){{
        scheduleCloudPush(localRec);
      }}
    }}).catch(function(err){{
      console.warn('[autosave] restore failed', err);
    }});
  }}

  function _schedSave(){{
    clearTimeout(_saveTimer);
    _saveTimer = setTimeout(saveNow, 400);
  }}

  document.addEventListener('input',  _schedSave);
  document.addEventListener('keyup',  _schedSave);
  document.addEventListener('change', _schedSave, true);
  document.addEventListener('blur',   _schedSave, true);
  document.addEventListener('visibilitychange', function(){{
    if(document.visibilityState === 'hidden') saveNow(true);
  }});
  window.addEventListener('pagehide', function(){{ saveNow(true); }});
  window.addEventListener('beforeunload', function(){{ saveNow(false); }});

  var mo = new MutationObserver(function(muts){{
    var changed = muts.some(function(m){{
      return m.type==='childList' || m.type==='characterData';
    }});
    if(changed) _schedSave();
  }});
  mo.observe(document.body, {{childList:true, subtree:true, characterData:true}});

  if(document.readyState === 'loading'){{
    document.addEventListener('DOMContentLoaded', restoreSaved);
  }} else {{
    restoreSaved();
  }}

  updateCloudBtnIdle();

  var cloudBtnEl = cloudBtn();
  if(cloudBtnEl){{
    cloudBtnEl.addEventListener('click', function(){{
      window.forceCloudSync();
    }});
  }}
}})();
</script>

</body>
</html>"""

    out_dir = DOCS_DIR / date_dir / shift
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "index.html").write_text(html, encoding="utf-8")







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
    # calendar is already imported at module level as _cal
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
        "shift2": {"label": "Afternoon", "ar": "ظهر",  "time": "13:00 – 22:00", "icon": "☀️"},
        "shift3": {"label": "Night",     "ar": "ليل",  "time": "21:00 – 06:00", "icon": "🌙"},
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

    # عد الرحلات (مع تطبيق فلتر التاريخ كما في التقرير)
    def _count_matching_flights(folder: Path, report_date: str) -> int:
        """Count JSON flight files whose date matches the report date."""
        if not folder.exists():
            return 0
        count = 0
        for p in folder.glob("*.json"):
            if p.name == "meta.json":
                continue
            try:
                flt = json.loads(p.read_text(encoding="utf-8"))
                fd = (flt.get("date") or "").strip().upper()
                if not fd:
                    count += 1  # no date = count it
                    continue
                # datetime is already imported at module level
                rd = datetime.strptime(report_date, "%Y-%m-%d")
                matched = False
                for fmt in ("%d%b%y", "%d%b%Y", "%d%b", "%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y"):
                    try:
                        parsed = datetime.strptime(fd, fmt)
                        if fmt == "%d%b":
                            parsed = parsed.replace(year=rd.year)
                        if parsed.day == rd.day and parsed.month == rd.month and parsed.year == rd.year:
                            matched = True
                        break
                    except ValueError:
                        continue
                else:
                    matched = True  # unknown format = count it
                if matched:
                    count += 1
            except Exception:
                count += 1
        return count

    days_html = ""
    for day in day_dirs:
        is_today   = day == today
        is_future  = day > today
        open_attr  = " open" if is_today else ""

        day_flights = 0
        for shift in ("shift1", "shift2", "shift3"):
            day_flights += _count_matching_flights(DATA_DIR / day / shift, day)

        badge        = '<span class="today-badge">TODAY</span>' if is_today else ('<span class="today-badge" style="background:#64748b;">UPCOMING</span>' if is_future else "")
        flights_pill = f'<span class="day-pill">{day_flights} flights</span>' if day_flights else ""

        rows = ""
        for shift in ("shift1", "shift2", "shift3"):
            shift_report = DOCS_DIR / day / shift / "index.html"
            meta_s = shift_meta.get(shift, {"label": shift, "ar": shift, "time": "", "icon": "✈"})
            ms_icon  = meta_s["icon"]
            ms_ar    = meta_s["ar"]
            ms_label = meta_s["label"]
            ms_time  = meta_s["time"]
            shift_flt_count = _count_matching_flights(DATA_DIR / day / shift, day)
            flt_txt = f"{shift_flt_count} flight{'s' if shift_flt_count != 1 else ''}" if shift_flt_count else ""

            if shift_report.exists():
                # مناوبة فيها تقرير — رابط
                rows += f"""
            <a class="shift-card" href="{day}/{shift}/">
                <div class="sc-icon">{ms_icon}</div>
                <div class="sc-body">
                    <div class="sc-title">{ms_ar} <span class="sc-en">/ {ms_label}</span></div>
                    <div class="sc-time">{ms_time}</div>
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
                <div class="sc-icon">{ms_icon}</div>
                <div class="sc-body">
                    <div class="sc-title">{ms_ar} <span class="sc-en">/ {ms_label}</span></div>
                    <div class="sc-time">{ms_time}</div>
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

    roster_base_url = ROSTER_PAGE_URL.rstrip('/') + "/date/"

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

        /* ── ACTION BAR ── */
        .action-bar {{
            max-width:860px; margin:16px auto 0; padding:0 14px;
            display:flex; flex-wrap:wrap; align-items:center; gap:10px;
        }}
        .action-btn {{
            display:inline-flex; align-items:center; gap:6px;
            font-family:'Inter',system-ui,sans-serif;
            font-size:13px; font-weight:700;
            color:#fff; background:var(--blue);
            border:none; border-radius:12px;
            padding:10px 18px; cursor:pointer;
            text-decoration:none;
            box-shadow:0 2px 8px rgba(37,99,235,.2);
            transition:filter .1s;
        }}
        .action-btn:hover {{ filter:brightness(1.1); }}

        /* ── MANPOWER MODAL ── */
        .mp-overlay {{
            display:none; position:fixed; inset:0;
            background:rgba(0,0,0,.50); z-index:1000;
            align-items:center; justify-content:center;
        }}
        .mp-overlay.open {{ display:flex; }}
        .mp-box {{
            background:#fff; border-radius:12px;
            width:94%; max-width:560px; max-height:82vh;
            overflow-y:auto; padding:0;
            box-shadow:0 24px 64px rgba(0,0,0,.30);
            position:relative; font-family:Calibri,Arial,sans-serif;
        }}
        /* ── Header bar (matches report navy) ── */
        .mp-header {{
            display:flex; align-items:center; justify-content:space-between;
            background:#0b3a78; border-radius:12px 12px 0 0;
            padding:14px 56px 14px 20px;
        }}
        .mp-title {{
            font-size:13px; font-weight:800; color:#fff;
            letter-spacing:.8px; text-transform:uppercase;
        }}
        .mp-copy {{
            font-size:12px; font-weight:700; color:#0b3a78;
            background:#fff; border:none; border-radius:7px;
            padding:5px 14px; cursor:pointer; white-space:nowrap;
        }}
        /* ── Close button: visible circle in header area ── */
        .mp-close {{
            position:absolute; top:10px; right:12px;
            width:32px; height:32px; border-radius:50%;
            background:rgba(255,255,255,.18); border:2px solid rgba(255,255,255,.60);
            font-size:17px; line-height:1; cursor:pointer;
            color:#fff; display:flex; align-items:center; justify-content:center;
        }}
        /* ── Body ── */
        .mp-body {{ padding:18px 22px 22px; }}
        /* ── Department header ── */
        .mp-dept {{
            font-size:11px; font-weight:800; color:#fff;
            background:#0b3a78; text-transform:uppercase;
            letter-spacing:.7px; padding:4px 10px;
            border-radius:5px; margin:14px 0 5px;
            display:inline-block;
        }}
        .mp-dept:first-child {{ margin-top:0; }}
        /* ── Employee rows ── */
        .mp-emp {{
            font-size:13px; color:#1b1f2a; text-align:left;
            padding:5px 10px; border-radius:5px;
            background:#f4f7fc; margin:3px 0;
            display:flex; align-items:center; gap:8px;
        }}
        .mp-emp:nth-child(even) {{ background:#ffffff; border:1px solid #e8edf5; }}
        .mp-emp .emp-sn {{
            font-size:11px; font-weight:700; color:#fff;
            background:#4a7bc4; border-radius:4px;
            padding:2px 7px; white-space:nowrap; flex-shrink:0;
        }}
        .mp-emp .emp-name {{ color:#1b1f2a; font-weight:600; }}
        .mp-loading {{ text-align:center; padding:36px; color:var(--muted); font-size:13px; }}

        @media (max-width:600px) {{
            .top {{ padding:14px 14px 12px; }}
            .top h1 {{ font-size:17px; }}
            .wrap {{ padding:0 10px; margin-top:14px; }}
            .day-date {{ font-size:13px; }}
            .shift-card {{ padding:10px 12px; }}
            .hdr-badge {{ display:none; }}
            .action-bar {{ padding:0 10px; margin-top:12px; }}
            .action-btn {{ font-size:12px; padding:8px 14px; }}
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
            <div class="hdr-badge" style="display:flex;flex-direction:column;align-items:center;line-height:1;">
                <strong id="todayDay" style="font-size:18px;"></strong>
                <span id="todayMonth" style="font-size:11px;opacity:.75;letter-spacing:.5px;"></span>
            </div>
        </div>
    </div>

    <!-- ═══ ACTION BAR ═══ -->
    <div class="action-bar">
        <div class="action-btn" style="cursor:default;">📅 {today}</div>
        <a class="action-btn" href="{ROSTER_PAGE_URL}" target="_blank" rel="noopener noreferrer">📋 Duty Roster</a>
        <button class="action-btn" id="btn-manpower" onclick="openManpower()">👥 Manpower</button>
    </div>

    <!-- ═══ MANPOWER MODAL ═══ -->
    <div class="mp-overlay" id="mp-overlay" onclick="if(event.target===this)closeManpower()">
        <div class="mp-box">
            <div class="mp-header">
                <span class="mp-title">6.&nbsp;&nbsp;MANPOWER</span>
                <button class="mp-copy" id="mp-copy-btn" onclick="copyManpower()">📋 Copy</button>
            </div>
            <button class="mp-close" onclick="closeManpower()" title="Close">&times;</button>
            <div class="mp-body">
                <div id="mp-content"><div class="mp-loading">جاري التحميل...</div></div>
            </div>
        </div>
    </div>

    <div class="wrap">
        {days_html}
        <div class="footer">Generated automatically by GitHub Actions · {today}</div>
    </div>

<script>
/* ── Manpower popup logic ── */
var ROSTER_DAILY_BASE = "{roster_base_url}";
var IMPORT_ROSTER_RAW_BASE = "https://raw.githubusercontent.com/khalidsaif912/roster-site/main/docs/import/";
var LEAVE_LABELS = ["annual leave","sick leave","emergency leave","off day","training"];
var SHIFT_MAP = {{"shift1":"Morning","shift2":"Afternoon","shift3":"Night"}};
var EXCLUDED_DEPTS = ["officers"];
var INVENTORY_SNS = ["82592","990737"];
var SUPPORT_SNS   = ["82653","82565"];
var mpData = null;

function getCurrentShift() {{
    var now = new Date();
    var m = (now.getUTCHours()*60 + now.getUTCMinutes() + 240) % 1440;
    if (m >= 330 && m < 870) return "shift1";
    if (m >= 870 && m < 1290) return "shift2";
    return "shift3";
}}

function getTodayDate() {{
    var now = new Date(Date.now() + 4*3600000);
    return now.toISOString().slice(0,10);
}}

function normShift(s) {{
    return (s || "").trim().toLowerCase();
}}

function parseRosterHtml(html, shiftKey) {{
    var parser = new DOMParser();
    var doc = parser.parseFromString(html, "text/html");
    var target = normShift(SHIFT_MAP[shiftKey] || shiftKey || "Morning");
    var onDuty = [];

    doc.querySelectorAll(".deptCard").forEach(function(card) {{
        var dept = ((card.querySelector(".deptTitle") || {{}}).textContent || "Unknown").trim();
        var deptNorm = dept.toLowerCase();

        card.querySelectorAll("details.shiftCard").forEach(function(sc) {{
            var label = normShift(sc.getAttribute("data-shift") || "");

            sc.querySelectorAll(".empRow").forEach(function(row) {{
                var ne = row.querySelector(".empName");
                if (!ne) return;

                var raw = ne.textContent.trim();
                var name = raw;
                var sn = "";
                var mx = raw.match(/^(.+?)\s*[-\u2013]\s*(\d+)(?:\s*\(.*?\))?\s*$/);
                if (mx) {{
                    name = mx[1].trim();
                    sn = mx[2].trim();
                }}

                if (LEAVE_LABELS.indexOf(label) !== -1) return;
                if (label !== target) return;
                if (EXCLUDED_DEPTS.indexOf(deptNorm) !== -1) return;

                onDuty.push({{name:name, sn:sn, dept:dept}});
            }});
        }});
    }});
    return onDuty;
}}

function parseImportFlightDispatchText(html, shiftKey) {{
    var target = normShift(SHIFT_MAP[shiftKey] || shiftKey || "Morning");
    var knownDepts = {{
        "documentation": "Documentation",
        "flight dispatch (export)": "Flight Dispatch",
        "flight dispatch (import)": "Flight Dispatch",
        "import checkers": "Import Checkers",
        "import operators": "Import Operators",
        "release control": "Release Control",
        "supervisors": "Supervisors"
    }};
    var result = [];
    var seen = {{}};
    var currentDept = "";
    var currentShift = "";

    (html || "")
        .split(/\r?\n/)
        .map(function(line) {{ return (line || "").replace(/\s+/g, " ").trim(); }})
        .filter(Boolean)
        .forEach(function(line) {{
            var low = normShift(line);

            if (knownDepts[low]) {{
                currentDept = low;
                currentShift = "";
                return;
            }}

            if (!currentDept) return;

            if (low.indexOf("morning") !== -1) {{ currentShift = "morning"; return; }}
            if (low.indexOf("afternoon") !== -1) {{ currentShift = "afternoon"; return; }}
            if (low.indexOf("night") !== -1) {{ currentShift = "night"; return; }}
            if (low.indexOf("off day") !== -1) {{ currentShift = "off day"; return; }}
            if (low.indexOf("annual leave") !== -1) {{ currentShift = "annual leave"; return; }}
            if (low.indexOf("sick leave") !== -1) {{ currentShift = "sick leave"; return; }}
            if (low.indexOf("emergency leave") !== -1) {{ currentShift = "emergency leave"; return; }}
            if (low.indexOf("training") !== -1) {{ currentShift = "training"; return; }}

            if (low.indexOf("total ") === 0) return;
            if (low.indexOf("view full roster") === 0 || low.indexOf("last updated:") === 0) return;
            if (currentDept !== "flight dispatch (export)" && currentDept !== "flight dispatch (import)") return;
            if (currentShift !== target) return;

            var mx = line.match(/^(.+?)\s*[·•\-–]\s*(\d{3,6})\b.*$/);
            if (!mx) return;

            var sn = mx[2].trim();
            var name = mx[1].trim();
            var key = sn + "|" + name.toLowerCase();
            if (seen[key]) return;
            seen[key] = true;

            result.push({{
                name: name,
                sn: sn,
                dept: "Flight Dispatch"
            }});
        }});

    return result;
}}

function buildEmpRow(name, sn) {{
    var snHtml = sn ? '<span class="emp-sn">' + sn + '</span>' : '';
    return '<div class="mp-emp">' + snHtml + '<span class="emp-name">' + name + '</span></div>';
}}

function loadManpower() {{
    var el = document.getElementById("mp-content");
    el.innerHTML = '<div class="mp-loading">\u062c\u0627\u0631\u064a \u0627\u0644\u062a\u062d\u0645\u064a\u0644...</div>';

    var dateStr = getTodayDate();
    var exportUrl = ROSTER_DAILY_BASE + dateStr + "/?t=" + Date.now();
    var importUrl = IMPORT_ROSTER_RAW_BASE + dateStr + "/index.html?t=" + Date.now();

    Promise.allSettled([
        fetch(exportUrl).then(function(r) {{
            if (!r.ok) throw new Error("Export HTTP " + r.status);
            return r.text();
        }}),
        fetch(importUrl).then(function(r) {{
            if (!r.ok) throw new Error("Import HTTP " + r.status);
            return r.text();
        }})
    ])
    .then(function(results) {{
        var shift = getCurrentShift();
        var exportHtml = results[0].status === "fulfilled" ? results[0].value : "";
        var importHtml = results[1].status === "fulfilled" ? results[1].value : "";

        var allEmps = exportHtml ? parseRosterHtml(exportHtml, shift) : [];
        var inventoryEmps = allEmps.filter(function(e) {{ return INVENTORY_SNS.indexOf(e.sn) !== -1; }});
        var supportEmps   = allEmps.filter(function(e) {{ return SUPPORT_SNS.indexOf(e.sn) !== -1 || (e.name + " " + e.dept).toLowerCase().indexOf("support") !== -1; }});
        var specialSNs = INVENTORY_SNS.concat(SUPPORT_SNS);
        var mainEmps = allEmps.filter(function(e) {{
            return specialSNs.indexOf(e.sn) === -1 && (e.name + " " + e.dept).toLowerCase().indexOf("support") === -1;
        }});

        var grouped = {{}};
        mainEmps.forEach(function(e) {{
            if (!grouped[e.dept]) grouped[e.dept] = [];
            grouped[e.dept].push(e);
        }});

        if (inventoryEmps.length) grouped["Inventory"] = inventoryEmps.slice();
        if (supportEmps.length) grouped["C) Support Team"] = supportEmps.slice();

        if (importHtml) {{
            var fdStaff = parseImportFlightDispatchText(importHtml, shift);
            if (fdStaff.length) {{
                grouped["Flight Dispatch"] = fdStaff.slice();
            }}
        }}

        mpData = grouped;

        var html = "";
        for (var dept in grouped) {{
            html += '<div class="mp-dept">' + dept + '</div>';
            grouped[dept].forEach(function(e) {{ html += buildEmpRow(e.name, e.sn); }});
        }}

        el.innerHTML = html || '<div class="mp-loading">\u0644\u0627 \u062a\u0648\u062c\u062f \u0628\u064a\u0627\u0646\u0627\u062a \u0644\u0644\u0645\u0646\u0627\u0648\u0628\u0629 \u0627\u0644\u062d\u0627\u0644\u064a\u0629</div>';
    }})
    .catch(function(err) {{
        el.innerHTML = '<div class="mp-loading">\u062e\u0637\u0623 \u0641\u064a \u062a\u062d\u0645\u064a\u0644 \u0627\u0644\u0631\u0648\u0633\u062a\u0631: ' + err.message + '</div>';
    }});
}}


function openManpower() {{
    document.getElementById("mp-overlay").classList.add("open");
    loadManpower();
}}

function closeManpower() {{
    document.getElementById("mp-overlay").classList.remove("open");
}}

function copyManpower() {{
    if (!mpData) return;
    var text = "6. MANPOWER\\n\\n";
    for (var dept in mpData) {{
        text += dept + ":\\n";
        mpData[dept].forEach(function(e) {{
            text += "  " + (e.sn ? e.sn + " — " : "") + e.name + "\\n";
        }});
        text += "\\n";
    }}
    navigator.clipboard.writeText(text.trim()).then(function() {{
        var btn = document.getElementById("mp-copy-btn");
        btn.textContent = "✅ Copied!";
        btn.style.background = "#059669";
        setTimeout(function() {{ btn.textContent = "📋 Copy"; btn.style.background = ""; }}, 2000);
    }});
}}
</script>

<script>
(function() {{
    try {{
        var now = new Date();
        var muscat = new Date(
            now.getTime() + (4 * 60 * 60 * 1000) + (now.getTimezoneOffset() * 60 * 1000)
        );

        var day = muscat.getDate();
        var months = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
        var month = months[muscat.getMonth()];

        var d = document.getElementById("todayDay");
        var m = document.getElementById("todayMonth");

        if (d) d.textContent = day;
        if (m) m.textContent = month;
    }} catch(e) {{
        console.log("date badge error:", e);
    }}
}})();
</script>
</body>
</html>"""

    (DOCS_DIR / "index.html").write_text(html, encoding="utf-8")


# ══════════════════════════════════════════════════════════════════
#  إعادة إثراء الرحلات السابقة بـ DEST الصحيح من AirLabs
# ══════════════════════════════════════════════════════════════════

def retroactive_enrich_all(now: datetime) -> None:
    """Re-check all saved flights and overwrite old wrong STD/ETD + DEST values.

    Enable with:
      RETRO_ENRICH=1 python offload_monitor.py

    This mode is intentionally aggressive:
      - it re-fetches each saved flight even if std_etd/destination already exist
      - it uses the fallback chain (AirLabs -> Flightradar -> MuscatAirport)
      - it overwrites old values only when a better value is found
      - it rebuilds all HTML reports afterwards
    """
    if not DATA_DIR.exists():
        print("[retroactive] data/ directory not found — skipping.")
        return

    updated_count = 0
    skipped_count = 0
    failed_count = 0

    for json_file in sorted(DATA_DIR.rglob("*.json")):
        if json_file.name == "meta.json":
            continue

        try:
            flight = json.loads(json_file.read_text(encoding="utf-8"))
        except Exception:
            failed_count += 1
            continue

        flt = normalize_flight_number((flight.get("flight") or "").strip())
        raw_date = (flight.get("date") or "").strip()
        if not flt:
            skipped_count += 1
            continue

        flight_date_iso = normalize_flight_date(raw_date, now) if raw_date else None
        old_dest = (flight.get("destination") or "").strip().upper()
        old_std_etd = (flight.get("std_etd") or "").strip()

        info, source_name = fetch_flight_info_with_fallbacks(
            flt,
            flight_date=flight_date_iso,
            dep_iata="MCT",
            arr_iata=old_dest or None,
        )
        if not info:
            print(f"  [retro] {flt}/{raw_date}: no data from fallback chain")
            skipped_count += 1
            continue

        changed = False

        new_dest = (info.get("dest") or "").strip().upper()
        if new_dest and new_dest != old_dest:
            print(f"  [retro] {flt}/{raw_date}: dest {old_dest!r} -> {new_dest!r} via {source_name}")
            flight["destination"] = new_dest
            changed = True

        std = (info.get("std") or "").strip()
        etd = (info.get("etd") or "").strip()
        new_std_etd = ""
        if std and etd and std != etd:
            new_std_etd = f"{std}|{etd}"
        elif std:
            new_std_etd = std
        elif etd:
            new_std_etd = etd

        if new_std_etd and new_std_etd != old_std_etd:
            print(f"  [retro] {flt}/{raw_date}: std_etd {old_std_etd!r} -> {new_std_etd!r} via {source_name}")
            flight["std_etd"] = new_std_etd
            changed = True

        if changed:
            flight["retro_enriched_at"] = now.isoformat()
            flight["retro_enriched_source"] = source_name or ""
            json_file.write_text(json.dumps(flight, ensure_ascii=False, indent=2), encoding="utf-8")
            updated_count += 1
        else:
            skipped_count += 1

    print(f"[retroactive] Done. Updated: {updated_count}, Skipped/unchanged: {skipped_count}, Failed: {failed_count}")

    print("[retroactive] Rebuilding HTML reports…")
    rebuilt = 0
    for date_dir_p in sorted(p for p in DATA_DIR.iterdir() if p.is_dir()):
        for shift in ("shift1", "shift2", "shift3"):
            if (date_dir_p / shift).exists():
                build_shift_report(date_dir_p.name, shift)
                print(f"  rebuilt: {date_dir_p.name}/{shift}")
                rebuilt += 1

    build_root_index(now)
    print(f"[retroactive] All reports rebuilt. {rebuilt} report(s). ✓")

# ══════════════════════════════════════════════════════════════════
#  إرسال التقرير بالبريد الإلكتروني
# ══════════════════════════════════════════════════════════════════

def _extract_report_content_html(page_html: str) -> str:
    """Return only the main report container without action buttons/scripts.
    Uses regex-based extraction to preserve nested table structure (avoids
    BeautifulSoup html.parser reordering nested tables).
    Strips contenteditable, tabindex, class attributes and removes the
    Back-to-Index link row.
    """
    # ── 1) Extract report-content table via regex (preserves nesting) ──
    # Find the opening tag with id="report-content"
    m_start = re.search(r'<table[^>]*id="report-content"[^>]*>', page_html, re.IGNORECASE)
    if not m_start:
        html = page_html
    else:
        start = m_start.start()
        # Walk forward counting <table> / </table> to find matching close
        depth = 0
        pos = start
        while pos < len(page_html):
            t_open  = re.search(r'<table[\s>]', page_html[pos:], re.IGNORECASE)
            t_close = re.search(r'</table\s*>', page_html[pos:], re.IGNORECASE)
            if t_close is None:
                break
            if t_open and t_open.start() < t_close.start():
                depth += 1
                pos += t_open.start() + 1
            else:
                depth -= 1
                pos += t_close.end()
                if depth == 0:
                    break
        html = page_html[start:pos]

    # ── 2) Remove Back-to-Index link row ──
    html = re.sub(
        r'<tr[^>]*id="back-link-row"[^>]*>.*?</tr>',
        '',
        html,
        flags=re.DOTALL | re.IGNORECASE,
    )

    # ── 3) Strip attributes invalid in email clients ──
    html = re.sub(r'\s+contenteditable="[^"]*"', '', html)
    html = re.sub(r'\s+tabindex="[^"]*"', '', html)
    html = re.sub(r'\s+class="[^"]*"', '', html)

    return html


def _build_email_html(page_html: str) -> str:
    """Build a mobile-friendly HTML email — left-aligned, no centering."""
    report_html = _extract_report_content_html(page_html)
    # Make the report table full-width regardless of inline width/max-width
    report_html = re.sub(r'width="(760|900|1100)"', 'width="900"', report_html)
    report_html = re.sub(
        r'style="width:(760|900|1100)px;[^"]*"',
        'style="width:900px; max-width:900px; background-color:#ffffff; border:none; table-layout:fixed;"',
        report_html,
    )

    return f"""<!doctype html>
<html dir="ltr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    html, body {{
      margin: 0 !important;
      padding: 0 !important;
      width: 100% !important;
      background: #ffffff !important;
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
    .mobile-wrap {{ width: 100%; padding: 8px 12px 18px; box-sizing: border-box; }}
    @media only screen and (max-width: 640px) {{
      body, table, td, div, p, a, li {{
        font-size: 16px !important;
        line-height: 1.65 !important;
      }}
      .mobile-wrap {{ padding: 4px 6px 14px !important; }}
      table[width="100%"], table[width="900"] {{ width: 100% !important; max-width: 900px !important; }}
      td[style*="font-size:20px"] div {{ font-size: 18px !important; }}
      td[style*="font-size:13.5px"], div[style*="font-size:13.5px"], span[style*="font-size:12px"], td[style*="font-size:12px"] {{
        font-size: 15px !important;
      }}
    }}
  </style>
</head>
<body>
  <div class="mobile-wrap">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="width:100%; background:#ffffff;">
      <tr>
        <td align="left" style="padding:0;">
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
    """نافذة الإرسال: التقرير يُرسل قبل ساعة من نهاية المناوبة.

    shift1 (06:00–15:00): يُرسل الساعة 14:00
    shift2 (13:00–22:00): يُرسل الساعة 21:00
    shift3 (21:00–06:00): يُرسل الساعة 05:00
    """
    # (start_h, start_m, end_h, end_m)
    windows = {
        "shift1": (14, 0, 15, 0),
        "shift2": (21, 0, 22, 0),
        "shift3": (5,  0,  6, 0),
    }
    w = windows.get(shift)
    if not w:
        return False

    current = now.hour * 60 + now.minute
    start_m = w[0] * 60 + w[1]
    end_m   = w[2] * 60 + w[3]

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

    المناوبات مع أوقات قطع الأوفلود:
      shift1 : 05:30 – 14:30 (أوفلود بعد 14:30 → shift2)
      shift2 : 14:30 – 21:30 (أوفلود بعد 21:30 → shift3)
      shift3 : 21:30 – 05:30 (أوفلود بعد 05:30 → shift1)
    """
    tz  = ZoneInfo(TIMEZONE)
    loc = ref_dt.astimezone(tz)
    mins = loc.hour * 60 + loc.minute

    base = loc.replace(hour=0, minute=0, second=0, microsecond=0)

    if 5 * 60 + 30 <= mins < 14 * 60 + 30:       # shift1
        start = base.replace(hour=5, minute=30)
        end   = base.replace(hour=14, minute=30)
    elif 14 * 60 + 30 <= mins < 21 * 60 + 30:    # shift2
        start = base.replace(hour=14, minute=30)
        end   = base.replace(hour=21, minute=30)
    else:                                          # shift3 يعبر منتصف الليل
        if loc.hour < 6:
            # بعد منتصف الليل — المناوبة بدأت أمس
            start = (base - timedelta(days=1)).replace(hour=21, minute=30)
        else:
            start = base.replace(hour=21, minute=30)
        end = (start + timedelta(hours=8))

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
        # الرحلة أقدم من بداية المناوبة بأكثر من يومين كاملين → قديمة
        # ملاحظة: استخدمنا 48h بدل 24h لاستيعاب الرحلات التي تأتي في اليوم التالي
        age_hours = (shift_start - flight_date).total_seconds() / 3600

        if age_hours > 48:
            skipped.append(f.get("flight","?") + "/" + raw_date)
        else:
            kept.append(f)

    if skipped:
        print(f"  [filter] Skipped {len(skipped)} old flight(s): {', '.join(skipped)}")
    print(f"  [filter] Kept {len(kept)} flight(s) for this shift window.")
    return kept


def filter_flights_already_in_other_shifts(flights: list[dict], now: datetime) -> list[dict]:
    """تمنع تكرار الرحلات بناءً على AWBs وليس فقط رقم الرحلة.

    المنطق الجديد:
    - إذا رحلة موجودة في مناوبة أخرى وتحتوي نفس AWBs → تكرار → تجاهل
    - إذا رحلة موجودة لكن بشحنات مختلفة (أوفلود جديد) → احتفظ بها
    - إذا رحلة بدون AWBs → قارن برقم الرحلة فقط كاحتياط
    """
    current_shift = get_shift(now)
    date_dir = get_shift_date(now, current_shift)

    # اجمع (flight_name → set of AWBs) من المناوبات الأخرى
    existing_flights: dict = {}
    for shift in ("shift1", "shift2", "shift3"):
        if shift == current_shift:
            continue
        shift_folder = DATA_DIR / date_dir / shift
        if not shift_folder.exists():
            continue
        for p in shift_folder.glob("*.json"):
            if p.name == "meta.json":
                continue
            try:
                flt_data = json.loads(p.read_text(encoding="utf-8"))
                flt_name = (flt_data.get("flight") or "").strip().upper()
                if not flt_name:
                    continue
                awbs = {
                    (it.get("awb") or "").strip()
                    for it in flt_data.get("items", [])
                    if (it.get("awb") or "").strip()
                }
                if flt_name not in existing_flights:
                    existing_flights[flt_name] = set()
                existing_flights[flt_name].update(awbs)
            except Exception:
                continue

    if not existing_flights:
        return flights

    kept = []
    skipped = []
    for f in flights:
        flt_name = (f.get("flight") or "").strip().upper()
        if not flt_name or flt_name not in existing_flights:
            kept.append(f)
            continue
        # نفس رقم الرحلة موجود — تحقق من AWBs
        new_awbs = {
            (it.get("awb") or "").strip()
            for it in f.get("items", [])
            if (it.get("awb") or "").strip()
        }
        existing_awbs = existing_flights[flt_name]
        if not new_awbs:
            # لا توجد AWBs → لا يمكن الحكم بالتكرار، نحتفظ بالرحلة
            kept.append(f)
        elif new_awbs.issubset(existing_awbs):
            # كل الشحنات الجديدة موجودة مسبقاً → تكرار حقيقي
            skipped.append(flt_name)
        else:
            # شحنات مختلفة أو جديدة → أوفلود جديد لنفس الرحلة
            kept.append(f)

    if skipped:
        print(f"  [dup-filter] Skipped {len(skipped)} exact duplicate(s): {', '.join(skipped)}")
    print(f"  [dup-filter] Kept {len(kept)} flight(s) for {current_shift}.")
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
        print("RETRO_ENRICH=1 detected. Re-checking and correcting all old saved flight times/destinations…")
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
    html, file_modified_time = download_file()
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
            today_str = get_shift_date(now)
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

    # ── حفظ وقت تعديل الملف في كل رحلة لاستخدامه كوقت الإيميل ──
    # واستخدام وقت الإيميل (وليس وقت تشغيل السكربت) لتحديد المناوبة الصحيحة
    email_dt = now  # fallback: وقت التشغيل
    if file_modified_time:
        for f in flights:
            f["email_time"] = file_modified_time
        print(f"  [email_time] Set to file Last-Modified: {file_modified_time}")
        # ── تحويل وقت الإيميل إلى datetime لاستخدامه في تحديد المناوبة ──
        try:
            _today = now.date()
            _h, _m = map(int, file_modified_time.split(":"))
            email_dt = datetime(_today.year, _today.month, _today.day, _h, _m, tzinfo=ZoneInfo(TIMEZONE))
            # إذا وقت الإيميل بعد منتصف الليل وقبل 06:00 والسكربت يشتغل بعد الظهر
            # فالتاريخ صحيح لأن الإيميل من نفس اليوم
            email_shift = get_shift(email_dt)
            print(f"  [shift-fix] Email time {file_modified_time} → shift={email_shift} (instead of {get_shift(now)} from script run time)")
        except Exception as exc:
            print(f"  [shift-fix] Failed to parse email time, using now: {exc}")
            email_dt = now

    # ── فلترة الرحلات القديمة بناءً على وقت الإيميل والمناوبة ──
    flights = filter_flights_by_shift(flights, email_dt)

    if not flights:
        print("WARNING: All flights filtered out as old/stale — no data for this shift.")
        STATE_FILE.write_text(new_hash, encoding="utf-8")
        build_root_index(now)
        return

    # ── Enrich with fallback chain: AirLabs -> Flightradar -> Muscat Airport ──
    enriched = 0
    for f in flights:
        flt = normalize_flight_number(f.get("flight") or "")
        if not flt:
            continue

        info, source_name = fetch_flight_info_with_fallbacks(
            flt,
            flight_date=normalize_flight_date(f.get('date', ''), now),
            dep_iata="MCT",
            arr_iata=(f.get("destination") or "").strip() or None,
        )
        if not info:
            continue

        std = (info.get("std") or "").strip()
        etd = (info.get("etd") or "").strip()
        if std and etd and std != etd:
            f["std_etd"] = f"{std}|{etd}"
        elif std:
            f["std_etd"] = std
        elif etd:
            f["std_etd"] = etd

        dest = (info.get("dest") or "").strip().upper()
        if dest:
            f["destination"] = dest
        enriched += 1
        print(f"  [enrich] {flt} enriched via {source_name}")

    print(f"Enriched {enriched} flight(s) via fallback chain.")

    # ── فلترة الرحلات المكررة من مناوبات سابقة ──
    flights = filter_flights_already_in_other_shifts(flights, email_dt)

    if not flights:
        print("WARNING: All flights filtered out as duplicates from other shifts — no new flights for this shift.")
        STATE_FILE.write_text(new_hash, encoding="utf-8")
        build_root_index(now)
        return

    print(f"Extracted {len(flights)} flight(s). Saving…")
    operational_date_dir, shift, _, affected_date_dirs = save_flights(flights, email_dt)

    ensure_email_recipients_file()
    for report_date_dir in affected_date_dirs or [operational_date_dir]:
        print(f"Building report: {report_date_dir}/{shift}…")
        build_shift_report(report_date_dir, shift)
    build_root_index(now)

    STATE_FILE.write_text(new_hash, encoding="utf-8")
    print(f"Done. ✓  ({len(flights)} flights saved)")

    # ── إرسال البريد قبل نهاية المناوبة ──
    today_str = get_shift_date(now)
    for _shift in ("shift1", "shift2", "shift3"):
        maybe_send_email(now, today_str, _shift)


if __name__ == "__main__":
    main()

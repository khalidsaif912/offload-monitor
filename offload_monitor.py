import os, re, json, hashlib, requests
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path

ONEDRIVE_URL = os.environ["ONEDRIVE_FILE_URL"]
TIMEZONE = "Asia/Muscat"

DATA_DIR = Path("data")
STATE_FILE = Path("state.txt")
PUBLIC_DIR = Path("public")

def download_file():
    url = ONEDRIVE_URL.strip()
    if "download=1" not in url:
        url += "&download=1" if "?" in url else "?download=1"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.text

def sha(content: str) -> str:
    return hashlib.sha256(content.encode("utf-8", errors="ignore")).hexdigest()

def get_shift(now):
    mins = now.hour * 60 + now.minute
    if 6*60 <= mins < 14*60 + 30:
        return "shift1"
    if 14*60 + 30 <= mins < 21*60 + 30:
        return "shift2"
    return "shift3"

def safe(s: str) -> str:
    s = (s or "UNKNOWN").strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^A-Za-z0-9_-]", "_", s)
    return s[:60] if s else "UNKNOWN"

def extract_flight_info(html: str):
    """
    يحاول يطلع FLIGHT / DATE / DESTINATION من HTML المصدر.
    مبني على نمط: FLIGHT<td>..</td> DATE<td>..</td> DESTINATION<td>..</td>
    """
    flight = re.search(r"FLIGHT\s*</td>\s*<td[^>]*>(.*?)</td>", html, re.I | re.S)
    date   = re.search(r"DATE\s*</td>\s*<td[^>]*>(.*?)</td>", html, re.I | re.S)
    dest   = re.search(r"DESTINATION\s*</td>\s*<td[^>]*>(.*?)</td>", html, re.I | re.S)

    f = safe(flight.group(1)) if flight else "UNKNOWN"
    d = safe(date.group(1))   if date else "UNKNOWN"
    t = safe(dest.group(1))   if dest else "UNKNOWN"
    return f, d, t

def load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except:
        return default

def write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def save_or_update_flight(content: str, now):
    date_dir = now.strftime("%Y-%m-%d")
    shift = get_shift(now)

    flight, fdate, dest = extract_flight_info(content)
    folder = DATA_DIR / date_dir / shift
    folder.mkdir(parents=True, exist_ok=True)

    # اسم ثابت: نفس الرحلة = نفس الملف (فيصير overwrite عند التعديل)
    file_name = f"{flight}_{fdate}_{dest}.html"
    file_path = folder / file_name

    was_existing = file_path.exists()

    file_path.write_text(content, encoding="utf-8")

    # meta.json لتتبع التعديلات
    meta_path = folder / "meta.json"
    meta = load_json(meta_path, {"flights": {}})

    key = file_name
    entry = meta["flights"].get(key, {
        "flight": flight, "date": fdate, "dest": dest,
        "created_at": now.isoformat(),
        "updated_at": now.isoformat(),
        "updates": 0
    })

    if was_existing:
        entry["updates"] = int(entry.get("updates", 0)) + 1
        entry["updated_at"] = now.isoformat()
    else:
        entry["created_at"] = now.isoformat()
        entry["updated_at"] = now.isoformat()

    meta["flights"][key] = entry
    write_json(meta_path, meta)

    return date_dir, shift

def build_shift_page(date_dir: str, shift: str):
    folder = DATA_DIR / date_dir / shift
    if not folder.exists():
        return

    meta = load_json(folder / "meta.json", {"flights": {}})
    flights_meta = meta.get("flights", {})

    files = sorted([p for p in folder.glob("*.html") if p.name != "meta.json"])

    # صفحة منظمة وبادج تحديث
    rows = []
    for p in files:
        info = flights_meta.get(p.name, {})
        updates = int(info.get("updates", 0))
        updated_at = info.get("updated_at", "")

        badge = ""
        if updates > 0:
            badge = f'<span style="margin-left:10px;padding:2px 8px;border-radius:999px;background:#f59e0b;color:#111827;font-weight:700;font-size:12px;">UPDATED x{updates}</span>'

        title = p.name.replace(".html", "")
        rows.append(
            f"""
            <div style="border:1px solid #d0d5e8;background:#fff;border-radius:12px;padding:12px 14px;margin:12px 0;">
              <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                <div style="font-weight:800;color:#0f172a;">{title}</div>
                {badge}
                <div style="color:#64748b;font-size:12px;">Last: {updated_at}</div>
              </div>
              <div style="margin-top:10px;border-top:1px dashed #e2e8f0;padding-top:10px;">
                {p.read_text(encoding="utf-8")}
              </div>
            </div>
            """
        )

    body = f"""
    <div style="max-width:1400px;margin:0 auto;padding:16px;font-family:Calibri,Arial,sans-serif;">
      <h1 style="margin:0 0 6px 0;">Offload - {date_dir} - {shift}</h1>
      <div style="color:#64748b;margin-bottom:12px;">This page contains all offloads for this shift. Updates replace the same flight.</div>
      {''.join(rows) if rows else '<div style="color:#94a3b8;">No entries yet.</div>'}
    </div>
    """

    out_dir = PUBLIC_DIR / date_dir / shift
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "index.html").write_text(f"<!doctype html><html><head><meta charset='utf-8'></head><body style='margin:0;background:#eef2ff;'>{body}</body></html>", encoding="utf-8")

def main():
    now = datetime.now(ZoneInfo(TIMEZONE))

    content = download_file()
    new_hash = sha(content)

    if STATE_FILE.exists():
        old_hash = STATE_FILE.read_text(encoding="utf-8").strip()
        if old_hash == new_hash:
            print("No change detected.")
            return

    print("Change detected. Saving flight snapshot (new or update)...")
    date_dir, shift = save_or_update_flight(content, now)
    build_shift_page(date_dir, shift)

    STATE_FILE.write_text(new_hash, encoding="utf-8")
    print("Done.")

if __name__ == "__main__":
    main()

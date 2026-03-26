import imaplib
import email
from email.utils import parsedate_to_datetime
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
import re

EMAIL = "SCOON80@gmail.com"
PASSWORD = "zcmn srnz xtln vkwe"
IMAP_SERVER = "imap.gmail.com"
TIMEZONE = "Asia/Muscat"
LABEL_NAME = "offload-reports"

BASE_DIR = Path("downloads")
BASE_DIR.mkdir(exist_ok=True)


# تنظيف اسم الملف
def clean_name(text: str) -> str:
    text = (text or "").strip().lower()
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^a-zA-Z0-9_-]", "", text)
    return text[:50] or "email"


# استخراج تاريخ الإيميل
def get_email_datetime(msg) -> datetime:
    raw_date = msg.get("Date")
    if raw_date:
        try:
            return parsedate_to_datetime(raw_date).astimezone(ZoneInfo(TIMEZONE))
        except Exception:
            pass
    return datetime.now(ZoneInfo(TIMEZONE))


# استخراج HTML من الإيميل
def get_html_content(msg) -> str | None:
    html_content = None

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition") or "").lower()

            if "attachment" in disposition:
                continue

            if content_type == "text/html":
                payload = part.get_payload(decode=True)
                if payload:
                    html_content = payload.decode(errors="ignore")
                    break
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            html_content = payload.decode(errors="ignore")

    return html_content


# 🔌 الاتصال
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, PASSWORD)

# 🔍 طباعة كل المجلدات
status, folders = mail.list()
print("\n[INFO] Available folders:")
for folder in folders:
    print(folder.decode(errors="ignore"))

# 🔄 تجربة فتح المجلد
candidate_labels = [
    "Offload Reports",
    '"Offload Reports"',
    "INBOX/Offload Reports",
    '"INBOX/Offload Reports"',
    "[Gmail]/Offload Reports",
    '"[Gmail]/Offload Reports"',
]

opened = False
for label in candidate_labels:
    status, _ = mail.select(label)
    print(f"[DEBUG] Trying: {label} -> {status}")
    if status == "OK":
        print(f"[SUCCESS] Opened folder: {label}")
        opened = True
        break

if not opened:
    raise RuntimeError("Cannot open any expected label/folder")

# 📥 قراءة الإيميلات
status, messages = mail.search(None, "ALL")
ids = messages[0].split()[-15:]

print(f"\n[INFO] Found {len(ids)} emails")

for num in ids:
    status, msg_data = mail.fetch(num, "(RFC822)")
    if status != "OK":
        print(f"[ERROR] Failed to fetch email {num.decode()}")
        continue

    msg = email.message_from_bytes(msg_data[0][1])

    subject = msg.get("Subject", "offload")
    print(f"\n[EMAIL] {subject}")

    email_dt = get_email_datetime(msg)
    date_folder = email_dt.strftime("%Y-%m-%d")
    time_part = email_dt.strftime("%Y-%m-%d_%H-%M")

    day_dir = BASE_DIR / date_folder
    day_dir.mkdir(parents=True, exist_ok=True)

    html_content = get_html_content(msg)

    if not html_content:
        print("[SKIP] No HTML content found")
        continue

    safe_subject = clean_name(subject)
    filename = f"{time_part}_{safe_subject}_{num.decode()}.html"
    file_path = day_dir / filename

    if file_path.exists():
        print(f"[SKIP] Already exists: {filename}")
        continue

    file_path.write_text(html_content, encoding="utf-8")
    print(f"[SAVED HTML] {file_path}")

mail.logout()

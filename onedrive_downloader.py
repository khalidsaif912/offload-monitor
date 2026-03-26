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
TARGET_FOLDER = "Offload Reports"

BASE_DIR = Path("downloads")
BASE_DIR.mkdir(exist_ok=True)


def clean_name(text: str) -> str:
    text = (text or "").strip().lower()
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^a-zA-Z0-9_-]", "", text)
    return text[:50] or "email"


def get_email_datetime(msg) -> datetime:
    raw_date = msg.get("Date")
    if raw_date:
        try:
            return parsedate_to_datetime(raw_date).astimezone(ZoneInfo(TIMEZONE))
        except Exception:
            pass
    return datetime.now(ZoneInfo(TIMEZONE))


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


def is_offload_email(subject: str) -> bool:
    s = (subject or "").lower()
    keywords = [
        "offload",
        "offloaded",
        "offloaded cargo",
        "fw_offloaded",
        "fw: offloaded",
        "fwd: offloaded",
    ]
    return any(k in s for k in keywords)


def extract_folder_names(folders) -> list[str]:
    names = []
    for folder in folders or []:
        line = folder.decode(errors="ignore")
        print(line)

        # نأخذ آخر جزء بين علامتَي اقتباس
        m = re.search(r'"([^"]+)"\s*$', line)
        if m:
            names.append(m.group(1))
    return names


def main():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)

    status, folders = mail.list()
    print("\n[INFO] Available folders:")
    if status != "OK":
        raise RuntimeError("Failed to list folders")

    folder_names = extract_folder_names(folders)

    actual_folder = None
    for name in folder_names:
        if name.strip().lower() == TARGET_FOLDER.lower():
            actual_folder = name
            break

    if not actual_folder:
        raise RuntimeError(f'Folder "{TARGET_FOLDER}" not found in IMAP list')

    print(f'\n[INFO] Matched folder: {actual_folder}')

    # مهم: نرسل الاسم بين اقتباسين
    try:
        status, data = mail.select(f'"{actual_folder}"')
        print(f'[DEBUG] Opening folder "{actual_folder}" -> {status}')
        print(f"[DEBUG] select() response: {data}")
    except Exception as e:
        raise RuntimeError(f'Cannot open folder "{actual_folder}": {e}')

    if status != "OK":
        raise RuntimeError(f'Cannot open folder "{actual_folder}"')

    folder_count = int(data[0].decode()) if data and data[0] else 0
    print(f'[INFO] Total emails in "{actual_folder}": {folder_count}')

    status, messages = mail.search(None, "ALL")
    if status != "OK":
        raise RuntimeError(f'Failed to search emails in folder "{actual_folder}"')

    ids = messages[0].split()
    print(f"[INFO] Search returned {len(ids)} email id(s)")

    ids = ids[-15:]
    saved_count = 0

    for num in ids:
        status, msg_data = mail.fetch(num, "(RFC822)")
        if status != "OK":
            print(f"[ERROR] Failed to fetch email {num.decode()}")
            continue

        msg = email.message_from_bytes(msg_data[0][1])

        subject = msg.get("Subject", "offload")
        print(f"\n[EMAIL] {subject}")

        if not is_offload_email(subject):
            print("[SKIP] Not an offload email")
            continue

        email_dt = get_email_datetime(msg)
        date_folder = email_dt.strftime("%Y-%m-%d")
        time_part = email_dt.strftime("%Y-%m-%d_%H-%M")

        day_dir = BASE_DIR / date_folder
        day_dir.mkdir(parents=True, exist_ok=True)

        html_content = get_html_content(msg)

        if not html_content:
            print("[SKIP] No HTML content found")
            continue

        print(f"[DEBUG] HTML length: {len(html_content)}")

        safe_subject = clean_name(subject)
        filename = f"{time_part}_{safe_subject}_{num.decode()}.html"
        file_path = day_dir / filename

        if file_path.exists():
            print(f"[SKIP] Already exists: {filename}")
            continue

        file_path.write_text(html_content, encoding="utf-8")
        print(f"[SAVED HTML] {file_path}")
        saved_count += 1

    print(f"\n[INFO] Saved {saved_count} offload HTML file(s)")

    mail.logout()


if __name__ == "__main__":
    main()

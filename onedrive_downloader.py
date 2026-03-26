import requests
from bs4 import BeautifulSoup
from pathlib import Path
from urllib.parse import urljoin
import os


# 🔧 إعدادات
ONEDRIVE_FOLDER_URL = "https://omanair-my.sharepoint.com/:f:/p/8715_hq/IgDdD8um6ShWSa7BONLOTNcXAYphb1AI98eW_NZjxjvbW0k?e=lEMoPT"
DOWNLOAD_DIR = Path("downloads")


def list_files(folder_url: str) -> list[str]:
    """Extract file links from SharePoint folder"""

    print("[INFO] Reading OneDrive/SharePoint folder...")

    r = requests.get(
        folder_url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Cache-Control": "no-cache",
        },
        timeout=30,
    )
    r.raise_for_status()

    soup = BeautifulSoup(r.text, "html.parser")

    links = []

    for a in soup.find_all("a", href=True):
        href = a["href"]

        # SharePoint patterns
        if any(x in href for x in ["download", "redir", ":u:", ":x:", "Doc.aspx"]):
            full_url = urljoin(folder_url, href)
            links.append(full_url)

    links = list(set(links))

    print(f"[INFO] Found {len(links)} candidate link(s)")
    return links


def download_file(url: str, save_dir: Path, index: int):
    """Download single file"""

    try:
        r = requests.get(
            url,
            timeout=60,
            allow_redirects=True,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        r.raise_for_status()

        # تجاهل صفحات HTML الصغيرة (ليست ملفات)
        content_type = r.headers.get("Content-Type", "").lower()
        if "text/html" in content_type and len(r.content) < 5000:
            print(f"[SKIP] Not a real file: {url}")
            return

        # اسم الملف
        filename = url.split("/")[-1].split("?")[0]
        if not filename:
            filename = f"file_{index}.html"

        file_path = save_dir / filename
        file_path.write_bytes(r.content)

        print(f"[DOWNLOADED] {filename}")

    except Exception as e:
        print(f"[ERROR] {url} -> {e}")


def main():
    print("[DEBUG] Current working dir:", os.getcwd())

    # إنشاء المجلد + ضمان ظهوره في GitHub
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    (DOWNLOAD_DIR / ".gitkeep").touch(exist_ok=True)

    print("[DEBUG] downloads path:", DOWNLOAD_DIR.resolve())
    print("[DEBUG] downloads exists:", DOWNLOAD_DIR.exists())

    links = list_files(ONEDRIVE_FOLDER_URL)

    if not links:
        print("[WARNING] No links found — SharePoint page may require JS")
        return

    for i, link in enumerate(links, start=1):
        download_file(link, DOWNLOAD_DIR, i)

    print("[DONE] All files processed.")


if __name__ == "__main__":
    main()

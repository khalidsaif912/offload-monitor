import requests
from bs4 import BeautifulSoup
from pathlib import Path
from urllib.parse import urljoin


# 🔧 إعدادات
ONEDRIVE_FOLDER_URL = "PUT_YOUR_FOLDER_LINK_HERE"
DOWNLOAD_DIR = Path("downloads")


def list_files(folder_url: str) -> list[str]:
    """Extract file download links from OneDrive shared folder"""
    
    print("[INFO] Reading OneDrive folder...")

    r = requests.get(folder_url)
    r.raise_for_status()

    soup = BeautifulSoup(r.text, "html.parser")

    links = []

    for a in soup.find_all("a", href=True):
        href = a["href"]

        if "download" in href or "redir" in href:
            full_url = urljoin("https://onedrive.live.com", href)
            links.append(full_url)

    links = list(set(links))

    print(f"[INFO] Found {len(links)} file(s)")
    return links


def download_file(url: str, save_dir: Path):
    """Download single file"""

    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()

        # نحاول نطلع اسم الملف
        filename = url.split("/")[-1].split("?")[0]
        if not filename:
            filename = "file.html"

        file_path = save_dir / filename

        file_path.write_bytes(r.content)

        print(f"[DOWNLOADED] {filename}")

    except Exception as e:
        print(f"[ERROR] {url} -> {e}")


def main():
    DOWNLOAD_DIR.mkdir(exist_ok=True)

    links = list_files(ONEDRIVE_FOLDER_URL)

    for link in links:
        download_file(link, DOWNLOAD_DIR)

    print("[DONE] All files downloaded.")


if __name__ == "__main__":
    main()
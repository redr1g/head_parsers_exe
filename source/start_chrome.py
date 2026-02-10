import subprocess
import time
import socket
import os
import sys

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
USER_DATA_DIR = r"C:\ChromeDebug"
DEBUG_PORT = 9222

SITES = {
    "1": ("CaseDrop", "https://casedrop.eu/shop"),
    "2": ("G4Skins", "https://g4skins.com/"),
    "3": ("CSGO-Skins", "https://csgo-skins.com/"),
    "4": ("CSGOCases", "https://csgocases.com/shop"),
}


def is_port_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0


def open_url_in_existing_chrome(url: str):
    subprocess.Popen(
        [
            CHROME_PATH,
            "--new-tab",
            url,
            f"--user-data-dir={USER_DATA_DIR}",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def start_chrome_with_url(url: str):
    if is_port_open(DEBUG_PORT):
        print(f"✅ Chrome already running on port {DEBUG_PORT}")
        print(f"🌍 Opening: {url}")
        open_url_in_existing_chrome(url)
        return

    print("🚀 Starting Chrome with remote debugging...")
    cmd = [
        CHROME_PATH,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--user-data-dir={USER_DATA_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        url,
    ]

    subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )

    for _ in range(30):
        if is_port_open(DEBUG_PORT):
            print("✅ Chrome debugger is ready")
            return
        time.sleep(1)

    raise RuntimeError("❌ Chrome did not start (port 9222 not opened)")


def show_menu() -> str:
    print("\n=== SELECT SITE ===")
    for key, (name, url) in SITES.items():
        print(f"{key}. {name} ({url})")
    print("0. Exit")

    return input("\nEnter number: ").strip()


def main():
    os.makedirs(USER_DATA_DIR, exist_ok=True)

    choice = show_menu()

    if choice == "0":
        print("👋 Bye")
        sys.exit(0)

    if choice not in SITES:
        print("❌ Invalid choice")
        sys.exit(1)

    name, url = SITES[choice]
    print(f"\n➡ Opening {name}")
    start_chrome_with_url(url)

    # ⬇️ ОДРАЗУ ЗАКІНЧУЄМО СКРИПТ І ЗАКРИВАЄМО КОНСОЛЬ
    sys.exit(0)


if __name__ == "__main__":
    main()

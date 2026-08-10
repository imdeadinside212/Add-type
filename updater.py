"""
updater.py
----------
Module tu-update mien phi, dung GitHub Releases lam noi luu file build.

Cach dung (xem main_example.py):
    from updater import check_for_update, download_update, apply_update

    info = check_for_update("imdeadinside212/add-type", "1.0.0")
    if info:
        new_file = download_update(info["download_url"])
        apply_update(new_file)   # app se tu thoat va restart voi ban moi

Yeu cau:
    pip install requests
"""

import os
import sys
import platform
import subprocess
import tempfile
import requests

# Repo GitHub chua cac ban Release cua app "Add type"
REPO = "imdeadinside212/add-type"


def _current_os():
    system = platform.system().lower()
    if system == "windows":
        return "windows"
    if system == "darwin":
        return "mac"
    return system


def check_for_update(repo: str = REPO, current_version: str = "0.0.0", timeout: int = 10):
    """
    Kiem tra GitHub Releases xem co ban moi hon current_version khong.

    repo: "imdeadinside212/add-type" (da dat san lam gia tri mac dinh)
    current_version: vd "1.0.0" (khong can chu 'v')

    Tra ve dict {"version": ..., "download_url": ..., "notes": ...}
    hoac None neu khong co ban moi / loi mang.
    """
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    try:
        resp = requests.get(url, timeout=timeout)
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        print(f"[updater] Khong the kiem tra update: {e}")
        return None

    latest_version = data.get("tag_name", "").lstrip("v")
    if not latest_version or latest_version == current_version:
        return None

    # so sanh dang so (1.10.0 > 1.9.0), fallback ve so sanh chuoi neu loi
    def parse(v):
        try:
            return tuple(int(x) for x in v.split("."))
        except ValueError:
            return None

    cur_t, new_t = parse(current_version), parse(latest_version)
    if cur_t is not None and new_t is not None:
        if new_t <= cur_t:
            return None

    os_name = _current_os()
    asset_url = None
    for asset in data.get("assets", []):
        name = asset["name"].lower()
        if os_name == "windows" and (name.endswith(".exe") or "win" in name):
            asset_url = asset["browser_download_url"]
            break
        if os_name == "mac" and (name.endswith(".dmg") or name.endswith(".zip") or "mac" in name):
            asset_url = asset["browser_download_url"]
            break

    if not asset_url:
        print(f"[updater] Co ban moi {latest_version} nhung khong tim thay file cho OS: {os_name}")
        return None

    return {
        "version": latest_version,
        "download_url": asset_url,
        "notes": data.get("body", ""),
    }


def download_update(url: str, timeout: int = 30) -> str:
    """
    Tai file ve thu muc temp, tra ve duong dan file da tai.
    """
    local_name = url.split("/")[-1]
    dest = os.path.join(tempfile.gettempdir(), local_name)

    with requests.get(url, stream=True, timeout=timeout) as r:
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

    return dest


def apply_update(new_file_path: str):
    """
    Ap dung ban update: thay the file/app hien tai bang ban moi roi khoi
    dong lai. Ham nay se goi sys.exit() o cuoi, tuc la app se thoat de
    script phu tien hanh ghi de.
    """
    os_name = _current_os()
    if os_name == "windows":
        _apply_update_windows(new_file_path)
    elif os_name == "mac":
        _apply_update_mac(new_file_path)
    else:
        raise NotImplementedError(f"Chua ho tro auto-update tren {os_name}")


def _apply_update_windows(new_file_path: str):
    """
    PyInstaller --onefile: sys.executable la duong dan file exe dang chay.
    Dung 1 file .bat de doi app thoat, roi ghi de, roi mo lai.
    """
    old_file = sys.executable
    bat_path = os.path.join(tempfile.gettempdir(), "app_update.bat")

    bat_script = f"""@echo off
:wait_loop
tasklist /FI "PID eq {os.getpid()}" 2>NUL | find "{os.getpid()}" >NUL
if not errorlevel 1 (
    timeout /t 1 /nobreak > NUL
    goto wait_loop
)
move /Y "{new_file_path}" "{old_file}"
start "" "{old_file}"
del "%~f0"
"""
    with open(bat_path, "w") as f:
        f.write(bat_script)

    subprocess.Popen(["cmd", "/c", bat_path], shell=True)
    sys.exit(0)


def _apply_update_mac(new_file_path: str):
    """
    new_file_path duoc ky vong la 1 file .zip chua ban .app moi,
    hoac 1 file .dmg. O day xu ly truong hop zip chua san .app
    (cach don gian nhat, khong can mount dmg).

    App bundle path duoc suy ra tu sys.executable:
    .../AddType.app/Contents/MacOS/AddType
    """
    app_bundle = sys.executable
    for _ in range(3):
        app_bundle = os.path.dirname(app_bundle)
    # app_bundle gio la .../AddType.app

    sh_path = os.path.join(tempfile.gettempdir(), "app_update.sh")
    unzip_dir = os.path.join(tempfile.gettempdir(), "app_update_extracted")

    sh_script = f"""#!/bin/bash
while kill -0 {os.getpid()} 2>/dev/null; do
    sleep 1
done
rm -rf "{unzip_dir}"
mkdir -p "{unzip_dir}"
unzip -o "{new_file_path}" -d "{unzip_dir}" > /dev/null
NEW_APP=$(find "{unzip_dir}" -maxdepth 1 -name "*.app" | head -n 1)
rm -rf "{app_bundle}"
cp -R "$NEW_APP" "{app_bundle}"
xattr -cr "{app_bundle}"
open "{app_bundle}"
rm -f "$0"
"""
    with open(sh_path, "w") as f:
        f.write(sh_script)
    os.chmod(sh_path, 0o755)

    subprocess.Popen(["/bin/bash", sh_path])
    sys.exit(0)

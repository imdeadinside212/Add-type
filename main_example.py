"""
main_example.py
----------------
Vi du gan updater.py vao app that "Add type". Chay check update tren 1
thread rieng de khong lam dung UI chinh.
"""

import threading
from updater import check_for_update, download_update, apply_update

APP_VERSION = "1.0.0"   # nho dong bo voi tag ban push, vd tag v1.0.0


def run_update_check(on_update_available):
    """
    Chay ngam trong thread rieng. Neu co ban moi, goi callback
    on_update_available(info) de UI hien popup hoi user.
    """
    def worker():
        info = check_for_update(current_version=APP_VERSION)
        if info:
            on_update_available(info)
    threading.Thread(target=worker, daemon=True).start()


def handle_update_available(info):
    print(f"Co ban moi: {info['version']}")
    print(info["notes"])

    # Trong app that: hien popup hoi "Co ban cap nhat moi, ban co
    # muon cai dat ngay khong?" - o day gia dinh user dong y luon.
    user_agree = True
    if user_agree:
        print("Dang tai ban moi...")
        new_file = download_update(info["download_url"])
        print("Dang cai dat va khoi dong lai...")
        apply_update(new_file)   # app se thoat o day


def start_main_app():
    # ... khoi tao UI chinh cua "Add type" o day ...
    print(f"Add type dang chay, phien ban {APP_VERSION}")


if __name__ == "__main__":
    run_update_check(handle_update_available)
    start_main_app()

    # giu app chay (vi du vong lap UI that su se o day thay vi input())
    input("Nhan Enter de thoat...\n")

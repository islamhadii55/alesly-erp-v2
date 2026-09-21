#!/usr/bin/env python3
import os
import sys
import threading
import time
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.chdir(ROOT)
PORT = int(os.environ.get("PORT", "5050"))
HOST = "127.0.0.1"


def run_server():
    from app import app, init_db
    init_db()
    app.run(host=HOST, port=PORT, debug=False, use_reloader=False, threaded=True)


def wait_ready(timeout=20):
    import urllib.request
    url = f"http://{HOST}:{PORT}/login"
    end = time.time() + timeout
    while time.time() < end:
        try:
            urllib.request.urlopen(url, timeout=1)
            return True
        except Exception:
            time.sleep(0.3)
    return False


def main():
    t = threading.Thread(target=run_server, daemon=True)
    t.start()
    wait_ready()
    url = f"http://{HOST}:{PORT}"
    try:
        import webview
        webview.create_window(
            "الاصلي لتجارة قطع الغيار",
            url,
            width=1360,
            height=860,
            min_size=(960, 640),
        )
        webview.start()
    except Exception:
        webbrowser.open(url)
        print("تم فتح النظام في المتصفح. أبقِ هذه النافذة مفتوحة.")
        t.join()


if __name__ == "__main__":
    main()

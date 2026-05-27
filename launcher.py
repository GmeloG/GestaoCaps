import sys
import os
import threading
import webbrowser
import time


def _open_browser():
    time.sleep(4)
    webbrowser.open("http://localhost:8501")


if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        app_script = os.path.join(sys._MEIPASS, "app.py")
    else:
        app_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")

    threading.Thread(target=_open_browser, daemon=True).start()

    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        app_script,
        "--global.developmentMode=false",
        "--server.port=8501",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
    ]
    sys.exit(stcli.main())

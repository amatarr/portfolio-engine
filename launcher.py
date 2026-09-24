import os
import sys
import threading
import webbrowser

from streamlit.web import cli as stcli


def _resource_path(relative_path):

    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))

    return os.path.join(base_path, relative_path)


def _open_browser():
    webbrowser.open("http://localhost:8501")


if __name__ == "__main__":

    threading.Timer(2.0, _open_browser).start()

    sys.argv = [
        "streamlit",
        "run",
        _resource_path("app.py"),
        "--global.developmentMode=false",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
    ]

    sys.exit(stcli.main())

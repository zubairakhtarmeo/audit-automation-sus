import os
import sys


def _base_dir() -> str:
    # In PyInstaller, files are extracted to sys._MEIPASS at runtime.
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def main() -> None:
    app_path = os.path.join(_base_dir(), "ui", "app.py")

    # Run Streamlit in-process (works for both python and PyInstaller builds)
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--server.headless=false",
        "--browser.gatherUsageStats=false",
    ]
    raise SystemExit(stcli.main())


if __name__ == "__main__":
    main()

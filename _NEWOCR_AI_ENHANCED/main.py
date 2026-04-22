"""
main.py — DocExtract Pro
=========================
Entry point. Run this file to launch the application:

    python main.py

Flow:
    1. LoginWindow appears (login.py)
    2. On successful sign-in → DocExtractorApp launches (app.py)

Dependencies (install with pip):
    customtkinter, pillow, paddleocr, groq, pdfplumber, pdf2image,
    openpyxl, pandas, python-docx, opencv-python, python-dotenv
"""

# ── DPI awareness — must be called FIRST, before any tkinter/ctk window ───────
# Prevents blurry UI on Windows high-DPI / scaled displays (125%, 150%, etc.)
import sys
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)   # Per-monitor DPI aware
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()    # Fallback: system DPI aware
        except Exception:
            pass

from dotenv import load_dotenv
load_dotenv()

import customtkinter as ctk
from login import LoginWindow


def launch_main_app(user_id=None, username=None):
    from app import DocExtractorApp
    DocExtractorApp._current_user_id = user_id
    DocExtractorApp._current_username = username
    app = DocExtractorApp()
    app.mainloop()


if __name__ == "__main__":
    login = LoginWindow(on_success=launch_main_app)
    login.mainloop()
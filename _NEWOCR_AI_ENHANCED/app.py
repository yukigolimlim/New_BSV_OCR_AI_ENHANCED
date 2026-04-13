"""
app.py — DocExtract Pro  (Banco San Vicente)
=============================================
Main entry point.  All UI methods live in the ui_* modules and are
attached to DocExtractorApp at import time via each module's attach().

File layout
-----------
  app_constants.py   — colours, fonts, API keys, rule helpers, gemini_chat
  ui_panels.py       — topbar, sidebar, right panel, search, display helpers
  ui_extraction.py   — file browsing, extraction, analysis trigger
  ui_summary.py      — summary tab, charts, parse helpers
  ui_chat.py         — AI chat tab, RAG, knowledge base
  ui_cibi.py         — CIBI 4-stage workflow
  doc_classifier_tab.py — smart document classifier (replaces plain Extracted tab)
  samples_tab.py     — few-shot samples management tab (SamplesTabMixin)
  lu_ui.py / lu_analysis_tab.py  — LU sector & expense risk scanner
  lu_analysis_row_format_patch.py — adds Look-Up Summary row-format support
  lu_simulator_patch.py      — risk simulator enhancements
  lu_loanbal_export_patch.py — export button for Sector vs Loan Balance tab
"""
import ctypes
import sys
import os
import threading
from pathlib import Path

import tkinter as tk
import customtkinter as ctk

from app_constants import (
    SIDEBAR_BG, OFF_WHITE, GEMINI_API_KEY, GEMINI_MODEL, FALLBACK_MODEL,
    _FONT_FAMILY, best_font, register_fonts, F, FMONO, hex_blend,
    gemini_chat, _ai_check_stage1,
    _rule_fuzzy_ncd, _rule_assess_text, _rule_extract_cic_tier,
    SCRIPT_DIR,
)

from widgets import GradientCanvas, Spinner
from samples_tab import SamplesTabMixin
from doc_classifier_tab import DocClassifierTabMixin

import ui_panels
import ui_extraction
import ui_summary
import ui_chat
import ui_cibi

# ── LU Analysis imports ───────────────────────────────────────────────────────
import lu_ui as lu_analysis_tab
import lu_simulator_patch
import lu_loanbal_export_patch

try:
    import lu_analysis_row_format_patch as _row_patch
    _HAS_ROW_PATCH = True
except ImportError:
    _HAS_ROW_PATCH = False

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION CLASS
# ══════════════════════════════════════════════════════════════════════════════

class DocExtractorApp(DocClassifierTabMixin, SamplesTabMixin, ctk.CTk):

    def __init__(self):
        super().__init__()
        register_fonts()
        global _FONT_FAMILY
        import app_constants as _ac
        _ac._FONT_FAMILY = best_font()

        self.title("DocExtract Pro — Banco San Vicente")
        self.configure(fg_color=SIDEBAR_BG)

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{sw}x{sh}+0+0")
        self.resizable(False, False)

        self.overrideredirect(True)
        self._drag_x = self._drag_y = 0

        # ── Core state ────────────────────────────────────────────────────
        self._selected_file  = None
        self._selected_files = []
        self._logo_img       = None
        self._extracted_text = ""

        self._search_matches = []
        self._search_cursor  = -1
        self._last_query     = ""

        self._chat_history:        list     = []
        self._pending_file_type:   str|None = None

        # ── CIBI state ────────────────────────────────────────────────────
        self._cibi_slots = {
            "CIC":     {"path": None, "text": None, "required": False},
            "BANK_CI": {"path": None, "text": None, "required": True},
            "PAYSLIP": {"path": None, "text": None, "required": True},
            "ITR":     {"path": None, "text": None, "required": False},
            "SALN":    {"path": None, "text": None, "required": False},
        }
        self._cibi_template_path: str|None  = None
        self._cibi_status_labels             = {}
        self._cibi_name_labels               = {}
        self._cibi_stage                     = "idle"
        self._cibi_loan_tier                 = "unknown"
        self._cibi_bank_ci_result            = {}
        self._cibi_has_cic                   = False

        # ── Summary / CIBI state ──────────────────────────────────────────
        self._summary_cibi_analysis_text: str      = ""
        self._summary_cibi_excel_path:    str|None = None
        self._summary_cibi_populated:     bool     = False
        self._last_cibi_path:             Path|None= None

        # ── Mixin init ────────────────────────────────────────────────────
        self._samples_init_state()   # SamplesTabMixin

        # ── Build UI ──────────────────────────────────────────────────────
        self._build_ui()
        self.after(100, self._fix_windows_taskbar)

    # ── Font helpers (used by SamplesTabMixin + DocClassifierTabMixin) ────
    def F(self, size: int, weight: str = "normal") -> tuple:
        import app_constants as _ac
        if _ac._FONT_FAMILY is None:
            _ac._FONT_FAMILY = best_font()
        return (_ac._FONT_FAMILY, size, weight)

    def FMONO(self, size: int, weight: str = "normal") -> tuple:
        import tkinter.font as tkfont
        available = set(tkfont.families())
        for f in ("JetBrains Mono", "Cascadia Code", "Consolas", "Courier New"):
            if f in available:
                return (f, size, weight)
        return ("Courier New", size, weight)

    # ── Window management ─────────────────────────────────────────────────
    def _fix_windows_taskbar(self):
        if sys.platform != "win32":
            return
        GWL_EXSTYLE      = -20
        WS_EX_APPWINDOW  = 0x00040000
        WS_EX_TOOLWINDOW = 0x00000080
        hwnd  = ctypes.windll.user32.GetParent(self.winfo_id())
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        style = (style & ~WS_EX_TOOLWINDOW) | WS_EX_APPWINDOW
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
        self.withdraw()
        self.after(150, self._restore_and_focus)

    def _restore_and_focus(self):
        self.deiconify()
        self.lift()
        self.attributes("-topmost", True)
        self.after(300, lambda: self.attributes("-topmost", False))
        self.focus_force()

    def _force_focus(self):
        if sys.platform != "win32":
            self.lift()
            self.attributes("-topmost", True)
            self.after(300, lambda: self.attributes("-topmost", False))
            self.focus_force()

    def _do_minimize(self):
        if sys.platform == "win32":
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            ctypes.windll.user32.ShowWindow(hwnd, 6)
        else:
            self.iconify()

    def _on_restore(self, e):
        self.unbind("<Map>")
        self._force_focus()

    def _drag_start(self, e):
        self._drag_x = e.x_root - self.winfo_x()
        self._drag_y = e.y_root - self.winfo_y()

    def _drag_move(self, e):
        self.geometry(f"+{e.x_root - self._drag_x}+{e.y_root - self._drag_y}")

    # ── UI build ──────────────────────────────────────────────────────────
    def _build_ui(self):
        self._build_topbar()

        body = tk.Frame(self, bg=SIDEBAR_BG)
        body.pack(fill="both", expand=True)

        # ── Sidebar — direct Frame, no canvas/scrollbar wrapper ───────────
        left = tk.Frame(body, bg=SIDEBAR_BG, width=280)
        left.pack(side="left", fill="both")
        left.pack_propagate(False)

        # ── Thin gradient divider between sidebar and content ─────────────
        from app_constants import LIME_BRIGHT, LIME_DARK
        div_canvas = tk.Canvas(body, bg=SIDEBAR_BG, width=2, highlightthickness=0)
        div_canvas.pack(side="left", fill="y")
        div_canvas.bind("<Configure>",
                        lambda e, c=div_canvas: self._vbar_full(c, LIME_BRIGHT, LIME_DARK))

        # ── Right content area ────────────────────────────────────────────
        right = tk.Frame(body, bg=OFF_WHITE)
        right.pack(side="left", fill="both", expand=True)

        self._build_left(left)
        self._build_right(right)
        self.after(2000, self._prewarm_rag)

    def _prewarm_rag(self):
        def _worker():
            try:
                from RAG import get_rag_engine
                get_rag_engine()
            except Exception:
                pass
            try:
                from extraction import prewarm_ocr_engines
                prewarm_ocr_engines()
            except Exception:
                pass
            try:
                from extraction import prewarm_trocr
                prewarm_trocr()
            except Exception:
                pass
        threading.Thread(target=_worker, daemon=True).start()


# ── Attach all UI modules ─────────────────────────────────────────────────────
ui_panels.attach(DocExtractorApp)
ui_extraction.attach(DocExtractorApp)
ui_summary.attach(DocExtractorApp)
ui_chat.attach(DocExtractorApp)
ui_cibi.attach(DocExtractorApp)

# ── Attach LU Analysis (order matters) ───────────────────────────────────────
lu_analysis_tab.attach(DocExtractorApp)          # 1. base tab (includes search)
if _HAS_ROW_PATCH:
    _row_patch.attach(DocExtractorApp)           # 2. row format support
lu_simulator_patch.attach(DocExtractorApp)       # 3. simulator enhancements
lu_loanbal_export_patch.attach(DocExtractorApp)  # 4. loan balance export button

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = DocExtractorApp()
    app.mainloop()
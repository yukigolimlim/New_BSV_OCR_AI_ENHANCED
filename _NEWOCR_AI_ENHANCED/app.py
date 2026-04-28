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
    _FONT_FAMILY, best_font, register_fonts, F, FMONO, hex_blend, set_ui_zoom,
    gemini_chat, _ai_check_stage1,
    _rule_fuzzy_ncd, _rule_assess_text, _rule_extract_cic_tier,
    SCRIPT_DIR,
)

from widgets import GradientCanvas, Spinner
from doc_classifier_tab import DocClassifierTabMixin

import ui_panels
import ui_extraction
import ui_summary
import ui_cibi

# ── LU Analysis imports ───────────────────────────────────────────────────────
import lu_ui as lu_analysis_tab
import lu_simulator_patch
import lu_loanbal_export_patch
import admin_logs
# app.py — at the top, after imports
import psycopg2
from dotenv import load_dotenv

load_dotenv()

# Global connection — created once when app launches
# app.py

def get_db_connection():
    """Create a fresh psycopg2 connection from .env settings."""
    return psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=int(os.getenv("DB_PORT", 5432)),
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
    )

# Test connection on startup
# Test connection on startup
def check_db_on_startup():
    try:
        _conn = get_db_connection()
        _conn.close()
        print("[OK] Database connected successfully")
        return True
    except Exception as e:
        print(f"[ERROR] Database connection failed: {e}")
        import tkinter.messagebox as mb
        mb.showerror(
            "Database Error",
            f"Cannot connect to the database.\n\n{e}\n\nCheck your .env file or server."
        )
        return False

DB_ONLINE = check_db_on_startup()

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

class DocExtractorApp(DocClassifierTabMixin, ctk.CTk):

    def __init__(self, user_id=None, username=None):
        super().__init__()
        self._current_user_id = user_id
        self._current_username = username
        register_fonts()
        global _FONT_FAMILY
        import app_constants as _ac
        _ac._FONT_FAMILY = best_font()
        # ── Database connection (shared across all tabs) ───────────────
        self.db_conn = None
        if DB_ONLINE:
            try:
                self.db_conn = get_db_connection()
            except Exception as e:
                print(f"✘ Could not create shared connection: {e}")

        self.title("DocExtract Pro — Banco San Vicente")
        self.configure(fg_color=SIDEBAR_BG)

        # Get true scaled screen size accounting for DPI awareness
        try:
            import ctypes
            user32 = ctypes.windll.user32
            # GetSystemMetrics with DPI-aware values
            sw = user32.GetSystemMetrics(0)   # SM_CXSCREEN
            sh = user32.GetSystemMetrics(1)   # SM_CYSCREEN
            # Scale back down by the system DPI factor so tkinter lays out correctly
            dpi = user32.GetDpiForSystem()
            scale = dpi / 96.0
            sw = int(sw / scale)
            sh = int(sh / scale)
        except Exception:
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
        self.geometry(f"{sw}x{sh}+0+0")
        self.resizable(False, False)

        self.overrideredirect(True)
        self._drag_x = self._drag_y = 0
        self._is_closing = False
        self.report_callback_exception = self._suppress_closing_errors
        self._ui_zoom = 1.0
        try:
            self._base_tk_scaling = float(self.tk.call("tk", "scaling"))
        except Exception:
            self._base_tk_scaling = 1.0
        # Reset tk scaling to 1.0 baseline — DPI is already handled at process level
        try:
            import ctypes
            dpi = ctypes.windll.user32.GetDpiForSystem()
            scale = dpi / 96.0
            self.tk.call("tk", "scaling", self._base_tk_scaling / scale)
        except Exception:
            pass
        self._apply_ui_zoom(1.0)

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

        # ── Build UI ──────────────────────────────────────────────────────
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._safe_close)
        self.after(100, self._fix_windows_taskbar)
        self.bind_all("<Control-plus>", lambda _e: self._zoom_in())
        self.bind_all("<Control-equal>", lambda _e: self._zoom_in())   # Ctrl+= is + on many layouts
        self.bind_all("<Control-minus>", lambda _e: self._zoom_out())
        self.bind_all("<Control-0>", lambda _e: self._zoom_reset())
        self.bind_all("<Control-KP_Add>", lambda _e: self._zoom_in())
        self.bind_all("<Control-KP_Subtract>", lambda _e: self._zoom_out())

    # ── Database ──────────────────────────────────────────────────────
    def get_conn(self):
        try:
            if self.db_conn is None or self.db_conn.closed:
                raise Exception("closed")
            cur = self.db_conn.cursor()
            cur.execute("SELECT 1")
            cur.close()
        except Exception:
            try:
                self.db_conn = get_db_connection()
            except Exception as e:
                print(f"✘ Reconnect failed: {e}")
                self.db_conn = None
        return self.db_conn

    # ── Font helpers (used by SamplesTabMixin + DocClassifierTabMixin) ────
    def F(self, size: int, weight: str = "normal") -> tuple:
        import app_constants as _ac
        if _ac._FONT_FAMILY is None:
            _ac._FONT_FAMILY = best_font()
        return (_ac._FONT_FAMILY, max(6, int(round(size * self._ui_zoom))), weight)

    def FMONO(self, size: int, weight: str = "normal") -> tuple:
        import tkinter.font as tkfont
        available = set(tkfont.families())
        for f in ("JetBrains Mono", "Cascadia Code", "Consolas", "Courier New"):
            if f in available:
                return (f, max(6, int(round(size * self._ui_zoom))), weight)
        return ("Courier New", max(6, int(round(size * self._ui_zoom))), weight)

    # ── Global UI zoom ─────────────────────────────────────────────────────
    def _apply_ui_zoom(self, value: float):
        z = max(0.8, min(1.6, float(value)))
        self._ui_zoom = z
        set_ui_zoom(z)
        try:
            self.tk.call("tk", "scaling", self._base_tk_scaling * z)
        except Exception:
            pass
        lbl = getattr(self, "_zoom_lbl", None)
        if lbl is not None:
            try:
                lbl.config(text=f"{int(round(z * 100))}%")
            except Exception:
                pass

    def _zoom_in(self):
        self._apply_ui_zoom(round(self._ui_zoom + 0.1, 2))
        return "break"

    def _zoom_out(self):
        self._apply_ui_zoom(round(self._ui_zoom - 0.1, 2))
        return "break"

    def _zoom_reset(self):
        self._apply_ui_zoom(1.0)
        return "break"

    def _ui_refresh(self):
        """
        Best-effort "re-render" of the active screen after zoom/theme changes.
        Avoids destroying global state; just re-calls the active view renderer(s).
        """
        try:
            self.update_idletasks()
        except Exception:
            pass

        tab = getattr(self, "_current_tab", "") or ""
        # Summary tab: refresh table/views
        if tab == "lookup_summary":
            try:
                from summary_tab import _refresh_summary
                _refresh_summary(self)
                return
            except Exception:
                pass

        # LU Analysis container: refresh the active LU sub-view if present
        if tab == "lu_analysis":
            try:
                # If LU panel exists, re-render whichever subtab is active
                if getattr(self, "_lu_all_data", None):
                    view_var = getattr(self, "_lu_active_view", None)
                    view = view_var.get() if view_var is not None else "analysis"
                    try:
                        from lu_ui import _lu_switch_view
                        _lu_switch_view(self, view)
                        return
                    except Exception:
                        pass
                    # Fallbacks (older attach paths)
                    if view == "charts" and hasattr(self, "_charts_render"):
                        self._charts_render()
                        return
                    if view == "loanbal" and hasattr(self, "_loanbal_render"):
                        self._loanbal_render()
                        return
                    if view == "report" and hasattr(self, "_report_render"):
                        self._report_render()
                        return
                    if view == "simulator" and hasattr(self, "_sim_populate"):
                        self._sim_populate()
                        return
            except Exception:
                pass

        # Generic fallback: re-pack current tab (no-op but safe)
        try:
            if hasattr(self, "_switch_tab"):
                self._switch_tab(tab)
        except Exception:
            pass

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

    def _safe_close(self):
        """Best-effort shutdown: cancel pending UI jobs, then destroy once."""
        if getattr(self, "_is_closing", False):
            return
        self._is_closing = True

        # Known recurring after() jobs across tabs/mixins.
        job_attrs = (
            "_sim_build_job",
            "_sum_search_after",
            "_gen_sum_search_after",
            "_job",
        )
        for attr in job_attrs:
            job = getattr(self, attr, None)
            if job:
                try:
                    self.after_cancel(job)
                except Exception:
                    pass
                try:
                    setattr(self, attr, None)
                except Exception:
                    pass
        # Cancel any remaining Tk/CustomTkinter scheduled callbacks so Tcl
        # does not try to run them after the root window has been destroyed.
        try:
            pending = self.tk.splitlist(self.tk.call("after", "info"))
        except Exception:
            pending = ()
        for job in pending:
            try:
                self.after_cancel(job)
            except Exception:
                pass
        # ── Close shared DB connection ─────────────────────────────
        # ── Close matplotlib figures (dashboard) ───────────────────
        try:
            import matplotlib.pyplot as plt
            plt.close("all")
        except Exception:
            pass

        # ── Close shared DB connection ─────────────────────────────
        if getattr(self, "db_conn", None) and not self.db_conn.closed:
            try:
                self.db_conn.close()
            except Exception:
                pass

        try:
            self.destroy()
        except Exception:
            # Last fallback to avoid hanging when widgets are half-destroyed.
            try:
                self.quit()
            except Exception:
                pass

    def _suppress_closing_errors(self, exc, val, tb):
        if getattr(self, '_is_closing', False):
            return  # silently swallow Tkinter errors during shutdown
        import traceback
        traceback.print_exception(exc, val, tb)

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
            if getattr(self, "_is_closing", False):
                return
            try:
                from extraction import prewarm_ocr_engines
                prewarm_ocr_engines()
            except Exception:
                pass
            if getattr(self, "_is_closing", False):
                return
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
ui_cibi.attach(DocExtractorApp)

# ── Attach LU Analysis (order matters) ───────────────────────────────────────
lu_analysis_tab.attach(DocExtractorApp)          # 1. base tab (includes search)
if _HAS_ROW_PATCH:
    _row_patch.attach(DocExtractorApp)           # 2. row format support
lu_simulator_patch.attach(DocExtractorApp)       # 3. simulator enhancements
lu_loanbal_export_patch.attach(DocExtractorApp)  # 4. loan balance export button

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    def _launch(user_id, username):
        app = DocExtractorApp()
        app._current_user_id = user_id
        app._current_username = username
        app.mainloop()

    from login import LoginWindow
    login = LoginWindow(on_success=_launch)
    login.mainloop()
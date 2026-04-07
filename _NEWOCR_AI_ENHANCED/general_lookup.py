import re
import io
import json
import time
import threading
import concurrent.futures
import tkinter as tk
import customtkinter as ctk
from pathlib import Path
from tkinter import filedialog
from datetime import datetime

from app_constants import (
    NAVY_DEEP, NAVY_LIGHT, NAVY_MID, NAVY_PALE, NAVY_MIST, NAVY_GHOST,
    WHITE, OFF_WHITE, CARD_WHITE,
    LIME_BRIGHT, LIME_DARK, LIME_MID, LIME_PALE, LIME_MIST,
    TXT_NAVY, TXT_SOFT, TXT_MUTED, TXT_ON_LIME,
    ACCENT_RED, ACCENT_GOLD, ACCENT_SUCCESS,
    BORDER_LIGHT, BORDER_MID,
    SIDEBAR_BG, SIDEBAR_ITEM, SIDEBAR_HVR,
    F, FF, FMONO,
    IMAGE_EXTS,
)

# ── Import from general_summary instead of summary_tab ───────────────
from general_summary import db_save_applicant, lookup_summary_notify

# ── Category master list ──────────────────────────────────────────────
LOOKUP_ROWS = [
    ("cibi_place_of_work",      "CI/BI Report",      "Office Address",                      "cibi"),
    ("cibi_temp_residence",     "CI/BI Report",      "Residence Address",                    "cibi"),
    ("cibi_spouse",             "CI/BI Report",      "Spouse / Employment",                  "cibi"),
    ("cibi_spouse_office",      "CI/BI Report",      "Spouse Office Address",                "cibi"),
    ("cibi_personal_assets",    "CI/BI Report",      "Personal Assets",                      "cibi"),
    ("cibi_business_assets",    "CI/BI Report",      "Business Assets",                      "cibi"),
    ("cibi_petrol_products",    "CI/BI Report",      "Petrol / Plastics / PVC Risk",         "cibi"),
    ("cibi_transport_services", "CI/BI Report",      "Transport Services Risk",              "cibi"),
    # ── Credit History amortization rows ────────────────────────
    ("credit_history_amort",    "CI/BI Report",      "Credit History Amort.",                "cibi"),
    ("income_remittance",       "Cashflow Analysis", "Source of Income",                     "cfa"),
    ("cfa_business_expenses",   "Cashflow Analysis", "Business Expenses",                    "cfa"),
    ("cfa_household_expenses",  "Cashflow Analysis", "Household / Personal Expenses",        "cfa"),
    ("ws_food_grocery",         "Worksheet",         "Food / Grocery",                       "ws"),
    ("ws_fuel_transport",       "Worksheet",         "Fuel and Transportation",              "ws"),
    ("ws_electricity",          "Worksheet",         "Electricity Expense",                  "ws"),
    ("ws_fertilizer",           "Worksheet",         "Fertilizer",                           "ws"),
    ("ws_forwarding",           "Worksheet",         "Forwarding / Trucking / Hauling",      "ws"),
    ("ws_fuel_diesel",          "Worksheet",         "Fuel / Gas / Diesel",                  "ws"),
    ("ws_equipment",            "Worksheet",         "Cost of Rent of Equipment",            "ws"),
]

SECTION_META = {
    "CI/BI Report":      "Credit Risk",
    "Cashflow Analysis": "Income / Expenses",
    "Worksheet":         "Business Worksheet",
}

DOC_TYPE_KEYWORDS = {
    "credit_scoring": [
        "credit scor", "scoring sheet", "credit rating", "risk score",
    ],
    "cibi": [
        "ci/bi", "cibi", "credit information", "bureau of internal",
        "place of work", "temporary residence", "employer",
        "trade references", "bank ci", "bank credit information",
        "character investigation",
        "credit history", "credit history & references",
        "bank/lending institution", "principal loan", "amort.",
        "amortization", "balance",
        "name of spouse", "spouse", "employed", "self-employed",
        "personal assets", "business assets", "personal and business assets",
        "serialized", "household assets", "personal vehicles",
        "business vehicles", "vehicle",
    ],
    "cfa": [
        "cash flow", "cashflow", "cash-flow",
        "source of income", "business income", "household expenses",
        "monthly expenses", "personal expenses", "net income",
        "remittance", "padala",
    ],
    "worksheet": [
        "worksheet", "work sheet",
        "fertilizer", "forwarding", "trucking", "hauling",
        "fuel", "diesel", "equipment rental", "food", "grocery",
        "electricity", "business expense worksheet",
    ],
}

PAD_VALUE = 28
QUEUE_COLORS = {
    "waiting": ("#F3F4F6", "#6B7280"),
    "running": ("#FFFBEB", "#92400E"),
    "done":    ("#F0FDF4", "#166534"),
    "error":   ("#FEF2F2", "#991B1B"),
}

# ── Concurrency / rate-limit settings ────────────────────────────────
MAX_CONCURRENT_FILES = 1
MAX_PARALLEL_CALLS   = 1
INTER_CALL_DELAY_S   = 10
INTER_FILE_DELAY_S   = 35

# Retry settings
GEMINI_MAX_RETRIES   = 3
GEMINI_RETRY_DELAYS  = [60, 90, 120]

# ── Global Gemini rate-limit gate ─────────────────────────────────────
_GEMINI_CALL_LOCK = threading.Lock()
_GEMINI_LAST_CALL = 0.0
_GEMINI_MIN_GAP_S = 10.0


# ═══════════════════════════════════════════════════════════════════════
#  PDF SPLITTING HELPER
# ═══════════════════════════════════════════════════════════════════════

def _auto_rotate_pdf(pdf_bytes: bytes) -> bytes:
    try:
        import fitz

        doc = fitz.open(stream=io.BytesIO(pdf_bytes), filetype="pdf")
        modified = False

        for page in doc:
            if page.rotation != 0:
                page.set_rotation(0)
                modified = True

            try:
                import pytesseract
                from PIL import Image
                import numpy as np

                mat  = fitz.Matrix(72 / 72, 72 / 72)
                pix  = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
                img  = Image.frombytes("L", (pix.width, pix.height), pix.samples)
                osd  = pytesseract.image_to_osd(
                    img, output_type=pytesseract.Output.DICT,
                    config="--psm 0")
                angle = int(osd.get("rotate", 0))
                if angle != 0:
                    page.set_rotation((360 - angle) % 360)
                    modified = True
            except Exception:
                pass

        if not modified:
            return pdf_bytes

        buf = io.BytesIO()
        doc.save(buf, garbage=0, deflate=False, clean=False)
        return buf.getvalue()

    except Exception:
        return pdf_bytes


def _extract_page_subset(pdf_bytes: bytes, page_numbers: list) -> bytes:
    if not page_numbers:
        return pdf_bytes
    try:
        import fitz
        src = fitz.open(stream=io.BytesIO(pdf_bytes), filetype="pdf")
        dst = fitz.open()
        for pg in page_numbers:
            zero = pg - 1
            if 0 <= zero < src.page_count:
                dst.insert_pdf(src, from_page=zero, to_page=zero)
        if dst.page_count == 0:
            return pdf_bytes
        buf = io.BytesIO()
        dst.save(buf, garbage=0, deflate=False, clean=False)
        return buf.getvalue()
    except Exception:
        return pdf_bytes


def _pages_for_type(page_map: dict, doc_type: str,
                    total_pages: int, padding: int = 1) -> list:
    matched = sorted(pg for pg, t in page_map.items() if t == doc_type)
    if not matched:
        return list(range(1, total_pages + 1))
    expanded = set()
    for pg in matched:
        for offset in range(-padding, padding + 1):
            n = pg + offset
            if 1 <= n <= total_pages:
                expanded.add(n)
    return sorted(expanded)


# ═══════════════════════════════════════════════════════════════════════
#  PANEL BUILDER
# ═══════════════════════════════════════════════════════════════════════

def _build_general_lookup_panel(self, parent):
    outer = tk.Frame(parent, bg=CARD_WHITE)
    self._general_lookup_frame = outer

    canvas = tk.Canvas(outer, bg=CARD_WHITE, highlightthickness=0, bd=0)
    vsb = tk.Scrollbar(outer, orient="vertical", command=canvas.yview,
                       relief="flat", troughcolor=OFF_WHITE,
                       bg=BORDER_LIGHT, width=7, bd=0)
    canvas.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    inner = tk.Frame(canvas, bg=CARD_WHITE)
    _win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

    canvas.bind("<Configure>",
                lambda e: canvas.itemconfig(_win_id, width=e.width))
    inner.bind("<Configure>",
               lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Enter>",
                lambda e: canvas.bind_all(
                    "<MouseWheel>",
                    lambda ev: canvas.yview_scroll(
                        int(-1 * (ev.delta / 120)), "units")))
    canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

    self._general_lookup_inner  = inner
    self._general_lookup_canvas = canvas
    _populate_general_lookup(self, inner)


def _populate_general_lookup(self, parent):
    PAD = PAD_VALUE

    # ── Header ─────────────────────────────────────────────────────────
    hdr = tk.Frame(parent, bg=CARD_WHITE)
    hdr.pack(fill="x", padx=PAD, pady=(28, 0))
    left = tk.Frame(hdr, bg=CARD_WHITE)
    left.pack(side="left")
    tk.Label(left, text="General Look-Up", font=F(20, "bold"),
             fg=NAVY_DEEP, bg=CARD_WHITE).pack(anchor="w")
    tk.Label(left,
             text="One PDF per applicant  ·  each processed independently"
                  "  ·  results saved to General Summary tab",
             font=F(9), fg=TXT_SOFT, bg=CARD_WHITE).pack(
                 anchor="w", pady=(2, 0))
    badge = tk.Frame(hdr, bg="#EEF6FF",
                     highlightbackground="#4F8EF7", highlightthickness=1)
    badge.pack(side="right", pady=4)
    tk.Label(badge, text="  Gemini 2.5 Flash · Parallel  ",
             font=F(8, "bold"), fg="#4F8EF7", bg="#EEF6FF", pady=4).pack()

    tk.Frame(parent, bg=BORDER_LIGHT, height=1).pack(
        fill="x", padx=PAD, pady=(16, 0))

    # ── Upload zone ────────────────────────────────────────────────────
    upload_sec = tk.Frame(parent, bg=CARD_WHITE)
    upload_sec.pack(fill="x", padx=PAD, pady=(20, 0))
    tk.Label(upload_sec, text="PDF FILES  (one per applicant)",
             font=F(7, "bold"), fg=TXT_MUTED, bg=CARD_WHITE).pack(
                 anchor="w", pady=(0, 6))

    drop = tk.Frame(upload_sec, bg="#F7FAFF",
                    highlightbackground="#C2D8F5",
                    highlightthickness=2, height=76)
    drop.pack(fill="x")
    drop.pack_propagate(False)
    drop_inner = tk.Frame(drop, bg="#F7FAFF")
    drop_inner.place(relx=0.5, rely=0.5, anchor="center")
    self._general_lookup_icon_lbl = tk.Label(drop_inner, text="📂",
                                     font=("Segoe UI Emoji", 18),
                                     fg="#4F8EF7", bg="#F7FAFF")
    self._general_lookup_icon_lbl.pack(side="left", padx=(0, 8))
    self._general_lookup_file_lbl = tk.Label(
        drop_inner, text="No files selected",
        font=F(9), fg=TXT_SOFT, bg="#F7FAFF")
    self._general_lookup_file_lbl.pack(side="left")

    btn_row = tk.Frame(upload_sec, bg=CARD_WHITE)
    btn_row.pack(fill="x", pady=(10, 0))

    self._general_lookup_browse_btn = ctk.CTkButton(
        btn_row, text="Choose PDF(s)",
        command=lambda: _general_lookup_browse(self),
        width=140, height=36, corner_radius=8,
        fg_color="#4F8EF7", hover_color="#3A7EE8",
        text_color=WHITE, font=FF(9, "bold"), border_width=0)
    self._general_lookup_browse_btn.pack(side="left")

    self._general_lookup_run_btn = ctk.CTkButton(
        btn_row, text="⚡  Run Look-Up",
        command=lambda: _general_lookup_run(self),
        width=150, height=36, corner_radius=8,
        fg_color=NAVY_LIGHT, hover_color=NAVY_PALE,
        text_color=WHITE, font=FF(9, "bold"),
        state="disabled", border_width=0)
    self._general_lookup_run_btn.pack(side="left", padx=(10, 0))

    self._general_lookup_clear_btn = ctk.CTkButton(
        btn_row, text="✕  Clear All",
        command=lambda: _general_lookup_clear(self),
        width=90, height=36, corner_radius=8,
        fg_color=CARD_WHITE, hover_color="#FEF2F2",
        text_color=ACCENT_RED, font=FF(9),
        border_width=1, border_color="#FCA5A5", state="disabled")
    self._general_lookup_clear_btn.pack(side="left", padx=(8, 0))

    self._general_lookup_overall_lbl = tk.Label(
        upload_sec, text="", font=F(8, "bold"),
        fg=TXT_SOFT, bg=CARD_WHITE)
    self._general_lookup_overall_lbl.pack(anchor="w", pady=(6, 0))

    # ── Overall progress bar ───────────────────────────────────────────
    self._general_lookup_prog_var = tk.DoubleVar(value=0.0)
    self._general_lookup_prog_bar = ctk.CTkProgressBar(
        parent, variable=self._general_lookup_prog_var,
        height=6, corner_radius=3,
        fg_color="#E8F0FA", progress_color="#4F8EF7", border_width=0)

    # ── Applicant queue ────────────────────────────────────────────────
    tk.Frame(parent, bg=BORDER_LIGHT, height=1).pack(
        fill="x", padx=PAD, pady=(18, 0))

    queue_hdr = tk.Frame(parent, bg=CARD_WHITE)
    queue_hdr.pack(fill="x", padx=PAD, pady=(12, 0))
    tk.Label(queue_hdr, text="APPLICANT QUEUE",
             font=F(7, "bold"), fg=TXT_MUTED, bg=CARD_WHITE).pack(side="left")
    self._general_lookup_queue_count_lbl = tk.Label(
        queue_hdr, text="", font=F(7), fg=TXT_MUTED, bg=CARD_WHITE)
    self._general_lookup_queue_count_lbl.pack(side="right")

    self._general_lookup_queue_frame = tk.Frame(parent, bg=CARD_WHITE)
    self._general_lookup_queue_frame.pack(fill="x", padx=PAD, pady=(6, 0))

    # ── Raw log (collapsible) ──────────────────────────────────────────
    raw_wrap = tk.Frame(parent, bg=CARD_WHITE)
    raw_wrap.pack(fill="x", padx=PAD, pady=(20, 0))
    raw_toggle_row = tk.Frame(raw_wrap, bg=CARD_WHITE)
    raw_toggle_row.pack(fill="x")
    tk.Label(raw_toggle_row, text="RAW LOG",
             font=F(7, "bold"), fg=TXT_MUTED, bg=CARD_WHITE).pack(side="left")
    self._general_lookup_raw_toggle_btn = ctk.CTkButton(
        raw_toggle_row, text="▼ Show",
        command=lambda: _general_toggle_raw(self),
        width=72, height=24, corner_radius=6,
        fg_color=OFF_WHITE, hover_color=NAVY_MIST,
        text_color=TXT_SOFT, font=FF(8),
        border_width=1, border_color=BORDER_LIGHT)
    self._general_lookup_raw_toggle_btn.pack(side="right")

    self._general_lookup_raw_frame = tk.Frame(raw_wrap, bg=CARD_WHITE)
    raw_sb = tk.Scrollbar(self._general_lookup_raw_frame, relief="flat",
                          troughcolor=OFF_WHITE, bg=BORDER_LIGHT,
                          width=6, bd=0)
    raw_sb.pack(side="right", fill="y")
    self._general_lookup_raw_box = tk.Text(
        self._general_lookup_raw_frame, wrap="word", font=FMONO(8),
        fg=TXT_SOFT, bg="#F7FAFF", relief="flat", bd=0,
        padx=12, pady=8, height=14, state="disabled",
        yscrollcommand=raw_sb.set,
        selectbackground=NAVY_GHOST, selectforeground=TXT_NAVY)
    self._general_lookup_raw_box.pack(side="left", fill="both", expand=True)
    raw_sb.config(command=self._general_lookup_raw_box.yview)
    self._general_lookup_raw_visible = False

    tk.Frame(parent, bg=CARD_WHITE, height=32).pack()

    # ── Internal state ─────────────────────────────────────────────────
    self._general_lookup_filepaths    = []
    self._general_lookup_cancel       = threading.Event()
    self._general_lookup_file_data    = {}
    self._general_lookup_raw_log      = []
    self._general_lookup_raw_lock     = threading.Lock()
    self._general_lookup_done_count   = 0
    self._general_lookup_done_lock    = threading.Lock()
    self._general_lookup_gemini_cache = {}


# ═══════════════════════════════════════════════════════════════════════
#  QUEUE ROW BUILDER
# ═══════════════════════════════════════════════════════════════════════

def _build_general_queue_row(self, path: str, index: int):
    parent = self._general_lookup_queue_frame
    bg, fg = QUEUE_COLORS["waiting"]

    row_outer = tk.Frame(parent, bg=bg,
                         highlightbackground=BORDER_LIGHT,
                         highlightthickness=1)
    row_outer.pack(fill="x", pady=(0, 5))

    top = tk.Frame(row_outer, bg=bg)
    top.pack(fill="x", padx=10, pady=(8, 2))

    tk.Label(top, text=f"{index+1:02d}",
             font=F(8, "bold"), fg=fg, bg=bg,
             width=3, anchor="e").pack(side="left")

    status_lbl = tk.Label(top, text="  WAITING  ",
                          font=F(7, "bold"), fg=fg, bg=bg)
    status_lbl.pack(side="left", padx=(6, 0))

    tk.Label(top, text=Path(path).name,
             font=F(9, "bold"), fg=NAVY_DEEP, bg=bg,
             anchor="w").pack(side="left", padx=(10, 0))

    expand_btn = ctk.CTkButton(
        top, text="▼ Details",
        command=lambda p=path: _toggle_general_applicant_detail(self, p),
        width=80, height=24, corner_radius=6,
        fg_color=OFF_WHITE, hover_color=NAVY_MIST,
        text_color=TXT_SOFT, font=FF(8),
        border_width=1, border_color=BORDER_LIGHT,
        state="disabled")
    expand_btn.pack(side="right")

    step_lbl = tk.Label(row_outer, text="Waiting…",
                        font=F(8), fg=TXT_MUTED, bg=bg,
                        padx=10, anchor="w")
    step_lbl.pack(fill="x", pady=(0, 1))

    name_lbl = tk.Label(row_outer, text="",
                        font=F(9, "bold"), fg=NAVY_DEEP, bg=bg,
                        padx=10, anchor="w")
    name_lbl.pack(fill="x")

    xlsx_lbl   = tk.Label(row_outer)
    folder_btn = tk.Label(row_outer)

    detail_frame = tk.Frame(row_outer, bg=CARD_WHITE,
                            highlightbackground=BORDER_LIGHT,
                            highlightthickness=1)

    self._general_lookup_file_data[path]["widgets"] = {
        "row_outer":    row_outer,
        "top":          top,
        "status_lbl":   status_lbl,
        "step_lbl":     step_lbl,
        "name_lbl":     name_lbl,
        "xlsx_lbl":     xlsx_lbl,
        "folder_btn":   folder_btn,
        "expand_btn":   expand_btn,
        "detail_frame": detail_frame,
        "expanded":     False,
        "bg":           bg,
    }


def _set_general_row_status(self, path: str, status: str, step_text: str = ""):
    d = self._general_lookup_file_data.get(path, {})
    w = d.get("widgets")
    if not w:
        return
    bg, fg = QUEUE_COLORS.get(status, QUEUE_COLORS["waiting"])
    w["bg"] = bg
    for widget in (w["row_outer"], w["top"], w["step_lbl"], w["name_lbl"]):
        try:
            widget.config(bg=bg)
        except Exception:
            pass
    for child in w["top"].winfo_children():
        try:
            child.config(bg=bg)
        except Exception:
            pass
    w["status_lbl"].config(text=f"  {status.upper()}  ", fg=fg, bg=bg)
    if step_text:
        w["step_lbl"].config(
            text=step_text,
            fg=fg if status != "waiting" else TXT_MUTED)


def _toggle_general_applicant_detail(self, path: str):
    d = self._general_lookup_file_data.get(path, {})
    w = d.get("widgets")
    if not w:
        return
    detail   = w["detail_frame"]
    expanded = w["expanded"]
    if expanded:
        detail.pack_forget()
        w["expand_btn"].configure(text="▼ Details")
        w["expanded"] = False
    else:
        if not detail.winfo_children():
            _build_general_detail_panel(self, path, detail)
        detail.pack(fill="x", padx=10, pady=(0, 8))
        w["expand_btn"].configure(text="▲ Hide")
        w["expanded"] = True


def _build_general_detail_panel(self, path: str, parent: tk.Frame):
    d       = self._general_lookup_file_data.get(path, {})
    results = d.get("results", {})
    pagemap = d.get("page_map", "")
    bg      = CARD_WHITE

    if pagemap:
        tk.Label(parent, text="Page map:",
                 font=F(7, "bold"), fg=TXT_MUTED, bg=bg,
                 padx=8).pack(anchor="w", pady=(6, 0))
        tk.Label(parent, text=pagemap,
                 font=FMONO(7), fg=TXT_SOFT, bg=bg,
                 padx=16, justify="left",
                 anchor="w").pack(anchor="w")

    tk.Frame(parent, bg=BORDER_LIGHT, height=1).pack(
        fill="x", padx=8, pady=(6, 4))

    last_section = None
    row_index    = 0

    for key, section, row_label, _src in LOOKUP_ROWS:
        if section != last_section:
            last_section = section
            sec_bar = tk.Frame(parent, bg=NAVY_MID)
            sec_bar.pack(fill="x", padx=8)
            tk.Label(sec_bar, text=f"  {section}",
                     font=F(7, "bold"), fg=WHITE, bg=NAVY_MID,
                     pady=3).pack(side="left")

        row_bg = "#F8FAFF" if row_index % 2 == 0 else WHITE
        row_f  = tk.Frame(parent, bg=row_bg,
                          highlightbackground="#E5EAF3",
                          highlightthickness=1)
        row_f.pack(fill="x", padx=8)

        tk.Label(row_f, text=row_label, font=F(8, "bold"),
                 fg=NAVY_DEEP, bg=row_bg,
                 padx=8, pady=5, anchor="w",
                 width=28).pack(side="left")

        data  = results.get(key, {})
        items = data.get("items", [])
        total = data.get("total")
        non_m = key in ("cibi_place_of_work", "cibi_temp_residence",
                        "cibi_spouse", "cibi_spouse_office",
                        "cibi_personal_assets", "cibi_business_assets")

        amt_txt = total if (total and not non_m) else "—"
        tk.Label(row_f, text=amt_txt, font=F(8),
                 fg=NAVY_DEEP if amt_txt != "—" else TXT_MUTED,
                 bg=row_bg, padx=6, width=12,
                 anchor="e").pack(side="left")

        det_txt = ("  |  ".join(items[:4]) +
                   (f"  (+{len(items)-4} more)" if len(items) > 4 else "")
                   if items else "No data found")
        tk.Label(row_f, text=det_txt, font=F(7),
                 fg=TXT_NAVY if items else TXT_MUTED,
                 bg=row_bg, padx=6, anchor="w",
                 wraplength=380,
                 justify="left").pack(side="left", fill="x", expand=True)

        row_index += 1


# ═══════════════════════════════════════════════════════════════════════
#  ACTIONS
# ═══════════════════════════════════════════════════════════════════════

def _general_lookup_browse(self):
    paths = filedialog.askopenfilenames(
        title="Select PDF file(s) — one per applicant",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
    if not paths:
        return

    existing  = set(self._general_lookup_filepaths)
    new_paths = [p for p in paths if p not in existing]
    self._general_lookup_filepaths.extend(new_paths)

    for p in new_paths:
        self._general_lookup_file_data[p] = {
            "status":    "waiting",
            "name":      "",
            "results":   {},
            "page_map":  "",
            "error":     "",
            "gate_data": {},
            "widgets":   {},
        }
        _build_general_queue_row(self, p, self._general_lookup_filepaths.index(p))

    n = len(self._general_lookup_filepaths)
    self._general_lookup_icon_lbl.config(text="📄")
    self._general_lookup_file_lbl.config(
        text=f"{n} applicant PDF{'s' if n > 1 else ''} loaded",
        fg=NAVY_DEEP)
    self._general_lookup_queue_count_lbl.config(text=f"{n} file(s)")
    self._general_lookup_run_btn.configure(state="normal")
    self._general_lookup_clear_btn.configure(state="normal")
    self._general_lookup_cancel.clear()


def _general_lookup_clear(self):
    self._general_lookup_cancel.set()
    self._general_lookup_filepaths    = []
    self._general_lookup_file_data    = {}
    self._general_lookup_raw_log      = []
    self._general_lookup_gemini_cache = {}
    for w in self._general_lookup_queue_frame.winfo_children():
        w.destroy()
    self._general_lookup_file_lbl.config(text="No files selected", fg=TXT_SOFT)
    self._general_lookup_icon_lbl.config(text="📂")
    self._general_lookup_queue_count_lbl.config(text="")
    self._general_lookup_run_btn.configure(state="disabled")
    self._general_lookup_clear_btn.configure(state="disabled")
    self._general_lookup_overall_lbl.config(text="", fg=TXT_SOFT)
    _general_set_raw(self, "")
    try:
        self._general_lookup_prog_bar.pack_forget()
    except Exception:
        pass


def _general_lookup_run(self):
    if not self._general_lookup_filepaths:
        return
    self._general_lookup_cancel.clear()
    self._general_lookup_raw_log    = []
    self._general_lookup_done_count = 0

    self._general_lookup_session_id = datetime.now().isoformat(timespec="seconds")

    self._general_lookup_run_btn.configure(state="disabled", text="Running…")
    self._general_lookup_prog_var.set(0.0)
    self._general_lookup_prog_bar.pack(fill="x", padx=PAD_VALUE, pady=(8, 0))
    self._general_lookup_overall_lbl.config(text="Processing…", fg=ACCENT_GOLD)

    try:
        client, gt = _get_gemini_client(self)
        self._general_lookup_gemini_cache = {"client": client, "gt": gt}
    except Exception as exc:
        self._general_lookup_overall_lbl.config(
            text=f"Gemini init failed: {exc}", fg=ACCENT_RED)
        self._general_lookup_run_btn.configure(state="normal", text="⚡  Run Look-Up")
        return

    threading.Thread(
        target=_general_lookup_worker, args=(self,), daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  MAIN WORKER
# ═══════════════════════════════════════════════════════════════════════

def _general_lookup_worker(self):
    paths = list(self._general_lookup_filepaths)
    total = len(paths)

    def cancelled():
        return self._general_lookup_cancel.is_set()

    with concurrent.futures.ThreadPoolExecutor(
            max_workers=min(MAX_CONCURRENT_FILES, total),
            thread_name_prefix="gen_lookup_file") as pool:

        futures = {}
        for i, path in enumerate(paths):
            if cancelled():
                break
            if i > 0 and INTER_FILE_DELAY_S > 0:
                self._general_lookup_cancel.wait(timeout=INTER_FILE_DELAY_S)
                if cancelled():
                    break
            futures[pool.submit(
                _general_process_single_file_safe, self, path, cancelled)] = path

        for future in concurrent.futures.as_completed(futures):
            if cancelled():
                for f in futures:
                    f.cancel()
                break

            try:
                future.result()
                with self._general_lookup_done_lock:
                    self._general_lookup_done_count += 1
            except Exception:
                pass

            with self._general_lookup_done_lock:
                done = self._general_lookup_done_count

            _general_ui(self, lambda v=done / total:
                self._general_lookup_prog_var.set(v))

    if not cancelled():
        with self._general_lookup_done_lock:
            done = self._general_lookup_done_count
        errors = total - done

        s = f"✓  {done}/{total} applicant(s) processed"
        if errors:
            s += f"  ·  {errors} error(s) — see queue"
        _general_ui(self, lambda msg=s:
            self._general_lookup_overall_lbl.config(
                text=msg,
                fg=ACCENT_SUCCESS if not errors else ACCENT_GOLD))

    _general_ui(self, lambda: self._general_lookup_run_btn.configure(
        state="normal", text="⚡  Run Look-Up"))
    _general_ui(self, lambda: self._general_lookup_prog_bar.pack_forget())


def _general_process_single_file_safe(self, path: str, cancelled) -> None:
    try:
        _general_process_single_file(self, path, cancelled)
        self._general_lookup_file_data[path]["status"] = "done"
    except Exception as exc:
        import traceback
        _general_log(self, f"\n[ERROR — {Path(path).name}]\n{traceback.format_exc()}")
        _general_ui(self, lambda p=path, e=str(exc):
            _set_general_row_status(self, p, "error", f"Error: {e[:120]}"))
        self._general_lookup_file_data[path]["error"] = str(exc)
        self._general_lookup_file_data[path]["status"] = "error"
        raise


# ═══════════════════════════════════════════════════════════════════════
#  SINGLE-FILE PROCESSOR
# ═══════════════════════════════════════════════════════════════════════

def _general_process_single_file(self, path: str, cancelled) -> None:
    fname  = Path(path).name
    client = self._general_lookup_gemini_cache["client"]
    gt     = self._general_lookup_gemini_cache["gt"]

    def step(msg: str):
        _general_ui(self, lambda m=msg:
            _set_general_row_status(self, path, "running", m))

    _general_log(self, f"\n{'═'*60}\nFILE: {fname}\n{'═'*60}")

    step("Reading PDF…")
    try:
        pdf_bytes = Path(path).read_bytes()
        _general_log(self, f"[{fname}] {len(pdf_bytes):,} bytes (raw)")
    except Exception as e:
        _general_log(self, f"[{fname}] Read failed: {e}")
        raise

    original_pdf_bytes = pdf_bytes

    step("Correcting page orientation…")
    pdf_bytes = _auto_rotate_pdf(pdf_bytes)
    _general_log(self, f"[{fname}] {len(pdf_bytes):,} bytes (after rotation fix)")

    step("Step 1/3 — Extracting & classifying pages…")
    try:
        pages_text    = _extract_pages_text(pdf_bytes)
        total_pages   = len(pages_text)
        page_map      = _classify_pages(pages_text)
        map_summary   = _format_page_map(page_map, total_pages)
        section_texts = _group_pages_by_section(pages_text, page_map)
        self._general_lookup_file_data[path]["page_map"] = map_summary
        _general_log(self, f"[{fname}] {total_pages} pages.\n{map_summary}")
    except Exception as e:
        _general_log(self, f"[{fname}] Classification failed: {e}")
        raise
    if cancelled(): return

    cibi_pages   = _pages_for_type(page_map, "cibi",      total_pages)
    cfa_pages    = _pages_for_type(page_map, "cfa",       total_pages)
    ws_pages     = _pages_for_type(page_map, "worksheet", total_pages)
    pdf_cibi     = _extract_page_subset(pdf_bytes, cibi_pages)
    cfa_ws_pages = sorted(set(cfa_pages) | set(ws_pages))
    pdf_cfa_ws   = _extract_page_subset(pdf_bytes, cfa_ws_pages)

    original_pdf_cibi   = _extract_page_subset(original_pdf_bytes, cibi_pages)
    original_pdf_cfa_ws = _extract_page_subset(original_pdf_bytes, cfa_ws_pages)

    _general_log(self,
         f"[{fname}] Slices — "
         f"cibi:{len(pdf_cibi):,}b  cfa_ws:{len(pdf_cfa_ws):,}b  "
         f"(full:{len(pdf_bytes):,}b)")

    pg_summary = (None if not (set(page_map.values()) - {"unknown"})
                  else map_summary)
    if cancelled(): return

    step("Steps 2-3/3 — Gemini calls running…")
    call_results      = {}
    call_results_lock = threading.Lock()

    call_defs = [
        ("cibi_combined", _gemini_extract_cibi_combined, self, client, gt,
         pdf_cibi, pg_summary,
         section_texts.get("cibi", ""), original_pdf_cibi),
        ("cfa_ws_combined", _gemini_extract_cfa_ws_combined, self, client, gt,
         pdf_cfa_ws, pg_summary,
         section_texts.get("cfa", ""), section_texts.get("worksheet", ""),
         original_pdf_cfa_ws),
    ]

    def _run(label, fn, *args):
        if cancelled():
            return
        try:
            result = fn(*args)
        except Exception as exc:
            import traceback
            _general_log(self, f"[{fname}] '{label}' FAILED: {exc}\n"
                               f"{traceback.format_exc()}")
            result = ("", {})
        with call_results_lock:
            call_results[label] = result

    if MAX_PARALLEL_CALLS == 1:
        for i, defn in enumerate(call_defs):
            if cancelled():
                break
            label = defn[0]; fn = defn[1]; args = defn[2:]
            step(f"Steps 2-3/3 — call {i+1}/2: {label}…")
            _run(label, fn, *args)
            if i < len(call_defs) - 1 and INTER_CALL_DELAY_S > 0:
                _general_log(self, f"[{fname}] Waiting {INTER_CALL_DELAY_S}s before next call…")
                self._general_lookup_cancel.wait(timeout=INTER_CALL_DELAY_S)
    else:
        with concurrent.futures.ThreadPoolExecutor(
                max_workers=MAX_PARALLEL_CALLS,
                thread_name_prefix=f"gem_{fname[:6]}") as pool:
            fs = [pool.submit(_run, defn[0], defn[1], *defn[2:])
                  for defn in call_defs]
            concurrent.futures.wait(fs)

    if cancelled(): return

    raw_cibi_c, data_cibi_c = call_results.get("cibi_combined",   ("", {}))
    raw_cfaws,  data_cfaws  = call_results.get("cfa_ws_combined", ("", {}))

    applicant_name = _sanitize_extracted_text(
        data_cibi_c.get("applicant_name", "").strip())
    gate_result = {
        "applicant_name":    applicant_name,
        "residence_address": _sanitize_extracted_text(
            data_cibi_c.get("residence_address", "")),
        "office_address":    _sanitize_extracted_text(
            data_cibi_c.get("office_address", "")),
    }
    _general_log(self,
         f"[{fname}] name={applicant_name or '[not found]'}  "
         f"cibi_keys={list(data_cibi_c)}  "
         f"cfaws_keys={list(data_cfaws)}")
    _general_log(self,
         f"[{fname}] Raw (first 3000 chars):\n"
         + "\n---\n".join(filter(None, [raw_cibi_c, raw_cfaws]))[:3000])

    combined = {}
    combined.update(data_cibi_c)
    combined.update(data_cfaws)
    results = _parse_extraction_response_from_dict(combined)

    fields_found = sum(1 for k, v in results.items()
                       if not k.startswith("_") and v.get("items"))
    _general_log(self, f"[{fname}] Fields with data: {fields_found}/{len(LOOKUP_ROWS)}")

    results["_applicant_name"] = applicant_name
    results["_gate_data"]      = gate_result
    results["_page_map"]       = map_summary
    results["_source_file"]    = fname
    results["_cfa_net_income"] = data_cfaws.get("cfa_net_income", "")

    self._general_lookup_file_data[path]["name"]      = applicant_name
    self._general_lookup_file_data[path]["gate_data"] = gate_result
    self._general_lookup_file_data[path]["results"]   = results

    display = applicant_name or fname
    _general_ui(self, lambda n=display:
        self._general_lookup_file_data[path]["widgets"]["name_lbl"].config(
            text=f"  {n}"))
    if cancelled(): return

    # ── Persist to general_summary SQLite ────────────────────────────
    step("Saving to General Summary…")
    try:
        session_id = getattr(self, "_general_lookup_session_id",
                             datetime.now().isoformat(timespec="seconds"))
        db_save_applicant(session_id, results)
        _general_log(self, f"[{fname}] Saved to General DB → session {session_id}")
        _general_ui(self, lambda: lookup_summary_notify(self))
    except Exception as exc:
        _general_log(self, f"[{fname}] DB save failed (non-fatal): {exc}")

    _general_ui(self, lambda:
        _set_general_row_status(self, path, "done", "Done  ·  Saved to General Summary"))
    _general_ui(self, lambda:
        self._general_lookup_file_data[path]["widgets"]["expand_btn"].configure(
            state="normal"))


# ═══════════════════════════════════════════════════════════════════════
#  TEXT EXTRACTION
# ═══════════════════════════════════════════════════════════════════════

def _extract_pages_text(pdf_bytes: bytes) -> list:
    pages = []
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pg in pdf.pages:
                t = pg.extract_text() or ""
                try:
                    for tbl in (pg.extract_tables() or []):
                        for row in tbl:
                            if row:
                                t += "\n" + "  |  ".join(c or "" for c in row)
                except Exception:
                    pass
                pages.append(t.strip())
        return pages
    except Exception:
        pass
    try:
        import fitz
        doc = fitz.open(stream=io.BytesIO(pdf_bytes), filetype="pdf")
        return [pg.get_text().strip() for pg in doc]
    except Exception:
        return []


# ═══════════════════════════════════════════════════════════════════════
#  PAGE CLASSIFIER
# ═══════════════════════════════════════════════════════════════════════

def _classify_pages(pages_text: list) -> dict:
    page_types = {}
    for i, text in enumerate(pages_text):
        pg      = i + 1
        t_lower = text.lower()
        if not t_lower.strip():
            page_types[pg] = "unknown"
            continue
        scores = {dt: sum(1 for kw in kws if kw in t_lower)
                  for dt, kws in DOC_TYPE_KEYWORDS.items()}
        best   = max(scores, key=scores.get)
        page_types[pg] = best if scores[best] > 0 else "unknown"

    for i in range(1, len(pages_text)):
        pg = i + 1
        if page_types[pg] == "unknown" and not pages_text[i].strip():
            page_types[pg] = page_types.get(pg - 1, "unknown")

    cibi_pgs = sorted(pg for pg, t in page_types.items() if t == "cibi")
    if len(cibi_pgs) == 1:
        nxt = cibi_pgs[0] + 1
        if page_types.get(nxt) in ("unknown", None):
            page_types[nxt] = "cibi"

    return page_types


def _format_page_map(page_map: dict, total_pages: int) -> str:
    labels = {
        "credit_scoring": "Credit Scoring",
        "cibi":           "CI/BI Report",
        "cfa":            "Cashflow Analysis",
        "worksheet":      "Worksheet",
        "unknown":        "Unclassified",
    }
    return "\n".join(
        f"  Page {pg:>2}: {labels.get(page_map.get(pg, 'unknown'), '?')}"
        for pg in range(1, total_pages + 1))


def _group_pages_by_section(pages_text: list, page_map: dict) -> dict:
    groups = {k: [] for k in
              ("credit_scoring", "cibi", "cfa", "worksheet", "unknown")}
    for i, text in enumerate(pages_text):
        pg    = i + 1
        dtype = page_map.get(pg, "unknown")
        groups.setdefault(dtype, []).append(f"[Page {pg}]\n{text}")
    return {k: "\n\n".join(v) for k, v in groups.items()}


# ═══════════════════════════════════════════════════════════════════════
#  GEMINI HELPERS  (shared logic, no prefix needed — internal only)
# ═══════════════════════════════════════════════════════════════════════

def _get_gemini_client(self):
    try:
        from google import genai
        from google.genai import types as gt
    except (ImportError, ModuleNotFoundError):
        raise RuntimeError(
            "google-genai not installed.  Run: pip install google-genai")
    api_key = None
    for attr in ("_gemini_api_key", "gemini_api_key", "_api_key"):
        if hasattr(self, attr):
            api_key = getattr(self, attr); break
    if not api_key:
        try:
            from app_constants import GEMINI_API_KEY
            api_key = GEMINI_API_KEY
        except Exception:
            pass
    if not api_key:
        try:
            import config; api_key = config.GEMINI_API_KEY
        except Exception:
            pass
    if not api_key:
        import os
        api_key = (os.environ.get("GEMINI_API_KEY") or
                   os.environ.get("GOOGLE_API_KEY"))
    if not api_key:
        raise RuntimeError(
            "Gemini API key not found.\n"
            "Add GEMINI_API_KEY to app_constants.py or set the "
            "GEMINI_API_KEY environment variable.")
    return genai.Client(api_key=api_key), gt


def _gemini_call(self, client, gt, contents, config,
                 max_retries=None, cancel_event=None):
    import random
    global _GEMINI_LAST_CALL

    if max_retries is None:
        max_retries = GEMINI_MAX_RETRIES

    for attempt in range(max_retries):
        if cancel_event and cancel_event.is_set():
            raise RuntimeError("Cancelled")

        with _GEMINI_CALL_LOCK:
            now     = time.time()
            elapsed = now - _GEMINI_LAST_CALL
            if elapsed < _GEMINI_MIN_GAP_S:
                wait = _GEMINI_MIN_GAP_S - elapsed
                time.sleep(wait)
            _GEMINI_LAST_CALL = time.time()

        try:
            resp = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=contents, config=config)
            if not resp.text or not resp.text.strip():
                raise ValueError("Empty response from Gemini")
            return resp
        except Exception as exc:
            msg = str(exc)
            is_transient = any(x in msg for x in (
                "500", "502", "503", "unavailable", "429", "quota",
                "RESOURCE_EXHAUSTED", "Empty response",
                "timeout", "timed out", "Connection", "RemoteDisconnected"))
            if is_transient and attempt < max_retries - 1:
                base  = GEMINI_RETRY_DELAYS[min(attempt, len(GEMINI_RETRY_DELAYS)-1)]
                delay = base + random.uniform(-base * 0.2, base * 0.2)
                _code = "ERR"
                for _candidate in ("429", "502", "503", "500"):
                    if _candidate in msg:
                        _code = _candidate
                        break
                else:
                    if "quota" in msg.lower() or "RESOURCE_EXHAUSTED" in msg:
                        _code = "QUOTA"
                    elif "timeout" in msg.lower() or "timed out" in msg.lower():
                        _code = "TIMEOUT"
                    elif "unavailable" in msg.lower():
                        _code = "UNAVAILABLE"
                _general_ui(self, lambda w=int(delay), a=attempt, r=max_retries, c=_code:
                    self._general_lookup_overall_lbl.config(
                        text=f"Gemini {c} — retry {a+1}/{r} in {w}s…",
                        fg=ACCENT_GOLD))
                if cancel_event:
                    cancel_event.wait(timeout=delay)
                else:
                    time.sleep(delay)
            else:
                raise


def _gemini_call_with_fallback(self, client, gt,
                                pdf_bytes, original_pdf_bytes,
                                prompt, config,
                                cancel_event=None):
    for attempt_bytes, label in [
            (pdf_bytes,          "processed"),
            (original_pdf_bytes, "original (fallback)"),
    ]:
        try:
            _general_log(self, f"  Sending {label} PDF bytes "
                               f"({len(attempt_bytes):,}b) to Gemini…")
            return _gemini_call(
                self, client, gt,
                [_pdf_part(gt, attempt_bytes), prompt],
                config,
                cancel_event=self._general_lookup_cancel)
        except Exception as exc:
            msg = str(exc)
            is_invalid = "400" in msg or "INVALID_ARGUMENT" in msg
            if is_invalid and label == "processed":
                _general_log(self, f"  400 INVALID_ARGUMENT on {label} PDF — "
                                   f"retrying with original bytes…")
                continue
            raise


def _pdf_part(gt, pdf_bytes: bytes):
    return gt.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")


def _parse_json_safe(text: str) -> dict:
    if not text or not text.strip():
        return {}
    cleaned = re.sub(r"```(?:json)?", "", text).strip().strip("`").strip()
    if "{" not in cleaned:
        return {}
    m = re.search(r"\{[\s\S]*\}", cleaned)
    if not m:
        return {}
    candidate = m.group(0)
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        pass
    try:
        partial = {}
        for km in re.finditer(
                r'"(\w+)"\s*:\s*(\[[\s\S]*?\])\s*(?=[,}]|$)', candidate):
            try:
                partial[km.group(1)] = json.loads(km.group(2))
            except json.JSONDecodeError:
                pass
        if partial:
            return partial
    except Exception:
        pass
    return {}


def _gemini_extract_cibi_combined(self, client, gt,
                                  pdf_bytes: bytes,
                                  pg_summary, hint: str,
                                  original_pdf_bytes: bytes = None) -> tuple:
    if pg_summary:
        scope = (f"The page classifier identified these sections:\n{pg_summary}\n\n"
                 f"Focus on the CI/BI Report pages identified above.")
    else:
        scope = ("The document is a fully scanned PDF. Search ALL pages for "
                 "CI/BI Report content — typically the first 1-2 pages.")

    hint_block = (f"\n\nPartial text from CI/BI pages:\n{hint[:3000]}"
                  if hint else "")

    prompt = f"""You are a credit analyst reviewing a Philippine rural bank loan application.
This is a SCANNED document — printed labels with handwritten values.
Read it visually as an image.

{scope}

Extract ALL of the following fields from the CI/BI Report section.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
IDENTITY FIELDS (plain strings, not arrays)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

applicant_name
  Full legal name of the loan applicant.
  Look for: "Name of Applicant", "Borrower", "Client", "Name".

residence_address
  Applicant's home address.
  Look for: "Residence Address", "Home Address", "Permanent Address".
  IMPORTANT — Philippine address notation:
    "P1", "P2" etc. means "Purok 1", "Purok 2" — write it in full.
    Never output Unicode box characters or Roman numerals for digits.

office_address
  Read ONLY from a field explicitly labelled "Office Address".
  If blank, N/A, or absent — return empty string.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SPOUSE FIELDS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

cibi_spouse
  Combine the spouse's name with their employment status into ONE entry.
  Look for fields labelled: "Name of Spouse", "Spouse", "Husband", "Wife".
  Look for employment status labelled: "Employed", "Self-Employed",
  "Occupation", "Nature of Work", "Position".

  Rules:
  - If BOTH name and employment status are present, combine as one item:
      description = "<Spouse Name> — <Employment Status/Occupation>"
      amount      = "N/A"
  - If ONLY the name is present (employment is blank/N/A), use just the name.
  - If ONLY employment is present (name is blank/N/A), use just the employment.
  - If BOTH are blank, N/A, or absent → return [].
  - Do NOT create separate items for name and employment — always ONE combined item.

cibi_spouse_office
  The spouse's office or employer address.
  Look for: "Office Address" near the spouse section, "Employer Address",
  "Place of Work" associated with the spouse (not the applicant).
  Return [] if blank, N/A, or absent.
  Each item: description = the address or employer name + address,
             amount = "N/A".

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ASSETS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

cibi_personal_assets
  FORMAT A: From the "PERSONAL ASSETS" section only.
  FORMAT B: From the combined "PERSONAL AND BUSINESS ASSETS" section.
  For each item:
    description = item name / description (include make, model, year if shown)
    amount      = estimated value or acquisition cost if written, else ""
  Return [] if the section is entirely blank.

cibi_business_assets
  FORMAT A: From the "BUSINESS ASSETS" section only.
  FORMAT B: If a combined section was detected and already extracted above,
    return [] here to avoid duplication.
  For each item:
    description = item name / description (include make, model, year if shown)
    amount      = estimated value or acquisition cost if written, else ""
  Return [] if the section is entirely blank or already captured above.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
RISK / ADDRESS FIELDS (arrays of objects)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

cibi_place_of_work
  Employer name and address, or own business name/address if self-employed.
  Set amount = "N/A".

cibi_temp_residence
  Applicant's residence address as written on the CI/BI form.
  Write "Purok 1", "Purok 2" etc. — never "P1", "P2".
  Set amount = "N/A". Return [] only if completely blank.

cibi_petrol_products
  Flag any employer, own business, or trade/bank reference involving:
  petroleum, oil depot, gasoline station, LPG, fuel, lubricants,
  plastic packaging, PVC, polypropylene, rubber, chemicals,
  fertilizer manufacturing.
  Return [] if none found.

cibi_transport_services
  Flag any employer, own business, or trade/bank reference involving:
  bus, jeepney, tricycle, forwarding, trucking, hauling, heavy equipment,
  crane, bulldozer, backhoe, freight, logistics, cargo, courier, shipping.
  Return [] if none found.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CREDIT HISTORY & REFERENCES TABLE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

credit_history_amort
  Find the table labelled "CREDIT HISTORY & REFERENCES" (or similar).
  Columns: Bank/Lending Institution | Principal Loan | Due Date | Amort. | Balance
  Extract EVERY filled-in data row (skip blank rows and the TOTAL row).
  For each row:
    description = Bank/Lending Institution name
    amount      = the value in the "Amort." column exactly as written
    date        = the Due Date value, or "" if blank
  If the table is completely blank or absent, return [].

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 OUTPUT — return ONLY this JSON, nothing else:
{{
  "applicant_name": "",
  "residence_address": "",
  "office_address": "",
  "cibi_spouse":             [{{"description":"...","amount":"N/A","date":""}}],
  "cibi_spouse_office":      [{{"description":"...","amount":"N/A","date":""}}],
  "cibi_personal_assets":    [{{"description":"...","amount":"...","date":""}}],
  "cibi_business_assets":    [{{"description":"...","amount":"...","date":""}}],
  "cibi_place_of_work":      [{{"description":"...","amount":"N/A","date":""}}],
  "cibi_temp_residence":     [{{"description":"...","amount":"N/A","date":""}}],
  "cibi_petrol_products":    [{{"description":"...","amount":"...","date":""}}],
  "cibi_transport_services": [{{"description":"...","amount":"...","date":""}}],
  "credit_history_amort":    [{{"description":"...","amount":"...","date":""}}]
}}

RULES:
- Return data even at 80% confidence — only return [] if genuinely absent.
- Never fabricate data not visible in the document.
- Use empty string "" for identity fields not found or marked N/A.
- For credit_history_amort: extract ONLY individual rows, never the TOTAL row.
- For cibi_spouse: always ONE combined item — never split name and employment.
- For assets: detect Format A vs Format B and avoid duplication.{hint_block}"""

    fallback = original_pdf_bytes if original_pdf_bytes is not None else pdf_bytes
    resp = _gemini_call_with_fallback(
        self, client, gt,
        pdf_bytes, fallback,
        prompt,
        gt.GenerateContentConfig(temperature=0.0),
        cancel_event=self._general_lookup_cancel)
    raw  = resp.text or ""
    data = _parse_json_safe(raw)
    _NA = {"N/A", "NA", "NONE", "NONE.", "-", "\u2014", "N/A.", "N.A."}
    for field in ("applicant_name", "residence_address", "office_address"):
        val = data.get(field, "").strip()
        data[field] = "" if val.upper() in _NA else val
    return raw, data


def _gemini_extract_cfa_ws_combined(self, client, gt,
                                    pdf_bytes: bytes,
                                    pg_summary,
                                    cfa_hint: str,
                                    ws_hint: str,
                                    original_pdf_bytes: bytes = None) -> tuple:
    if pg_summary:
        scope = (f"The page classifier identified these sections:\n{pg_summary}\n\n"
                 f"Focus on the Cashflow Analysis and Worksheet pages above.")
    else:
        scope = ("The document is a fully scanned PDF. Search ALL pages for "
                 "Cashflow Analysis and Worksheet content.")

    hint_block = ""
    if cfa_hint:
        hint_block += f"\n\nPartial text from CFA pages:\n{cfa_hint[:2000]}"
    if ws_hint:
        hint_block += f"\n\nPartial text from Worksheet pages:\n{ws_hint[:2000]}"

    prompt = f"""You are a credit analyst reviewing a Philippine rural bank loan application.
This is a SCANNED document — printed labels with handwritten values.
Read it visually as an image.

{scope}

Extract ALL of the following fields. The document has two sections:
the Cashflow Analysis (CFA) and the Business Expense Worksheet.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FROM THE CASHFLOW ANALYSIS (CFA):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

income_remittance
  Every filled-in row in the SOURCE OF INCOME section.
  amount = MONTHLY TOTALS column value.
  Return [] only if the entire section is blank.

cfa_business_expenses
  Every row in the BUSINESS EXPENSES section.
  Do NOT include food, electricity, or household rows.
  Use amount="" if the amount cell is empty.

cfa_household_expenses
  Every row in the HOUSEHOLD / PERSONAL / FAMILY EXPENSES section.
  Use amount="" if the amount cell is empty.

cfa_net_income
  The TOTAL NET INCOME (or NET CASH FLOW / NET SURPLUS) value at the bottom.
  Return the value exactly as written, including the peso sign if present.
  Return "" if the field is blank or absent.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FROM THE BUSINESS EXPENSE WORKSHEET:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ws_food_grocery      — Row labelled "Food / Grocery" or similar.
ws_fuel_transport    — Personal transport row in HOUSEHOLD section.
ws_electricity       — Electricity / utility bill row.
ws_fertilizer        — Fertilizer / Farm Inputs row.
ws_forwarding        — Forwarding / Trucking / Hauling row in BUSINESS section.
ws_fuel_diesel       — Fuel / Gas / Diesel row in BUSINESS section.
ws_equipment         — Equipment rental row.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 OUTPUT — return ONLY this JSON, nothing else:
{{
  "income_remittance":      [{{"description":"...","amount":"...","date":""}}],
  "cfa_business_expenses":  [{{"description":"...","amount":"...","date":""}}],
  "cfa_household_expenses": [{{"description":"...","amount":"...","date":""}}],
  "cfa_net_income":         "",
  "ws_food_grocery":        [{{"description":"...","amount":"...","date":""}}],
  "ws_fuel_transport":      [{{"description":"...","amount":"...","date":""}}],
  "ws_electricity":         [{{"description":"...","amount":"...","date":""}}],
  "ws_fertilizer":          [{{"description":"...","amount":"...","date":""}}],
  "ws_forwarding":          [{{"description":"...","amount":"...","date":""}}],
  "ws_fuel_diesel":         [{{"description":"...","amount":"...","date":""}}],
  "ws_equipment":           [{{"description":"...","amount":"...","date":""}}]
}}

RULES:
- Extract ALL rows with any content — do not skip rows.
- Write amounts exactly as written.
- Return [] ONLY if the entire section is absent from the document.
- cfa_net_income is a plain string, not an array.
- Never fabricate data not visible in the document.{hint_block}"""

    fallback = original_pdf_bytes if original_pdf_bytes is not None else pdf_bytes
    resp = _gemini_call_with_fallback(
        self, client, gt,
        pdf_bytes, fallback,
        prompt,
        gt.GenerateContentConfig(temperature=0.0),
        cancel_event=self._general_lookup_cancel)
    raw  = resp.text or ""
    return raw, _parse_json_safe(raw)


# ═══════════════════════════════════════════════════════════════════════
#  PARSE EXTRACTION RESPONSE
# ═══════════════════════════════════════════════════════════════════════

def _sanitize_extracted_text(text: str) -> str:
    if not text:
        return text
    import unicodedata

    _DIGIT_SUBS = {
        "Ⅰ": "1", "Ⅱ": "2", "Ⅲ": "3", "Ⅳ": "4", "Ⅴ": "5",
        "Ⅵ": "6", "Ⅶ": "7", "Ⅷ": "8", "Ⅸ": "9", "Ⅹ": "10",
        "│": "1",
    }
    for bad, good in _DIGIT_SUBS.items():
        text = text.replace(bad, good)

    text = re.sub(r"P[■-◿▀-▟]", "Purok ", text)
    text = re.sub(r"[■-◿▀-▟─-╿]", "", text)
    text = re.sub(r"P(\d{1,2})",
                  lambda m: f"Purok {m.group(1)}", text)

    KEEP_WS = {" ", chr(10), chr(9)}
    text = "".join(
        ch for ch in text
        if unicodedata.category(ch)[0] != "C" or ch in KEEP_WS)

    return text.strip()


def _parse_extraction_response_from_dict(data: dict) -> dict:
    all_keys = [r[0] for r in LOOKUP_ROWS]
    results  = {k: {"total": None, "items": []} for k in all_keys}
    if not data:
        return results

    raw_entries = {}
    for key in all_keys:
        entries = data.get(key, [])
        raw_entries[key] = entries if isinstance(entries, list) else []

    cfa_biz_fingerprints = set()
    for entry in raw_entries.get("cfa_business_expenses", []):
        if isinstance(entry, dict):
            cfa_biz_fingerprints.add((
                entry.get("description", "").strip().lower(),
                entry.get("amount", "").strip()))

    WS_DEDUP_KEYS = {"ws_fuel_diesel", "ws_forwarding", "ws_fertilizer"}

    NO_TOTAL_KEYS = {
        "cibi_place_of_work", "cibi_temp_residence",
        "cibi_spouse", "cibi_spouse_office",
        "cibi_personal_assets", "cibi_business_assets",
    }

    for key in all_keys:
        entries    = raw_entries[key]
        items_text = []
        total_sum  = 0.0
        has_total  = False

        for entry in entries:
            if not isinstance(entry, dict):
                continue
            desc = entry.get("description", "").strip()
            amt  = entry.get("amount",      "").strip()
            date = entry.get("date",        "").strip()

            if key in WS_DEDUP_KEYS:
                if (desc.lower(), amt) in cfa_biz_fingerprints:
                    continue

            freq = ""
            if date:
                _d = date.upper().replace("TOTALS", "").strip()
                _freq_map = {
                    "MONTHLY":      "Monthly",
                    "WEEKLY":       "Weekly",
                    "DAILY":        "Daily",
                    "SEMI-MONTHLY": "Semi-Monthly",
                    "SEMI MONTHLY": "Semi-Monthly",
                    "ANNUAL":       "Annual",
                    "YEARLY":       "Yearly",
                    "QUARTERLY":    "Quarterly",
                }
                freq = _freq_map.get(_d, date.title())

            amt_part  = f"[{amt}]" if amt and amt.upper() not in ("N/A", "P0.00", "0", "") else ""
            freq_part = f"({freq})" if freq else ""
            parts = [p for p in [desc, amt_part, freq_part] if p]
            label = "  ".join(parts).strip()
            if label:
                items_text.append(label)

            if key not in NO_TOTAL_KEYS and amt and amt.upper() != "N/A":
                nums = re.findall(r"[\d,]+\.?\d*",
                                  re.sub(r"[^\d.,]", " ", amt))
                if nums:
                    try:
                        total_sum += float(nums[0].replace(",", ""))
                        has_total  = True
                    except ValueError:
                        pass

        results[key]["items"] = items_text
        if has_total and key not in NO_TOTAL_KEYS:
            results[key]["total"] = f"P{total_sum:,.2f}"

    return results


# ═══════════════════════════════════════════════════════════════════════
#  MISC HELPERS
# ═══════════════════════════════════════════════════════════════════════

def _general_log(self, msg: str):
    with self._general_lookup_raw_lock:
        self._general_lookup_raw_log.append(msg)
        snapshot = "\n".join(self._general_lookup_raw_log)
    _general_ui(self, lambda s=snapshot: _general_set_raw(self, s))


def _general_set_raw(self, text: str):
    box = self._general_lookup_raw_box
    box.config(state="normal")
    box.delete("1.0", "end")
    if text:
        box.insert("end", text)
    box.config(state="disabled")


def _general_toggle_raw(self):
    if self._general_lookup_raw_visible:
        self._general_lookup_raw_frame.pack_forget()
        self._general_lookup_raw_toggle_btn.configure(text="▼ Show")
        self._general_lookup_raw_visible = False
    else:
        self._general_lookup_raw_frame.pack(fill="x", pady=(8, 0))
        self._general_lookup_raw_toggle_btn.configure(text="▲ Hide")
        self._general_lookup_raw_visible = True


def _general_ui(self, fn):
    self.after(0, fn)


# ═══════════════════════════════════════════════════════════════════════
#  ATTACH
# ═══════════════════════════════════════════════════════════════════════

def attach(cls):
    cls._build_general_lookup_panel = _build_general_lookup_panel
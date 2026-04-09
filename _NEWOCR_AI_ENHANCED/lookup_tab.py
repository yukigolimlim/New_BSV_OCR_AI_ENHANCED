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

from summary_tab import db_save_applicant, lookup_summary_notify

# ── Category master list ──────────────────────────────────────────────
LOOKUP_ROWS = [
    ("cibi_place_of_work",      "CI/BI Report",      "Office Address",                      "cibi"),
    ("cibi_temp_residence",     "CI/BI Report",      "Residence Address",                    "cibi"),
    ("cibi_spouse",             "CI/BI Report",      "Spouse / Employment",                  "cibi"),
    ("cibi_spouse_office",      "CI/BI Report",      "Spouse Office Address",                "cibi"),
    ("cibi_personal_assets",    "CI/BI Report",      "Personal Assets",                      "cibi"),
    ("cibi_business_assets",    "CI/BI Report",      "Business Assets",                      "cibi"),
    ("cibi_business_inventory", "CI/BI Report",      "Business Inventory",                   "cibi"),
    ("cibi_petrol_products",    "CI/BI Report",      "Petrol / Plastics / PVC Risk",         "cibi"),
    ("cibi_transport_services", "CI/BI Report",      "Transport Services Risk",              "cibi"),
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
        "balance sheet", "business inventory", "inventory",
        "stocks on hand", "merchandise",
    ],
    "cfa": [
        "cash flow", "cashflow", "cash-flow",
        "source of income", "business income", "household expenses",
        "monthly expenses", "personal expenses", "net income",
        "remittance", "padala",
        "cashflow analysis", "cash flow analysis", "income analysis",
        "net surplus", "net cash flow", "total income", "total expenses",
        "gross income", "business expenses", "household / personal",
        "family expenses", "monthly totals", "semi-monthly",
        "income source", "salary", "farming income", "farm income",
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

# ── Concurrency settings ──────────────────────────────────────────────
MAX_PARALLEL_CALLS = 2

# Retry settings — shortened delays for non-quota transient errors
GEMINI_MAX_RETRIES  = 3
GEMINI_RETRY_DELAYS = [15, 30, 60]   # FIX: was [60, 90, 120] — far too long

# ── Global Gemini rate-limit gate ─────────────────────────────────────
_GEMINI_CALL_LOCK = threading.Lock()
_GEMINI_LAST_CALL = 0.0
_GEMINI_MIN_GAP_S = 0.3   # FIX: was 0.5 — shaved a bit for parallel calls

# ── Ordinal helper (used for multi-CFA labelling) ─────────────────────
_ORDINALS = ["First", "Second", "Third", "Fourth", "Fifth",
             "Sixth", "Seventh", "Eighth", "Ninth", "Tenth"]

def _ordinal(n: int) -> str:
    """Return English ordinal word for 1-based index n."""
    return _ORDINALS[n - 1] if 1 <= n <= len(_ORDINALS) else f"#{n}"


# ═══════════════════════════════════════════════════════════════════════
#  PDF SPLITTING HELPER
# ═══════════════════════════════════════════════════════════════════════

def _auto_rotate_pdf(pdf_bytes: bytes) -> bytes:
    # FIX: Skip the slow pytesseract OSD pass — it runs on every page and
    # adds significant latency.  Only correct fitz-detected rotation flags,
    # which is near-instant.
    try:
        import fitz
        doc = fitz.open(stream=io.BytesIO(pdf_bytes), filetype="pdf")
        modified = False
        for page in doc:
            if page.rotation != 0:
                page.set_rotation(0)
                modified = True
        if not modified:
            return pdf_bytes
        buf = io.BytesIO()
        doc.save(buf, garbage=2, deflate=True, clean=False)
        return buf.getvalue()
    except Exception:
        return pdf_bytes


def _extract_page_subset(pdf_bytes: bytes, page_numbers: list) -> bytes:
    if not page_numbers:
        return pdf_bytes
    try:
        import fitz
        src = fitz.open(stream=io.BytesIO(pdf_bytes), filetype="pdf")
        if sorted(page_numbers) == list(range(1, src.page_count + 1)):
            return pdf_bytes
        dst = fitz.open()
        for pg in page_numbers:
            zero = pg - 1
            if 0 <= zero < src.page_count:
                dst.insert_pdf(src, from_page=zero, to_page=zero)
        if dst.page_count == 0:
            return pdf_bytes
        return dst.tobytes(garbage=0, deflate=True, clean=False)
    except Exception:
        return pdf_bytes


def _pages_for_type(page_map: dict, doc_type: str,
                    total_pages: int, padding: int = 1) -> list:
    matched = sorted(pg for pg, t in page_map.items() if t == doc_type)
    if not matched:
        return []
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

def _build_lookup_panel(self, parent):
    outer = tk.Frame(parent, bg=CARD_WHITE)
    self._lookup_frame = outer

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

    self._lookup_inner  = inner
    self._lookup_canvas = canvas
    _populate_lookup(self, inner)


def _populate_lookup(self, parent):
    PAD = PAD_VALUE

    # ── Header ─────────────────────────────────────────────────────────
    hdr = tk.Frame(parent, bg=CARD_WHITE)
    hdr.pack(fill="x", padx=PAD, pady=(28, 0))
    left = tk.Frame(hdr, bg=CARD_WHITE)
    left.pack(side="left")
    tk.Label(left, text="Look-Up", font=F(20, "bold"),
             fg=NAVY_DEEP, bg=CARD_WHITE).pack(anchor="w")
    tk.Label(left,
             text="One PDF per applicant  ·  processed one at a time for accuracy"
                  "  ·  results saved to Summary tab",
             font=F(9), fg=TXT_SOFT, bg=CARD_WHITE).pack(
                 anchor="w", pady=(2, 0))
    badge = tk.Frame(hdr, bg="#EEF6FF",
                     highlightbackground="#4F8EF7", highlightthickness=1)
    badge.pack(side="right", pady=4)
    tk.Label(badge, text="  Gemini 2.5 Flash · Sequential  ",
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
    self._lookup_icon_lbl = tk.Label(drop_inner, text="📂",
                                     font=("Segoe UI Emoji", 18),
                                     fg="#4F8EF7", bg="#F7FAFF")
    self._lookup_icon_lbl.pack(side="left", padx=(0, 8))
    self._lookup_file_lbl = tk.Label(
        drop_inner, text="No files selected",
        font=F(9), fg=TXT_SOFT, bg="#F7FAFF")
    self._lookup_file_lbl.pack(side="left")

    btn_row = tk.Frame(upload_sec, bg=CARD_WHITE)
    btn_row.pack(fill="x", pady=(10, 0))

    self._lookup_browse_btn = ctk.CTkButton(
        btn_row, text="Choose PDF(s)",
        command=lambda: _lookup_browse(self),
        width=140, height=36, corner_radius=8,
        fg_color="#4F8EF7", hover_color="#3A7EE8",
        text_color=WHITE, font=FF(9, "bold"), border_width=0)
    self._lookup_browse_btn.pack(side="left")

    self._lookup_run_btn = ctk.CTkButton(
        btn_row, text="⚡  Run Look-Up",
        command=lambda: _lookup_run(self),
        width=150, height=36, corner_radius=8,
        fg_color=NAVY_LIGHT, hover_color=NAVY_PALE,
        text_color=WHITE, font=FF(9, "bold"),
        state="disabled", border_width=0)
    self._lookup_run_btn.pack(side="left", padx=(10, 0))

    self._lookup_clear_btn = ctk.CTkButton(
        btn_row, text="✕  Clear All",
        command=lambda: _lookup_clear(self),
        width=90, height=36, corner_radius=8,
        fg_color=CARD_WHITE, hover_color="#FEF2F2",
        text_color=ACCENT_RED, font=FF(9),
        border_width=1, border_color="#FCA5A5", state="disabled")
    self._lookup_clear_btn.pack(side="left", padx=(8, 0))

    self._lookup_overall_lbl = tk.Label(
        upload_sec, text="", font=F(8, "bold"),
        fg=TXT_SOFT, bg=CARD_WHITE)
    self._lookup_overall_lbl.pack(anchor="w", pady=(6, 0))

    # ── Overall progress bar ───────────────────────────────────────────
    self._lookup_prog_var = tk.DoubleVar(value=0.0)
    self._lookup_prog_bar = ctk.CTkProgressBar(
        parent, variable=self._lookup_prog_var,
        height=6, corner_radius=3,
        fg_color="#E8F0FA", progress_color="#4F8EF7", border_width=0)

    # ── Applicant queue ────────────────────────────────────────────────
    tk.Frame(parent, bg=BORDER_LIGHT, height=1).pack(
        fill="x", padx=PAD, pady=(18, 0))

    queue_hdr = tk.Frame(parent, bg=CARD_WHITE)
    queue_hdr.pack(fill="x", padx=PAD, pady=(12, 0))
    tk.Label(queue_hdr, text="APPLICANT QUEUE",
             font=F(7, "bold"), fg=TXT_MUTED, bg=CARD_WHITE).pack(side="left")
    self._lookup_queue_count_lbl = tk.Label(
        queue_hdr, text="", font=F(7), fg=TXT_MUTED, bg=CARD_WHITE)
    self._lookup_queue_count_lbl.pack(side="right")

    self._lookup_queue_frame = tk.Frame(parent, bg=CARD_WHITE)
    self._lookup_queue_frame.pack(fill="x", padx=PAD, pady=(6, 0))

    # ── Raw log (collapsible) ──────────────────────────────────────────
    raw_wrap = tk.Frame(parent, bg=CARD_WHITE)
    raw_wrap.pack(fill="x", padx=PAD, pady=(20, 0))
    raw_toggle_row = tk.Frame(raw_wrap, bg=CARD_WHITE)
    raw_toggle_row.pack(fill="x")
    tk.Label(raw_toggle_row, text="RAW LOG",
             font=F(7, "bold"), fg=TXT_MUTED, bg=CARD_WHITE).pack(side="left")
    self._lookup_raw_toggle_btn = ctk.CTkButton(
        raw_toggle_row, text="▼ Show",
        command=lambda: _toggle_raw(self),
        width=72, height=24, corner_radius=6,
        fg_color=OFF_WHITE, hover_color=NAVY_MIST,
        text_color=TXT_SOFT, font=FF(8),
        border_width=1, border_color=BORDER_LIGHT)
    self._lookup_raw_toggle_btn.pack(side="right")

    self._lookup_raw_frame = tk.Frame(raw_wrap, bg=CARD_WHITE)
    raw_sb = tk.Scrollbar(self._lookup_raw_frame, relief="flat",
                          troughcolor=OFF_WHITE, bg=BORDER_LIGHT,
                          width=6, bd=0)
    raw_sb.pack(side="right", fill="y")
    self._lookup_raw_box = tk.Text(
        self._lookup_raw_frame, wrap="word", font=FMONO(8),
        fg=TXT_SOFT, bg="#F7FAFF", relief="flat", bd=0,
        padx=12, pady=8, height=14, state="disabled",
        yscrollcommand=raw_sb.set,
        selectbackground=NAVY_GHOST, selectforeground=TXT_NAVY)
    self._lookup_raw_box.pack(side="left", fill="both", expand=True)
    raw_sb.config(command=self._lookup_raw_box.yview)
    self._lookup_raw_visible = False

    tk.Frame(parent, bg=CARD_WHITE, height=32).pack()

    # ── Internal state ─────────────────────────────────────────────────
    self._lookup_filepaths   = []
    self._lookup_cancel      = threading.Event()
    self._lookup_file_data   = {}
    self._lookup_raw_log     = []
    self._lookup_raw_lock    = threading.Lock()
    self._lookup_done_count  = 0
    self._lookup_error_count = 0


# ═══════════════════════════════════════════════════════════════════════
#  QUEUE ROW BUILDER
# ═══════════════════════════════════════════════════════════════════════

def _build_queue_row(self, path: str, index: int):
    parent = self._lookup_queue_frame
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
        command=lambda p=path: _toggle_applicant_detail(self, p),
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

    detail_frame = tk.Frame(row_outer, bg=CARD_WHITE,
                            highlightbackground=BORDER_LIGHT,
                            highlightthickness=1)

    self._lookup_file_data[path]["widgets"] = {
        "row_outer":    row_outer,
        "top":          top,
        "status_lbl":   status_lbl,
        "step_lbl":     step_lbl,
        "name_lbl":     name_lbl,
        "expand_btn":   expand_btn,
        "detail_frame": detail_frame,
        "expanded":     False,
        "bg":           bg,
    }


def _set_row_status(self, path: str, status: str, step_text: str = ""):
    d = self._lookup_file_data.get(path, {})
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


def _toggle_applicant_detail(self, path: str):
    d = self._lookup_file_data.get(path, {})
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
            _build_detail_panel(self, path, detail)
        detail.pack(fill="x", padx=10, pady=(0, 8))
        w["expand_btn"].configure(text="▲ Hide")
        w["expanded"] = True


def _build_detail_panel(self, path: str, parent: tk.Frame):
    d       = self._lookup_file_data.get(path, {})
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
                        "cibi_personal_assets", "cibi_business_assets",
                        "cibi_business_inventory")

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

def _lookup_browse(self):
    paths = filedialog.askopenfilenames(
        title="Select PDF file(s) — one per applicant",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
    if not paths:
        return

    existing  = set(self._lookup_filepaths)
    new_paths = [p for p in paths if p not in existing]
    self._lookup_filepaths.extend(new_paths)

    for p in new_paths:
        self._lookup_file_data[p] = {
            "status":    "waiting",
            "name":      "",
            "results":   {},
            "page_map":  "",
            "error":     "",
            "gate_data": {},
            "widgets":   {},
        }
        _build_queue_row(self, p, self._lookup_filepaths.index(p))

    n = len(self._lookup_filepaths)
    self._lookup_icon_lbl.config(text="📄")
    self._lookup_file_lbl.config(
        text=f"{n} applicant PDF{'s' if n > 1 else ''} loaded",
        fg=NAVY_DEEP)
    self._lookup_queue_count_lbl.config(text=f"{n} file(s)")
    self._lookup_run_btn.configure(state="normal")
    self._lookup_clear_btn.configure(state="normal")
    self._lookup_cancel.clear()


def _lookup_clear(self):
    self._lookup_cancel.set()
    self._lookup_filepaths   = []
    self._lookup_file_data   = {}
    self._lookup_raw_log     = []
    self._lookup_done_count  = 0
    self._lookup_error_count = 0
    for w in self._lookup_queue_frame.winfo_children():
        w.destroy()
    self._lookup_file_lbl.config(text="No files selected", fg=TXT_SOFT)
    self._lookup_icon_lbl.config(text="📂")
    self._lookup_queue_count_lbl.config(text="")
    self._lookup_run_btn.configure(state="disabled")
    self._lookup_clear_btn.configure(state="disabled")
    self._lookup_overall_lbl.config(text="", fg=TXT_SOFT)
    _set_raw(self, "")
    try:
        self._lookup_prog_bar.pack_forget()
    except Exception:
        pass


def _lookup_run(self):
    if not self._lookup_filepaths:
        return
    self._lookup_cancel.clear()
    self._lookup_raw_log     = []
    self._lookup_done_count  = 0
    self._lookup_error_count = 0

    self._lookup_session_id = datetime.now().isoformat(timespec="seconds")

    self._lookup_run_btn.configure(state="disabled", text="Running…")
    self._lookup_prog_var.set(0.0)
    self._lookup_prog_bar.pack(fill="x", padx=PAD_VALUE, pady=(8, 0))
    self._lookup_overall_lbl.config(text="Processing…", fg=ACCENT_GOLD)

    try:
        client, gt = _get_gemini_client(self)
    except Exception as exc:
        self._lookup_overall_lbl.config(
            text=f"Gemini init failed: {exc}", fg=ACCENT_RED)
        self._lookup_run_btn.configure(state="normal", text="⚡  Run Look-Up")
        return

    threading.Thread(
        target=_lookup_worker,
        args=(self, client, gt),
        daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  MAIN WORKER — sequential file processing
# ═══════════════════════════════════════════════════════════════════════

def _lookup_worker(self, client, gt):
    paths = list(self._lookup_filepaths)
    total = len(paths)

    for i, path in enumerate(paths):
        if self._lookup_cancel.is_set():
            break

        fname = Path(path).name
        _ui(self, lambda n=fname, idx=i, t=total:
            self._lookup_overall_lbl.config(
                text=f"Processing {idx+1}/{t}: {n}…",
                fg=ACCENT_GOLD))

        try:
            _process_single_file(self, path, client, gt)
            self._lookup_done_count += 1
            self._lookup_file_data[path]["status"] = "done"
        except Exception as exc:
            import traceback
            self._lookup_error_count += 1
            err_msg = str(exc)
            _log(self, f"\n[ERROR — {fname}]\n{traceback.format_exc()}")
            _ui(self, lambda p=path, e=err_msg:
                _set_row_status(self, p, "error", f"Error: {e[:120]}"))
            self._lookup_file_data[path]["error"] = err_msg
            self._lookup_file_data[path]["status"] = "error"

        done_so_far = self._lookup_done_count + self._lookup_error_count
        _ui(self, lambda v=done_so_far / total:
            self._lookup_prog_var.set(v))

    if not self._lookup_cancel.is_set():
        done   = self._lookup_done_count
        errors = self._lookup_error_count
        s = f"✓  {done}/{total} applicant(s) processed"
        if errors:
            s += f"  ·  {errors} error(s) — see queue"
        _ui(self, lambda msg=s:
            self._lookup_overall_lbl.config(
                text=msg,
                fg=ACCENT_SUCCESS if not errors else ACCENT_GOLD))
    else:
        _ui(self, lambda:
            self._lookup_overall_lbl.config(
                text="Cancelled.", fg=TXT_MUTED))

    _ui(self, lambda: self._lookup_run_btn.configure(
        state="normal", text="⚡  Run Look-Up"))
    _ui(self, lambda: self._lookup_prog_bar.pack_forget())


# ═══════════════════════════════════════════════════════════════════════
#  SINGLE-FILE PROCESSOR
# ═══════════════════════════════════════════════════════════════════════

def _process_single_file(self, path: str, client, gt) -> None:
    fname = Path(path).name

    file_cancel = threading.Event()

    def cancelled():
        if self._lookup_cancel.is_set():
            file_cancel.set()
        return file_cancel.is_set()

    def step(msg: str):
        _ui(self, lambda m=msg:
            _set_row_status(self, path, "running", m))

    _log(self, f"\n{'═'*60}\nFILE: {fname}\n{'═'*60}")

    step("Reading PDF…")
    try:
        pdf_bytes = Path(path).read_bytes()
        _log(self, f"[{fname}] {len(pdf_bytes):,} bytes (raw)")
    except Exception as e:
        _log(self, f"[{fname}] Read failed: {e}")
        raise

    original_pdf_bytes = pdf_bytes

    step("Correcting page orientation…")
    pdf_bytes = _auto_rotate_pdf(pdf_bytes)
    _log(self, f"[{fname}] {len(pdf_bytes):,} bytes (after rotation fix)")

    step("Step 1/3 — Extracting & classifying pages…")
    try:
        pages_text    = _extract_pages_text(pdf_bytes)
        total_pages   = len(pages_text)
        page_map      = _classify_pages(pages_text)
        map_summary   = _format_page_map(page_map, total_pages)
        section_texts = _group_pages_by_section(pages_text, page_map)
        self._lookup_file_data[path]["page_map"] = map_summary
        _log(self, f"[{fname}] {total_pages} pages.\n{map_summary}")
    except Exception as e:
        _log(self, f"[{fname}] Classification failed: {e}")
        raise
    if cancelled():
        return

    cibi_pages = _pages_for_type(page_map, "cibi", total_pages)
    ws_pages   = _pages_for_type(page_map, "worksheet", total_pages)

    cfa_run_keys = sorted(
        {t for t in page_map.values() if re.match(r"cfa_\d+$", str(t))},
        key=lambda k: int(k.split("_")[1]))
    _log(self, f"[{fname}] CFA runs detected: {cfa_run_keys or ['none']}")

    if not cibi_pages:
        _log(self, f"[{fname}] WARNING: No CI/BI pages detected — using full PDF for cibi call.")
        cibi_pages = list(range(1, total_pages + 1))

    pdf_cibi = _extract_page_subset(pdf_bytes, cibi_pages)
    original_pdf_cibi = _extract_page_subset(original_pdf_bytes, cibi_pages)

    if ws_pages:
        pdf_ws          = _extract_page_subset(pdf_bytes,          ws_pages)
        original_pdf_ws = _extract_page_subset(original_pdf_bytes, ws_pages)
    else:
        pdf_ws = original_pdf_ws = None

    cfa_slices = []
    if cfa_run_keys:
        for ck in cfa_run_keys:
            run_pages = _pages_for_type(page_map, ck, total_pages)
            combined_pages = sorted(set(run_pages) | set(ws_pages))
            if not combined_pages:
                combined_pages = list(range(1, total_pages + 1))
            cfa_slices.append((
                ck,
                _extract_page_subset(pdf_bytes,          combined_pages),
                _extract_page_subset(original_pdf_bytes, combined_pages),
                section_texts.get(ck, ""),
            ))
    else:
        _log(self, f"[{fname}] WARNING: No CFA pages detected — using full PDF for cfa_ws call.")
        all_pages = list(range(1, total_pages + 1))
        cfa_slices.append((
            "cfa_1",
            _extract_page_subset(pdf_bytes,          all_pages),
            _extract_page_subset(original_pdf_bytes, all_pages),
            section_texts.get("cfa", ""),
        ))

    _log(self,
         f"[{fname}] Slices — "
         f"cibi:{len(pdf_cibi):,}b  "
         + "  ".join(f"{ck}:{len(sl):,}b" for ck, sl, _, _ in cfa_slices)
         + f"  (full:{len(pdf_bytes):,}b)")

    pg_summary = (None if not (set(page_map.values()) - {"unknown"})
                  else map_summary)
    if cancelled():
        return

    step(f"Gemini calls — 1 CI/BI + {len(cfa_slices)} CFA run(s)…")
    call_results      = {}
    call_results_lock = threading.Lock()

    call_defs = [
        ("cibi_combined", _gemini_extract_cibi_combined,
         self, client, gt, pdf_cibi, pg_summary,
         section_texts.get("cibi", ""), original_pdf_cibi, file_cancel),
    ]
    for ck, cfa_pdf, cfa_orig, cfa_hint in cfa_slices:
        ws_hint = section_texts.get("worksheet", "")
        call_defs.append((
            f"cfa_ws_{ck}",
            _gemini_extract_cfa_ws_combined,
            self, client, gt, cfa_pdf, pg_summary,
            cfa_hint, ws_hint, cfa_orig, file_cancel,
        ))

    def _run(label, fn, *args):
        if cancelled():
            return
        try:
            result = fn(*args)
        except Exception as exc:
            import traceback
            _log(self, f"[{fname}] '{label}' FAILED: {exc}\n"
                       f"{traceback.format_exc()}")
            result = ("", {})
        with call_results_lock:
            call_results[label] = result

    if MAX_PARALLEL_CALLS >= 2:
        with concurrent.futures.ThreadPoolExecutor(
                max_workers=max(2, len(call_defs)),
                thread_name_prefix=f"gem_{fname[:6]}") as pool:
            fs = [pool.submit(_run, defn[0], defn[1], *defn[2:])
                  for defn in call_defs]
            concurrent.futures.wait(fs)
    else:
        for defn in call_defs:
            if cancelled():
                break
            _run(defn[0], defn[1], *defn[2:])

    if cancelled():
        return

    raw_cibi_c, data_cibi_c = call_results.get("cibi_combined", ("", {}))

    # FIX: Guard against Gemini returning non-dict for data_cibi_c
    if not isinstance(data_cibi_c, dict):
        _log(self, f"[{fname}] WARNING: cibi_combined returned non-dict — resetting")
        data_cibi_c = {}

    CFA_MULTI_FIELDS = (
        "income_remittance",
        "cfa_business_expenses",
        "cfa_household_expenses",
    )
    merged_cfa: dict = {}
    all_cfa_net_incomes: list = []
    raw_cfa_parts: list = []

    for idx, (ck, _, _, _) in enumerate(cfa_slices, start=1):
        label_key  = f"cfa_ws_{ck}"
        raw_r, data_r = call_results.get(label_key, ("", {}))

        # FIX: Guard against non-dict CFA results
        if not isinstance(data_r, dict):
            _log(self, f"[{fname}] WARNING: {label_key} returned non-dict — skipping")
            data_r = {}

        raw_cfa_parts.append(raw_r)

        ordinal = _ordinal(idx)
        prefix  = f"{ordinal} Cashflow"

        for field in CFA_MULTI_FIELDS:
            entries = data_r.get(field, [])
            if not isinstance(entries, list):
                entries = []
            tagged = []
            for entry in entries:
                if isinstance(entry, dict):
                    entry = dict(entry)
                    orig_desc = entry.get("description", "").strip()
                    entry["description"] = (
                        f"{prefix}: {orig_desc}" if orig_desc
                        else prefix)
                    tagged.append(entry)
            merged_cfa.setdefault(field, []).extend(tagged)

        net = str(data_r.get("cfa_net_income", "")).strip()
        if net:
            all_cfa_net_incomes.append(f"{prefix}: {net}")

        WS_FIELDS = (
            "ws_food_grocery", "ws_fuel_transport", "ws_electricity",
            "ws_fertilizer",   "ws_forwarding",     "ws_fuel_diesel",
            "ws_equipment",
        )
        for wf in WS_FIELDS:
            if wf not in merged_cfa and data_r.get(wf):
                merged_cfa[wf] = data_r[wf]

    applicant_name = _sanitize_extracted_text(
        data_cibi_c.get("applicant_name", "").strip())
    gate_result = {
        "applicant_name":    applicant_name,
        "residence_address": _sanitize_extracted_text(
            data_cibi_c.get("residence_address", "")),
        "office_address":    _sanitize_extracted_text(
            data_cibi_c.get("office_address", "")),
    }
    _log(self,
         f"[{fname}] name={applicant_name or '[not found]'}  "
         f"cibi_keys={list(data_cibi_c)}  "
         f"cfa_runs={len(cfa_slices)}  merged_cfa_keys={list(merged_cfa)}")
    _log(self,
         f"[{fname}] Raw (first 3000 chars):\n"
         + "\n---\n".join(filter(None, [raw_cibi_c] + raw_cfa_parts))[:3000])

    # FIX: Build combined dict carefully — start fresh, layer CIBI first,
    # then CFA.  Never let an empty merged_cfa key clobber a CIBI key that
    # shares the same name.  CIBI-only keys are never overwritten by CFA.
    combined = {}
    # Layer 1: all CIBI data
    for k, v in data_cibi_c.items():
        combined[k] = v
    # Layer 2: CFA/WS data — only overwrite if CFA has actual content
    for k, v in merged_cfa.items():
        if v:  # only overwrite if non-empty list
            combined[k] = v

    results = _parse_extraction_response_from_dict(combined)

    fields_found = sum(1 for k, v in results.items()
                       if not k.startswith("_") and v.get("items"))
    _log(self, f"[{fname}] Fields with data: {fields_found}/{len(LOOKUP_ROWS)}")

    results["_applicant_name"]    = applicant_name
    results["_gate_data"]         = gate_result
    results["_page_map"]          = map_summary
    results["_source_file"]       = fname
    results["_cfa_net_income"]    = "  /  ".join(all_cfa_net_incomes)
    results["_cfa_run_count"]     = len(cfa_slices)

    self._lookup_file_data[path]["name"]      = applicant_name
    self._lookup_file_data[path]["gate_data"] = gate_result
    self._lookup_file_data[path]["results"]   = results

    display = applicant_name or fname
    _ui(self, lambda n=display:
        self._lookup_file_data[path]["widgets"]["name_lbl"].config(
            text=f"  {n}"))
    if cancelled():
        return

    step("Saving to Summary…")
    try:
        session_id = getattr(self, "_lookup_session_id",
                             datetime.now().isoformat(timespec="seconds"))
        db_save_applicant(session_id, results)
        _log(self, f"[{fname}] Saved to DB → session {session_id}")
        _ui(self, lambda: lookup_summary_notify(self))
    except Exception as exc:
        _log(self, f"[{fname}] DB save failed (non-fatal): {exc}")

    _ui(self, lambda:
        _set_row_status(self, path, "done", "Done  ·  Saved to Summary"))
    _ui(self, lambda:
        self._lookup_file_data[path]["widgets"]["expand_btn"].configure(
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

_CFA_RUN_BOUNDARY_KEYWORDS = [
    "cashflow analysis",
    "cash flow analysis",
    "cash-flow analysis",
    "income analysis",
    "statement of cash flow",
    "monthly cash flow",
]

_CFA_PAGES_PER_FORM = 2


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

    # FIX: A worksheet page immediately before a CFA page must NOT trigger
    # a new CFA run — only a genuine non-CFA/non-worksheet gap should.
    # Expanded the "previous was not CFA" check to also allow worksheet.
    cfa_run   = 0
    prev_type = None
    for pg in sorted(page_types):
        t = page_types[pg]
        if t == "cfa":
            t_lower = pages_text[pg - 1].lower()

            # A worksheet page immediately before CFA is still the same block
            prev_is_cfa_adjacent = (
                prev_type is not None and
                (re.match(r"cfa_\d+$", str(prev_type)) or prev_type == "worksheet")
            )
            is_new_run_by_gap      = not prev_is_cfa_adjacent
            is_new_run_by_boundary = (
                cfa_run > 0 and
                any(kw in t_lower for kw in _CFA_RUN_BOUNDARY_KEYWORDS)
            )

            if is_new_run_by_gap or is_new_run_by_boundary:
                cfa_run += 1

            page_types[pg] = f"cfa_{cfa_run}"
        prev_type = page_types[pg]

    all_unknown = all(v == "unknown" for v in page_types.values())
    if all_unknown:
        page_types = _heuristic_layout_split(page_types, pages_text)

    return page_types


def _heuristic_layout_split(page_types: dict, pages_text: list) -> dict:
    total  = len(pages_text)
    result = dict(page_types)
    cpf    = _CFA_PAGES_PER_FORM

    if total == 4:
        assign = ["cibi"] * 2 + ["cfa_1"] * 2
    elif total == 5:
        assign = ["cibi"] * 2 + ["cfa_1"] * 2 + ["worksheet"] * 1
    elif total == 6:
        assign = ["cibi"] * 2 + ["cfa_1"] * cpf + ["cfa_2"] * cpf
    elif total == 7:
        assign = ["cibi"] * 2 + ["cfa_1"] * cpf + ["cfa_2"] * cpf + ["worksheet"] * 1
    elif total == 8:
        assign = ["cibi"] * 2 + ["cfa_1"] * cpf + ["cfa_2"] * cpf + ["worksheet"] * 2
    elif total == 9:
        assign = ["cibi"] * 2 + ["cfa_1"] * cpf + ["cfa_2"] * cpf + ["worksheet"] * 3
    else:
        cibi_count = max(1, round(total * 0.25))
        remaining  = total - cibi_count
        ws_count   = 1 if total > 4 else 0
        cfa_total  = remaining - ws_count
        cfa1_count = cfa_total // 2
        cfa2_count = cfa_total - cfa1_count
        assign = (
            ["cibi"]        * cibi_count
            + ["cfa_1"]     * cfa1_count
            + ["cfa_2"]     * cfa2_count
            + ["worksheet"] * ws_count
        )

    assign = (assign + ["unknown"] * total)[:total]
    for i, label in enumerate(assign):
        result[i + 1] = label
    return result


def _format_page_map(page_map: dict, total_pages: int) -> str:
    labels = {
        "credit_scoring": "Credit Scoring",
        "cibi":           "CI/BI Report",
        "worksheet":      "Worksheet",
        "unknown":        "Unclassified",
    }

    def _label(t: str) -> str:
        if t in labels:
            return labels[t]
        m = re.match(r"cfa_(\d+)$", t)
        if m:
            n = int(m.group(1))
            ord_ = _ORDINALS[n - 1] if 1 <= n <= len(_ORDINALS) else f"#{n}"
            return f"Cashflow Analysis ({ord_})"
        return "?"

    return "\n".join(
        f"  Page {pg:>2}: {_label(page_map.get(pg, 'unknown'))}"
        for pg in range(1, total_pages + 1))


def _group_pages_by_section(pages_text: list, page_map: dict) -> dict:
    groups = {k: [] for k in
              ("credit_scoring", "cibi", "worksheet", "unknown")}
    for i, text in enumerate(pages_text):
        pg    = i + 1
        dtype = page_map.get(pg, "unknown")
        groups.setdefault(dtype, []).append(f"[Page {pg}]\n{text}")

    cfa_keys = sorted(
        {t for t in page_map.values() if re.match(r"cfa_\d+$", str(t))},
        key=lambda k: int(k.split("_")[1]))
    for ck in cfa_keys:
        groups.setdefault(ck, [])
    all_cfa = []
    for ck in cfa_keys:
        all_cfa.extend(groups[ck])
    groups["cfa"] = all_cfa

    return {k: "\n\n".join(v) for k, v in groups.items()}


# ═══════════════════════════════════════════════════════════════════════
#  GEMINI HELPERS
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
            api_key = getattr(self, attr)
            break
    if not api_key:
        try:
            from app_constants import GEMINI_API_KEY
            api_key = GEMINI_API_KEY
        except Exception:
            pass
    if not api_key:
        try:
            import config
            api_key = config.GEMINI_API_KEY
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
            # FIX: Differentiate quota exhaustion (needs long wait) from
            # transient server errors (short wait is fine).
            is_quota     = any(x in msg for x in ("429", "quota", "RESOURCE_EXHAUSTED"))
            is_transient = is_quota or any(x in msg for x in (
                "500", "502", "503", "unavailable",
                "Empty response", "timeout", "timed out",
                "Connection", "RemoteDisconnected"))

            if is_transient and attempt < max_retries - 1:
                if is_quota:
                    # Quota errors still need a longer backoff
                    base = [60, 90, 120][min(attempt, 2)]
                else:
                    # Transient 5xx / network errors: short retry
                    base = GEMINI_RETRY_DELAYS[min(attempt, len(GEMINI_RETRY_DELAYS)-1)]
                delay = base + random.uniform(-base * 0.1, base * 0.1)

                _code = "ERR"
                for _candidate in ("429", "502", "503", "500"):
                    if _candidate in msg:
                        _code = _candidate
                        break
                else:
                    if is_quota:
                        _code = "QUOTA"
                    elif "timeout" in msg.lower() or "timed out" in msg.lower():
                        _code = "TIMEOUT"
                    elif "unavailable" in msg.lower():
                        _code = "UNAVAILABLE"

                _ui(self, lambda w=int(delay), a=attempt, r=max_retries, c=_code:
                    self._lookup_overall_lbl.config(
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
            _log(self, f"  Sending {label} PDF bytes "
                       f"({len(attempt_bytes):,}b) to Gemini…")
            return _gemini_call(
                self, client, gt,
                [_pdf_part(gt, attempt_bytes), prompt],
                config,
                cancel_event=cancel_event)
        except Exception as exc:
            msg = str(exc)
            is_invalid = "400" in msg or "INVALID_ARGUMENT" in msg
            if is_invalid and label == "processed":
                _log(self, f"  400 INVALID_ARGUMENT on {label} PDF — "
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


# ═══════════════════════════════════════════════════════════════════════
#  GEMINI CALL 1 — CI/BI + ASSETS
# ═══════════════════════════════════════════════════════════════════════

def _gemini_extract_cibi_combined(self, client, gt,
                                  pdf_bytes: bytes,
                                  pg_summary, hint: str,
                                  original_pdf_bytes: bytes = None,
                                  cancel_event=None) -> tuple:
    if pg_summary:
        scope = (f"The page classifier identified these sections:\n{pg_summary}\n\n"
                 f"Focus on the CI/BI Report and Assets pages identified above.")
    else:
        scope = ("The document is a fully scanned PDF. Search ALL pages for "
                 "CI/BI Report content and the Assets page.")

    hint_block = (f"\n\nPartial text extracted from CI/BI pages (use only as hint):\n{hint[:3000]}"
                  if hint else "")

    prompt = f"""You are a credit analyst extracting data from a Philippine rural bank loan application.
This is a SCANNED document — printed form labels with handwritten/typed values.
Read the document VISUALLY as an image.
 
{scope}
 
══════════════════════════════════════════════════════
CORE READING RULES  (apply to every field below)
══════════════════════════════════════════════════════
 
RULE 1 — VALUE BELONGS TO THE LABEL IT SITS BESIDE, NOT ABOVE OR BELOW.
  Forms use a two-column layout: label on the left, value blank on the right.
  A value written on a line belongs ONLY to the label printed on THAT SAME LINE.
  Never "look up" to the label above or "look down" to the label below to
  explain a value. The row boundary is sacred.
 
RULE 2 — STACKED RADIO-STYLE LABELS ("Employed / Self-employed").
  Some forms show two or three sub-labels stacked vertically, each with its
  own value blank to the right:
      Occupation/Business: _______________
            Employed:      _______________   ← value on THIS line → Employed
            Self-employed: ___[T...]______   ← value on THIS line → Self-employed
  Rule: whichever line has the handwritten entry is the one that is ticked/filled.
  The OTHER lines stay blank (their blanks are empty). Never move a value from
  one sub-label line to another. If "Self-employed:" has "T..." written beside it,
  the value is Self-employed = "T...", and Employed = blank.
 
RULE 3 — NEVER USE A PRINTED LABEL AS A VALUE.
  "Occupation/Business", "Employed", "Self-employed", "Office Address" etc. are
  labels — they are pre-printed on the form. Only handwritten/typed text in the
  blank area to the right of (or below) the label is the value.
 
RULE 4 — SECTION BOUNDARY: APPLICANT vs SPOUSE.
  The CI/BI form has two mirrored blocks: one for the APPLICANT and one for the
  SPOUSE. Each block has its own "Occupation/Business", "Employed",
  "Self-employed", and "Office Address" rows.
  • Applicant's "Office Address" → cibi_place_of_work
  • Spouse's "Office Address"    → cibi_spouse_office
  Never mix values across the section boundary.
 
RULE 5 — DASH OR BLANK = NO DATA.
  If a cell contains only a dash ( - or — ), is empty, or is zero → return []
  or "" for that field. Do not substitute with a neighbouring row's value.
 
══════════════════════════════════════════════════════
 
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PAGE 1 (or first CI/BI page): CI/BI REPORT
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 
① applicant_name  (plain string)
   Label: "Name of Applicant", "Borrower", "Client", or "Name".
   Return the handwritten full name only. Never return the label itself.
 
② residence_address  (plain string)
   Label: "Residence Address", "Home Address", or "Permanent Address".
   Philippine notation: expand "P1"→"Purok 1", "P2"→"Purok 2", etc.
   Return only the handwritten address.
 
③ office_address  (plain string)
   The APPLICANT'S Office Address — the "Office Address" row that appears
   in the APPLICANT's block (NOT the spouse block). Return "" if blank.
 
④ cibi_spouse  (array)
   Location: the SPOUSE block of the CI/BI page.
   Sub-labels in the spouse block (stacked vertically, each with its own blank):
       Name of Spouse:   _______________
       Occupation/Business: ____________
             Employed:   _______________
             Self-employed: ____________
 
   Step A — Spouse name: read the value on the "Name of Spouse" line.
   Step B — Employment: look at the Employed and Self-employed lines
             (apply RULE 2 above). Whichever line has a handwritten entry is the
             active employment type and its value is the employer/business name.
             If both are blank, use the Occupation/Business line value if present.
   Step C — Combine into ONE item:
             description = "<Spouse Name> — <Employment Type>: <Value>"
             e.g. "Maria Santos — Self-employed: Tindahan"
             If only name is present → use name alone.
             If only employment is present → use that alone.
   Return [] if both spouse name and employment are entirely blank.
   NEVER create separate items for name and employment — always ONE combined item.
 
⑤ cibi_spouse_office  (array)
   The "Office Address" row that is inside the SPOUSE block (below the spouse
   employment lines). description = the address or employer name written there.
   amount = "N/A". Return [] if blank or absent.
 
⑥ cibi_place_of_work  (array)
   The APPLICANT's employer or own business name and address.
   Source rows in the APPLICANT block:
     • "Occupation/Business" line (and its Employed / Self-employed sub-rows,
       applying RULE 2 — only the filled sub-row counts)
     • "Office Address" in the applicant block
     • Any "Name of Employer / Business", "Nature of Business", "Position" rows
   Also check TRADE REFERENCES and BANK REFERENCES for employer details.
   amount = "N/A".
   Do NOT pull anything from the SPOUSE block.
 
⑦ cibi_temp_residence  (array)
   The applicant's residence address stored as an array item.
   Same value as residence_address. description = full address. amount = "N/A".
   Write "Purok 1" etc. — never abbreviate as "P1".
 
⑧ cibi_petrol_products  (array)
   Flag if the applicant's employer, own business, or any trade/bank reference
   involves: petroleum, oil depot, gasoline station, LPG, fuel, lubricants,
   plastics, PVC, polypropylene, rubber, chemicals, fertilizer manufacturing.
   Return [] if none found.
 
⑨ cibi_transport_services  (array)
   Flag if the applicant's employer, own business, or any trade/bank reference
   involves: bus, jeepney, tricycle, forwarding, trucking, hauling, heavy
   equipment, crane, bulldozer, backhoe, freight, logistics, cargo, courier.
   Return [] if none found.
 
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STILL ON CI/BI PAGE: CREDIT HISTORY & REFERENCES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 
⑩ credit_history_amort  (array)
   Find the section/table labelled "CREDIT HISTORY & REFERENCES".
 
   Columns (left to right):
     Bank/Lending Institution | Principal Loan | Due Date | Amort. | Balance
 
   COLUMN IDENTIFICATION (important for rotated or skewed scans):
   • Identify each column by its printed header text, not by position alone.
   • "Amort." is the 4th column — it is NEITHER Principal (2nd) NOR Balance (5th).
   • If headers are hard to read, the Amort. column is always narrower than
     Principal and Balance, and contains monthly payment amounts.
 
   WHAT TO EXTRACT:
   • Extract individual data rows (one per lending institution).
     description = institution name, amount = Amort. value, date = Due Date.
   • Use amount="" for rows where Amort. cell is blank or a dash.
   • Return [] only if the entire table is blank.
 
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STILL ON CI/BI PAGE: BALANCE SHEET (Assets column only)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 
⑪ cibi_business_inventory  (array)
   Source: "BALANCE SHEET" section → ASSETS column → "Business Inventory" row.
 
   The ASSETS column rows in ORDER (top to bottom):
     1. Cash on Hand
     2. Bank Deposits
     3. Accounts Receivable
     4. Real Properties
     5. Personal Assets
     6. Business Assets          ← row 6
     7. Business Inventory       ← row 7  ← THIS is the target row
     8. TOTAL ASSETS
 
   ANCHORING RULE: The value for "Business Inventory" is the amount written on
   row 7 — the row IMMEDIATELY below "Business Assets". These two rows are
   adjacent and their value cells sit directly beside their respective labels.
   Do NOT use the row 6 (Business Assets) value for row 7, even if row 7 appears
   blank at first glance. Re-examine the scan carefully before concluding blank.
 
   Return [] if: Business Inventory row value is a dash, blank, or zero.
   For a real value:
     description = "Business Inventory"
     amount      = peso amount exactly as written (e.g. "50,000.00")
 
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
NEXT PAGE: ASSETS PAGE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 
FORMAT DETECTION — look at the heading on the Assets page:
  → Separate "PERSONAL ASSETS" AND "BUSINESS ASSETS" headings → Format A
  → Single "PERSONAL & BUSINESS ASSETS" heading → Format B
  → Uncertain → default to Format B
 
FORMAT A — Two separate categories:
  "PERSONAL ASSETS (Non-Productive Assets)" and "BUSINESS ASSETS (Productive Assets)"
  Each has a "SERIALIZED HOUSEHOLD ASSETS" sub-section with columns:
    Item | Description | Serial No. | Acquisition Cost
 
  cibi_personal_assets:
    Extract filled rows from PERSONAL ASSETS → SERIALIZED HOUSEHOLD ASSETS only.
    description = item name + brand/model + serial number
    amount = Acquisition Cost column value, or "" if blank.
    Return [] if entirely blank.
 
  cibi_business_assets:
    Extract filled rows from BUSINESS ASSETS → SERIALIZED HOUSEHOLD ASSETS only.
    description = item name + brand/model + serial number
    amount = Acquisition Cost column value, or "" if blank.
    Return [] if entirely blank.
 
FORMAT B — Single combined category:
  "PERSONAL & BUSINESS ASSETS" with sub-sections:
    - "SERIALIZED ASSETS" (appliances, equipment, electronics)
    - "VEHICLES" (motorcycles, motor vehicles, boats, tractors)
  Columns: Item | Description | Serial/Plate No. | Acquisition Cost
 
  cibi_personal_assets:
    Extract ALL rows from BOTH sub-sections (Serialized + Vehicles).
    description = item type + brand/model + serial or plate number
    amount = Acquisition Cost as written, or "" if blank.
    Return [] only if both sub-sections are entirely blank.
 
  cibi_business_assets:
    Return [] — everything goes in cibi_personal_assets for Format B.
    DO NOT duplicate.
 
ROW-LEVEL ANCHORING FOR ASSETS:
  Each item row has its own serial number and cost cell. The cost on row N
  belongs ONLY to the item on row N. If two items are listed consecutively
  and only one has a cost filled in, the other item has amount="".
 
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT — return ONLY valid JSON, nothing else:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{{
  "applicant_name":          "",
  "residence_address":       "",
  "office_address":          "",
  "cibi_spouse":             [{{"description":"...","amount":"N/A","date":""}}],
  "cibi_spouse_office":      [{{"description":"...","amount":"N/A","date":""}}],
  "cibi_place_of_work":      [{{"description":"...","amount":"N/A","date":""}}],
  "cibi_temp_residence":     [{{"description":"...","amount":"N/A","date":""}}],
  "cibi_petrol_products":    [{{"description":"...","amount":"N/A","date":""}}],
  "cibi_transport_services": [{{"description":"...","amount":"N/A","date":""}}],
  "credit_history_amort":    [{{"description":"...","amount":"...","date":""}}],
  "cibi_business_inventory": [{{"description":"Business Inventory","amount":"...","date":""}}],
  "cibi_personal_assets":    [{{"description":"...","amount":"...","date":""}}],
  "cibi_business_assets":    [{{"description":"...","amount":"...","date":""}}]
}}
 
FINAL CHECKS BEFORE RESPONDING:
  1. cibi_spouse — is it exactly ONE combined item (name + employment)?
     Did I apply RULE 2 to determine which employment sub-label is filled?
  2. cibi_place_of_work vs cibi_spouse_office — did I respect the
     APPLICANT / SPOUSE section boundary (RULE 4)?
  3. cibi_business_inventory — did I read row 7 (Business Inventory),
     not row 6 (Business Assets)?
  4. For every stacked sub-label group — did I assign the value only to
     the line it sits beside (RULE 2)?
  5. Did any printed label text accidentally end up as a value? If so, remove it.{hint_block}"""

    fallback = original_pdf_bytes if original_pdf_bytes is not None else pdf_bytes
    resp = _gemini_call_with_fallback(
        self, client, gt,
        pdf_bytes, fallback,
        prompt,
        gt.GenerateContentConfig(temperature=0.0),
        cancel_event=cancel_event)
    raw  = resp.text or ""
    data = _parse_json_safe(raw)

    # FIX: Guard against non-dict parse result
    if not isinstance(data, dict):
        data = {}

    _NA = {"N/A", "NA", "NONE", "NONE.", "-", "\u2014", "N/A.", "N.A."}
    for field in ("applicant_name", "residence_address", "office_address"):
        val = data.get(field, "").strip()
        data[field] = "" if val.upper() in _NA else val
    return raw, data


# ═══════════════════════════════════════════════════════════════════════
#  GEMINI CALL 2 — CFA + WORKSHEET
# ═══════════════════════════════════════════════════════════════════════

def _gemini_extract_cfa_ws_combined(self, client, gt,
                                    pdf_bytes: bytes,
                                    pg_summary,
                                    cfa_hint: str,
                                    ws_hint: str,
                                    original_pdf_bytes: bytes = None,
                                    cancel_event=None) -> tuple:
    if pg_summary:
        scope = (f"The page classifier identified these sections:\n{pg_summary}\n\n"
                 f"Focus on the Cashflow Analysis and Worksheet pages above.")
    else:
        scope = ("The document is a fully scanned PDF. Search ALL pages for "
                 "Cashflow Analysis and Worksheet content.")

    hint_block = ""
    if cfa_hint:
        hint_block += f"\n\nPartial text from CFA pages (use as hint only):\n{cfa_hint[:2000]}"
    if ws_hint:
        hint_block += f"\n\nPartial text from Worksheet pages (use as hint only):\n{ws_hint[:2000]}"

    prompt = f"""You are a credit analyst extracting data from a Philippine rural bank loan application.
This is a SCANNED document — printed form labels with handwritten/typed values.
Read the document VISUALLY as an image.
 
{scope}
 
══════════════════════════════════════════════════════
CORE READING RULES  (apply to every field below)
══════════════════════════════════════════════════════
 
RULE 1 — VALUE BELONGS TO THE LABEL ON THE SAME LINE/ROW.
  These forms use a table layout: label in the left column, value(s) in the
  right column(s). A value written in row N belongs ONLY to the label in row N.
  Never attribute a value to the row above or below it.
 
RULE 2 — COLUMN IDENTITY: USE HEADER TEXT, NOT COLUMN POSITION ALONE.
  Tables have multiple numeric columns (Daily, Weekly, Monthly, etc.).
  Always identify which column you are reading by its printed header text.
  For income rows: use the "Monthly Totals" (rightmost summary) column only —
  not the Daily or Weekly columns, even if they are larger numbers.
 
RULE 3 — SECTION BOUNDARY.
  CFA page has distinct labelled sections: "SOURCE OF INCOME",
  "BUSINESS EXPENSES", "HOUSEHOLD / PERSONAL EXPENSES".
  A row belongs to the section whose heading it falls under.
  Never move a row from one section to another.
 
RULE 4 — WORKSHEET vs CFA.
  The Worksheet page is SEPARATE from the CFA page. Worksheet rows give
  per-unit breakdowns (Qty × Unit Cost = Total). CFA rows give monthly totals.
  ws_* fields must come from the Worksheet page, not the CFA page.
  cfa_* fields must come from the CFA page, not the Worksheet page.
 
RULE 5 — BLANK OR DASH = NO DATA.
  Use amount="" for any row whose value cell is blank or contains only a dash.
 
══════════════════════════════════════════════════════
 
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CASHFLOW ANALYSIS (CFA) PAGE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 
① income_remittance — SOURCE OF INCOME section
   Table columns: Income Source | Daily | Weekly | Semi-Monthly | Monthly | Monthly Totals
   Extract every row that has any content written.
   description = income source label (leftmost column)
   amount      = "Monthly Totals" column value (rightmost filled column)
                 ← this is the column with the aggregated monthly figure.
                 DO NOT use Daily or Weekly column values as the amount.
   Include all income types: salary, farming, sari-sari, remittance/padala,
   tricycle operation, school service, pension, any other.
   Return [] only if the entire section is blank.
 
② cfa_business_expenses — BUSINESS EXPENSES section
   Extract every row under the "BUSINESS EXPENSES" heading.
   description = expense label. amount = amount written on that same row.
   Use amount="" for rows with a label but no amount filled in.
   STOP at the section boundary — do NOT read rows from the
   Household/Personal Expenses section below.
 
③ cfa_household_expenses — HOUSEHOLD / PERSONAL EXPENSES section
   Extract every row under the "HOUSEHOLD EXPENSES", "PERSONAL EXPENSES",
   or "FAMILY EXPENSES" heading.
   Includes: food, electricity, water, clothing, school fees, medical,
   personal transportation, loan payments, and all other listed items.
   description = expense label. amount = amount on that same row.
   Use amount="" for rows with a label but no amount filled in.
   STOP at the section boundary — do NOT read rows from Business Expenses.
 
④ cfa_net_income — single bottom-line summary
   Labels: "Total Net Income", "Net Income", "Net Cash Flow",
   "Net Surplus", "NET INCOME", "TOTAL NET INCOME".
   This is ONE amount at the very bottom of the CFA — the final surplus figure.
   Return the value exactly as written (e.g. "P 8,500.00" or "12,000").
   Return "" if absent or blank.
 
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BUSINESS EXPENSE WORKSHEET PAGE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 
This is a SEPARATE page from the CFA. It lists individual expense line items
with columns: Item/Description | Qty | Unit | Unit Cost | Monthly Total
 
ROW ANCHORING RULE FOR WORKSHEET:
  Each expense row has its own label and its own Qty/Unit/Cost/Total cells.
  The monthly total on row N belongs ONLY to the item labelled on row N.
  Adjacent rows (e.g. "Fuel/Diesel" directly below "Forwarding") each have
  their own separate total — never assign one row's total to the other.
 
⑤ ws_food_grocery
   Row label: "Food / Grocery", "Food and Grocery", or "Food".
   amount = monthly total on THAT row only.
 
⑥ ws_fuel_transport
   Row label: "Fuel and Transportation", "Transportation", or "Gasoline / Fare"
   in the HOUSEHOLD section of the worksheet.
   This is PERSONAL transport cost (fare, commute) — not business fuel/diesel.
   amount = monthly total on THAT row only.
 
⑦ ws_electricity
   Row label: "Electricity", "Electric Bill", or a Philippine cooperative name
   (ANTECO, MORESCO, MERALCO, CASURECO, FICELCO, BUSECO, or similar).
   description = include the utility/co-op name if written on that row.
   amount = monthly bill amount on THAT row only.
 
⑧ ws_fertilizer
   Row label: "Fertilizer", "Fertilizer / Pesticide", or "Farm Inputs"
   in the BUSINESS section of the worksheet.
   Include type (Urea, Complete, Organic), quantity, unit cost, total if written.
   Combine into the description. amount = monthly total on THAT row.
 
⑨ ws_forwarding
   Row label: "Forwarding", "Trucking / Hauling", "Hauling", or "Freight"
   in the BUSINESS EXPENSE section (not the household section).
   amount = monthly total on THAT row only.
   ANCHORING: This row's total is SEPARATE from the Fuel/Diesel row immediately
   adjacent to it. Do not swap their totals.
 
⑩ ws_fuel_diesel
   Row label: "Fuel / Gas / Diesel", "Diesel", "Gasoline", or "Fuel Cost"
   in the BUSINESS EXPENSE section (not personal transportation).
   Include fuel type, liters, unit price, total if written.
   Combine into description. amount = monthly total on THAT row.
   ANCHORING: This row's total is SEPARATE from the Forwarding row adjacent to it.
 
⑪ ws_equipment
   Row label: "Cost of Rent of Equipment", "Equipment Rental",
   "Tractor Rental", "Backhoe Rental", "Thresher Rental".
   Include equipment type, rate, period, total if written.
   amount = monthly total on THAT row only.
 
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT — return ONLY valid JSON, nothing else:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
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
 
FINAL CHECKS BEFORE RESPONDING:
  1. income_remittance — did I use ONLY the Monthly Totals column (not Daily/Weekly)?
  2. cfa_business_expenses / cfa_household_expenses — are all rows in the correct
     section with no cross-contamination?
  3. cfa_net_income — is it the single final surplus figure, not a section subtotal?
  4. ws_* fields — did they all come from the WORKSHEET page, not the CFA page?
  5. ws_fuel_transport vs ws_fuel_diesel — personal transport → ws_fuel_transport;
     business fuel → ws_fuel_diesel. Check which section each row belongs to.
  6. ws_forwarding vs ws_fuel_diesel — are their monthly totals correctly assigned
     to their own rows (not swapped due to adjacency)?
  7. Did any printed label text end up as a value? If so, remove it.{hint_block}"""

    fallback = original_pdf_bytes if original_pdf_bytes is not None else pdf_bytes
    resp = _gemini_call_with_fallback(
        self, client, gt,
        pdf_bytes, fallback,
        prompt,
        gt.GenerateContentConfig(temperature=0.0),
        cancel_event=cancel_event)
    raw  = resp.text or ""
    data = _parse_json_safe(raw)

    # FIX: Guard against non-dict parse result
    if not isinstance(data, dict):
        data = {}

    return raw, data


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

    # FIX: The original regex  P(\d{1,2})  incorrectly converts peso amounts
    # like "PHP 50,000" or "P 8,500" → "PHurok  50,000" / "Purok 8,500".
    # Restrict the Purok expansion to standalone "P" not preceded by letters
    # and not followed by a space or currency context.
    text = re.sub(r"(?<![A-Za-z])P(\d{1,2})(?!\d)",
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
        # FIX: Entries might arrive as a dict (single item) rather than a
        # list when Gemini omits the surrounding brackets.  Normalise early.
        if isinstance(entries, dict):
            entries = [entries]
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
        "cibi_business_inventory",
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
                # FIX: Strip leading peso/currency prefix before numeric parse
                # so "P 8,500.00", "PHP8500", "₱ 1,200" all parse correctly.
                amt_clean = re.sub(r"^(?:PHP?|₱)\s*", "", amt, flags=re.IGNORECASE).strip()
                nums = re.findall(r"[\d,]+\.?\d*", re.sub(r"[^\d.,]", " ", amt_clean))
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

def _log(self, msg: str):
    with self._lookup_raw_lock:
        self._lookup_raw_log.append(msg)
        snapshot = "\n".join(self._lookup_raw_log)
    _ui(self, lambda s=snapshot: _set_raw(self, s))


def _set_raw(self, text: str):
    box = self._lookup_raw_box
    box.config(state="normal")
    box.delete("1.0", "end")
    if text:
        box.insert("end", text)
    box.config(state="disabled")


def _toggle_raw(self):
    if self._lookup_raw_visible:
        self._lookup_raw_frame.pack_forget()
        self._lookup_raw_toggle_btn.configure(text="▼ Show")
        self._lookup_raw_visible = False
    else:
        self._lookup_raw_frame.pack(fill="x", pady=(8, 0))
        self._lookup_raw_toggle_btn.configure(text="▲ Hide")
        self._lookup_raw_visible = True


def _ui(self, fn):
    self.after(0, fn)


# ═══════════════════════════════════════════════════════════════════════
#  ATTACH
# ═══════════════════════════════════════════════════════════════════════

def attach(cls):
    cls._build_lookup_panel = _build_lookup_panel
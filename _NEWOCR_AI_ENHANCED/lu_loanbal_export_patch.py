"""
lu_loanbal_export_patch.py
==========================
Adds an Export button to the Sector vs Loan Balance tab.

HOW TO APPLY
------------
In app.py, after lu_analysis_tab.attach() and any other patch calls:

    import lu_loanbal_export_patch
    lu_loanbal_export_patch.attach(DocExtractorApp)

WHAT IT DOES
------------
• Replaces _build_loanbal_panel with a version that adds a
  "💾 Export" button in the tab header bar.
• The button opens a small popup menu with two options:
    – 📄 Export PDF   → calls _loanbal_export_pdf()
    – 📊 Export Excel → calls _loanbal_export_excel()
• Both export functions write a dedicated "Sector vs Loan Balance"
  report (they do NOT call the existing general _export_pdf / _export_excel
  — those are for the Risk Analysis report).
• All colour constants and helpers are self-contained so this patch
  has no dependency on lu_ui internals beyond what is already exported
  via lu_analysis_tab.

COMPATIBILITY
-------------
Fully compatible with:
  • lu_core.py        (reads .general, .sector_map, .totals from run_lu_analysis)
  • lu_ui.py          (monkey-patches only _build_loanbal_panel; does not touch
                       any other method)
  • lu_analysis_tab.py (re-export shim)
  • lu_client_search_patch.py (no conflicts — that patch only touches
                               _build_lu_analysis_panel,
                               _lu_populate_client_dropdown,
                               _lu_on_client_change)
"""

import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
from pathlib import Path

# ── Optional deps (same guards as lu_ui) ─────────────────────────────────────
try:
    import openpyxl
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False

try:
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as plt
    import matplotlib.ticker
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    _HAS_MPL = True
except ImportError:
    _HAS_MPL = False

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        HRFlowable, PageBreak,
    )
    _HAS_RL = True
except ImportError:
    _HAS_RL = False


# ── Colour palette (mirrors lu_ui) ───────────────────────────────────────────
_NAVY_DEEP    = "#0A1628"
_NAVY_MID     = "#1A3A6B"
_NAVY_LIGHT   = "#1E4080"
_NAVY_MIST    = "#EEF3FB"
_NAVY_GHOST   = "#D6E4F7"
_NAVY_PALE    = "#4A6FA5"
_WHITE        = "#FFFFFF"
_CARD_WHITE   = "#FAFBFD"
_OFF_WHITE    = "#F5F7FA"
_BORDER_LIGHT = "#E2E8F5"
_BORDER_MID   = "#C5D0E8"
_TXT_NAVY     = "#1A2B4A"
_TXT_SOFT     = "#6B7FA3"
_TXT_MUTED    = "#9AAACE"
_TXT_ON_LIME  = "#0A2010"
_LIME_BRIGHT  = "#B8FF57"
_LIME_MID     = "#8FD14F"
_LIME_DARK    = "#5A9E28"
_LIME_PALE    = "#D4F5A0"
_ACCENT_RED   = "#E53E3E"
_ACCENT_GOLD  = "#D4A017"
_ACCENT_SUCCESS = "#2E7D32"

_RISK_COLOR    = {"HIGH": _ACCENT_RED, "MODERATE": _ACCENT_GOLD, "LOW": _ACCENT_SUCCESS}
_RISK_BG       = {"HIGH": "#FFF5F5", "MODERATE": "#FFFBF0", "LOW": "#F0FBE8"}
_RISK_BADGE_BG = {"HIGH": "#FFE8E8", "MODERATE": "#FFF3CD", "LOW": "#DCEDC8"}

# Import sector constants from lu_core via lu_analysis_tab
from lu_analysis_tab import (
    SECTOR_WHOLESALE, SECTOR_AGRICULTURE, SECTOR_TRANSPORT,
    SECTOR_REMITTANCE, SECTOR_CONSUMER, SECTOR_OTHER,
    _compute_risk_score,
)

_SECTOR_COLORS = {
    SECTOR_WHOLESALE:   "#1A3A6B",
    SECTOR_AGRICULTURE: "#2E7D32",
    SECTOR_TRANSPORT:   "#D4A017",
    SECTOR_REMITTANCE:  "#8B5CF6",
    SECTOR_CONSUMER:    "#E53E3E",
    SECTOR_OTHER:       "#9AAACE",
}
_SECTOR_ICON = {
    SECTOR_WHOLESALE:   "🏪",
    SECTOR_AGRICULTURE: "🐟",
    SECTOR_TRANSPORT:   "🚛",
    SECTOR_REMITTANCE:  "💸",
    SECTOR_CONSUMER:    "🏦",
    SECTOR_OTHER:       "🏠",
}
_CHART_SECTORS = [
    SECTOR_WHOLESALE, SECTOR_AGRICULTURE,
    SECTOR_TRANSPORT, SECTOR_REMITTANCE, SECTOR_CONSUMER,
]

_MPL_BG   = "#FAFBFD"
_MPL_NAVY = "#1A3A6B"

# ── Pagination & filter state constants ───────────────────────────────────────
_LOANBAL_PAGE_SIZE = 15
_LOANBAL_SENTINEL  = object()

# ── Column specs (must match lu_ui.py) ────────────────────────────────────────
_SECTOR_COLS = [
    ("Sector",              3, 200, "w"),
    ("# Clients",           1,  80, "center"),
    ("Total Loan Balance",  2, 170, "center"),
    ("% of Total",          2, 130, "center"),
    ("Avg Loan per Client", 2, 160, "center"),
    ("Avg Net Income",      2, 150, "center"),
    ("Risk Profile",        1, 100, "center"),
]

_CLIENT_COLS = [
    ("Client",          3, 200, "w"),
    ("ID",              1,  60, "center"),
    ("Sector",          2, 160, "center"),
    ("Principal Loan",  2, 130, "center"),
    ("Loan Balance",    2, 130, "center"),
    ("% of Total",      1, 100, "center"),
    ("Net Income",      2, 120, "center"),
    ("Current Amort",   2, 130, "center"),
    ("Risk",            1,  90, "center"),
]

# ── Import helpers from lu_ui ─────────────────────────────────────────────────
from lu_ui import (
    _lu_get_filtered_all_data, _lu_get_active_sectors,
    _make_table_frame, _table_header, _table_divider,
)


def F(size, weight="normal"):
    return ("Segoe UI", size, weight)

def FF(size, weight="normal"):
    return ctk.CTkFont(family="Segoe UI", size=size, weight=weight)


# ── Scrollable helper ────────────────────────────────────────────────────────

def _bind_mousewheel(canvas: tk.Canvas):
    def _on_enter(e):
        canvas.bind_all(
            "<MouseWheel>",
            lambda ev: canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units"))
    def _on_leave(e):
        canvas.unbind_all("<MouseWheel>")
    canvas.bind("<Enter>", _on_enter)
    canvas.bind("<Leave>", _on_leave)


def _make_scrollable(parent, bg=None):
    bg = bg or _CARD_WHITE
    outer  = tk.Frame(parent, bg=bg)
    outer.pack(fill="both", expand=True)
    sb     = tk.Scrollbar(outer, relief="flat", troughcolor=_OFF_WHITE,
                          bg=_BORDER_LIGHT, width=8, bd=0)
    sb.pack(side="right", fill="y")
    canvas = tk.Canvas(outer, bg=bg, highlightthickness=0, yscrollcommand=sb.set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.config(command=canvas.yview)
    inner  = tk.Frame(canvas, bg=bg)
    win    = canvas.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>",
               lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>",
                lambda e: canvas.itemconfig(win, width=e.width))
    _bind_mousewheel(canvas)
    return outer, inner, canvas


# ══════════════════════════════════════════════════════════════════════════════
#  PATCHED _build_loanbal_panel
# ══════════════════════════════════════════════════════════════════════════════

def _build_loanbal_panel_patched(self, parent):
    hdr = tk.Frame(parent, bg=_NAVY_MID, height=46)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)

    self._loanbal_hdr_lbl = tk.Label(
        hdr,
        text="📊  Sector vs Total Loan Balance  —  Exposure Analysis",
        font=F(10, "bold"), fg=_WHITE, bg=_NAVY_MID,
    )
    self._loanbal_hdr_lbl.pack(side="left", padx=20, pady=12)

    self._loanbal_export_btn = ctk.CTkButton(
        hdr,
        text="💾  Export",
        command=lambda: _loanbal_show_export_menu(self),
        width=110, height=30, corner_radius=6,
        fg_color=_LIME_DARK, hover_color=_LIME_MID,
        text_color=_TXT_ON_LIME, font=FF(9, "bold"),
        state="disabled",
    )
    self._loanbal_export_btn.pack(side="right", padx=16, pady=8)

    self._loanbal_body = tk.Frame(parent, bg=_CARD_WHITE)
    self._loanbal_body.pack(fill="both", expand=True)
    tk.Label(
        self._loanbal_body,
        text="Run an analysis first to view loan balance exposure.",
        font=F(10), fg=_TXT_MUTED, bg=_CARD_WHITE,
    ).pack(pady=60)


# ══════════════════════════════════════════════════════════════════════════════
#  CLIENT TABLE HELPER — paginated + risk-filtered + name-searched
# ══════════════════════════════════════════════════════════════════════════════

def _loanbal_render_client_table(pad, clients_sorted, grand_lb,
                                 page, risk_filter, name_search, canvas,
                                 on_page_change, on_risk_change, on_name_search):
    # ── Heading row: title left, risk pills right ─────────────────────────────
    hdr_row = tk.Frame(pad, bg=_CARD_WHITE)
    hdr_row.pack(fill="x", pady=(14, 4))

    tk.Label(hdr_row, text="Individual Client Loan Balance",
             font=F(11, "bold"), fg=_TXT_NAVY, bg=_CARD_WHITE
             ).pack(side="left", anchor="w")

    pill_frame = tk.Frame(hdr_row, bg=_CARD_WHITE)
    pill_frame.pack(side="right", anchor="e", padx=(0, 4))

    PILLS = [
        (None,       "All",      _NAVY_MID,       _WHITE),
        ("HIGH",     "🔴 High",  _ACCENT_RED,     _WHITE),
        ("MODERATE", "🟡 Mod",   _ACCENT_GOLD,    _WHITE),
        ("LOW",      "🟢 Low",   _ACCENT_SUCCESS, _WHITE),
    ]
    for rv, lbl_text, act_bg, act_fg in PILLS:
        is_active = (rv == risk_filter)
        pill = tk.Label(
            pill_frame, text=lbl_text,
            font=F(8, "bold" if is_active else "normal"),
            fg=act_fg if is_active else _TXT_NAVY,
            bg=act_bg if is_active else _BORDER_LIGHT,
            padx=10, pady=4, relief="flat", cursor="hand2",
        )
        pill.pack(side="left", padx=(0, 4))
        pill.bind("<Button-1>", lambda e, r=rv: on_risk_change(r))

    # ── Name search bar – triggers only on Enter ──────────────────────────────
    search_row = tk.Frame(pad, bg=_CARD_WHITE)
    search_row.pack(fill="x", pady=(0, 6))

    tk.Label(search_row, text="🔍", font=F(10),
             fg=_NAVY_PALE, bg=_CARD_WHITE).pack(side="left", padx=(0, 4))

    search_var = tk.StringVar(value=name_search)
    entry_widget = ctk.CTkEntry(
        search_row, textvariable=search_var,
        placeholder_text="Search client name or ID…  then press Enter",
        width=300, height=26, corner_radius=4,
        fg_color=_WHITE, text_color=_TXT_NAVY,
        border_color=_BORDER_MID, font=FF(9),
    )
    entry_widget.pack(side="left")
    entry_widget.bind("<Return>",
        lambda e: on_name_search(search_var.get().strip().lower()))

    clr = tk.Label(search_row, text=" ✕ ", font=F(8, "bold"),
                   fg=_ACCENT_RED, bg=_CARD_WHITE, cursor="hand2")
    clr.pack(side="left", padx=(2, 0))
    clr.bind("<Button-1>", lambda e: (search_var.set(""), on_name_search("")))

    # ── Apply both filters ────────────────────────────────────────────────────
    filtered = clients_sorted
    if risk_filter:
        filtered = [r for r in filtered
                    if r.get("score_label", "N/A") == risk_filter]
    if name_search:
        filtered = [r for r in filtered
                    if name_search in r.get("client", "").lower()
                    or name_search in str(r.get("client_id", "")).lower()]

    total_clients = len(filtered)
    total_pages   = max(1, (total_clients + _LOANBAL_PAGE_SIZE - 1)
                           // _LOANBAL_PAGE_SIZE)
    page          = max(1, min(page, total_pages))
    start         = (page - 1) * _LOANBAL_PAGE_SIZE
    end           = start + _LOANBAL_PAGE_SIZE
    page_clients  = filtered[start:end]

    # ── Filter description ────────────────────────────────────────────────────
    parts = []
    if risk_filter:
        parts.append(f"{risk_filter} risk")
    if name_search:
        parts.append(f'"{name_search}"')
    filter_note = ("  ·  filter: " + ", ".join(parts)) if parts else ""
    tk.Label(pad,
             text=f"{total_clients} client{'s' if total_clients != 1 else ''}"
                  f"{filter_note}  —  page {page} of {total_pages}",
             font=F(7), fg=_TXT_MUTED, bg=_CARD_WHITE
             ).pack(anchor="e", padx=4, pady=(0, 2))

    # ── Table header ──────────────────────────────────────────────────────────
    cl_tf = _make_table_frame(pad, _CLIENT_COLS)
    _table_header(cl_tf, _CLIENT_COLS)
    _table_divider(cl_tf, 1, len(_CLIENT_COLS), _NAVY_MID)

    grid_row = 2
    for idx, rec in enumerate(page_clients):
        lb      = rec.get("loan_balance") or 0
        pl      = rec.get("principal_loan") or 0
        net     = rec.get("net_income") or 0
        amrt    = rec.get("current_amort") or 0
        pct     = (lb / grand_lb * 100) if grand_lb > 0 else 0.0
        rl      = rec.get("score_label", "N/A")
        sec     = rec.get("sector", "—")
        row_bg  = _CARD_WHITE if idx % 2 == 0 else _OFF_WHITE
        col_clr = _SECTOR_COLORS.get(sec, _NAVY_MID)
        risk_fg = _RISK_COLOR.get(rl, _TXT_SOFT)
        risk_bg = _RISK_BADGE_BG.get(rl, _OFF_WHITE)

        stripe = tk.Frame(cl_tf, bg=row_bg)
        stripe.grid(row=grid_row, column=0,
                    columnspan=len(_CLIENT_COLS), sticky="nsew")
        stripe.lower()

        # Column 0: Client name (left-aligned)
        tk.Label(cl_tf, text=f"  {rec['client'][:30]}", font=F(9, "bold"),
                 fg=_TXT_NAVY, bg=row_bg, anchor="w", padx=8, pady=10
                 ).grid(row=grid_row, column=0, sticky="nsew")
        # Column 1: ID (center)
        tk.Label(cl_tf, text=rec.get("client_id", "—"), font=F(8),
                 fg=_TXT_SOFT, bg=row_bg, anchor="center", padx=4, pady=10
                 ).grid(row=grid_row, column=1, sticky="nsew")
        # Column 2: Sector (center)
        tk.Label(cl_tf,
                 text=f"{_SECTOR_ICON.get(sec, '')} {sec[:22]}",
                 font=F(8), fg=col_clr, bg=row_bg, anchor="center", padx=6, pady=10
                 ).grid(row=grid_row, column=2, sticky="nsew")
        # Column 3: Principal Loan (center)
        tk.Label(cl_tf, text=f"₱{pl:,.2f}" if pl else "—", font=F(9),
                 fg=_TXT_NAVY, bg=row_bg, anchor="center", padx=14, pady=10
                 ).grid(row=grid_row, column=3, sticky="nsew")
        # Column 4: Loan Balance (center)
        tk.Label(cl_tf, text=f"₱{lb:,.2f}", font=F(9, "bold"),
                 fg=_TXT_NAVY, bg=row_bg, anchor="center", padx=14, pady=10
                 ).grid(row=grid_row, column=4, sticky="nsew")

        # Column 5: % of Total (center with bar)
        pct_cell = tk.Frame(cl_tf, bg=row_bg)
        pct_cell.grid(row=grid_row, column=5, sticky="nsew", padx=8, pady=5)
        tk.Label(pct_cell, text=f"{pct:.2f}%", font=F(8, "bold"),
                 fg=col_clr, bg=row_bg, anchor="center").pack(anchor="center", pady=(3, 1))
        bar_outer = tk.Frame(pct_cell, bg=_BORDER_LIGHT, height=4)
        bar_outer.pack(fill="x", pady=(0, 2))
        bar_outer.pack_propagate(False)
        bw = max(2, int(80 * pct / 100))
        tk.Frame(bar_outer, bg=col_clr, height=4, width=bw).place(x=0, y=0, relheight=1)

        # Column 6: Net Income (center)
        tk.Label(cl_tf, text=f"₱{net:,.2f}" if net else "—",
                 font=F(9), fg=_TXT_NAVY, bg=row_bg, anchor="center", padx=14, pady=10
                 ).grid(row=grid_row, column=6, sticky="nsew")
        # Column 7: Current Amort (center)
        tk.Label(cl_tf, text=f"₱{amrt:,.2f}" if amrt else "—",
                 font=F(9), fg=_TXT_NAVY, bg=row_bg, anchor="center", padx=14, pady=10
                 ).grid(row=grid_row, column=7, sticky="nsew")

        # Column 8: Risk badge (center)
        badge_cell = tk.Frame(cl_tf, bg=row_bg)
        badge_cell.grid(row=grid_row, column=8, sticky="nsew", pady=9, padx=10)
        tk.Label(badge_cell, text=rl, font=F(8, "bold"),
                 fg=risk_fg, bg=risk_bg, padx=10, pady=4,
                 highlightbackground=risk_fg, highlightthickness=1
                 ).pack(anchor="center")

        grid_row += 1
        div = tk.Frame(cl_tf, bg=_BORDER_LIGHT, height=1)
        div.grid(row=grid_row, column=0, columnspan=len(_CLIENT_COLS), sticky="ew")
        grid_row += 1

    if not page_clients:
        msg = ("No clients match the current filter."
               if (risk_filter or name_search) else "No clients to display.")
        tk.Label(pad, text=msg, font=F(9), fg=_TXT_MUTED,
                 bg=_CARD_WHITE).pack(pady=20)

    # ── Pagination bar ────────────────────────────────────────────────────────
    if total_pages > 1:
        pg_bar = tk.Frame(pad, bg=_CARD_WHITE)
        pg_bar.pack(pady=(10, 4))

        def _page_btn(parent, text, cmd, enabled):
            cfg = dict(text=text, font=F(8, "bold"), padx=10, pady=4,
                       relief="flat", cursor="hand2" if enabled else "arrow")
            if enabled:
                lbl = tk.Label(parent, fg=_WHITE, bg=_NAVY_MID, **cfg)
                lbl.bind("<Button-1>", lambda e: cmd())
                lbl.bind("<Enter>",
                         lambda e, l=lbl: l.config(bg=_LIME_DARK, fg=_TXT_ON_LIME))
                lbl.bind("<Leave>",
                         lambda e, l=lbl: l.config(bg=_NAVY_MID, fg=_WHITE))
            else:
                lbl = tk.Label(parent, fg=_TXT_MUTED, bg=_BORDER_LIGHT, **cfg)
            lbl.pack(side="left", padx=2)

        _page_btn(pg_bar, "◀◀ First", lambda: on_page_change(1),           page > 1)
        _page_btn(pg_bar, "◀ Prev",   lambda: on_page_change(page - 1),    page > 1)

        half    = 3
        p_start = max(1, page - half)
        p_end   = min(total_pages, page + half)
        if p_start > 1:
            tk.Label(pg_bar, text="…", font=F(8), fg=_TXT_MUTED,
                     bg=_CARD_WHITE, padx=4).pack(side="left")
        for pn in range(p_start, p_end + 1):
            is_cur = (pn == page)
            num_lbl = tk.Label(
                pg_bar, text=str(pn),
                font=F(8, "bold" if is_cur else "normal"),
                fg=_TXT_ON_LIME if is_cur else _TXT_NAVY,
                bg=_LIME_DARK   if is_cur else _BORDER_LIGHT,
                padx=9, pady=4, relief="flat",
                cursor="arrow" if is_cur else "hand2",
            )
            num_lbl.pack(side="left", padx=2)
            if not is_cur:
                num_lbl.bind("<Button-1>", lambda e, p=pn: on_page_change(p))
                num_lbl.bind("<Enter>",
                             lambda e, l=num_lbl: l.config(bg=_NAVY_GHOST))
                num_lbl.bind("<Leave>",
                             lambda e, l=num_lbl: l.config(bg=_BORDER_LIGHT))
        if p_end < total_pages:
            tk.Label(pg_bar, text="…", font=F(8), fg=_TXT_MUTED,
                     bg=_CARD_WHITE, padx=4).pack(side="left")

        _page_btn(pg_bar, "Next ▶",  lambda: on_page_change(page + 1),     page < total_pages)
        _page_btn(pg_bar, "Last ▶▶", lambda: on_page_change(total_pages),   page < total_pages)

        end_display = min(end, total_clients)
        tk.Label(pg_bar,
                 text=f"  rows {start + 1}–{end_display} of {total_clients}",
                 font=F(7), fg=_TXT_MUTED, bg=_CARD_WHITE
                 ).pack(side="left", padx=(8, 0))

        # ── Scroll to top unconditionally ─────────────────────────────────────
        canvas.after(50, lambda: canvas.yview_moveto(0))


# ══════════════════════════════════════════════════════════════════════════════
#  PATCHED _loanbal_render
# ══════════════════════════════════════════════════════════════════════════════

def _loanbal_render_patched(self):
    for w in self._loanbal_body.winfo_children():
        w.destroy()
    plt.close("all")

    all_data   = _lu_get_filtered_all_data(self)
    general    = all_data.get("general", [])
    sector_map = all_data.get("sector_map", {})
    totals     = all_data.get("totals", {})
    grand_lb   = totals.get("loan_balance", 0) or 0

    active_sectors = _lu_get_active_sectors(self)
    prev_sectors   = getattr(self, "_loanbal_prev_sectors", _LOANBAL_SENTINEL)
    if prev_sectors is _LOANBAL_SENTINEL or prev_sectors != active_sectors:
        self._loanbal_page        = 1
        self._loanbal_risk        = None
        self._loanbal_name_search = ""
    self._loanbal_prev_sectors = active_sectors

    if not hasattr(self, "_loanbal_page"):        self._loanbal_page        = 1
    if not hasattr(self, "_loanbal_risk"):        self._loanbal_risk        = None
    if not hasattr(self, "_loanbal_name_search"): self._loanbal_name_search = ""

    if active_sectors:
        sector_text = " · ".join(active_sectors)
        self._loanbal_hdr_lbl.config(
            text=f"📊  Loan Balance — Filtered: {sector_text}", fg=_LIME_MID)
    else:
        self._loanbal_hdr_lbl.config(
            text="📊  Sector vs Total Loan Balance  —  Exposure Analysis",
            fg=_WHITE)

    btn = getattr(self, "_loanbal_export_btn", None)
    if btn:
        btn.configure(state="normal" if general else "disabled")

    if not general:
        tk.Label(self._loanbal_body, text="No data available for this filter.",
                 font=F(10), fg=_TXT_MUTED, bg=_CARD_WHITE).pack(pady=60)
        return

    _, inner, canvas_scroll = _make_scrollable(self._loanbal_body, _CARD_WHITE)
    pad = tk.Frame(inner, bg=_CARD_WHITE)
    pad.pack(fill="both", expand=True, padx=24, pady=16)

    # ── Grand total card ──────────────────────────────────────────────────────
    grand_card = tk.Frame(pad, bg=_NAVY_DEEP,
                          highlightbackground=_NAVY_MID, highlightthickness=1)
    grand_card.pack(fill="x", pady=(0, 16))
    gc_inner = tk.Frame(grand_card, bg=_NAVY_DEEP)
    gc_inner.pack(fill="both", padx=22, pady=16)

    # Left — main balance figure
    left_gc = tk.Frame(gc_inner, bg=_NAVY_DEEP)
    left_gc.pack(side="left", fill="y")

    lbl_icon_text = (
        f"💰  FILTERED LOAN BALANCE  ·  {' · '.join(active_sectors)}"
        if active_sectors else "💰  GRAND TOTAL LOAN BALANCE"
    )
    lbl_icon_fg = _LIME_MID if active_sectors else _TXT_MUTED
    tk.Label(left_gc, text=lbl_icon_text,
             font=F(9, "bold"), fg=lbl_icon_fg, bg=_NAVY_DEEP).pack(anchor="w")

    tk.Label(left_gc, text=f"₱{grand_lb:,.2f}",
             font=F(22, "bold"), fg=_LIME_MID, bg=_NAVY_DEEP).pack(anchor="w")

    tk.Label(left_gc,
             text=f"{len(general)} clients  ·  {len(sector_map)} sectors detected",
             font=F(9), fg=_TXT_SOFT, bg=_NAVY_DEEP).pack(anchor="w")

    if active_sectors:
        tk.Label(left_gc, text="⚠  Export uses the full unfiltered dataset",
                 font=F(7), fg=_ACCENT_GOLD, bg=_NAVY_DEEP).pack(anchor="w", pady=(4, 0))

    # Right — secondary stats
    right_gc = tk.Frame(gc_inner, bg=_NAVY_DEEP)
    right_gc.pack(side="right", anchor="e", pady=(4, 0))
    total_net = sum(r.get("net_income") or 0 for r in general)
    avg_loan  = (grand_lb / len(general)) if general else 0.0
    for lbl_text, val_text in [
        ("Total Net Income",    f"₱{total_net:,.2f}"),
        ("Avg Loan per Client", f"₱{avg_loan:,.2f}"),
    ]:
        row_f = tk.Frame(right_gc, bg=_NAVY_DEEP)
        row_f.pack(anchor="e", pady=3)
        tk.Label(row_f, text=lbl_text, font=F(8),
                 fg=_TXT_MUTED, bg=_NAVY_DEEP).pack(side="left", padx=(0, 12))
        tk.Label(row_f, text=val_text, font=F(11, "bold"),
                 fg=_WHITE, bg=_NAVY_DEEP).pack(side="left")

    # ── Per-sector breakdown table ────────────────────────────────────────────
    tk.Label(pad, text="Sector Loan Balance Breakdown",
             font=F(11, "bold"), fg=_TXT_NAVY, bg=_CARD_WHITE).pack(anchor="w", pady=(0, 8))

    from lu_ui import _make_table_frame as _mtf, _table_header as _th, _table_divider as _td
    sector_tf = _mtf(pad, _SECTOR_COLS)
    _th(sector_tf, _SECTOR_COLS)
    _td(sector_tf, 1, len(_SECTOR_COLS), _NAVY_MID)

    all_sectors = [s for s in _CHART_SECTORS if s in sector_map]
    if SECTOR_OTHER in sector_map:
        all_sectors.append(SECTOR_OTHER)

    sector_rows_data = []
    for sector in all_sectors:
        recs    = sector_map.get(sector, [])
        n       = len(recs)
        s_lb    = sum(r.get("loan_balance") or 0 for r in recs)
        s_net   = sum(r.get("net_income")   or 0 for r in recs)
        pct     = (s_lb / grand_lb * 100) if grand_lb > 0 else 0.0
        avg_lb  = s_lb / n  if n > 0 else 0.0
        avg_net = s_net / n if n > 0 else 0.0
        all_exp = [e for r in recs for e in r.get("expenses", [])]
        _, risk_label, _, _ = _compute_risk_score(all_exp)
        sector_rows_data.append((sector, n, s_lb, pct, avg_lb, avg_net, risk_label))

    sector_rows_data.sort(key=lambda x: -x[2])

    grid_row = 2
    for idx, (sector, n, s_lb, pct, avg_lb, avg_net, risk_label) in \
            enumerate(sector_rows_data):
        row_bg    = _CARD_WHITE if idx % 2 == 0 else _OFF_WHITE
        col_color = _SECTOR_COLORS.get(sector, _NAVY_MID)
        icon      = _SECTOR_ICON.get(sector, "📋")
        risk_fg   = _RISK_COLOR.get(risk_label, _TXT_SOFT)
        risk_bg   = _RISK_BADGE_BG.get(risk_label, _OFF_WHITE)

        if active_sectors and sector in active_sectors:
            row_bg = "#0E2040"
        name_fg = _LIME_MID if (active_sectors and sector in active_sectors) else col_color
        val_fg  = _WHITE    if (active_sectors and sector in active_sectors) else _TXT_NAVY

        stripe = tk.Frame(sector_tf, bg=row_bg)
        stripe.grid(row=grid_row, column=0, columnspan=len(_SECTOR_COLS), sticky="nsew")
        stripe.lower()

        # Col 0: Sector name — left-aligned
        tk.Label(sector_tf, text=f"  {icon}  {sector}", font=F(9, "bold"),
                 fg=name_fg, bg=row_bg, anchor="w", padx=8, pady=11
                 ).grid(row=grid_row, column=0, sticky="nsew")
        # Col 1: # Clients — center
        tk.Label(sector_tf, text=str(n), font=F(9),
                 fg=val_fg, bg=row_bg, anchor="center", padx=6, pady=11
                 ).grid(row=grid_row, column=1, sticky="nsew")
        # Col 2: Total Loan Balance — center
        tk.Label(sector_tf, text=f"₱{s_lb:,.2f}", font=F(9, "bold"),
                 fg=val_fg, bg=row_bg, anchor="center", padx=14, pady=11
                 ).grid(row=grid_row, column=2, sticky="nsew")
        # Col 3: % with mini bar — center
        pct_cell = tk.Frame(sector_tf, bg=row_bg)
        pct_cell.grid(row=grid_row, column=3, sticky="nsew", padx=8, pady=6)
        tk.Label(pct_cell, text=f"{pct:.1f}%", font=F(9, "bold"),
                 fg=name_fg, bg=row_bg, anchor="center").pack(anchor="center", pady=(3, 1))
        bar_outer = tk.Frame(pct_cell, bg=_BORDER_LIGHT, height=5)
        bar_outer.pack(fill="x", pady=(0, 2))
        bar_outer.pack_propagate(False)
        fill_w = max(3, int(100 * pct / 100))
        tk.Frame(bar_outer, bg=col_color, height=5, width=fill_w).place(x=0, y=0, relheight=1)
        # Col 4: Avg Loan per Client — center
        tk.Label(sector_tf, text=f"₱{avg_lb:,.2f}", font=F(9),
                 fg=val_fg, bg=row_bg, anchor="center", padx=14, pady=11
                 ).grid(row=grid_row, column=4, sticky="nsew")
        # Col 5: Avg Net Income — center
        tk.Label(sector_tf, text=f"₱{avg_net:,.2f}" if avg_net else "—",
                 font=F(9), fg=val_fg, bg=row_bg, anchor="center", padx=14, pady=11
                 ).grid(row=grid_row, column=5, sticky="nsew")
        # Col 6: Risk badge — center
        badge_cell = tk.Frame(sector_tf, bg=row_bg)
        badge_cell.grid(row=grid_row, column=6, sticky="nsew", pady=9, padx=10)
        tk.Label(badge_cell, text=risk_label, font=F(8, "bold"),
                 fg=risk_fg, bg=risk_bg, padx=12, pady=5,
                 highlightbackground=risk_fg, highlightthickness=1
                 ).pack(anchor="center")

        grid_row += 1
        div = tk.Frame(sector_tf, bg=_BORDER_LIGHT, height=1)
        div.grid(row=grid_row, column=0, columnspan=len(_SECTOR_COLS), sticky="ew")
        grid_row += 1

    # ── Pie + Bar chart ───────────────────────────────────────────────────────
    if _HAS_MPL and sector_rows_data:
        tk.Label(pad, text="Loan Balance Share by Sector",
                 font=F(11, "bold"), fg=_TXT_NAVY, bg=_CARD_WHITE
                 ).pack(anchor="w", pady=(20, 8))
        fig_pie = None
        try:
            fig_pie, (ax_pie, ax_bar) = plt.subplots(1, 2, figsize=(10, 4.5))
            fig_pie.patch.set_facecolor(_MPL_BG)
            for ax in (ax_pie, ax_bar):
                ax.set_facecolor(_MPL_BG)

            valid = [(s, lb, pct) for s, n, lb, pct, *_ in sector_rows_data if lb > 0]
            if valid:
                snames  = [x[0] for x in valid]
                svals   = [x[1] for x in valid]
                spcts   = [x[2] for x in valid]
                scolors = [_SECTOR_COLORS.get(s, _MPL_NAVY) for s in snames]
                labels = [f"{s}\n{p:.1f}%"
                           for s, p in zip(snames, spcts)]
                wedges, _ = ax_pie.pie(
                    svals, colors=scolors, startangle=90,
                    wedgeprops=dict(width=0.55, edgecolor=_MPL_BG, linewidth=2))
                ax_pie.legend(wedges, labels, loc="lower center", fontsize=7,
                              frameon=False, ncol=2, bbox_to_anchor=(0.5, -0.18))
                total_str = (f"₱{grand_lb/1e6:.2f}M"
                             if grand_lb >= 1e6 else f"₱{grand_lb:,.0f}")
                ax_pie.text(0, 0.1, total_str, ha="center", va="center",
                            fontsize=10, fontweight="bold", color=_MPL_NAVY)
                ax_pie.text(0, -0.15, "total", ha="center", va="center",
                            fontsize=8, color="#6B7FA3")
                ax_pie.set_title("Loan Balance Share", fontsize=10,
                                 color=_MPL_NAVY, fontweight="bold", pad=8)
                short_names = [s[:22] for s in snames]
                bars = ax_bar.barh(short_names, svals, color=scolors,
                                   edgecolor=_MPL_BG, linewidth=1.2, height=0.6)
                ax_bar.invert_yaxis()
                ax_bar.xaxis.set_major_formatter(
                    matplotlib.ticker.FuncFormatter(
                        lambda x, _: f"₱{x/1e6:.1f}M" if x >= 1e6 else f"₱{x:,.0f}"))
                ax_bar.tick_params(axis="x", labelsize=7)
                ax_bar.tick_params(axis="y", labelsize=8)
                ax_bar.spines[["top", "right"]].set_visible(False)
                max_v = max(svals) if svals else 1
                for bar, val, pct in zip(bars, svals, spcts):
                    ax_bar.text(val + max_v * 0.01,
                                bar.get_y() + bar.get_height() / 2,
                                f"{pct:.1f}%", va="center",
                                fontsize=8, fontweight="bold", color=_MPL_NAVY)
                ax_bar.set_title("Loan Balance by Sector", fontsize=10,
                                 color=_MPL_NAVY, fontweight="bold", pad=8)
            fig_pie.tight_layout(pad=1.5)
            c_pie = tk.Frame(pad, bg=_WHITE,
                             highlightbackground=_BORDER_MID, highlightthickness=1)
            c_pie.pack(fill="x", pady=(0, 16))
            FigureCanvasTkAgg(fig_pie, master=c_pie).get_tk_widget().pack(
                fill="both", expand=True, padx=4, pady=4)
        except Exception:
            pass
        finally:
            if fig_pie:
                plt.close(fig_pie)

    # ── Per-client table — paginated + risk-filtered + name-searched ──────────
    clients_sorted = sorted(general, key=lambda r: -(r.get("loan_balance") or 0))

    client_section = tk.Frame(pad, bg=_CARD_WHITE)
    client_section.pack(fill="x")

    def _rebuild(new_page, new_risk, new_name):
        self._loanbal_page        = new_page
        self._loanbal_risk        = new_risk
        self._loanbal_name_search = new_name
        for w in client_section.winfo_children():
            w.destroy()
        _loanbal_render_client_table(
            client_section, clients_sorted, grand_lb,
            new_page, new_risk, new_name, canvas_scroll,
            on_page_change=lambda p: _rebuild(p, self._loanbal_risk,
                                              self._loanbal_name_search),
            on_risk_change=lambda r: _rebuild(1, r,
                                              self._loanbal_name_search),
            on_name_search=lambda q: _rebuild(1, self._loanbal_risk, q),
        )

    _rebuild(self._loanbal_page, self._loanbal_risk, self._loanbal_name_search)
    plt.close("all")


# ══════════════════════════════════════════════════════════════════════════════
#  EXPORT MENU POPUP
# ══════════════════════════════════════════════════════════════════════════════

def _loanbal_show_export_menu(self):
    risk   = getattr(self, "_loanbal_risk", None)
    name   = getattr(self, "_loanbal_name_search", "")
    sector = _lu_get_active_sectors(self)

    filter_parts = []
    if sector:
        filter_parts.append(" · ".join(sector))
    if risk:
        filter_parts.append(risk)
    if name:
        filter_parts.append(f'"{name}"')
    has_filter = bool(filter_parts)
    filter_desc = "  [" + ", ".join(filter_parts) + "]" if has_filter else ""

    menu = tk.Menu(
        self._loanbal_body, tearoff=0,
        font=F(9), bg=_WHITE, fg=_TXT_NAVY,
        activebackground=_NAVY_GHOST, activeforeground=_NAVY_DEEP,
        relief="flat", bd=1,
    )
    menu.add_command(
        label="📄  Export PDF — Loan Balance Report",
        command=lambda: _loanbal_export_pdf(self),
    )
    menu.add_separator()
    if has_filter:
        menu.add_command(
            label=f"📊  Export Excel — Filtered{filter_desc}",
            command=lambda: _loanbal_export_excel(self, filtered=True),
        )
    menu.add_command(
        label="📊  Export Excel — Full Workbook (all clients)",
        command=lambda: _loanbal_export_excel(self, filtered=False),
    )
    btn = self._loanbal_export_btn
    try:
        menu.tk_popup(
            btn.winfo_rootx(),
            btn.winfo_rooty() + btn.winfo_height(),
        )
    finally:
        menu.grab_release()


# ══════════════════════════════════════════════════════════════════════════════
#  PDF EXPORT
# ══════════════════════════════════════════════════════════════════════════════

def _loanbal_export_pdf(self):
    if not self._lu_all_data:
        messagebox.showwarning("No Data", "Run an analysis first.")
        return
    if not _HAS_RL:
        messagebox.showerror(
            "Missing Library",
            "reportlab is not installed.\nRun:  pip install reportlab")
        return

    default_name = (
        f"LU_LoanBalance_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
    path = filedialog.asksaveasfilename(
        title="Save Loan Balance PDF",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        initialfile=default_name,
    )
    if not path:
        return

    try:
        _generate_loanbal_pdf(self._lu_all_data,
                              path,
                              filepath=self._lu_filepath or "")
        messagebox.showinfo("Export Complete", f"PDF saved to:\n{path}")
    except Exception as ex:
        messagebox.showerror("PDF Export Error", str(ex))


def _generate_loanbal_pdf(all_data, out_path, filepath=""):
    from reportlab.lib.pagesizes import A4, landscape as rl_landscape

    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(
        out_path,
        pagesize=rl_landscape(A4),
        leftMargin=1.5 * cm, rightMargin=1.5 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
    )

    navy    = rl_colors.HexColor("#1A3A6B")
    lime    = rl_colors.HexColor("#5A9E28")
    white   = rl_colors.white
    off     = rl_colors.HexColor("#F5F7FA")
    border  = rl_colors.HexColor("#C5D0E8")
    mist    = rl_colors.HexColor("#EEF3FB")
    dark    = rl_colors.HexColor("#0A1628")

    RISK_C = {
        "HIGH":     rl_colors.HexColor("#E53E3E"),
        "MODERATE": rl_colors.HexColor("#D4A017"),
        "LOW":      rl_colors.HexColor("#2E7D32"),
        "N/A":      rl_colors.HexColor("#9AAACE"),
    }
    SEC_C = {
        SECTOR_WHOLESALE:   rl_colors.HexColor("#1A3A6B"),
        SECTOR_AGRICULTURE: rl_colors.HexColor("#2E7D32"),
        SECTOR_TRANSPORT:   rl_colors.HexColor("#D4A017"),
        SECTOR_REMITTANCE:  rl_colors.HexColor("#8B5CF6"),
        SECTOR_CONSUMER:    rl_colors.HexColor("#E53E3E"),
        SECTOR_OTHER:       rl_colors.HexColor("#9AAACE"),
    }

    title_s = ParagraphStyle("LBTitle", parent=styles["Title"],
                              fontSize=16, textColor=navy, spaceAfter=4)
    sub_s   = ParagraphStyle("LBSub",   parent=styles["Normal"],
                              fontSize=9,  textColor=rl_colors.HexColor("#6B7FA3"), leading=12)
    body_s  = ParagraphStyle("LBBody",  parent=styles["Normal"],
                              fontSize=8,  textColor=rl_colors.HexColor("#1A2B4A"), leading=11)
    hdr_s   = ParagraphStyle("LBHdr",   parent=styles["Normal"],
                              fontSize=8,  textColor=white, leading=11)
    bold_s  = ParagraphStyle("LBBold",  parent=styles["Normal"],
                              fontSize=9,  textColor=navy, leading=12)

    general    = all_data.get("general", [])
    sector_map = all_data.get("sector_map", {})
    totals     = all_data.get("totals", {})
    grand_lb   = totals.get("loan_balance", 0) or 0
    fname      = Path(filepath).name if filepath else "—"
    now        = datetime.now().strftime("%B %d, %Y  %H:%M")

    story = [
        Paragraph("LU Analysis — Sector vs Loan Balance Report", title_s),
        Paragraph(f"File: {fname}    Generated: {now}", sub_s),
        Spacer(1, 0.3 * cm),
        HRFlowable(width="100%", thickness=1, color=border),
        Spacer(1, 0.3 * cm),
    ]

    grand_str = f"₱{grand_lb:,.2f}"
    story.append(Paragraph(
        f"<b>Grand Total Loan Balance: {grand_str}</b>   |   "
        f"{len(general)} clients  ·  {len(sector_map)} sectors",
        bold_s))
    story.append(Spacer(1, 0.4 * cm))

    story.append(Paragraph("<b>Sector Loan Balance Breakdown</b>", bold_s))
    story.append(Spacer(1, 0.15 * cm))

    sec_hdr = [Paragraph(f"<b>{h}</b>", hdr_s) for h in [
        "Sector", "# Clients", "Total Loan Balance",
        "% of Total", "Avg Loan/Client", "Avg Net Income", "Risk Profile",
    ]]

    all_sectors = [s for s in _CHART_SECTORS if s in sector_map]
    if SECTOR_OTHER in sector_map:
        all_sectors.append(SECTOR_OTHER)

    sector_rows_data = []
    for sector in all_sectors:
        recs    = sector_map.get(sector, [])
        n       = len(recs)
        s_lb    = sum(r.get("loan_balance") or 0 for r in recs)
        s_net   = sum(r.get("net_income")   or 0 for r in recs)
        pct     = (s_lb / grand_lb * 100) if grand_lb > 0 else 0.0
        avg_lb  = s_lb / n  if n > 0 else 0.0
        avg_net = s_net / n if n > 0 else 0.0
        all_exp = [e for r in recs for e in r.get("expenses", [])]
        _, risk_label, _, _ = _compute_risk_score(all_exp)
        sector_rows_data.append(
            (sector, n, s_lb, pct, avg_lb, avg_net, risk_label))
    sector_rows_data.sort(key=lambda x: -x[2])

    sec_tbl_data = [sec_hdr]
    sec_tbl_style = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, -1), 8),
        ("LEADING",    (0, 0), (-1, -1), 11),
        ("BOX",        (0, 0), (-1, -1), 0.5, border),
        ("INNERGRID",  (0, 0), (-1, -1), 0.3, border),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("VALIGN",    (0, 0), (-1, -1), "MIDDLE"),
    ]
    for i, (sector, n, s_lb, pct, avg_lb, avg_net, risk_label) in \
            enumerate(sector_rows_data, start=1):
        row_c = SEC_C.get(sector, rl_colors.HexColor("#9AAACE"))
        icon  = _SECTOR_ICON.get(sector, "")
        rc    = RISK_C.get(risk_label, rl_colors.HexColor("#2E7D32"))
        sec_tbl_data.append([
            Paragraph(f"<font color='#{_rl_hex(row_c)}'><b>{icon} {sector}</b></font>", body_s),
            Paragraph(str(n), body_s),
            Paragraph(f"<b>₱{s_lb:,.2f}</b>", body_s),
            Paragraph(f"{pct:.1f}%", body_s),
            Paragraph(f"₱{avg_lb:,.2f}", body_s),
            Paragraph(f"₱{avg_net:,.2f}" if avg_net else "—", body_s),
            Paragraph(
                f"<font color='#{_rl_hex(rc)}'><b>{risk_label}</b></font>",
                body_s),
        ])
        bg = mist if i % 2 == 0 else white
        sec_tbl_style.append(("BACKGROUND", (0, i), (-1, i), bg))

    sec_col_w = [4.8*cm, 1.8*cm, 4.0*cm, 2.2*cm, 4.0*cm, 3.5*cm, 2.5*cm]
    sec_tbl = Table(sec_tbl_data, colWidths=sec_col_w, repeatRows=1)
    sec_tbl.setStyle(TableStyle(sec_tbl_style))
    story += [sec_tbl, Spacer(1, 0.5 * cm)]

    story.append(PageBreak())
    story.append(Paragraph("<b>Individual Client Loan Balance</b>", bold_s))
    story.append(Spacer(1, 0.15 * cm))

    cl_hdr_row = [Paragraph(f"<b>{h}</b>", hdr_s) for h in [
        "Client", "ID", "Sector", "Loan Balance", "% of Total",
        "Net Income", "Current Amort", "Risk",
    ]]
    cl_tbl_data  = [cl_hdr_row]
    cl_tbl_style = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, -1), 7),
        ("LEADING",    (0, 0), (-1, -1), 10),
        ("BOX",        (0, 0), (-1, -1), 0.5, border),
        ("INNERGRID",  (0, 0), (-1, -1), 0.3, border),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("VALIGN",    (0, 0), (-1, -1), "MIDDLE"),
    ]

    clients_sorted = sorted(general, key=lambda r: -(r.get("loan_balance") or 0))
    for i, rec in enumerate(clients_sorted, start=1):
        lb   = rec.get("loan_balance") or 0
        net  = rec.get("net_income")   or 0
        amrt = rec.get("current_amort") or 0
        pct  = (lb / grand_lb * 100) if grand_lb > 0 else 0.0
        rl   = rec.get("score_label", "N/A")
        sec  = rec.get("sector", "—")
        rc   = RISK_C.get(rl, rl_colors.HexColor("#9AAACE"))
        icon = _SECTOR_ICON.get(sec, "")
        cl_tbl_data.append([
            Paragraph(rec["client"][:32], body_s),
            Paragraph(rec.get("client_id", "—"), body_s),
            Paragraph(f"{icon} {sec[:20]}", body_s),
            Paragraph(f"₱{lb:,.2f}", body_s),
            Paragraph(f"{pct:.2f}%", body_s),
            Paragraph(f"₱{net:,.2f}" if net else "—", body_s),
            Paragraph(f"₱{amrt:,.2f}" if amrt else "—", body_s),
            Paragraph(
                f"<font color='#{_rl_hex(rc)}'><b>{rl}</b></font>",
                body_s),
        ])
        bg = mist if i % 2 == 0 else white
        cl_tbl_style.append(("BACKGROUND", (0, i), (-1, i), bg))

    cl_col_w = [4.5*cm, 1.8*cm, 3.5*cm, 3.2*cm, 2.2*cm, 3.2*cm, 3.2*cm, 2.0*cm]
    cl_tbl = Table(cl_tbl_data, colWidths=cl_col_w, repeatRows=1)
    cl_tbl.setStyle(TableStyle(cl_tbl_style))
    story.append(cl_tbl)

    doc.build(story)


def _rl_hex(color) -> str:
    try:
        return f"{int(color.red*255):02X}{int(color.green*255):02X}{int(color.blue*255):02X}"
    except Exception:
        return "000000"


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════

def _loanbal_export_excel(self, filtered=False):
    if not self._lu_all_data:
        messagebox.showwarning("No Data", "Run an analysis first.")
        return
    if not _HAS_OPENPYXL:
        messagebox.showerror(
            "Missing Library",
            "openpyxl is not installed.\nRun:  pip install openpyxl")
        return

    if filtered:
        risk        = getattr(self, "_loanbal_risk", None)
        name_search = getattr(self, "_loanbal_name_search", "").lower()
        base_data   = _lu_get_filtered_all_data(self)
        general_all = base_data.get("general", [])

        export_clients = sorted(general_all,
                                key=lambda r: -(r.get("loan_balance") or 0))
        if risk:
            export_clients = [r for r in export_clients
                              if r.get("score_label", "N/A") == risk]
        if name_search:
            export_clients = [r for r in export_clients
                              if name_search in r.get("client", "").lower()
                              or name_search in str(r.get("client_id", "")).lower()]

        filter_parts = []
        active_sec = _lu_get_active_sectors(self)
        if active_sec:
            filter_parts.append("Sector: " + " · ".join(active_sec))
        if risk:
            filter_parts.append(f"Risk: {risk}")
        if name_search:
            filter_parts.append(f'Name: "{name_search}"')
        filter_label = "  |  Filter: " + ", ".join(filter_parts) if filter_parts else ""
    else:
        export_clients = None
        filter_label   = ""

    suffix = "_filtered" if filtered and export_clients is not None else "_full"
    default_name = f"LU_LoanBalance{suffix}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    path = filedialog.asksaveasfilename(
        title="Save Loan Balance Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=default_name,
    )
    if not path:
        return

    try:
        _generate_loanbal_excel(
            self._lu_all_data,
            path,
            filepath=self._lu_filepath or "",
            export_clients=export_clients,
            filter_label=filter_label,
        )
        n = len(export_clients) if export_clients is not None else \
            len(self._lu_all_data.get("general", []))
        messagebox.showinfo("Export Complete",
                            f"Excel saved to:\n{path}\n({n} client(s) exported)")
    except Exception as ex:
        messagebox.showerror("Excel Export Error", str(ex))


def _generate_loanbal_excel(all_data, out_path, filepath="",
                            export_clients=None, filter_label=""):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    general    = all_data.get("general", [])
    sector_map = all_data.get("sector_map", {})
    totals     = all_data.get("totals", {})
    grand_lb   = totals.get("loan_balance", 0) or 0
    fname      = Path(filepath).name if filepath else "—"
    now        = datetime.now().strftime("%Y-%m-%d %H:%M")

    clients_to_export = (export_clients if export_clients is not None
                         else sorted(general, key=lambda r: -(r.get("loan_balance") or 0)))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def fill(hex_col):
        return PatternFill("solid", fgColor=hex_col.lstrip("#"))

    def thin_border():
        s = Side(style="thin", color="C5D0E8")
        return Border(left=s, right=s, top=s, bottom=s)

    FILLS = {
        "navy":     fill("#1A3A6B"),
        "mist":     fill("#EEF3FB"),
        "white":    fill("#FFFFFF"),
        "cd_row":   fill("#FFFBF0"),
    }
    RISK_FC = {"HIGH": "E53E3E", "MODERATE": "D4A017", "LOW": "2E7D32", "N/A": "9AAACE"}
    SEC_FC  = {
        SECTOR_WHOLESALE:   "1A3A6B",
        SECTOR_AGRICULTURE: "2E7D32",
        SECTOR_TRANSPORT:   "D4A017",
        SECTOR_REMITTANCE:  "8B5CF6",
        SECTOR_CONSUMER:    "E53E3E",
        SECTOR_OTHER:       "9AAACE",
    }
    NUM_FMT = "#,##0.00"

    # ── Sheet 1: Sector Summary ───────────────────────────────────────────────
    ws1 = wb.create_sheet("Sector Summary")

    ws1.merge_cells("A1:G1")
    c = ws1["A1"]
    c.value     = "LU Analysis — Sector vs Loan Balance"
    c.font      = Font(bold=True, size=13, color="0A1628")
    c.fill      = FILLS["mist"]
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[1].height = 22.05

    ws1.merge_cells("A2:G2")
    c = ws1["A2"]
    c.value     = (f"File: {fname}    Generated: {now}    "
                   f"Grand Total Loan Balance: ₱{grand_lb:,.2f}    "
                   f"Clients: {len(general)}{filter_label}")
    c.font      = Font(size=8, color="6B7FA3")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[2].height = 16.05

    hdr_cols_s = [
        ("Sector",             22),
        ("# Clients",          10),
        ("Total Loan Balance",  22),
        ("% of Total",          12),
        ("Avg Loan / Client",   22),
        ("Avg Net Income",      20),
        ("Risk Profile",        14),
    ]
    for ci, (hdr, w) in enumerate(hdr_cols_s, 1):
        ws1.column_dimensions[get_column_letter(ci)].width = w
        c = ws1.cell(3, ci, hdr)
        c.fill      = FILLS["navy"]
        c.font      = Font(bold=True, color="FFFFFF", size=9)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = thin_border()
    ws1.row_dimensions[3].height = 18.0
    ws1.freeze_panes = "A2"

    all_sectors = [s for s in _CHART_SECTORS if s in sector_map]
    if SECTOR_OTHER in sector_map and SECTOR_OTHER not in all_sectors:
        all_sectors.append(SECTOR_OTHER)

    sector_rows_data = []
    for sector in all_sectors:
        recs    = sector_map.get(sector, [])
        n       = len(recs)
        s_lb    = sum(r.get("loan_balance") or 0 for r in recs)
        s_net   = sum(r.get("net_income")   or 0 for r in recs)
        pct     = (s_lb / grand_lb * 100) if grand_lb > 0 else 0.0
        avg_lb  = s_lb / n  if n > 0 else 0.0
        avg_net = s_net / n if n > 0 else 0.0
        all_exp = [e for r in recs for e in r.get("expenses", [])]
        _, risk_label, _, _ = _compute_risk_score(all_exp)
        sector_rows_data.append((sector, n, s_lb, pct, avg_lb, avg_net, risk_label))
    sector_rows_data.sort(key=lambda x: -x[2])

    for idx, (sector, n, s_lb, pct, avg_lb, avg_net, risk_label) in \
            enumerate(sector_rows_data):
        ri      = 4 + idx
        icon    = _SECTOR_ICON.get(sector, "")
        sec_fc  = SEC_FC.get(sector, "1A2B4A")
        risk_fc = RISK_FC.get(risk_label, "9AAACE")
        row_fill = FILLS["mist"] if idx % 2 == 0 else FILLS["white"]

        row_def = [
            (f"{icon} {sector}",            sec_fc,  True,  "left",   None),
            (n,                             "1A2B4A", False, "center", "0"),
            (s_lb,                          "1A2B4A", True,  "center", NUM_FMT),
            (f"{pct:.1f}%",                 sec_fc,  True,  "center", None),
            (avg_lb,                        "1A2B4A", False, "center", NUM_FMT),
            (avg_net if avg_net else "—",   "1A2B4A", False, "center",
             NUM_FMT if avg_net else None),
            (risk_label,                    risk_fc,  True,  "center", None),
        ]
        for ci, (val, fc, bold, align, fmt) in enumerate(row_def, 1):
            c = ws1.cell(ri, ci, val)
            c.fill      = row_fill
            c.font      = Font(size=9, bold=bold, color=fc)
            c.border    = thin_border()
            c.alignment = Alignment(horizontal=align, vertical="center", indent=1)
            if fmt:
                c.number_format = fmt
        ws1.row_dimensions[ri].height = 16.05

    # ── Sheet 2: Client Detail ────────────────────────────────────────────────
    ws2 = wb.create_sheet("Client Detail")
    ws2.freeze_panes = "A2"

    hdr_cols_c = [
        ("Client",         28),
        ("Client ID",      14),
        ("PN",             12),
        ("Sector",         22),
        ("Loan Balance",   20),
        ("% of Total",     14),
        ("Net Income",     18),
        ("Current Amort",  18),
        ("Amort History",  18),
        ("Total Source",   18),
        ("Risk Label",     14),
        ("Risk Score",     12),
    ]
    for ci, (hdr, w) in enumerate(hdr_cols_c, 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
        c = ws2.cell(1, ci, hdr)
        c.fill      = FILLS["navy"]
        c.font      = Font(bold=True, color="FFFFFF", size=9)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = thin_border()
    ws2.row_dimensions[1].height = 18.0

    for ri, rec in enumerate(clients_to_export, start=2):
        lb   = rec.get("loan_balance") or 0
        net  = rec.get("net_income")   or 0
        amrt = rec.get("current_amort") or 0
        hist = rec.get("amort_history") or 0
        src  = rec.get("total_source")  or 0
        pct  = (lb / grand_lb * 100) if grand_lb > 0 else 0.0
        rl   = rec.get("score_label", "N/A")
        sec  = rec.get("sector", "—")
        icon = _SECTOR_ICON.get(sec, "")
        risk_fc = RISK_FC.get(rl, "9AAACE")
        sec_fc  = SEC_FC.get(sec, "1A2B4A")
        row_fill = FILLS["cd_row"]

        row_def = [
            (rec.get("client", ""),          "1A2B4A", False, "left",   None),
            (str(rec.get("client_id", "")),  "6B7FA3", False, "center", None),
            (str(rec.get("pn", "")),         "6B7FA3", False, "center", None),
            (f"{icon} {sec}",                sec_fc,  False, "center", None),
            (lb,                             "1A2B4A", True,  "center", NUM_FMT),
            (f"{pct:.2f}%",                  sec_fc,  True,  "center", None),
            (net  if net  else "—",          "1A2B4A", False, "center",
             NUM_FMT if net  else None),
            (amrt if amrt else "—",          "1A2B4A", False, "center",
             NUM_FMT if amrt else None),
            (hist if hist else "—",          "1A2B4A", False, "center",
             NUM_FMT if hist else None),
            (src  if src  else "—",          "1A2B4A", False, "center",
             NUM_FMT if src  else None),
            (rl,                             risk_fc,  True,  "center", None),
            (rec.get("score", 0),            "1A2B4A", False, "center", "0.0"),
        ]
        for ci, (val, fc, bold, align, fmt) in enumerate(row_def, 1):
            c = ws2.cell(ri, ci, val)
            c.fill      = row_fill
            c.font      = Font(size=9, bold=bold, color=fc)
            c.border    = thin_border()
            c.alignment = Alignment(horizontal=align, vertical="center", indent=1)
            if fmt:
                c.number_format = fmt
        ws2.row_dimensions[ri].height = 15.0

    wb.save(out_path)


# ══════════════════════════════════════════════════════════════════════════════
#  ATTACH
# ══════════════════════════════════════════════════════════════════════════════

def attach(cls):
    """
    Call AFTER lu_analysis_tab.attach() and any other patch (including
    lu_client_search_patch if used).
    """
    cls._build_loanbal_panel      = _build_loanbal_panel_patched
    cls._loanbal_render           = _loanbal_render_patched
    cls._loanbal_show_export_menu = _loanbal_show_export_menu
    cls._loanbal_export_pdf       = _loanbal_export_pdf
    cls._loanbal_export_excel     = _loanbal_export_excel
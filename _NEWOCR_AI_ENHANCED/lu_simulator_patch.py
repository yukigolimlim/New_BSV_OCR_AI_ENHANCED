"""
lu_simulator_patch.py
======================
Enhanced Risk Simulator panel — replaces the original in lu_analysis_tab.

Changes vs original:
  • Per-client selector inside the simulator
  • Net Income input with live surplus / deficit computation
  • Surplus / Deficit hero banner
  • "Stress Report" button and generator
  • Cost Impact Chart REMOVED (cleaner layout, more room for expense rows)
  • Rebuild fix: _rebuild_simulator_panel_patched() tears down and rebuilds
    the already-constructed panel on the first mainloop idle tick, because
    lu_analysis_tab calls the module-level builder directly during panel
    construction (before this patch's attach() runs).
  • Column alignment fix: static header with fixed column widths using
    columnconfigure(minsize) to perfectly align with CTkEntry in expense rows.

HOW TO APPLY
------------
In app.py, after lu_client_search_patch.attach():

    import lu_simulator_patch
    lu_simulator_patch.attach(DocExtractorApp)
"""

import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
from pathlib import Path

# ── colours ───────────────────────────────────────────────────────────────────
_NAVY_DEEP      = "#0A1628"
_NAVY_MID       = "#1A3A6B"
_NAVY_LIGHT     = "#1E4080"
_NAVY_MIST      = "#EEF3FB"
_NAVY_GHOST     = "#D6E4F7"
_NAVY_PALE      = "#4A6FA5"
_WHITE          = "#FFFFFF"
_CARD_WHITE     = "#FAFBFD"
_OFF_WHITE      = "#F5F7FA"
_BORDER_LIGHT   = "#E2E8F5"
_BORDER_MID     = "#C5D0E8"
_TXT_NAVY       = "#1A2B4A"
_TXT_SOFT       = "#6B7FA3"
_TXT_MUTED      = "#9AAACE"
_TXT_ON_LIME    = "#0A2010"
_LIME_BRIGHT    = "#B8FF57"
_LIME_MID       = "#8FD14F"
_LIME_DARK      = "#5A9E28"
_ACCENT_RED     = "#E53E3E"
_ACCENT_GOLD    = "#D4A017"
_ACCENT_SUCCESS = "#2E7D32"
_DEFICIT_BG     = "#2D0A0A"
_SURPLUS_BG     = "#0A1A0A"

_RISK_COLOR = {
    "HIGH":     _ACCENT_RED,
    "MODERATE": _ACCENT_GOLD,
    "LOW":      _ACCENT_SUCCESS,
}
_RISK_BG = {
    "HIGH":     "#FFF5F5",
    "MODERATE": "#FFFBF0",
    "LOW":      "#F0FBE8",
}
_RISK_BADGE_BG = {
    "HIGH":     "#FFE8E8",
    "MODERATE": "#FFF3CD",
    "LOW":      "#DCEDC8",
}

GENERAL_CLIENT = "📊  General (All Clients)"


def _F(size, weight="normal"):
    return ("Segoe UI", size, weight)

def _FF(size, weight="normal"):
    return ctk.CTkFont(family="Segoe UI", size=size, weight=weight)


# ══════════════════════════════════════════════════════════════════════════════
#  THE KEY FIX — rebuild the already-constructed simulator panel
# ══════════════════════════════════════════════════════════════════════════════

def _rebuild_simulator_panel_patched(self):
    """
    lu_analysis_tab._build_lu_analysis_panel() calls the bare module-level
    _build_simulator_panel(self, self._lu_simulator_view) directly, so the
    panel is built with the OLD method before attach() ever runs.

    Fix: destroy all children of _lu_simulator_view and re-run the patched
    builder.  The frame itself and its place() geometry are untouched.
    """
    view = getattr(self, "_lu_simulator_view", None)
    if view is None:
        return

    for child in view.winfo_children():
        child.destroy()

    self._sim_sliders       = {}
    self._sim_expenses      = []
    self._sim_active_client = GENERAL_CLIENT

    _build_simulator_panel_patched(self, view)

    if getattr(self, "_lu_results", None):
        _sim_populate_patched(self)

    # Hook the global search bar so typing a name updates the simulator
    _sim_bind_search_bar(self)


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN PANEL BUILDER  (cost impact chart removed)
# ══════════════════════════════════════════════════════════════════════════════

def _build_simulator_panel_patched(self, parent):

    # ── Header bar ────────────────────────────────────────────────────────────
    sim_hdr = tk.Frame(parent, bg=_NAVY_MID, height=44)
    sim_hdr.pack(fill="x")
    sim_hdr.pack_propagate(False)

    tk.Label(sim_hdr, text="🎛️  Inflation / Cost-Shock Simulator",
             font=_F(10, "bold"), fg=_WHITE, bg=_NAVY_MID
             ).pack(side="left", padx=20, pady=10)

    tk.Label(sim_hdr, text="Global %:", font=_F(8), fg=_TXT_MUTED, bg=_NAVY_MID
             ).pack(side="right", padx=(0, 4))
    self._sim_global_var = tk.StringVar(value="0")
    global_entry = ctk.CTkEntry(
        sim_hdr, textvariable=self._sim_global_var,
        width=52, height=28, corner_radius=4,
        font=_FF(9), fg_color=_NAVY_DEEP, text_color=_WHITE,
        border_color=_NAVY_PALE,
    )
    global_entry.pack(side="right", padx=(0, 8), pady=8)
    global_entry.bind("<Return>",   lambda e: _sim_apply_global(self))
    global_entry.bind("<FocusOut>", lambda e: _sim_apply_global(self))

    ctk.CTkButton(
        sim_hdr, text="Apply All",
        command=lambda: _sim_apply_global(self),
        width=80, height=28, corner_radius=4,
        fg_color=_LIME_DARK, hover_color=_LIME_MID,
        text_color=_TXT_ON_LIME, font=_FF(8, "bold"),
    ).pack(side="right", padx=(0, 6), pady=8)

    ctk.CTkButton(
        sim_hdr, text="Reset",
        command=lambda: _sim_reset(self),
        width=60, height=28, corner_radius=4,
        fg_color=_NAVY_LIGHT, hover_color=_NAVY_DEEP,
        text_color=_WHITE, font=_FF(8, "bold"),
    ).pack(side="right", padx=(0, 4), pady=8)

    # ── Per-simulator client selector strip ───────────────────────────────────
    client_strip = tk.Frame(parent, bg=_NAVY_GHOST, height=40)
    client_strip.pack(fill="x")
    client_strip.pack_propagate(False)

    tk.Label(client_strip, text="👤  Client:", font=_F(8, "bold"),
             fg=_NAVY_MID, bg=_NAVY_GHOST
             ).pack(side="left", padx=(16, 6), pady=10)

    self._sim_client_var = tk.StringVar(value=GENERAL_CLIENT)
    self._sim_client_menu = ctk.CTkOptionMenu(
        client_strip,
        variable=self._sim_client_var,
        values=[GENERAL_CLIENT],
        command=lambda v: _sim_on_client_change(self, v),
        width=280, height=26, corner_radius=4,
        fg_color=_WHITE, button_color=_NAVY_MID,
        button_hover_color=_NAVY_LIGHT,
        text_color=_TXT_NAVY, font=_FF(9),
        dropdown_fg_color=_WHITE,
        dropdown_text_color=_TXT_NAVY,
        dropdown_hover_color=_NAVY_GHOST,
    )
    self._sim_client_menu.pack(side="left", pady=7)

    self._sim_client_badge = tk.Label(
        client_strip, text="GENERAL",
        font=_F(7, "bold"), fg=_WHITE, bg=_NAVY_MID,
        padx=8, pady=3
    )
    self._sim_client_badge.pack(side="left", padx=10, pady=10)

    # ── Net Income input strip ────────────────────────────────────────────────
    income_strip = tk.Frame(parent, bg="#E8F5E9", height=44)
    income_strip.pack(fill="x")
    income_strip.pack_propagate(False)

    tk.Label(income_strip, text="💰  Total Source of Income (₱):",
             font=_F(9, "bold"), fg=_ACCENT_SUCCESS, bg="#E8F5E9"
             ).pack(side="left", padx=(16, 8), pady=10)

    self._sim_income_var = tk.StringVar(value="")
    income_entry = ctk.CTkEntry(
        income_strip,
        textvariable=self._sim_income_var,
        placeholder_text="e.g. 50,000",
        width=180, height=28, corner_radius=4,
        font=_FF(10), fg_color=_WHITE, text_color=_TXT_NAVY,
        border_color=_ACCENT_SUCCESS,
    )
    income_entry.pack(side="left", pady=8)
    income_entry.bind("<Return>",   lambda e: _sim_refresh_patched(self))
    income_entry.bind("<FocusOut>", lambda e: _sim_refresh_patched(self))

    tk.Label(income_strip,
             text="  ← Enter client's total source of income to compute surplus / deficit",
             font=_F(8), fg=_TXT_SOFT, bg="#E8F5E9"
             ).pack(side="left", padx=6)

    ctk.CTkButton(
        income_strip, text="📄  Stress Report",
        command=lambda: _sim_generate_report(self),
        width=130, height=28, corner_radius=4,
        fg_color=_ACCENT_RED, hover_color="#C53030",
        text_color=_WHITE, font=_FF(8, "bold"),
    ).pack(side="right", padx=12, pady=8)

    # ── Summary cards ─────────────────────────────────────────────────────────
    self._sim_summary_bar = tk.Frame(parent, bg=_OFF_WHITE)
    self._sim_summary_bar.pack(fill="x")
    _build_sim_summary_cards_patched(self, self._sim_summary_bar)

    # ── Net income hero banner ────────────────────────────────────────────────
    self._sim_income_hero = tk.Frame(parent, bg=_SURPLUS_BG, height=58)
    self._sim_income_hero.pack(fill="x")
    self._sim_income_hero.pack_propagate(False)

    hero_inner = tk.Frame(self._sim_income_hero, bg=_SURPLUS_BG)
    hero_inner.pack(fill="both", expand=True, padx=20)

    self._sim_income_status_lbl = tk.Label(
        hero_inner, text="NET INCOME",
        font=_F(7, "bold"), fg=_LIME_MID, bg=_SURPLUS_BG
    )
    self._sim_income_status_lbl.pack(side="left", anchor="center")

    self._sim_net_result_lbl = tk.Label(
        hero_inner, text="Enter total source of income above",
        font=_F(14, "bold"), fg=_TXT_MUTED, bg=_SURPLUS_BG
    )
    self._sim_net_result_lbl.pack(side="left", padx=20, anchor="center")

    self._sim_surplus_lbl = tk.Label(
        hero_inner, text="",
        font=_F(11, "bold"), fg=_LIME_MID, bg=_SURPLUS_BG
    )
    self._sim_surplus_lbl.pack(side="left", anchor="center")

    ratio_frame = tk.Frame(hero_inner, bg=_SURPLUS_BG)
    ratio_frame.pack(side="right", anchor="center")
    tk.Label(ratio_frame, text="Expense Ratio",
             font=_F(7), fg=_TXT_MUTED, bg=_SURPLUS_BG
             ).pack(anchor="e")
    self._sim_ratio_lbl = tk.Label(
        ratio_frame, text="—",
        font=_F(10, "bold"), fg=_WHITE, bg=_SURPLUS_BG
    )
    self._sim_ratio_lbl.pack(anchor="e")

    tk.Frame(parent, bg=_BORDER_MID, height=1).pack(fill="x")

    # ── Column header row (fixed, does not scroll) ────────────────────────────
    header_row = tk.Frame(parent, bg=_NAVY_DEEP)
    header_row.pack(fill="x", padx=0)

    # Define column headers: (text, width_chars, anchor)
    headers = [
        ("Expense Item", 26, "w"),
        ("Risk", 6, "center"),
        ("Base Amount", 13, "e"),
        ("% Input", 9, "center"),
        ("Extra Cost", 13, "e"),
        ("Simulated", 13, "e"),
    ]

    for col, (text, width, anchor) in enumerate(headers):
        if col == 3:  # % Input column - use fixed-width frame to match CTkEntry
            f = tk.Frame(header_row, bg=_NAVY_DEEP, width=80)
            f.grid(row=0, column=col, padx=8)
            f.pack_propagate(False)
            tk.Label(f, text=text, font=_F(8, "bold"), fg=_TXT_MUTED,
                     bg=_NAVY_DEEP, anchor="center").pack(fill="x")
        else:
            tk.Label(
                header_row, text=text,
                font=_F(8, "bold"), fg=_TXT_MUTED, bg=_NAVY_DEEP,
                anchor=anchor, padx=6 if anchor != "center" else 0,
                width=width
            ).grid(row=0, column=col, sticky="ew", padx=(4 if col == 3 else 0))

    # Lock column widths using minsize (pixels)
    header_row.columnconfigure(0, minsize=200)  # Expense Item
    header_row.columnconfigure(1, minsize=50)   # Risk
    header_row.columnconfigure(2, minsize=100)  # Base Amount
    header_row.columnconfigure(3, minsize=96)   # % Input (80px entry + 16px pad)
    header_row.columnconfigure(4, minsize=100)  # Extra Cost
    header_row.columnconfigure(5, minsize=100)  # Simulated

    # Uniform group to keep all columns consistent across rows
    for i in range(6):
        header_row.columnconfigure(i, uniform="simcol")

    # ── Scrollable expense list (full width — no chart panel) ─────────────────
    body = tk.Frame(parent, bg=_CARD_WHITE)
    body.pack(fill="both", expand=True)

    sim_sb = tk.Scrollbar(body, relief="flat",
                          troughcolor=_OFF_WHITE, bg=_BORDER_LIGHT, width=8, bd=0)
    sim_sb.pack(side="right", fill="y")

    self._sim_canvas = tk.Canvas(
        body, bg=_CARD_WHITE, highlightthickness=0,
        yscrollcommand=sim_sb.set
    )
    self._sim_canvas.pack(side="left", fill="both", expand=True)
    sim_sb.config(command=self._sim_canvas.yview)

    self._sim_scroll_frame = tk.Frame(self._sim_canvas, bg=_CARD_WHITE)
    self._sim_canvas_win = self._sim_canvas.create_window(
        (0, 0), window=self._sim_scroll_frame, anchor="nw"
    )
    self._sim_scroll_frame.bind(
        "<Configure>",
        lambda e: self._sim_canvas.configure(
            scrollregion=self._sim_canvas.bbox("all"))
    )
    self._sim_canvas.bind(
        "<Configure>",
        lambda e: self._sim_canvas.itemconfig(self._sim_canvas_win, width=e.width)
    )
    self._sim_canvas.bind_all(
        "<MouseWheel>",
        lambda e: (
            self._sim_canvas.yview_scroll(int(-1*(e.delta/120)), "units")
            if self._lu_active_view.get() == "simulator" else None
        )
    )

    # State
    self._sim_sliders       = {}
    self._sim_expenses      = []
    self._sim_active_client = GENERAL_CLIENT

    _sim_show_placeholder_patched(self)


# ══════════════════════════════════════════════════════════════════════════════
#  SUMMARY CARDS
# ══════════════════════════════════════════════════════════════════════════════

def _build_sim_summary_cards_patched(self, parent):
    for w in parent.winfo_children():
        w.destroy()
    for title, attr, color in [
        ("Base Total Expenses", "_sim_lbl_base",  _TXT_NAVY),
        ("Simulated Total",     "_sim_lbl_sim",   _TXT_NAVY),
        ("Total Increase (₱)",  "_sim_lbl_inc",   _ACCENT_RED),
        ("Increase (%)",        "_sim_lbl_pct",   _ACCENT_GOLD),
    ]:
        card = tk.Frame(parent, bg=_NAVY_MIST,
                        highlightbackground=_NAVY_GHOST, highlightthickness=1)
        card.pack(side="left", fill="x", expand=True, padx=6, pady=8)
        tk.Label(card, text=title, font=_F(7), fg=_TXT_SOFT,
                 bg=_NAVY_MIST).pack(anchor="w", padx=10, pady=(6, 0))
        lbl = tk.Label(card, text="—", font=_F(13, "bold"), fg=color, bg=_NAVY_MIST)
        lbl.pack(anchor="w", padx=10, pady=(0, 6))
        setattr(self, attr, lbl)


# ══════════════════════════════════════════════════════════════════════════════
#  CLIENT CHANGE (simulator-local dropdown)
# ══════════════════════════════════════════════════════════════════════════════

def _sim_on_client_change(self, value):
    self._sim_active_client = value
    is_general = (value == GENERAL_CLIENT)

    self._sim_client_badge.config(
        text="GENERAL" if is_general else "PER-CLIENT",
        bg=_NAVY_MID if is_general else _ACCENT_RED
    )

    if is_general:
        results = self._lu_all_data.get("general", [])
    else:
        client_rec = self._lu_all_data.get("clients", {}).get(value)
        results = [client_rec] if client_rec else []

    # Sync _lu_results and _lu_active_client so all renderers stay in sync
    self._lu_results       = results
    self._lu_active_client = value
    self._sim_sliders      = {}

    # Auto-fill total source of income from the Excel data if available
    total_source = None
    gross_income = None
    if not is_general and results:
        total_source = results[0].get("total_source", None)
        gross_income = results[0].get("gross_income", None)

    # Fallback: check income_map directly (single-sheet summary files)
    if total_source is None and not is_general:
        income_map = getattr(self, "_lu_all_data", {}).get("income_map", {})
        income = income_map.get(value)
        if not income:
            for k, v in income_map.items():
                if k.strip().upper() == value.strip().upper():
                    income = v
                    break
        if income:
            total_source = income.get("total_source") or income.get("gross")
            gross_income = income.get("gross")

    if total_source is not None:
        self._sim_income_var.set(f"{total_source:,.2f}")
    elif is_general:
        self._sim_income_var.set("")

    _sim_populate_patched(self)
    _sim_refresh_patched(self)


# ══════════════════════════════════════════════════════════════════════════════
#  POPULATE (header removed – now static; rows only)
# ══════════════════════════════════════════════════════════════════════════════

def _sim_populate_patched(self):
    # Update dropdown options
    clients = sorted(self._lu_all_data.get("clients", {}).keys())
    options = [GENERAL_CLIENT] + clients
    self._sim_client_menu.configure(values=options)

    # Always reload expenses from current _lu_results (client may have changed)
    seen = {}
    for r in self._lu_results:
        for e in r["expenses"]:
            if e["total"] > 0 and e["name"] not in seen:
                seen[e["name"]] = e
    self._sim_expenses = list(seen.values())

    self._sim_sliders = {}

    for w in self._sim_scroll_frame.winfo_children():
        w.destroy()

    if not self._sim_expenses:
        tk.Label(self._sim_scroll_frame,
                 text="No numeric expense data found.\n"
                      "Ensure expense columns contain numeric values.",
                 font=_F(9), fg=_TXT_MUTED, bg=_CARD_WHITE, justify="center"
                 ).pack(pady=60)
        return

    # Expense rows (no header inside scroll area)
    for idx, exp in enumerate(self._sim_expenses):
        var = tk.StringVar(value="0")
        self._sim_sliders[exp["name"]] = var
        _sim_build_expense_row_patched(self, self._sim_scroll_frame, exp, var, idx)

    _sim_refresh_patched(self)


# ══════════════════════════════════════════════════════════════════════════════
#  EXPENSE ROW (ALIGNED with static header using fixed column widths)
# ══════════════════════════════════════════════════════════════════════════════

def _sim_build_expense_row_patched(self, parent, exp, var, idx):
    risk   = exp["risk"]
    row_bg = _RISK_BG.get(risk, _WHITE) if idx % 2 == 0 else _WHITE
    row    = tk.Frame(parent, bg=row_bg)
    row.pack(fill="x")

    # Apply same column minsizes and uniform group as header row
    row.columnconfigure(0, minsize=200)  # Expense Item
    row.columnconfigure(1, minsize=50)   # Risk
    row.columnconfigure(2, minsize=100)  # Base Amount
    row.columnconfigure(3, minsize=96)   # % Input (80px entry + 16px pad)
    row.columnconfigure(4, minsize=100)  # Extra Cost
    row.columnconfigure(5, minsize=100)  # Simulated
    for i in range(6):
        row.columnconfigure(i, uniform="simcol")

    # Column 0: Expense name (fixed width 26 characters)
    tk.Label(row, text=exp["name"], font=_F(9, "bold"),
             fg=_TXT_NAVY, bg=row_bg, anchor="w", padx=8, pady=6, width=26
             ).grid(row=0, column=0, sticky="ew")

    # Column 1: Risk badge (fixed width 6 characters)
    tk.Label(row, text=risk[:3], font=_F(7, "bold"), width=6,
             fg=_RISK_COLOR.get(risk, _TXT_SOFT),
             bg=_RISK_BADGE_BG.get(risk, _OFF_WHITE),
             padx=4, pady=3
             ).grid(row=0, column=1, padx=4, pady=6)

    # Column 2: Base amount (fixed width 13 characters, right-aligned)
    tk.Label(row, text=f"₱{exp['total']:,.2f}", font=_F(9), width=13,
             fg=_TXT_NAVY, bg=row_bg, anchor="e", padx=6
             ).grid(row=0, column=2, sticky="ew")

    # Column 3: Percentage entry (fixed pixel width 80)
    pct_entry = ctk.CTkEntry(
        row, textvariable=var,
        width=80, height=26, corner_radius=4,
        font=_FF(9), fg_color=_WHITE, text_color=_TXT_NAVY,
        border_color=_BORDER_MID,
    )
    pct_entry.grid(row=0, column=3, padx=8, pady=6, sticky="w")

    pct_entry.bind("<Return>",   lambda e, exp=exp: _sim_on_slide_patched(self, exp, var.get()))
    pct_entry.bind("<FocusOut>", lambda e, exp=exp: _sim_on_slide_patched(self, exp, var.get()))

    # Column 4: Extra cost (fixed width 13 characters, right-aligned)
    extra_lbl = tk.Label(row, text="—", font=_F(9), width=13,
                         fg=_ACCENT_RED, bg=row_bg, anchor="e", padx=6)
    extra_lbl.grid(row=0, column=4, sticky="ew")

    # Column 5: Simulated amount (fixed width 13 characters, right-aligned)
    sim_lbl = tk.Label(row, text="—", font=_F(9, "bold"), width=13,
                       fg=_TXT_NAVY, bg=row_bg, anchor="e", padx=6)
    sim_lbl.grid(row=0, column=5, sticky="ew")

    var._extra_lbl = extra_lbl
    var._sim_lbl   = sim_lbl
    var._base      = exp["total"]
    var._row       = row
    var._row_bg    = row_bg

    tk.Frame(parent, bg=_BORDER_LIGHT, height=1).pack(fill="x")


# ══════════════════════════════════════════════════════════════════════════════
#  REFRESH
# ══════════════════════════════════════════════════════════════════════════════

def _sim_refresh_patched(self):
    base_total = 0.0
    sim_total  = 0.0

    income_raw = self._sim_income_var.get().replace(",", "").replace("₱", "").strip()
    try:
        net_income = float(income_raw) if income_raw else None
    except ValueError:
        net_income = None

    for exp in self._sim_expenses:
        var = self._sim_sliders.get(exp["name"])
        pct = 0.0
        if var:
            try:
                pct = float(var.get())
            except Exception:
                pct = 0.0
        base  = exp["total"]
        extra = base * pct / 100.0
        sim   = base + extra
        base_total += base
        sim_total  += sim

        if var and hasattr(var, "_extra_lbl"):
            var._extra_lbl.config(
                text=f"+₱{extra:,.2f}" if extra > 0 else "—",
                fg=_ACCENT_RED if extra > 0 else _TXT_MUTED
            )
            var._sim_lbl.config(
                text=f"₱{sim:,.2f}" if base > 0 else "—",
                fg=_TXT_NAVY
            )
            if net_income is not None and sim_total > net_income and extra > 0:
                try:
                    var._row.config(bg="#FFF0F0")
                    var._extra_lbl.config(bg="#FFF0F0")
                    var._sim_lbl.config(bg="#FFF0F0")
                except Exception:
                    pass
            else:
                try:
                    var._row.config(bg=var._row_bg)
                    var._extra_lbl.config(bg=var._row_bg)
                    var._sim_lbl.config(bg=var._row_bg)
                except Exception:
                    pass

    increase = sim_total - base_total
    pct_inc  = (increase / base_total * 100) if base_total > 0 else 0.0

    if hasattr(self, "_sim_lbl_base"):
        self._sim_lbl_base.config(text=f"₱{base_total:,.2f}" if base_total else "—")
        self._sim_lbl_sim.config( text=f"₱{sim_total:,.2f}"  if base_total else "—")
        self._sim_lbl_inc.config(
            text=f"+₱{increase:,.2f}" if increase > 0 else "₱0.00",
            fg=_ACCENT_RED if increase > 0 else _TXT_NAVY
        )
        self._sim_lbl_pct.config(
            text=f"+{pct_inc:.1f}%" if increase > 0 else "0.0%",
            fg=_ACCENT_GOLD if increase > 0 else _TXT_NAVY
        )

    # Hero banner
    if net_income is not None:
        remaining  = net_income - sim_total
        is_deficit = remaining < 0
        ratio      = (sim_total / net_income * 100) if net_income > 0 else 0

        hero_bg   = _DEFICIT_BG if is_deficit else _SURPLUS_BG
        result_fg = _ACCENT_RED if is_deficit else _LIME_MID
        status    = "⚠️  DEFICIT" if is_deficit else "✅  SURPLUS"
        status_fg = _ACCENT_RED if is_deficit else _LIME_MID

        self._sim_income_hero.config(bg=hero_bg)
        self._sim_income_status_lbl.config(text=status, fg=status_fg, bg=hero_bg)
        sign = "-" if is_deficit else "+"
        self._sim_net_result_lbl.config(
            text=f"Net: {sign}₱{abs(remaining):,.2f}", fg=result_fg, bg=hero_bg
        )
        self._sim_surplus_lbl.config(
            text=(f"Expenses {ratio:.1f}% of income  |  "
                  f"Total Source: ₱{net_income:,.2f}  |  "
                  f"Simulated: ₱{sim_total:,.2f}"),
            fg=_TXT_MUTED, bg=hero_bg
        )
        ratio_color = (
            _ACCENT_RED  if ratio > 100 else
            _ACCENT_GOLD if ratio > 80  else
            _LIME_MID
        )
        self._sim_ratio_lbl.config(text=f"{ratio:.1f}%", fg=ratio_color, bg=hero_bg)
    else:
        self._sim_income_hero.config(bg=_SURPLUS_BG)
        self._sim_income_status_lbl.config(text="NET INCOME", fg=_LIME_MID, bg=_SURPLUS_BG)
        self._sim_net_result_lbl.config(
            text="Enter total source of income above", fg=_TXT_MUTED, bg=_SURPLUS_BG
        )
        self._sim_surplus_lbl.config(text="", bg=_SURPLUS_BG)
        self._sim_ratio_lbl.config(text="—", fg=_WHITE, bg=_SURPLUS_BG)


# ══════════════════════════════════════════════════════════════════════════════
#  SLIDE / GLOBAL / RESET
# ══════════════════════════════════════════════════════════════════════════════

def _sim_on_slide_patched(self, exp, value):
    try:
        pct = float(value)
        if pct < 0.0:
            pct = 0.0
    except (ValueError, TypeError):
        pct = 0.0
    self._sim_sliders[exp["name"]].set(str(pct))
    _sim_refresh_patched(self)


def _sim_apply_global(self):
    try:
        pct = float(self._sim_global_var.get())
        if pct < 0.0:
            pct = 0.0
    except (ValueError, TypeError):
        pct = 0.0
    if not self._sim_sliders and self._lu_results:
        _sim_populate_patched(self)
    for var in self._sim_sliders.values():
        var.set(str(pct))
    _sim_refresh_patched(self)


def _sim_reset(self):
    self._sim_global_var.set("0")
    if not self._sim_sliders and self._lu_results:
        _sim_populate_patched(self)
    for var in self._sim_sliders.values():
        var.set("0")
    _sim_refresh_patched(self)


# ══════════════════════════════════════════════════════════════════════════════
#  PLACEHOLDER
# ══════════════════════════════════════════════════════════════════════════════

def _sim_show_placeholder_patched(self):
    for w in self._sim_scroll_frame.winfo_children():
        w.destroy()
    tk.Label(self._sim_scroll_frame,
             text="Run an analysis first to unlock the simulator.",
             font=_F(10), fg=_TXT_MUTED, bg=_CARD_WHITE
             ).pack(pady=60)


# ══════════════════════════════════════════════════════════════════════════════
#  SEARCH BAR INTEGRATION
# ══════════════════════════════════════════════════════════════════════════════

def _sim_on_search_patched(self, text: str):
    """
    Called whenever the global search bar changes while the simulator is active.
    Finds the best-matching client name, updates the simulator dropdown,
    and triggers the full client-change flow (income auto-fill, expense reload, etc.).
    """
    text = text.strip()

    if not text:
        # Empty search → revert to General
        self._sim_client_var.set(GENERAL_CLIENT)
        _sim_on_client_change(self, GENERAL_CLIENT)
        return

    clients = sorted(self._lu_all_data.get("clients", {}).keys())
    if not clients:
        return

    text_up = text.upper()

    # 1. Exact match
    exact = next((c for c in clients if c.strip().upper() == text_up), None)
    if exact:
        _sim_select_client(self, exact)
        return

    # 2. Starts-with match
    starts = [c for c in clients if c.strip().upper().startswith(text_up)]
    if starts:
        _sim_select_client(self, starts[0])
        return

    # 3. Contains match
    contains = [c for c in clients if text_up in c.strip().upper()]
    if contains:
        _sim_select_client(self, contains[0])
        return

    # 4. No match — keep current selection, do nothing
    # (avoids jarring reset while the user is still typing)


def _sim_select_client(self, client_name: str):
    """Update the dropdown and fire the full client-change handler."""
    self._sim_client_var.set(client_name)
    _sim_on_client_change(self, client_name)


def _sim_bind_search_bar(self):
    """
    Find the app's global search Entry widget and attach a trace so the
    simulator reacts when the user types in it.

    Tries several attribute names that the host app may use.
    Also attempts to locate the widget by walking the widget tree if
    no known attribute is found.
    """
    # Common attribute names used by the host app for the search field
    _SEARCH_ATTR_CANDIDATES = [
        "_lu_search_var",       # StringVar — preferred
        "_search_var",
        "_lu_search_entry",     # Entry widget
        "_search_entry",
        "_lu_client_search_var",
        "_client_search_var",
    ]

    for attr in _SEARCH_ATTR_CANDIDATES:
        obj = getattr(self, attr, None)
        if obj is None:
            continue

        if isinstance(obj, tk.StringVar):
            # Trace on StringVar: fires on every keystroke
            obj.trace_add("write", lambda *_: _sim_on_search_patched(
                self, obj.get()))
            self._sim_search_bound_var = obj
            return

        if isinstance(obj, (tk.Entry, ctk.CTkEntry)):
            # Bind directly on an Entry widget
            var = tk.StringVar()
            obj.configure(textvariable=var)
            var.trace_add("write", lambda *_: _sim_on_search_patched(
                self, var.get()))
            self._sim_search_bound_var = var
            return

    # Fallback: walk widget tree looking for an Entry in the top-level frame
    # that is NOT inside the simulator panel itself.
    _sim_search_bind_by_walk(self)


def _sim_search_bind_by_walk(self):
    """
    Last-resort: walk the widget tree from the root window and bind to the
    first Entry (or CTkEntry) that is NOT a child of _lu_simulator_view.
    """
    sim_view = getattr(self, "_lu_simulator_view", None)

    def _is_inside_sim(w):
        try:
            p = w
            while p:
                if p == sim_view:
                    return True
                p = p.master
        except Exception:
            pass
        return False

    def _walk(widget):
        for child in widget.winfo_children():
            if _is_inside_sim(child):
                continue
            if isinstance(child, (tk.Entry, ctk.CTkEntry)):
                var = tk.StringVar()
                try:
                    child.configure(textvariable=var)
                    var.trace_add("write", lambda *_: _sim_on_search_patched(
                        self, var.get()))
                    self._sim_search_bound_var = var
                    return True
                except Exception:
                    pass
            if _walk(child):
                return True
        return False

    try:
        _walk(self)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════════
#  STRESS REPORT
# ══════════════════════════════════════════════════════════════════════════════

def _rl_hex(color) -> str:
    """Convert a reportlab color to a hex string for use in Paragraph markup."""
    try:
        return f"{int(color.red*255):02X}{int(color.green*255):02X}{int(color.blue*255):02X}"
    except Exception:
        return "000000"


def _sim_generate_report(self):
    """Generate and save a PDF stress report."""
    if not self._sim_expenses:
        messagebox.showwarning("No Data", "Run an analysis and load expenses first.")
        return

    try:
        from reportlab.lib.pagesizes import A4, landscape as rl_landscape
        from reportlab.lib import colors as rl_colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
            HRFlowable,
        )
    except ImportError:
        messagebox.showerror(
            "Missing Library",
            "reportlab is not installed.\nRun:  pip install reportlab"
        )
        return

    # ── Gather data ───────────────────────────────────────────────────────────
    income_raw = self._sim_income_var.get().replace(",", "").replace("₱", "").strip()
    try:
        net_income = float(income_raw) if income_raw else None
    except ValueError:
        net_income = None

    client     = getattr(self, "_sim_active_client", GENERAL_CLIENT)
    now        = datetime.now().strftime("%B %d, %Y  %H:%M")
    fname      = Path(self._lu_filepath).name if getattr(self, "_lu_filepath", None) else "—"

    base_total   = sum(e["total"] for e in self._sim_expenses)
    sim_total    = 0.0
    expense_rows = []

    for exp in self._sim_expenses:
        var = self._sim_sliders.get(exp["name"])
        try:
            pct = float(var.get()) if var else 0.0
        except (ValueError, TypeError):
            pct = 0.0
        base  = exp["total"]
        extra = base * pct / 100.0
        sim   = base + extra
        sim_total += sim
        expense_rows.append((exp["name"], exp["risk"], base, sim, extra, pct))

    increase   = sim_total - base_total
    pct_inc    = (increase / base_total * 100) if base_total > 0 else 0.0
    remaining  = (net_income - sim_total) if net_income is not None else None
    is_deficit = (remaining is not None and remaining < 0)
    ratio      = (sim_total / net_income * 100) if net_income and net_income > 0 else None

    deficit_items = sorted(
        [(name, risk, extra) for name, risk, base, sim, extra, pct in expense_rows
         if pct > 0 and is_deficit],
        key=lambda x: -x[2]
    )

    # ── File dialog ───────────────────────────────────────────────────────────
    safe_client  = client.replace(" ", "_").replace("/", "_")[:40]
    default_name = f"StressReport_{safe_client}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    path = filedialog.asksaveasfilename(
        title="Save Stress Report PDF",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        initialfile=default_name,
    )
    if not path:
        return

    # ── Colours ───────────────────────────────────────────────────────────────
    c_navy   = rl_colors.HexColor("#1A3A6B")
    c_navy_m = rl_colors.HexColor("#EEF3FB")
    c_red    = rl_colors.HexColor("#E53E3E")
    c_gold   = rl_colors.HexColor("#D4A017")
    c_green  = rl_colors.HexColor("#2E7D32")
    c_white  = rl_colors.white
    c_off    = rl_colors.HexColor("#F5F7FA")
    c_border = rl_colors.HexColor("#C5D0E8")
    c_muted  = rl_colors.HexColor("#9AAACE")
    c_hdr_bg = c_red if is_deficit else c_navy
    RISK_C   = {"HIGH": c_red, "MODERATE": c_gold, "LOW": c_green}

    # ── Styles ────────────────────────────────────────────────────────────────
    styles  = getSampleStyleSheet()
    s_title = ParagraphStyle("SRTitle", parent=styles["Normal"],
                             fontSize=14, textColor=c_white, leading=18,
                             fontName="Helvetica-Bold")
    s_sub   = ParagraphStyle("SRSub",   parent=styles["Normal"],
                             fontSize=8,  textColor=c_muted,  leading=11)
    s_h2    = ParagraphStyle("SRH2",    parent=styles["Normal"],
                             fontSize=10, textColor=c_navy,   leading=14,
                             fontName="Helvetica-Bold", spaceBefore=8)
    s_body  = ParagraphStyle("SRBody",  parent=styles["Normal"],
                             fontSize=8,  textColor=rl_colors.HexColor("#1A2B4A"), leading=11)
    s_muted = ParagraphStyle("SRMuted", parent=styles["Normal"],
                             fontSize=7,  textColor=c_muted,  leading=10)
    s_warn  = ParagraphStyle("SRWarn",  parent=styles["Normal"],
                             fontSize=9,  textColor=c_red,    leading=12,
                             fontName="Helvetica-Bold")

    # ── Build PDF ─────────────────────────────────────────────────────────────
    PAGE = rl_landscape(A4)
    doc  = SimpleDocTemplate(
        path, pagesize=PAGE,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm,  bottomMargin=1.5*cm,
    )
    pw    = PAGE[0] - 3*cm   # usable page width
    story = []

    # Title banner
    status_text = "DEFICIT STRESS REPORT" if is_deficit else "SURPLUS STRESS REPORT"
    banner = Table(
        [[Paragraph(f"INFLATION / COST-SHOCK  —  {status_text}", s_title)]],
        colWidths=[pw]
    )
    banner.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), c_hdr_bg),
        ("TOPPADDING",    (0, 0), (-1, -1), 12),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
        ("LEFTPADDING",   (0, 0), (-1, -1), 14),
    ]))
    story.append(banner)
    story.append(Spacer(1, 0.25*cm))
    story.append(Paragraph(
        f"File: <b>{fname}</b>&nbsp;&nbsp; Generated: <b>{now}</b>&nbsp;&nbsp; Client: <b>{client}</b>",
        s_sub))
    story.append(Spacer(1, 0.25*cm))
    story.append(HRFlowable(width="100%", thickness=1.5, color=c_navy))
    story.append(Spacer(1, 0.25*cm))

    # Summary cards
    card_items = [
        ("Base Total Expenses", f"P{base_total:,.2f}", c_navy),
        ("Simulated Total",     f"P{sim_total:,.2f}",  c_navy),
        ("Total Increase",      f"+P{increase:,.2f}  (+{pct_inc:.1f}%)",
         c_red if increase > 0 else c_navy),
    ]
    if net_income is not None:
        card_items.append(("Total Source of Income", f"P{net_income:,.2f}", c_green))

    card_col_w = pw / len(card_items)
    card_data  = [[
        Paragraph(
            f"<font size='7' color='#{_rl_hex(c_muted)}'>{lbl}</font><br/>"
            f"<font size='12' color='#{_rl_hex(fg)}'><b>{val}</b></font>",
            s_body)
        for lbl, val, fg in card_items
    ]]
    card_tbl = Table(card_data, colWidths=[card_col_w] * len(card_items))
    card_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), c_navy_m),
        ("BOX",           (0, 0), (-1, -1), 0.5, c_border),
        ("INNERGRID",     (0, 0), (-1, -1), 0.5, c_border),
        ("TOPPADDING",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(card_tbl)
    story.append(Spacer(1, 0.25*cm))

    # Net result hero (only if income entered)
    if net_income is not None and remaining is not None:
        ratio_c     = c_red if ratio > 100 else (c_gold if ratio > 80 else c_green)
        result_fg   = c_red if is_deficit else c_green
        result_sign = "-" if is_deficit else "+"
        hero_bg     = rl_colors.HexColor("#2D0A0A") if is_deficit else rl_colors.HexColor("#0A1A0A")
        hero = Table([[Paragraph(
            f"<font color='#{_rl_hex(result_fg)}'><b>"
            f"{'DEFICIT' if is_deficit else 'SURPLUS'}  "
            f"{result_sign}P{abs(remaining):,.2f}"
            f"</b></font>"
            f"&nbsp;&nbsp;&nbsp; Expense Ratio: "
            f"<font color='#{_rl_hex(ratio_c)}'><b>{ratio:.1f}%</b></font>"
            f"&nbsp;&nbsp; Simulated P{sim_total:,.2f} vs "
            f"Total Source P{net_income:,.2f}",
            s_body)]], colWidths=[pw])
        hero.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, -1), hero_bg),
            ("TOPPADDING",    (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ("LEFTPADDING",   (0, 0), (-1, -1), 14),
        ]))
        story.append(hero)
        story.append(Spacer(1, 0.25*cm))

    # Expense table
    story.append(Paragraph("Expense Breakdown", s_h2))
    story.append(Spacer(1, 0.15*cm))

    hdr_s   = ParagraphStyle("EHdr", parent=styles["Normal"],
                             fontSize=7, textColor=c_white, leading=9)
    e_cols  = [pw*0.28, pw*0.09, pw*0.14, pw*0.14, pw*0.14, pw*0.10, pw*0.11]
    tbl_data = [[
        Paragraph("<b>Expense Item</b>",  hdr_s),
        Paragraph("<b>Risk</b>",          hdr_s),
        Paragraph("<b>Base Amount</b>",   hdr_s),
        Paragraph("<b>Simulated</b>",     hdr_s),
        Paragraph("<b>Extra Cost</b>",    hdr_s),
        Paragraph("<b>Rate (%)</b>",      hdr_s),
        Paragraph("<b>Flag</b>",          hdr_s),
    ]]

    tbl_style = TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  c_navy),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("LEADING",       (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("BOX",           (0, 0), (-1, -1), 0.5, c_border),
        ("INNERGRID",     (0, 0), (-1, -1), 0.3, c_border),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ])

    for idx, (name, risk, base, sim, extra, pct) in enumerate(expense_rows):
        risk_c   = RISK_C.get(risk, c_muted)
        stressed = pct > 0
        row_bg   = (rl_colors.HexColor("#FFF5F5") if risk == "HIGH" else
                    rl_colors.HexColor("#FFFBF0") if risk == "MODERATE" else
                    c_white) if idx % 2 == 0 else c_off
        tbl_style.add("BACKGROUND", (0, idx+1), (-1, idx+1), row_bg)
        tbl_data.append([
            Paragraph(f"<b>{name[:42]}</b>", s_body),
            Paragraph(f"<font color='#{_rl_hex(risk_c)}'><b>{risk}</b></font>", s_body),
            Paragraph(f"P{base:,.2f}", s_body),
            Paragraph(f"<b>P{sim:,.2f}</b>", s_body),
            Paragraph(
                f"<font color='#{_rl_hex(c_red)}'><b>+P{extra:,.2f}</b></font>"
                if stressed else "—", s_body),
            Paragraph(
                f"<font color='#{_rl_hex(c_red)}'><b>{pct:.1f}%</b></font>"
                if stressed else "0.0%", s_body),
            Paragraph(
                f"<font color='#{_rl_hex(c_red)}'><b>STRESS</b></font>"
                if stressed else "", s_body),
        ])

    # Totals row
    tbl_data.append([
        Paragraph("<b>TOTAL</b>", s_body),
        Paragraph("", s_body),
        Paragraph(f"<b>P{base_total:,.2f}</b>", s_body),
        Paragraph(f"<b>P{sim_total:,.2f}</b>",  s_body),
        Paragraph(f"<b>+P{increase:,.2f}</b>",  s_body),
        Paragraph(f"<b>{pct_inc:.1f}%</b>",     s_body),
        Paragraph("", s_body),
    ])
    tbl_style.add("BACKGROUND", (0, -1), (-1, -1), c_navy_m)
    tbl_style.add("FONTNAME",   (0, -1), (-1, -1), "Helvetica-Bold")
    tbl_style.add("LINEABOVE",  (0, -1), (-1, -1), 1.2, c_navy)

    exp_tbl = Table(tbl_data, colWidths=e_cols, repeatRows=1)
    exp_tbl.setStyle(tbl_style)
    story.append(exp_tbl)

    # Deficit contributing items
    if is_deficit and deficit_items:
        story.append(Spacer(1, 0.35*cm))
        story.append(HRFlowable(width="100%", thickness=1, color=c_red))
        story.append(Paragraph("Deficit Contributing Items", s_warn))
        story.append(Spacer(1, 0.15*cm))

        di_cols = [pw*0.40, pw*0.12, pw*0.18, pw*0.30]
        di_s    = ParagraphStyle("DIHdr", parent=styles["Normal"],
                                 fontSize=7, textColor=c_white, leading=9)
        di_data = [[
            Paragraph("<b>Expense Item</b>",    di_s),
            Paragraph("<b>Risk</b>",            di_s),
            Paragraph("<b>Extra Cost</b>",      di_s),
            Paragraph("<b>Recommendation</b>",  di_s),
        ]]
        for i, (name, risk, extra) in enumerate(deficit_items):
            risk_c = RISK_C.get(risk, c_muted)
            rec    = ("Renegotiate or hedge immediately"    if risk == "HIGH"     else
                      "Monitor closely, plan mitigation"   if risk == "MODERATE" else
                      "Review for cost savings")
            di_data.append([
                Paragraph(f"<b>{name[:55]}</b>", s_body),
                Paragraph(f"<font color='#{_rl_hex(risk_c)}'><b>{risk}</b></font>", s_body),
                Paragraph(f"<font color='#{_rl_hex(c_red)}'><b>+P{extra:,.2f}</b></font>", s_body),
                Paragraph(rec, s_muted),
            ])
        di_tbl = Table(di_data, colWidths=di_cols, repeatRows=1)
        di_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0),  c_red),
            ("FONTSIZE",      (0, 0), (-1, -1), 8),
            ("LEADING",       (0, 0), (-1, -1), 10),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
            ("BOX",           (0, 0), (-1, -1), 0.5, c_border),
            ("INNERGRID",     (0, 0), (-1, -1), 0.3, c_border),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            *[("BACKGROUND", (0, i), (-1, i), c_off if i % 2 == 0 else c_white)
              for i in range(1, len(di_data))],
        ]))
        story.append(di_tbl)

    story.append(Spacer(1, 0.35*cm))
    story.append(HRFlowable(width="100%", thickness=1, color=c_border))
    story.append(Paragraph("END OF STRESS REPORT", s_muted))

    try:
        doc.build(story)
        messagebox.showinfo("Export Complete", f"Stress Report PDF saved to:\n{path}")
    except Exception as ex:
        messagebox.showerror("PDF Export Error", str(ex))


# ══════════════════════════════════════════════════════════════════════════════
#  ATTACH
# ══════════════════════════════════════════════════════════════════════════════

def attach(cls):
    """Call AFTER lu_client_search_patch.attach()."""
    cls._build_simulator_panel   = _build_simulator_panel_patched
    cls._build_sim_summary_cards = _build_sim_summary_cards_patched
    cls._sim_show_placeholder    = _sim_show_placeholder_patched
    cls._sim_populate            = _sim_populate_patched
    cls._sim_build_expense_row   = _sim_build_expense_row_patched
    cls._sim_on_slide            = _sim_on_slide_patched
    cls._sim_apply_global        = _sim_apply_global
    cls._sim_reset               = _sim_reset
    cls._sim_refresh             = _sim_refresh_patched
    cls._sim_generate_report     = _sim_generate_report
    cls._sim_on_client_change    = _sim_on_client_change
    cls._rebuild_simulator_panel = _rebuild_simulator_panel_patched
    cls._sim_on_search           = _sim_on_search_patched
    cls._sim_bind_search_bar     = _sim_bind_search_bar

    # THE FIX: wrap __init__ to schedule rebuild on first idle tick
    original_init = cls.__init__

    def patched_init(self, *args, **kwargs):
        original_init(self, *args, **kwargs)
        self.after(0, lambda: _rebuild_simulator_panel_patched(self))

    cls.__init__ = patched_init
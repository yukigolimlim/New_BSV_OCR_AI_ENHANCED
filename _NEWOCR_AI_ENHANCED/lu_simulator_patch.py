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

    tk.Label(income_strip, text="💰  Declared Net Income (₱):",
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
             text="  ← Enter client's net income to compute surplus / deficit",
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
        hero_inner, text="Enter net income above",
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

    # Auto-fill net income from the Excel data if available
    net_income   = None
    gross_income = None
    if not is_general and results:
        net_income   = results[0].get("net_income",   None)
        gross_income = results[0].get("gross_income", None)

    # Fallback: check income_map directly (single-sheet summary files)
    if net_income is None and not is_general:
        income_map = getattr(self, "_lu_all_data", {}).get("income_map", {})
        # Try exact match first, then case-insensitive
        income = income_map.get(value)
        if not income:
            for k, v in income_map.items():
                if k.strip().upper() == value.strip().upper():
                    income = v
                    break
        if income:
            net_income   = income.get("net")
            gross_income = income.get("gross")

    if net_income is not None:
        self._sim_income_var.set(f"{net_income:,.2f}")
    elif is_general:
        self._sim_income_var.set("")

    _sim_populate_patched(self)
    _sim_refresh_patched(self)


# ══════════════════════════════════════════════════════════════════════════════
#  POPULATE
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

    # Column header
    hdr = tk.Frame(self._sim_scroll_frame, bg=_OFF_WHITE)
    hdr.pack(fill="x", pady=(8, 0))
    for text, col, w in [
        ("Expense Item",   0, 220),
        ("Risk",           1,  60),
        ("Base Amount",    2, 110),
        ("% Input",        3,  80),
        ("Extra Cost",     4, 110),
        ("Simulated",      5, 110),
    ]:
        tk.Label(hdr, text=text, font=_F(8, "bold"), fg=_NAVY_PALE,
                 bg=_OFF_WHITE, width=w//8, anchor="w", padx=6, pady=5
                 ).grid(row=0, column=col, sticky="ew", padx=(0, 2))

    tk.Frame(self._sim_scroll_frame, bg=_BORDER_MID, height=1).pack(fill="x")

    for idx, exp in enumerate(self._sim_expenses):
        var = tk.StringVar(value="0")
        self._sim_sliders[exp["name"]] = var
        _sim_build_expense_row_patched(self, self._sim_scroll_frame, exp, var, idx)

    _sim_refresh_patched(self)


# ══════════════════════════════════════════════════════════════════════════════
#  EXPENSE ROW
# ══════════════════════════════════════════════════════════════════════════════

def _sim_build_expense_row_patched(self, parent, exp, var, idx):
    risk   = exp["risk"]
    row_bg = _RISK_BG.get(risk, _WHITE) if idx % 2 == 0 else _WHITE
    row    = tk.Frame(parent, bg=row_bg)
    row.pack(fill="x")

    tk.Label(row, text=exp["name"], font=_F(9, "bold"),
             fg=_TXT_NAVY, bg=row_bg, anchor="w", padx=8, pady=6, width=26
             ).grid(row=0, column=0, sticky="ew")

    tk.Label(row, text=risk[:3], font=_F(7, "bold"),
             fg=_RISK_COLOR.get(risk, _TXT_SOFT),
             bg=_RISK_BADGE_BG.get(risk, _OFF_WHITE),
             padx=4, pady=3
             ).grid(row=0, column=1, padx=4, pady=6)

    tk.Label(row, text=f"₱{exp['total']:,.2f}", font=_F(9),
             fg=_TXT_NAVY, bg=row_bg, anchor="e", padx=6, width=13
             ).grid(row=0, column=2, sticky="ew")

    pct_entry = ctk.CTkEntry(
        row, textvariable=var,
        width=80, height=26, corner_radius=4,
        font=_FF(9), fg_color=_WHITE, text_color=_TXT_NAVY,
        border_color=_BORDER_MID,
    )
    pct_entry.grid(row=0, column=3, padx=8, pady=6)
    pct_entry.bind("<Return>",   lambda e, exp=exp: _sim_on_slide_patched(self, exp, var.get()))
    pct_entry.bind("<FocusOut>", lambda e, exp=exp: _sim_on_slide_patched(self, exp, var.get()))

    extra_lbl = tk.Label(row, text="—", font=_F(9),
                         fg=_ACCENT_RED, bg=row_bg, anchor="e", padx=6, width=13)
    extra_lbl.grid(row=0, column=4, sticky="ew")

    sim_lbl = tk.Label(row, text="—", font=_F(9, "bold"),
                       fg=_TXT_NAVY, bg=row_bg, anchor="e", padx=6, width=13)
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
                  f"Declared: ₱{net_income:,.2f}  |  "
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
            text="Enter net income above", fg=_TXT_MUTED, bg=_SURPLUS_BG
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
#  STRESS REPORT
# ══════════════════════════════════════════════════════════════════════════════

def _sim_generate_report(self):
    if not self._sim_expenses:
        messagebox.showwarning("No Data", "Run an analysis and load expenses first.")
        return

    income_raw = self._sim_income_var.get().replace(",", "").replace("₱", "").strip()
    try:
        net_income = float(income_raw) if income_raw else None
    except ValueError:
        net_income = None

    client     = getattr(self, "_sim_active_client", GENERAL_CLIENT)
    now        = datetime.now().strftime("%B %d, %Y  %H:%M")
    fname      = Path(self._lu_filepath).name if self._lu_filepath else "—"

    base_total = sum(e["total"] for e in self._sim_expenses)
    sim_total  = sum(
        e["total"] + e["total"] * (
            self._sim_sliders.get(e["name"], tk.DoubleVar()).get() or 0
        ) / 100
        for e in self._sim_expenses
    )
    increase   = sim_total - base_total
    pct_inc    = (increase / base_total * 100) if base_total > 0 else 0.0
    remaining  = (net_income - sim_total) if net_income is not None else None
    is_deficit = (remaining is not None and remaining < 0)

    dbl  = "═" * 72
    rule = "─" * 72
    lines = []

    lines.append(dbl)
    lines.append("  INFLATION / COST-SHOCK STRESS REPORT")
    lines.append(dbl)
    lines.append(f"  File      : {fname}")
    lines.append(f"  Generated : {now}")
    lines.append(f"  Client    : {client}")
    if net_income is not None:
        lines.append(f"  Net Income: ₱{net_income:,.2f}")
    lines.append(rule)
    lines.append(f"  Base Total Expenses : ₱{base_total:,.2f}")
    lines.append(f"  Simulated Total     : ₱{sim_total:,.2f}")
    lines.append(f"  Total Increase      : +₱{increase:,.2f}  (+{pct_inc:.1f}%)")
    if net_income is not None:
        sign   = "-" if is_deficit else "+"
        status = "⚠️  DEFICIT" if is_deficit else "✅  SURPLUS"
        ratio  = sim_total / net_income * 100 if net_income > 0 else 0
        lines.append(f"  Net Result          : {sign}₱{abs(remaining):,.2f}  [{status}]")
        lines.append(f"  Expense / Income    : {ratio:.1f}%")
    lines.append(dbl)
    lines.append("")
    lines.append(f"  {'EXPENSE ITEM':<28} {'RISK':<10} {'BASE':>12} {'SIMULATED':>12} {'EXTRA':>12} {'%':>6}")
    lines.append("  " + rule)

    deficit_items = []
    for exp in self._sim_expenses:
        var = self._sim_sliders.get(exp["name"])
        pct = 0.0
        if var:
            try:
                pct = float(var.get())
            except Exception:
                pass
        base  = exp["total"]
        extra = base * pct / 100.0
        sim   = base + extra
        flag  = " ◄ STRESS" if pct > 0 else ""
        lines.append(
            f"  {exp['name'][:27]:<28} {exp['risk']:<10} ₱{base:>11,.2f} "
            f"₱{sim:>11,.2f} +₱{extra:>10,.2f} {pct:>5.1f}%{flag}"
        )
        if pct > 0 and is_deficit:
            deficit_items.append((exp["name"], exp["risk"], extra))

    lines.append("")
    lines.append(dbl)

    if is_deficit and deficit_items:
        lines.append("")
        lines.append("  ⚠️  DEFICIT CONTRIBUTING ITEMS")
        lines.append("  " + rule)
        for name, risk, extra in sorted(deficit_items, key=lambda x: -x[2]):
            lines.append(f"  • {name}  [{risk}]  +₱{extra:,.2f} extra cost")
        lines.append("")
        lines.append("  RECOMMENDATION: Review HIGH and MODERATE risk expenses above.")
        lines.append("  Consider renegotiating or hedging these cost items.")
        lines.append(dbl)

    lines.append("")
    lines.append("  END OF STRESS REPORT")
    lines.append(dbl)

    report_text = "\n".join(lines)

    win = tk.Toplevel(self)
    win.title("Stress Report")
    win.configure(bg=_NAVY_DEEP)
    win.geometry("860x620")
    win.grab_set()

    hdr_bg = _ACCENT_RED if is_deficit else _NAVY_MID
    title_bar = tk.Frame(win, bg=hdr_bg, height=44)
    title_bar.pack(fill="x")
    title_bar.pack_propagate(False)
    tk.Label(title_bar,
             text="⚠️  DEFICIT STRESS REPORT" if is_deficit else "✅  SURPLUS STRESS REPORT",
             font=_F(11, "bold"), fg=_WHITE, bg=hdr_bg
             ).pack(side="left", padx=20, pady=10)

    def _save():
        path = filedialog.asksaveasfilename(
            title="Save Stress Report", defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"StressReport_{client.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(report_text)
            messagebox.showinfo("Saved", f"Report saved to:\n{path}", parent=win)

    ctk.CTkButton(title_bar, text="💾  Save Report", command=_save,
                  width=120, height=30, corner_radius=4,
                  fg_color=_LIME_DARK, hover_color=_LIME_MID,
                  text_color=_TXT_ON_LIME, font=_FF(9, "bold")
                  ).pack(side="right", padx=12, pady=7)
    ctk.CTkButton(title_bar, text="✕  Close", command=win.destroy,
                  width=80, height=30, corner_radius=4,
                  fg_color=_NAVY_LIGHT, hover_color=_NAVY_DEEP,
                  text_color=_WHITE, font=_FF(9)
                  ).pack(side="right", padx=(0, 4), pady=7)

    body = tk.Frame(win, bg=_NAVY_DEEP)
    body.pack(fill="both", expand=True, padx=8, pady=8)
    sb = tk.Scrollbar(body, relief="flat", troughcolor=_NAVY_MID,
                      bg=_NAVY_PALE, width=8, bd=0)
    sb.pack(side="right", fill="y")
    txt = tk.Text(body, font=("Consolas", 9), fg=_WHITE, bg=_NAVY_DEEP,
                  relief="flat", bd=0, padx=16, pady=12,
                  wrap="none", yscrollcommand=sb.set)
    txt.pack(side="left", fill="both", expand=True)
    sb.config(command=txt.yview)

    txt.tag_configure("deficit", foreground=_ACCENT_RED,  font=("Consolas", 9, "bold"))
    txt.tag_configure("surplus", foreground=_LIME_MID,    font=("Consolas", 9, "bold"))
    txt.tag_configure("stress",  foreground=_ACCENT_GOLD, font=("Consolas", 9))
    txt.tag_configure("rule",    foreground="#3A5A8A",     font=("Consolas", 9))
    txt.tag_configure("title",   foreground=_LIME_BRIGHT, font=("Consolas", 12, "bold"))
    txt.tag_configure("normal",  foreground="#C8D8F0",     font=("Consolas", 9))
    txt.tag_configure("warning", foreground=_ACCENT_RED,  font=("Consolas", 9, "bold"))

    for line in lines:
        if "═" in line or "─" in line:
            txt.insert("end", line + "\n", "rule")
        elif "STRESS REPORT" in line:
            txt.insert("end", line + "\n", "title")
        elif "DEFICIT" in line or "⚠️" in line:
            txt.insert("end", line + "\n", "deficit" if is_deficit else "normal")
        elif "SURPLUS" in line or "✅" in line:
            txt.insert("end", line + "\n", "surplus")
        elif "◄ STRESS" in line:
            txt.insert("end", line + "\n", "stress")
        elif "RECOMMENDATION" in line or "Consider" in line:
            txt.insert("end", line + "\n", "warning")
        else:
            txt.insert("end", line + "\n", "normal")

    txt.config(state="disabled")
    txt.yview_moveto(0)


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

    # THE FIX: wrap __init__ to schedule rebuild on first idle tick
    original_init = cls.__init__

    def patched_init(self, *args, **kwargs):
        original_init(self, *args, **kwargs)
        self.after(0, lambda: _rebuild_simulator_panel_patched(self))

    cls.__init__ = patched_init
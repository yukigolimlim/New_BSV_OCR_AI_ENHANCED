"""
lu_simulator_patch.py — Risk Simulator Tab
============================================
Inflation / Cost-Shock Simulator.

This file is the CANONICAL simulator implementation.
It supersedes _sim_* functions in lu_ui.py and the simulator patches
in lu_analysis_patch_v8.py.  Apply it AFTER lu_tab_analysis.attach(cls).

Standalone: imports only lu_core and lu_shared.
Attached to app class via attach(cls).

Key improvements over original lu_ui
--------------------------------------
  • Expense mix shown as a matplotlib pie (% of total simulated spend).
  • Chunked row construction via after() to keep UI responsive.
  • SIM_MAX_ROWS cap (50) to prevent widget explosion.
  • Respects active sector filter from lu_shared.
"""

import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import re

from lu_core import GENERAL_CLIENT, _RISK_ORDER, _fmt_value, get_high_risk_industries
from lu_shared import (
    F, FF, _bind_mousewheel,
    _NAVY_DEEP, _NAVY_MID, _NAVY_LIGHT, _NAVY_MIST, _NAVY_GHOST, _NAVY_PALE,
    _WHITE, _CARD_WHITE, _OFF_WHITE, _BORDER_LIGHT, _BORDER_MID,
    _TXT_NAVY, _TXT_SOFT, _TXT_MUTED, _TXT_ON_LIME,
    _LIME_MID, _LIME_DARK, _LIME_PALE,
    _ACCENT_RED, _ACCENT_GOLD, _ACCENT_SUCCESS,
    _RISK_COLOR, _RISK_BG, _RISK_BADGE_BG,
    _lu_filter_data_by_query,
    _lu_get_active_sectors, _lu_get_filtered_all_data,
)

# ── Tuneable constants ──────────────────────────────────────────────
SIM_MAX_ROWS       = 50   # max expense rows shown
SIM_CHUNK_SIZE     = 10   # rows built per after() tick
SIM_CHART_MAX_BARS = 20   # max expense rows considered for chart
PIE_MAX_SLICES     = 10   # pie slices before aggregating to "Other"
SIM_TABLE_COLUMNS = (
    # (title, min_width_px, weight)
    ("Expense Item", 220, 5),
    ("Risk", 72, 1),
    ("Base Amount", 120, 2),
    ("Inflation Rate (%)", 100, 2),
    ("Extra Cost", 120, 2),
    ("Simulated", 120, 2),
)

try:
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    _HAS_MPL = True
except ImportError:
    _HAS_MPL = False


# ══════════════════════════════════════════════════════════════════════
#  PANEL BUILDER (sets fixed canvas height)
# ══════════════════════════════════════════════════════════════════════

def _build_simulator_panel(self, parent):
    hdr = tk.Frame(parent, bg=_NAVY_MID, height=38)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)
    self._sim_hdr_lbl = tk.Label(
        hdr, text="⚙️  Inflation / Cost-Shock Simulator",
        font=F(10, "bold"), fg=_WHITE, bg=_NAVY_MID)
    self._sim_hdr_lbl.pack(side="left", padx=20, pady=8)

    ctrl = tk.Frame(parent, bg=_OFF_WHITE, height=58)
    ctrl.pack(fill="x")
    ctrl.pack_propagate(False)
    tk.Label(ctrl, text="Global %:", font=F(11, "bold"),
            fg=_NAVY_MID, bg=_OFF_WHITE).pack(side="left", padx=(16, 4), pady=14)
    self._sim_global_var = tk.StringVar(value="0")
    ctk.CTkEntry(ctrl, textvariable=self._sim_global_var, width=84, height=32,
                corner_radius=5, fg_color=_WHITE, text_color=_TXT_NAVY,
                border_color=_BORDER_MID, font=FF(11)
                ).pack(side="left", pady=13)
    ctk.CTkButton(ctrl, text="Apply All", command=lambda: _sim_apply_global(self),
                width=96, height=32, corner_radius=5,
                fg_color=_NAVY_LIGHT, hover_color=_NAVY_MID,
                text_color=_WHITE, font=FF(10, "bold")
                ).pack(side="left", padx=6, pady=13)
    ctk.CTkButton(ctrl, text="Reset", command=lambda: _sim_reset(self),
                width=84, height=32, corner_radius=5,
                fg_color=_ACCENT_RED, hover_color="#C53030",
                text_color=_WHITE, font=FF(10, "bold")
                ).pack(side="left", padx=(0, 6), pady=13)
    tk.Label(ctrl, text="Filter:", font=F(10, "bold"),
            fg=_NAVY_MID, bg=_OFF_WHITE).pack(side="left", padx=(14, 4), pady=14)
    self._sim_search_var = tk.StringVar()
    sim_search_entry = ctk.CTkEntry(
        ctrl, textvariable=self._sim_search_var, width=260, height=32, corner_radius=5,
        fg_color=_WHITE, text_color=_TXT_NAVY, border_color=_BORDER_MID, font=FF(13),
        placeholder_text="client, ID, PN, industry, sector..."
    )
    sim_search_entry.pack(side="left", pady=10)
    sim_search_entry.bind(
        "<Return>",
        lambda _e: _sim_populate(self) if getattr(self, "_lu_all_data", None) else None,
    )
    self._sim_match_lbl = tk.Label(
        ctrl, text="", font=F(8, "bold"), fg=_WHITE, bg=_OFF_WHITE, padx=8, pady=3)
    self._sim_match_lbl.pack(side="left", padx=(8, 0), pady=10)
    self._sim_industry_filter_btn = ctk.CTkButton(
        ctrl,
        text="Industry Checklist",
        command=lambda: _sim_open_industry_checklist(self),
        width=150,
        height=32,
        corner_radius=5,
        fg_color=_NAVY_LIGHT,
        hover_color=_NAVY_MID,
        text_color=_WHITE,
        font=FF(10, "bold"),
    )
    self._sim_industry_filter_btn.pack(side="left", padx=(8, 0), pady=10)
    self._sim_industry_filter_lbl = tk.Label(
        ctrl, text="", font=F(8, "bold"), fg=_TXT_SOFT, bg=_OFF_WHITE
    )
    self._sim_industry_filter_lbl.pack(side="left", padx=(8, 0), pady=10)
    tk.Frame(parent, bg=_BORDER_LIGHT, height=1).pack(fill="x")

    cards_frame = tk.Frame(parent, bg=_NAVY_MIST)
    cards_frame.pack(fill="x")
    _build_sim_summary_cards(self, cards_frame)

    inc_bar = tk.Frame(parent, bg=_NAVY_DEEP, height=38)
    inc_bar.pack(fill="x")
    inc_bar.pack_propagate(False)
    self._sim_income_lbl = tk.Label(
        inc_bar, text="TOTAL SOURCE OF INCOME  —  Load a file to begin",
        font=F(9, "bold"), fg=_TXT_MUTED, bg=_NAVY_DEEP)
    self._sim_income_lbl.pack(side="left", padx=20, pady=10)
    self._sim_surplus_lbl = tk.Label(inc_bar, text="", font=F(9, "bold"),
                                     fg=_LIME_MID, bg=_NAVY_DEEP)
    self._sim_surplus_lbl.pack(side="right", padx=20, pady=10)
    tk.Frame(parent, bg=_BORDER_LIGHT, height=1).pack(fill="x")

    split = tk.Frame(parent, bg=_CARD_WHITE)
    split.pack(fill="both", expand=True)

    left_frame = tk.Frame(split, bg=_CARD_WHITE)
    left_frame.pack(side="left", fill="both", expand=True)

    # Scrollable expense list
    sim_sb = tk.Scrollbar(left_frame, relief="flat",
                          troughcolor=_OFF_WHITE, bg=_BORDER_LIGHT, width=8, bd=0)
    sim_sb.pack(side="right", fill="y")
    self._sim_canvas = tk.Canvas(left_frame, bg=_CARD_WHITE, highlightthickness=0,
                                 yscrollcommand=sim_sb.set)
    self._sim_canvas.pack(side="left", fill="both", expand=True)
    sim_sb.config(command=self._sim_canvas.yview)
    self._sim_scroll_frame = tk.Frame(self._sim_canvas, bg=_CARD_WHITE)
    self._sim_canvas_win   = self._sim_canvas.create_window(
        (0, 0), window=self._sim_scroll_frame, anchor="nw")
    self._sim_scroll_frame.bind(
        "<Configure>",
        lambda e: self._sim_canvas.configure(scrollregion=self._sim_canvas.bbox("all")))
    self._sim_canvas.bind(
        "<Configure>",
        lambda e: self._sim_canvas.itemconfig(self._sim_canvas_win, width=e.width))
    _bind_mousewheel(self._sim_canvas)

    # Pie chart panel (matplotlib) — shares use simulated amounts after slider %
    right_frame = tk.Frame(split, bg=_CARD_WHITE,
                           highlightbackground=_BORDER_MID, highlightthickness=1,
                           width=390)
    right_frame.pack(side="right", fill="y")
    right_frame.pack_propagate(False)
    tk.Label(right_frame, text="Expense Breakdown", font=F(9, "bold"),
             fg=_TXT_SOFT, bg=_CARD_WHITE).pack(pady=(8, 0))
    tk.Label(
        right_frame,
        text="Percentage of Total Simulated Expenses",
        font=F(7), fg=_TXT_MUTED, bg=_CARD_WHITE,
    ).pack(pady=(0, 4))
    self._sim_chart_holder = tk.Frame(right_frame, bg=_CARD_WHITE)
    self._sim_chart_holder.pack(fill="both", expand=True, padx=4, pady=(0, 8))

    self._sim_sliders    = {}
    self._sim_expenses   = []
    self._sim_build_job  = None
    self._sim_net_income = 0.0
    self._sim_selected_industries = set()
    self._sim_manual_industries = set()
    _sim_show_placeholder(self)
def _sim_filter_data_by_industry_checklist(all_data, selected_industries):
    selected = {str(x).strip().lower() for x in (selected_industries or set()) if str(x).strip()}
    if not selected:
        return all_data

    splitter = re.compile(r"\s*(?:,|/|;|&|\band\b)\s*", re.I)

    def _industry_tokens(rec: dict) -> set[str]:
        tags = rec.get("industry_tags") or []
        if tags:
            return {str(x).strip().lower() for x in tags if str(x).strip()}
        raw = str((rec or {}).get("industry") or "").strip()
        if not raw:
            return set()
        return {tok.strip().lower() for tok in splitter.split(raw) if tok.strip()}

    base_general = list((all_data or {}).get("general", []))
    kept_general = [
        rec for rec in base_general
        if _industry_tokens(rec) & selected
    ]
    kept_clients = {str((r or {}).get("client") or ""): r for r in kept_general if (r or {}).get("client")}

    patched = dict(all_data or {})
    patched["general"] = kept_general
    patched["clients"] = kept_clients
    return patched


def _sim_open_industry_checklist(self):
    all_data = getattr(self, "_lu_all_data", None) or {}
    base_industries = {str(x).strip() for x in all_data.get("unique_industries", []) if str(x).strip()}
    manual_industries = set(getattr(self, "_sim_manual_industries", set()) or set())
    industries = sorted(base_industries | manual_industries, key=str.lower)
    if not industries:
        messagebox.showwarning("No Data", "Load and run LU analysis first.")
        return

    selected = set(getattr(self, "_sim_selected_industries", set()) or set())
    high_defaults = {str(x).strip().lower() for x in get_high_risk_industries() if str(x).strip()}
    if not selected:
        selected = {name for name in industries if name.lower() in high_defaults}

    dialog = ctk.CTkToplevel(self)
    dialog.title("Risk Simulator Industry Checklist")
    dialog.geometry("620x620")
    dialog.minsize(520, 480)
    dialog.transient(self)
    dialog.grab_set()
    dialog.configure(fg_color=_CARD_WHITE)

    hdr = tk.Frame(dialog, bg=_NAVY_MID, height=52)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)
    tk.Label(
        hdr,
        text="☑  Risk Simulator Industry Checklist",
        font=F(11, "bold"),
        fg=_WHITE,
        bg=_NAVY_MID,
    ).pack(side="left", padx=16, pady=12)

    info = tk.Frame(dialog, bg=_NAVY_MIST, highlightbackground=_BORDER_MID, highlightthickness=1)
    info.pack(fill="x", padx=16, pady=(10, 6))
    tk.Label(
        info,
        text=(
            "Use checkboxes to filter the Risk Simulator by industry. "
            "By default, this follows HIGH industries from Risk Settings."
        ),
        font=F(8),
        fg=_TXT_SOFT,
        bg=_NAVY_MIST,
        anchor="w",
        justify="left",
    ).pack(fill="x", padx=10, pady=8)

    search_row = tk.Frame(dialog, bg=_CARD_WHITE)
    search_row.pack(fill="x", padx=16, pady=(4, 6))
    tk.Label(search_row, text="🔍", font=F(10), fg=_TXT_SOFT, bg=_CARD_WHITE).pack(side="left")
    search_var = tk.StringVar(value="")
    search_entry = ctk.CTkEntry(
        search_row,
        textvariable=search_var,
        width=380,
        height=28,
        corner_radius=4,
        fg_color=_WHITE,
        text_color=_TXT_NAVY,
        border_color=_BORDER_MID,
        font=FF(9),
        placeholder_text="Search industry...",
    )
    search_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

    add_row = tk.Frame(dialog, bg=_CARD_WHITE)
    add_row.pack(fill="x", padx=16, pady=(0, 6))
    tk.Label(
        add_row, text="Add Industry:", font=F(8, "bold"),
        fg=_NAVY_MID, bg=_CARD_WHITE
    ).pack(side="left", padx=(0, 6))
    add_var = tk.StringVar(value="")
    add_entry = ctk.CTkEntry(
        add_row,
        textvariable=add_var,
        width=260,
        height=28,
        corner_radius=4,
        fg_color=_WHITE,
        text_color=_TXT_NAVY,
        border_color=_BORDER_MID,
        font=FF(9),
        placeholder_text="e.g. Logistics",
    )
    add_entry.pack(side="left", padx=(0, 6))
    add_hint_lbl = tk.Label(add_row, text="", font=F(7), fg=_TXT_MUTED, bg=_CARD_WHITE)
    add_hint_lbl.pack(side="left", padx=(4, 0))

    list_wrap = tk.Frame(dialog, bg=_CARD_WHITE)
    list_wrap.pack(fill="both", expand=True, padx=16, pady=(0, 8))
    sb = tk.Scrollbar(list_wrap, relief="flat", troughcolor=_OFF_WHITE, bg=_BORDER_LIGHT, width=8, bd=0)
    sb.pack(side="right", fill="y")
    canvas = tk.Canvas(list_wrap, bg=_CARD_WHITE, highlightthickness=0, yscrollcommand=sb.set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.config(command=canvas.yview)
    rows_frame = tk.Frame(canvas, bg=_CARD_WHITE)
    win = canvas.create_window((0, 0), window=rows_frame, anchor="nw")
    rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
    canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", lambda ev: canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")))
    canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

    row_widgets = []
    var_map = {}
    row_by_industry = {}

    def _add_industry_row(industry: str, preselect: bool = False):
        if industry in var_map:
            if preselect:
                var_map[industry].set(True)
            return
        idx = len(row_widgets)
        row_bg = _WHITE if idx % 2 == 0 else _OFF_WHITE
        row = tk.Frame(rows_frame, bg=row_bg)
        row.pack(fill="x")
        tk.Frame(row, bg=_BORDER_LIGHT, height=1).pack(fill="x")
        inner = tk.Frame(row, bg=row_bg)
        inner.pack(fill="x", padx=8, pady=4)
        var = tk.BooleanVar(value=preselect)
        var_map[industry] = var
        chk = tk.Checkbutton(
            inner,
            text=industry,
            variable=var,
            onvalue=True,
            offvalue=False,
            font=F(9),
            fg=_TXT_NAVY,
            bg=row_bg,
            activebackground=row_bg,
            anchor="w",
            justify="left",
            relief="flat",
            highlightthickness=0,
        )
        chk.pack(side="left", fill="x", expand=True, padx=(4, 0))
        row_widgets.append((row, industry))
        row_by_industry[industry] = row

    for industry in industries:
        _add_industry_row(industry, preselect=(industry in selected))

    def _apply_search(*_args):
        q = search_var.get().strip().lower()
        for row, industry in row_widgets:
            show = (not q) or (q in industry.lower())
            if show and not row.winfo_ismapped():
                row.pack(fill="x")
            elif (not show) and row.winfo_ismapped():
                row.pack_forget()

    search_var.trace_add("write", _apply_search)

    def _add_manual_industry():
        name = add_var.get().strip()
        if not name:
            add_hint_lbl.config(text="Enter an industry name.", fg=_ACCENT_RED)
            return
        existing = {k.lower(): k for k in var_map.keys()}
        if name.lower() in existing:
            _add_industry_row(existing[name.lower()], preselect=True)
            add_hint_lbl.config(text="Already exists; selected.", fg=_ACCENT_SUCCESS)
        else:
            _add_industry_row(name, preselect=True)
            add_hint_lbl.config(text=f"Added '{name}'.", fg=_ACCENT_SUCCESS)
        add_var.set("")
        _apply_search()
        add_entry.focus_set()

    def _remove_manual_industry():
        name = add_var.get().strip()
        if not name:
            add_hint_lbl.config(text="Enter an industry to remove.", fg=_ACCENT_RED)
            return
        existing = {k.lower(): k for k in var_map.keys()}
        canonical = existing.get(name.lower())
        if not canonical:
            add_hint_lbl.config(text="Industry not found.", fg=_ACCENT_RED)
            return
        if canonical in base_industries:
            add_hint_lbl.config(text="Cannot remove base industry from data.", fg=_ACCENT_RED)
            return
        row = row_by_industry.pop(canonical, None)
        if row is not None:
            try:
                row.destroy()
            except Exception:
                pass
        row_widgets[:] = [(r, n) for (r, n) in row_widgets if n != canonical]
        var_map.pop(canonical, None)
        selected.discard(canonical)
        add_hint_lbl.config(text=f"Removed '{canonical}'.", fg=_ACCENT_SUCCESS)
        add_var.set("")
        _apply_search()
        add_entry.focus_set()

    tk.Button(
        add_row,
        text="Add",
        font=F(8, "bold"),
        fg=_WHITE,
        bg=_NAVY_MID,
        activebackground=_NAVY_LIGHT,
        activeforeground=_WHITE,
        relief="flat",
        bd=0,
        padx=10,
        pady=5,
        cursor="hand2",
        command=_add_manual_industry,
    ).pack(side="left")
    tk.Button(
        add_row,
        text="Remove",
        font=F(8, "bold"),
        fg=_WHITE,
        bg=_ACCENT_RED,
        activebackground="#C53030",
        activeforeground=_WHITE,
        relief="flat",
        bd=0,
        padx=10,
        pady=5,
        cursor="hand2",
        command=_remove_manual_industry,
    ).pack(side="left", padx=(4, 0))
    add_entry.bind("<Return>", lambda _e: _add_manual_industry())

    def _set_all(v: bool):
        for _name, vv in var_map.items():
            vv.set(v)

    def _use_high_defaults():
        for name, vv in var_map.items():
            vv.set(name.lower() in high_defaults)

    def _apply_and_close():
        chosen = {name for name, vv in var_map.items() if vv.get()}
        self._sim_selected_industries = chosen
        self._sim_manual_industries = set(var_map.keys()) - base_industries
        _sim_populate(self)
        dialog.destroy()

    footer = tk.Frame(dialog, bg=_OFF_WHITE, highlightbackground=_BORDER_MID, highlightthickness=1)
    footer.pack(fill="x", padx=16, pady=(2, 14))
    tk.Button(
        footer,
        text="Select All",
        font=F(8, "bold"),
        fg=_TXT_NAVY,
        bg=_WHITE,
        relief="flat",
        bd=0,
        padx=10,
        pady=6,
        cursor="hand2",
        command=lambda: _set_all(True),
    ).pack(side="left", padx=(12, 4), pady=8)
    tk.Button(
        footer,
        text="Clear All",
        font=F(8, "bold"),
        fg=_TXT_SOFT,
        bg=_WHITE,
        relief="flat",
        bd=0,
        padx=10,
        pady=6,
        cursor="hand2",
        command=lambda: _set_all(False),
    ).pack(side="left", padx=4, pady=8)
    tk.Button(
        footer,
        text="Use HIGH from Settings",
        font=F(8, "bold"),
        fg=_ACCENT_RED,
        bg="#FFE8E8",
        relief="flat",
        bd=0,
        padx=10,
        pady=6,
        cursor="hand2",
        command=_use_high_defaults,
    ).pack(side="left", padx=4, pady=8)
    tk.Button(
        footer,
        text="Cancel",
        font=F(9),
        fg=_TXT_SOFT,
        bg=_OFF_WHITE,
        relief="flat",
        bd=0,
        padx=10,
        pady=8,
        cursor="hand2",
        command=dialog.destroy,
    ).pack(side="right", padx=(0, 4), pady=8)
    tk.Button(
        footer,
        text="  ✔  Apply Filter  ",
        font=F(9, "bold"),
        fg=_WHITE,
        bg=_NAVY_MID,
        activebackground=_NAVY_LIGHT,
        activeforeground=_WHITE,
        relief="flat",
        bd=0,
        padx=14,
        pady=8,
        cursor="hand2",
        command=_apply_and_close,
    ).pack(side="right", padx=12, pady=8)



def _build_sim_summary_cards(self, parent):
    for title, attr, color in [
        ("Total Source of Income", "_sim_lbl_income",  _ACCENT_SUCCESS),
        ("Base Total Expenses",    "_sim_lbl_base",    _TXT_NAVY),
        ("Simulated Total",        "_sim_lbl_sim",     _TXT_NAVY),
        ("Total Increase (₱)",     "_sim_lbl_inc",     _ACCENT_RED),
        ("Surplus / Deficit",      "_sim_lbl_surplus", _ACCENT_SUCCESS),
    ]:
        card = tk.Frame(parent, bg=_NAVY_MIST,
                        highlightbackground="#D6E4F7", highlightthickness=1)
        card.pack(side="left", fill="x", expand=True, padx=6, pady=8)
        tk.Label(card, text=title, font=F(7), fg=_TXT_SOFT, bg=_NAVY_MIST
                 ).pack(anchor="w", padx=10, pady=(6, 0))
        lbl = tk.Label(card, text="—", font=F(13, "bold"), fg=color, bg=_NAVY_MIST)
        lbl.pack(anchor="w", padx=10, pady=(0, 6))
        setattr(self, attr, lbl)


# ══════════════════════════════════════════════════════════════════════
#  PLACEHOLDER
# ══════════════════════════════════════════════════════════════════════

def _sim_show_placeholder(self):
    for w in self._sim_scroll_frame.winfo_children():
        w.destroy()
    tk.Label(self._sim_scroll_frame,
             text="Run an analysis first to unlock the simulator.",
             font=F(10), fg=_TXT_MUTED, bg=_CARD_WHITE).pack(pady=60)
    _sim_draw_chart(self)


# ══════════════════════════════════════════════════════════════════════
#  POPULATE  (chunked — FIX 2)
# ══════════════════════════════════════════════════════════════════════

def _sim_populate(self):
    if not hasattr(self, "_sim_hdr_lbl") or not self._sim_hdr_lbl.winfo_exists():
        return

    # Cancel any in-progress build
    if getattr(self, "_sim_build_job", None):
        try:
            self._sim_hdr_lbl.after_cancel(self._sim_build_job)
        except Exception:
            pass
        self._sim_build_job = None

    filtered_data  = _lu_get_filtered_all_data(self)
    selected_inds  = getattr(self, "_sim_selected_industries", set()) or set()
    filtered_data  = _sim_filter_data_by_industry_checklist(filtered_data, selected_inds)
    q              = getattr(self, "_sim_search_var", tk.StringVar(value="")).get().strip()
    filtered_data  = _lu_filter_data_by_query(filtered_data, q)
    match_count    = len(filtered_data.get("general", []))
    match_lbl      = getattr(self, "_sim_match_lbl", None)
    filter_lbl     = getattr(self, "_sim_industry_filter_lbl", None)
    if filter_lbl is not None:
        if selected_inds:
            selected_sorted = sorted(selected_inds, key=str.lower)
            preview = ", ".join(selected_sorted[:2])
            extra_count = len(selected_sorted) - 2
            suffix = f" (+{extra_count} more)" if extra_count > 0 else ""
            filter_lbl.config(text=f"Industry filter: {preview}{suffix}")
        else:
            filter_lbl.config(text="")
    if match_lbl is not None:
        if q:
            client_names = sorted({
                (r.get("client") or "").strip()
                for r in filtered_data.get("general", [])
                if r.get("client")
            })
            if len(client_names) == 1:
                match_lbl.config(text=client_names[0][:28], bg="#4A6FA5")
            else:
                match_lbl.config(text=f"{match_count} CLIENTS MATCHED", bg="#4A6FA5")
        else:
            match_lbl.config(text="", bg=_OFF_WHITE)
    active_sectors = _lu_get_active_sectors(self)
    client         = self._lu_active_client
    is_general     = (client == GENERAL_CLIENT)
    all_clients    = filtered_data.get("clients", {})

    # If the search narrows to a single client (or exact client key match),
    # treat it as a per-client simulator view so expenses always show.
    chosen_client = None
    if q:
        ql = q.strip().lower()
        exact = next((name for name in all_clients.keys() if str(name).strip().lower() == ql), None)
        if exact:
            chosen_client = exact
        else:
            client_names = sorted({
                (r.get("client") or "").strip()
                for r in filtered_data.get("general", [])
                if r.get("client")
            })
            if len(client_names) == 1:
                chosen_client = client_names[0]

    if chosen_client and chosen_client in all_clients:
        try:
            self._lu_active_client = chosen_client
        except Exception:
            pass
        recs = [all_clients[chosen_client]]
    else:
        recs = (list(all_clients.values())
                if (is_general or active_sectors)
                else ([all_clients[client]] if client in all_clients else []))

    # Update header
    if q:
        self._sim_hdr_lbl.config(
            text=f"⚙️  Simulator — Search: {q[:30]}",
            fg=_LIME_MID)
    elif active_sectors:
        self._sim_hdr_lbl.config(
            text=f"⚙️  Simulator — Filtered: {' · '.join(active_sectors)}",
            fg=_LIME_MID)
    else:
        self._sim_hdr_lbl.config(text="⚙️  Inflation / Cost-Shock Simulator", fg=_WHITE)

    net_income           = sum((r.get("total_source") or 0) for r in recs)
    self._sim_net_income = net_income

    accumulated: dict = {}
    for rec in recs:
        for exp in rec.get("expenses", []):
            name = str((exp or {}).get("name") or "").strip()
            if not name:
                continue
            try:
                total = float((exp or {}).get("total") or 0.0)
            except (TypeError, ValueError):
                total = 0.0
            if total <= 0:
                continue
            risk = str((exp or {}).get("risk") or "LOW").upper()
            if risk not in _RISK_ORDER:
                risk = "LOW"
            reason = str((exp or {}).get("reason") or "").strip()
            if name not in accumulated:
                accumulated[name] = {
                    "name": name,
                    "total": total,
                    "risk": risk,
                    "reason": reason,
                    "value_str": _fmt_value([total]),
                }
            else:
                accumulated[name]["total"] += total
                if _RISK_ORDER.get(risk, 9) < _RISK_ORDER.get(accumulated[name]["risk"], 9):
                    accumulated[name]["risk"] = risk
                    accumulated[name]["reason"] = reason
                accumulated[name]["value_str"] = _fmt_value([accumulated[name]["total"]])

    all_expenses = sorted(accumulated.values(),
                          key=lambda e: _RISK_ORDER.get(e["risk"], 9))

    # Apply row cap
    capped = False
    if len(all_expenses) > SIM_MAX_ROWS:
        all_expenses = all_expenses[:SIM_MAX_ROWS]
        capped       = True

    self._sim_expenses = all_expenses
    self._sim_sliders  = {}

    for w in list(self._sim_scroll_frame.winfo_children()):
        try:
            w.destroy()
        except Exception:
            pass

    if not all_expenses:
        tk.Label(self._sim_scroll_frame,
                 text="No numeric expense data found.",
                 font=F(9), fg=_TXT_MUTED, bg=_CARD_WHITE, justify="center").pack(pady=60)
        return

    if capped:
        tk.Label(self._sim_scroll_frame,
                 text=f"ℹ  Showing top {SIM_MAX_ROWS} expense rows (file has more).",
                 font=F(8), fg=_ACCENT_GOLD, bg=_OFF_WHITE,
                 padx=10, pady=4).pack(fill="x")

    hdr = tk.Frame(self._sim_scroll_frame, bg=_OFF_WHITE)
    hdr.pack(fill="x", pady=(8, 0))
    for col, (_title, min_px, _wt) in enumerate(SIM_TABLE_COLUMNS):
        hdr.grid_columnconfigure(col, weight=1, minsize=min_px, uniform="sim_col")
    for col, (text, _min_px, _wt) in enumerate(SIM_TABLE_COLUMNS):
        tk.Label(hdr, text=text, font=F(8, "bold"), fg=_NAVY_PALE, bg=_OFF_WHITE,
                 anchor="w" if col == 0 else "center",
                 justify="left" if col == 0 else "center",
                 padx=6, pady=5
                 ).grid(row=0, column=col, sticky="ew", padx=(0, 2))
    tk.Frame(self._sim_scroll_frame, bg=_BORDER_MID, height=1).pack(fill="x")

    # Pre‑create all DoubleVar objects (fast)
    for exp in all_expenses:
        var = tk.DoubleVar(value=0.0)
        self._sim_sliders[exp["name"]] = var

    # Chunked build
    built_rows = {"count": 0}
    def _build_chunk(start: int):
        chunk = all_expenses[start: start + SIM_CHUNK_SIZE]
        for idx, exp in enumerate(chunk, start=start):
            try:
                var = self._sim_sliders[exp["name"]]
                _sim_build_expense_row(self, self._sim_scroll_frame, exp, var, idx)
                built_rows["count"] += 1
            except Exception:
                continue
        next_start = start + SIM_CHUNK_SIZE
        if next_start < len(all_expenses):
            self._sim_build_job = self._sim_scroll_frame.after(
                16, lambda s=next_start: _build_chunk(s))
        else:
            self._sim_build_job = None
            if built_rows["count"] == 0 and all_expenses:
                tk.Label(
                    self._sim_scroll_frame,
                    text="Unable to render expense rows for this selection.",
                    font=F(9), fg=_ACCENT_RED, bg=_CARD_WHITE, justify="center"
                ).pack(pady=20)
            _sim_refresh(self)

    _build_chunk(0)


# ══════════════════════════════════════════════════════════════════════
#  ROW BUILDER
# ══════════════════════════════════════════════════════════════════════

def _sim_build_expense_row(self, parent, exp, var, idx):
    risk   = str(exp.get("risk") or "LOW").upper()
    if risk not in _RISK_ORDER:
        risk = "LOW"
    name = str(exp.get("name") or "Unnamed Expense")
    try:
        base_total = float(exp.get("total") or 0.0)
    except (TypeError, ValueError):
        base_total = 0.0
    row_bg = _RISK_BG.get(risk, _WHITE) if idx % 2 == 0 else _WHITE
    row    = tk.Frame(parent, bg=row_bg)
    row.pack(fill="x")
    for ci, (_title, min_px, _wt) in enumerate(SIM_TABLE_COLUMNS):
        row.grid_columnconfigure(ci, weight=1, minsize=min_px, uniform="sim_col")

    tk.Label(row, text=name, font=F(9, "bold"),
             fg=_TXT_NAVY, bg=row_bg, anchor="w", padx=8, pady=6
             ).grid(row=0, column=0, sticky="ew")
    tk.Label(row, text=risk, font=F(7, "bold"),
             fg=_RISK_COLOR.get(risk, _TXT_SOFT),
             bg=_RISK_BADGE_BG.get(risk, _OFF_WHITE),
             anchor="center", justify="center",
             padx=10, pady=3).grid(row=0, column=1, padx=4, pady=6, sticky="")
    tk.Label(row, text=f"₱{base_total:,.2f}" if base_total > 0 else "—",
             font=F(9), fg=_TXT_NAVY, bg=row_bg, anchor="center", justify="center", padx=6
             ).grid(row=0, column=2, sticky="ew")

    rate_entry = ctk.CTkEntry(row, textvariable=var, width=80, height=26, corner_radius=4,
                              font=FF(9), fg_color=_WHITE, text_color=_TXT_NAVY,
                              border_color=_RISK_COLOR.get(risk, _BORDER_MID),
                              placeholder_text="0")
    rate_entry.grid(row=0, column=3, padx=8, pady=6)
    rate_entry.bind("<Return>",   lambda e, ex=exp: _sim_on_slide(self, ex, var.get()))
    rate_entry.bind("<FocusOut>", lambda e, ex=exp: _sim_on_slide(self, ex, var.get()))

    extra_lbl = tk.Label(row, text="—", font=F(9), fg=_ACCENT_RED,
                         bg=row_bg, anchor="center", justify="center", padx=6)
    extra_lbl.grid(row=0, column=4, sticky="ew")
    sim_lbl = tk.Label(row, text="—", font=F(9, "bold"), fg=_TXT_NAVY,
                       bg=row_bg, anchor="center", justify="center", padx=6)
    sim_lbl.grid(row=0, column=5, sticky="ew")

    var._extra_lbl = extra_lbl
    var._sim_lbl   = sim_lbl
    var._base      = base_total
    tk.Frame(parent, bg=_BORDER_LIGHT, height=1).pack(fill="x")


# ══════════════════════════════════════════════════════════════════════
#  INTERACTION CALLBACKS
# ══════════════════════════════════════════════════════════════════════

def _sim_on_slide(self, exp, value):
    try:
        pct = float(value)
    except (ValueError, TypeError):
        pct = 0.0
    if pct < 0.0:
        pct = 0.0
    self._sim_sliders[exp["name"]].set(str(pct))
    _sim_refresh(self)


def _sim_apply_global(self):
    try:
        pct = float(self._sim_global_var.get())
    except (ValueError, TypeError):
        pct = 0.0
    if pct < 0.0:
        pct = 0.0
    if not self._sim_sliders and self._lu_all_data:
        _sim_populate(self)
    for var in self._sim_sliders.values():
        var.set(str(pct))
    _sim_refresh(self)


def _sim_reset(self):
    self._sim_global_var.set("0")
    if not self._sim_sliders and self._lu_all_data:
        _sim_populate(self)
    for var in self._sim_sliders.values():
        var.set("0")
    _sim_refresh(self)


def _sim_refresh(self):
    base_total = sim_total = 0.0
    for exp in getattr(self, "_sim_expenses", []):
        pct = 0.0
        var = self._sim_sliders.get(exp["name"])
        if var:
            try:
                pct = float(var.get())
            except (ValueError, TypeError):
                pass
        base  = exp["total"]
        extra = base * pct / 100.0
        sim   = base + extra
        base_total += base
        sim_total  += sim
        if var and hasattr(var, "_extra_lbl"):
            try:
                var._extra_lbl.config(
                    text=f"+₱{extra:,.2f}" if extra > 0 else "—",
                    fg=_ACCENT_RED if extra > 0 else _TXT_MUTED)
                var._sim_lbl.config(
                    text=f"₱{sim:,.2f}" if base > 0 else "—", fg=_TXT_NAVY)
            except Exception:
                pass

    increase   = sim_total - base_total
    net_income = getattr(self, "_sim_net_income", 0.0) or 0.0
    surplus    = net_income - sim_total

    if hasattr(self, "_sim_lbl_base"):
        try:
            self._sim_lbl_income.config(
                text=f"₱{net_income:,.2f}" if net_income else "—",
                fg=_ACCENT_SUCCESS)
            self._sim_lbl_base.config(
                text=f"₱{base_total:,.2f}" if base_total else "—")
            self._sim_lbl_sim.config(
                text=f"₱{sim_total:,.2f}" if base_total else "—")
            self._sim_lbl_inc.config(
                text=f"+₱{increase:,.2f}" if increase > 0 else "₱0.00",
                fg=_ACCENT_RED if increase > 0 else _TXT_NAVY)
            if net_income:
                surplus_txt = f"{'▲' if surplus >= 0 else '▼'} ₱{abs(surplus):,.2f}"
                self._sim_lbl_surplus.config(
                    text=surplus_txt,
                    fg=_ACCENT_SUCCESS if surplus >= 0 else _ACCENT_RED)
            else:
                self._sim_lbl_surplus.config(text="—", fg=_TXT_MUTED)
        except Exception:
            pass

    if hasattr(self, "_sim_income_lbl"):
        try:
            if net_income:
                self._sim_income_lbl.config(
                    text=f"TOTAL SOURCE OF INCOME  ₱{net_income:,.2f}",
                    fg=_LIME_MID)
                self._sim_surplus_lbl.config(
                    text=(f"SURPLUS  ₱{surplus:,.2f}" if surplus >= 0
                          else f"DEFICIT  ▲ ₱{abs(surplus):,.2f}"),
                    fg=_LIME_MID if surplus >= 0 else _ACCENT_RED)
            else:
                self._sim_income_lbl.config(
                    text="TOTAL SOURCE OF INCOME  —  Load a file to begin",
                    fg=_TXT_MUTED)
                self._sim_surplus_lbl.config(text="", fg=_LIME_MID)
        except Exception:
            pass

    _sim_draw_chart(self)


# ══════════════════════════════════════════════════════════════════════
#  PIE CHART  (simulated expense mix — % of total)
# ══════════════════════════════════════════════════════════════════════

def _sim_draw_chart(self):
    holder = getattr(self, "_sim_chart_holder", None)
    if holder is None:
        return
    try:
        if not holder.winfo_exists():
            return
    except Exception:
        return

    for w in holder.winfo_children():
        w.destroy()

    expenses = [e for e in getattr(self, "_sim_expenses", []) if e["total"] > 0]
    if not expenses:
        tk.Label(
            holder,
            text="No numeric data\nto chart.",
            font=F(9),
            fg=_TXT_MUTED,
            bg=_CARD_WHITE,
            justify="center",
        ).pack(pady=40)
        return

    expenses = expenses[:SIM_CHART_MAX_BARS]

    def _sim_amount(exp):
        pct = 0.0
        var = self._sim_sliders.get(exp["name"])
        if var:
            try:
                pct = float(var.get() or 0)
            except Exception:
                pass
        base = float(exp["total"] or 0)
        return max(0.0, base + base * (pct / 100.0))

    pairs = [(e["name"], _sim_amount(e)) for e in expenses]
    pairs.sort(key=lambda x: -x[1])
    names = [p[0] for p in pairs]
    vals = [p[1] for p in pairs]

    if sum(vals) <= 0:
        tk.Label(
            holder,
            text="No simulated amounts\nto chart.",
            font=F(9),
            fg=_TXT_MUTED,
            bg=_CARD_WHITE,
            justify="center",
        ).pack(pady=40)
        return

    if len(pairs) > PIE_MAX_SLICES:
        top = pairs[: PIE_MAX_SLICES - 1]
        other_sum = sum(p[1] for p in pairs[PIE_MAX_SLICES - 1 :])
        names = [p[0] for p in top] + (["Other"] if other_sum > 0 else [])
        vals = [p[1] for p in top] + ([other_sum] if other_sum > 0 else [])

    if not _HAS_MPL:
        lines = [f"{n[:22]}{'…' if len(n) > 22 else ''}: {v/sum(vals)*100:.1f}%"
                 for n, v in zip(names, vals)]
        tk.Label(
            holder,
            text="matplotlib unavailable.\n\n" + "\n".join(lines[:12]),
            font=F(7),
            fg=_TXT_SOFT,
            bg=_CARD_WHITE,
            justify="left",
        ).pack(padx=6, pady=8)
        return

    def _short(n: str, w: int = 18) -> str:
        n = str(n or "").strip()
        return n if len(n) <= w else n[: w - 1] + "…"

    try:
        fig, ax = plt.subplots(figsize=(4.7, 4.9))
        fig.patch.set_facecolor(_CARD_WHITE)
        ax.set_facecolor(_CARD_WHITE)

        colors = [plt.cm.Pastel2(i % 8) for i in range(len(vals))]
        wedges, _texts, autotexts = ax.pie(
            vals,
            labels=None,
            colors=colors,
            startangle=90,
            counterclock=False,
            autopct=lambda p: f"{p:.1f}%" if p >= 4.5 else "",
            pctdistance=0.80,
            textprops={"fontsize": 8, "color": "#243B64", "fontweight": "bold"},
            wedgeprops={"width": 0.44, "linewidth": 1.0, "edgecolor": _CARD_WHITE},
        )
        for t in autotexts:
            t.set_fontsize(8)
        total_sim = sum(vals)
        ax.text(
            0, 0,
            f"Total\nP{total_sim:,.0f}",
            ha="center",
            va="center",
            fontsize=9,
            color="#365B8C",
            fontweight="bold",
        )
        ax.set_title("Share of Total (Simulated)", fontsize=9, color="#4A6FA5", pad=6)

        leg_labels = [_short(n, 22) for n in names]
        ax.legend(
            wedges,
            leg_labels,
            loc="upper center",
            bbox_to_anchor=(0.5, -0.07),
            ncol=2,
            fontsize=6.5,
            frameon=False,
        )
        fig.subplots_adjust(left=0.06, right=0.94, top=0.88, bottom=0.30)

        canvas = FigureCanvasTkAgg(fig, master=holder)
        widget = canvas.get_tk_widget()
        widget.config(width=370, height=370)
        widget.pack_propagate(False)
        widget.pack(fill="none", expand=False)
        plt.close(fig)
    except Exception:
        tk.Label(
            holder,
            text="Could not draw chart.",
            font=F(9),
            fg=_TXT_MUTED,
            bg=_CARD_WHITE,
        ).pack(pady=24)


# ══════════════════════════════════════════════════════════════════════
#  ATTACH
# ══════════════════════════════════════════════════════════════════════

def attach(cls):
    """
    Attach Risk Simulator methods to the app class.
    Call AFTER lu_tab_analysis.attach(cls).
    """
    cls._build_simulator_panel    = _build_simulator_panel
    cls._build_sim_summary_cards  = _build_sim_summary_cards
    cls._sim_show_placeholder     = _sim_show_placeholder
    cls._sim_populate             = _sim_populate
    cls._sim_build_expense_row    = _sim_build_expense_row
    cls._sim_on_slide             = _sim_on_slide
    cls._sim_apply_global         = _sim_apply_global
    cls._sim_reset                = _sim_reset
    cls._sim_refresh              = _sim_refresh
    cls._sim_draw_chart           = _sim_draw_chart
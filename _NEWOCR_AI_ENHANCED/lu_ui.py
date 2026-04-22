"""
lu_ui.py — LU Analysis: main panel orchestrator
=================================================
This file is now a thin orchestrator.  All tab logic lives in
dedicated files:

  lu_shared.py              — shared constants, helpers, filter state
  lu_tab_analysis.py        — Analysis tab (client search + scorecard)
  lu_tab_charts.py          — Charts tab (sector charts)
  lu_simulator_patch.py     — Risk Simulator tab
  lu_loanbal_export_patch.py — Sector vs Loan Balance tab + export
  lu_tab_report.py          — Report tab + PDF/Excel export

How to use
----------
The existing entry point is unchanged:

    import lu_analysis_tab
    lu_analysis_tab.attach(MyAppClass)

Or call attach() directly from this module:

    import lu_ui
    lu_ui.attach(MyAppClass)
"""

import tkinter as tk
import customtkinter as ctk
import re
from pathlib import Path
from tkinter import filedialog, messagebox

from lu_core import (
    GENERAL_CLIENT,
    run_lu_analysis,
    get_high_risk_industries,
    set_high_risk_industries,
    get_product_risk_overrides,
    set_product_risk_overrides,
    get_expense_risk_overrides,
    set_expense_risk_overrides,
    lookup_product_risk_override,
)
import lu_core as _lu_core

# ── Tab modules ────────────────────────────────────────────────────────
import lu_tab_analysis
import lu_tab_charts
import lu_simulator_patch
import lu_loanbal_export_patch
import lu_tab_report

# ── Re-export shared helpers expected by lu_analysis_tab shim ─────────
from lu_shared import (
    F, FF, _bind_mousewheel,
    _NAVY_DEEP, _NAVY_MID, _NAVY_LIGHT, _NAVY_MIST, _NAVY_GHOST, _NAVY_PALE,
    _WHITE, _CARD_WHITE, _OFF_WHITE, _BORDER_LIGHT, _BORDER_MID,
    _TXT_NAVY, _TXT_SOFT, _TXT_MUTED, _TXT_ON_LIME,
    _LIME_MID, _LIME_DARK, _LIME_PALE, _LIME_BRIGHT,
    _ACCENT_RED, _ACCENT_GOLD, _ACCENT_SUCCESS,
    _RISK_COLOR, _RISK_BG, _RISK_BADGE_BG,
    _CLIENT_HERO_BG, _CLIENT_HERO_ACCENT,
    _SECTOR_COLORS, _SECTOR_ICON, _CHART_SECTORS, _ALL_SECTORS,
    _SIM_BAR_BASE, _SIM_BAR_SIM,
    _lu_get_active_sectors, _lu_get_filtered_all_data,
)
from lu_tab_analysis import (
    _lu_update_filter_pill, _lu_clear_sector_filter,
    _lu_on_client_change, _lu_filter_by_search, _lu_populate_client_dropdown,
    _lu_render_results, _lu_render_general_view, _lu_render_client_view,
    _lu_show_placeholder, _lu_show_error, _tv_sort,
)
from lu_tab_charts import _build_charts_panel, _charts_show_placeholder, _charts_render
from lu_simulator_patch import (
    _build_simulator_panel, _build_sim_summary_cards, _sim_show_placeholder,
    _sim_populate, _sim_build_expense_row, _sim_on_slide,
    _sim_apply_global, _sim_reset, _sim_refresh, _sim_draw_chart,
)
from lu_loanbal_export_patch import (
    _build_loanbal_panel, _loanbal_render,
    _loanbal_show_export_menu, _loanbal_export_pdf, _loanbal_export_excel,
    _generate_loanbal_pdf, _generate_loanbal_excel,
)
from lu_tab_report import (
    _build_report_panel, _report_show_placeholder, _report_render, _report_print,
    _lu_show_export_menu, _export_pdf, _export_excel,
    _generate_pdf, _generate_excel,
)


# ══════════════════════════════════════════════════════════════════════
#  MAIN PANEL BUILDER
# ══════════════════════════════════════════════════════════════════════

def _build_lu_analysis_panel(self, parent):
    self._lu_analysis_frame = tk.Frame(parent, bg=_CARD_WHITE)

    # ── Top header bar ─────────────────────────────────────────────────
    hdr_bar = tk.Frame(self._lu_analysis_frame, bg=_NAVY_DEEP, height=56)
    hdr_bar.pack(fill="x")
    hdr_bar.pack_propagate(False)

    hdr_inner = tk.Frame(hdr_bar, bg=_NAVY_DEEP)
    hdr_inner.pack(side="left", fill="y", padx=(28, 0))
    tk.Label(hdr_inner, text="📈  LU Analysis",
             font=F(14, "bold"), fg=_WHITE, bg=_NAVY_DEEP).pack(side="left", anchor="center")
    tk.Label(hdr_inner, text=" — Industry Risk Scanner",
             font=F(9), fg=_TXT_MUTED, bg=_NAVY_DEEP).pack(side="left", anchor="center")

    self._lu_active_view = tk.StringVar(value="analysis")
    tab_frame = tk.Frame(hdr_bar, bg=_NAVY_DEEP)
    tab_frame.pack(side="left", padx=16, fill="y")
    for label, view in [("Analysis",             "analysis"),
                        ("Charts",               "charts"),
                        ("Risk Simulator",        "simulator"),
                        ("Sector vs Loan Balance","loanbal"),
                        ("Report",               "report")]:
        tk.Button(tab_frame, text=label, font=F(8, "bold"),
                  bg=_NAVY_MID, fg=_WHITE,
                  activebackground=_LIME_MID, activeforeground=_TXT_ON_LIME,
                  relief="flat", padx=10, pady=0, cursor="hand2",
                  command=lambda v=view: _lu_switch_view(self, v)
                  ).pack(side="left", padx=2, pady=12, ipady=4)

    # Export button intentionally removed per request.
    self._lu_export_btn = None

    self._lu_load_btn = ctk.CTkButton(
        hdr_bar, text="📂  Load Excel File",
        command=lambda: _lu_browse_file(self),
        width=160, height=34, corner_radius=6,
        fg_color=_LIME_MID, hover_color=_LIME_BRIGHT,
        text_color=_TXT_ON_LIME, font=FF(9, "bold"))
    self._lu_load_btn.pack(side="right", padx=(0, 4), pady=11)

    self._lu_rescan_btn = ctk.CTkButton(
        hdr_bar, text="🔄  Re-Scan",
        command=lambda: _lu_run_analysis(self),
        width=90, height=34, corner_radius=6,
        fg_color=_NAVY_LIGHT, hover_color=_NAVY_MID,
        text_color=_WHITE, font=FF(9, "bold"), state="disabled")
    self._lu_rescan_btn.pack(side="right", padx=(0, 4), pady=11)

    # ── File info strip ────────────────────────────────────────────────
    self._lu_file_strip = tk.Frame(self._lu_analysis_frame, bg=_OFF_WHITE, height=30)
    self._lu_file_strip.pack(fill="x")
    self._lu_file_strip.pack_propagate(False)
    self._lu_file_lbl = tk.Label(
        self._lu_file_strip,
        text="No file loaded  —  click  📂 Load Excel File  to begin",
        font=F(8), fg=_TXT_SOFT, bg=_OFF_WHITE)
    self._lu_file_lbl.pack(side="left", padx=28, pady=6)
    self._lu_status_lbl = tk.Label(self._lu_file_strip, text="",
                                   font=F(8, "bold"), fg=_LIME_DARK, bg=_OFF_WHITE)
    self._lu_status_lbl.pack(side="right", padx=28)
    tk.Frame(self._lu_analysis_frame, bg=_BORDER_LIGHT, height=1).pack(fill="x")

    # ── Client search bar ──────────────────────────────────────────────
    self._lu_client_bar = tk.Frame(self._lu_analysis_frame, bg=_NAVY_MIST, height=50)
    self._lu_client_bar.pack(fill="x")
    self._lu_client_bar.pack_propagate(False)

    self._lu_mode_badge = tk.Label(
        self._lu_client_bar, text="  GENERAL VIEW  ",
        font=F(8, "bold"), fg=_WHITE, bg=_NAVY_MID, padx=10, pady=4)
    self._lu_mode_badge.pack(side="left", padx=(14, 8), pady=12)

    tk.Label(self._lu_client_bar, text="🔍", font=F(11),
             fg=_NAVY_PALE, bg=_NAVY_MIST).pack(side="left", padx=(4, 2), pady=12)
    self._lu_search_var = tk.StringVar()
    self._lu_search_entry = ctk.CTkEntry(
        self._lu_client_bar, textvariable=self._lu_search_var,
        placeholder_text=(
            "Search client, ID, PN, industry, product…  "
            "Try: high risk | low risk  (blank = all)"
        ),
        width=340, height=28, corner_radius=4,
        fg_color=_WHITE, text_color=_TXT_NAVY,
        border_color=_BORDER_MID, font=FF(9))
    self._lu_search_entry.pack(side="left", pady=10)
    self._lu_search_var.trace_add("write", lambda *_: _lu_filter_by_search(self))

    ctk.CTkButton(self._lu_client_bar, text="✕", width=28, height=28,
                  corner_radius=4, fg_color=_BORDER_MID, hover_color=_ACCENT_RED,
                  text_color=_TXT_NAVY, font=FF(9, "bold"),
                  command=lambda: _lu_clear_sector_filter(self)
                  ).pack(side="left", padx=(2, 0), pady=10)

    self._lu_filter_pill      = None
    self._lu_client_var       = tk.StringVar(value=GENERAL_CLIENT)
    self._lu_client_dropdown  = None
    self._lu_client_count_lbl = tk.Label(self._lu_client_bar, text="",
                                         font=F(8), fg=_TXT_SOFT, bg=_NAVY_MIST)
    self._lu_client_count_lbl.pack(side="right", padx=20)
    self._lu_client_bar_divider = tk.Frame(self._lu_analysis_frame, bg=_BORDER_MID, height=1)
    self._lu_client_bar_divider.pack(fill="x")

    # ── View container ─────────────────────────────────────────────────
    self._lu_view_container = tk.Frame(self._lu_analysis_frame, bg=_CARD_WHITE)
    self._lu_view_container.pack(fill="both", expand=True)

    # Build each sub-panel (each tab module creates its own frame)
    lu_tab_analysis._build_analysis_view(self, self._lu_view_container)

    self._lu_charts_view = tk.Frame(self._lu_view_container, bg=_CARD_WHITE)
    _build_charts_panel(self, self._lu_charts_view)

    self._lu_simulator_view = tk.Frame(self._lu_view_container, bg=_CARD_WHITE)
    _build_simulator_panel(self, self._lu_simulator_view)

    self._lu_loanbal_view = tk.Frame(self._lu_view_container, bg=_CARD_WHITE)
    _build_loanbal_panel(self, self._lu_loanbal_view)

    self._lu_report_view = tk.Frame(self._lu_view_container, bg=_CARD_WHITE)
    _build_report_panel(self, self._lu_report_view)

    # State
    self._lu_filepath         = None
    self._lu_results          = []
    self._lu_all_data         = {}
    # Shared (cross-tab) selection/filter state.
    self._lu_active_client    = GENERAL_CLIENT
    self._lu_filtered_sectors = None
    # Analysis-tab-local selection/filter state (does not affect other tabs).
    self._lu_analysis_active_client    = GENERAL_CLIENT
    self._lu_analysis_filtered_sectors = None
    self._lu_analysis_risk_filter      = None   # None | "HIGH" | "LOW" (all-clients tier filter)
    self._lu_analysis_product_substr   = None   # substring filter on Product Name

    _lu_show_placeholder(self)
    _lu_switch_view(self, "analysis")


# ══════════════════════════════════════════════════════════════════════
#  VIEW SWITCHER
# ══════════════════════════════════════════════════════════════════════

def _lu_switch_view(self, view: str):
    self._lu_active_view.set(view)
    views = {
        "analysis":  self._lu_analysis_view,
        "charts":    self._lu_charts_view,
        "simulator": self._lu_simulator_view,
        "loanbal":   self._lu_loanbal_view,
        "report":    self._lu_report_view,
    }
    for name, frame in views.items():
        if name == view:
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        else:
            frame.place_forget()

    # Main Analysis toolbar (global search row) is visible only on Analysis tab.
    bar = getattr(self, "_lu_client_bar", None)
    divider = getattr(self, "_lu_client_bar_divider", None)
    if bar is not None:
        if view == "analysis":
            if not bar.winfo_ismapped():
                bar.pack(fill="x", before=self._lu_view_container)
        else:
            if bar.winfo_ismapped():
                bar.pack_forget()
    if divider is not None:
        if view == "analysis":
            if not divider.winfo_ismapped():
                divider.pack(fill="x", before=self._lu_view_container)
        else:
            if divider.winfo_ismapped():
                divider.pack_forget()

    # Bottom Risk Settings button is analysis-only.
    rs_btn = getattr(self, "_risk_settings_btn", None)
    if rs_btn is not None:
        if view == "analysis":
            if not rs_btn.winfo_ismapped():
                rs_btn.pack(side="right", padx=(0, 14), pady=10)
        else:
            if rs_btn.winfo_ismapped():
                rs_btn.pack_forget()

    if view == "simulator" and self._lu_all_data:
        _sim_populate(self)
    if view == "charts"    and self._lu_all_data:
        _charts_render(self)
    if view == "report"    and self._lu_all_data:
        _report_render(self)
    if view == "loanbal"   and self._lu_all_data:
        _loanbal_render(self)


# ══════════════════════════════════════════════════════════════════════
#  FILE LOAD + ANALYSIS RUNNER
# ══════════════════════════════════════════════════════════════════════

def _lu_browse_file(self):
    path = filedialog.askopenfilename(
        title="Select Excel File for LU Analysis",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
    if not path:
        return
    self._lu_filepath = path
    self._lu_file_lbl.config(text=f"📊  {Path(path).name}", fg=_TXT_NAVY)
    self._lu_rescan_btn.configure(state="normal")
    _lu_run_analysis(self)


def _lu_run_analysis(self):
    if not self._lu_filepath:
        return
    self._lu_status_lbl.config(text="⏳  Scanning…", fg=_ACCENT_GOLD)
    self._lu_load_btn.configure(state="disabled")
    self._lu_rescan_btn.configure(state="disabled")
    if getattr(self, "_lu_export_btn", None) is not None:
        self._lu_export_btn.configure(state="disabled")
    self.update_idletasks()
    try:
        all_data = run_lu_analysis(self._lu_filepath)
        self._lu_all_data         = all_data
        self._lu_filtered_sectors = None
        self._lu_analysis_filtered_sectors = None
        self._lu_analysis_risk_filter = None
        self._lu_analysis_product_substr = None
        self._lu_analysis_active_client = GENERAL_CLIENT
        _lu_update_filter_pill(self)
        _lu_populate_client_dropdown(self)

        # Re-apply any active settings overrides so that score_label,
        # score_fg, score_bg are all in sync before the first render.
        _lu_rescore_all(self)

        general = all_data.get("general", [])
        self._lu_results = general
        _lu_render_results(self, general)

        n_clients = len(all_data.get("clients", {}))
        n_industries = len(all_data.get("unique_industries", []))
        totals    = all_data.get("totals", {})
        self._lu_status_lbl.config(
            text=(f"✅  {n_clients} client(s) · {n_industries} industries · "
                  f"₱{totals.get('loan_balance', 0):,.0f} total loan balance"),
            fg=_LIME_DARK)
        if getattr(self, "_lu_export_btn", None) is not None:
            self._lu_export_btn.configure(state="normal")
        if getattr(self, "_loanbal_export_btn", None) is not None:
            self._loanbal_export_btn.configure(state="normal")
        _sim_populate(self)
    except Exception as exc:
        _lu_show_error(self, str(exc))
        self._lu_status_lbl.config(text="❌  Error during scan", fg=_ACCENT_RED)
    finally:
        self._lu_load_btn.configure(state="normal")
        self._lu_rescan_btn.configure(state="normal")


def _lu_rescore_all(self):
    """
    Re-score every client record in-memory using the current lu_core settings
    (high-risk industries, product overrides, expense overrides).
    Updates score, score_label, score_fg, score_bg, risk, risk_reasoning on each record.
    Call this after any settings change instead of re-parsing the file.
    """
    all_data = getattr(self, "_lu_all_data", None)
    if not all_data:
        return
    for rec in all_data.get("general", []):
        industry = rec.get("industry", "")
        expenses = rec.get("expenses", [])
        product_name = str(rec.get("product_name") or "")

        # Determine product override (any atomic product in the cell may match)
        pr, _pr_matched = lookup_product_risk_override(product_name)

        # Determine industry high
        high_ind_set = {str(x).strip().lower() for x in get_high_risk_industries()}
        tags = rec.get("industry_tags") or _lu_core._extract_industry_tags(industry)
        high_ind = any(str(t or "").strip().lower() in high_ind_set for t in tags)

        # Determine expense override
        exp_overrides = get_expense_risk_overrides()
        high_exp_name = ""
        exp_override_result = None
        for e in expenses or []:
            nm = str((e or {}).get("name") or "").strip()
            if not nm:
                continue
            lvl = next(
                (exp_overrides.get(k) for k in _lu_core._expense_override_lookup_keys(nm) if exp_overrides.get(k) == "HIGH"),
                None,
            )
            if lvl == "HIGH":
                high_exp_name = nm
                exp_override_result = "HIGH"
                e["risk"] = "HIGH"
                break

        # Apply precedence: Product > user expense HIGH > Industry
        if pr == "HIGH":
            label_val = "HIGH"
        elif exp_override_result == "HIGH":
            label_val = "HIGH"
        elif high_ind:
            label_val = "HIGH"
        else:
            label_val = "LOW"

        score_val = 1.8 if label_val == "HIGH" else 0.0
        rec["score"]       = score_val
        rec["score_label"] = label_val
        rec["risk"]        = label_val
        rec["score_fg"]    = "#E53E3E" if label_val == "HIGH" else "#2E7D32"
        rec["score_bg"]    = "#FFF5F5" if label_val == "HIGH" else "#F0FBE8"
        rec["risk_reasoning"] = _lu_core._compute_risk_reasoning(
            industry=industry,
            product_name=product_name,
            product_override=pr,
            expense_high_name=high_exp_name,
            is_high_industry=high_ind,
            product_matched_token=_pr_matched,
        )
    # Keep the clients dict in sync
    all_data["clients"] = {r["client"]: r for r in all_data.get("general", [])}


# ══════════════════════════════════════════════════════════════════════
#  RISK SETTINGS DIALOG
# ══════════════════════════════════════════════════════════════════════

def _open_industry_risk_dialog(self):
    if not self._lu_all_data or not self._lu_all_data.get("unique_industries"):
        messagebox.showwarning("No Data", "Load an Excel file first to see industries.")
        return

    industries = sorted(self._lu_all_data["unique_industries"], key=str.lower)
    high_set = {str(x).strip().lower() for x in get_high_risk_industries()}

    dialog = ctk.CTkToplevel(self)
    dialog.title("Industry Risk Settings")
    dialog.geometry("700x620")
    dialog.minsize(620, 520)
    dialog.transient(self)
    dialog.grab_set()
    dialog.configure(fg_color=_CARD_WHITE)

    # Header
    hdr = tk.Frame(dialog, bg=_NAVY_MID, height=52)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)
    tk.Label(
        hdr,
        text="⚙  Industry Risk Settings",
        font=F(11, "bold"),
        fg=_WHITE,
        bg=_NAVY_MID
    ).pack(side="left", padx=16, pady=12)

    note = tk.Frame(dialog, bg=_NAVY_MIST, highlightbackground=_BORDER_MID, highlightthickness=1)
    note.pack(fill="x", padx=16, pady=(10, 6))
    tk.Label(
        note,
        text=(
            "Set industry overrides. HIGH forces HIGH risk. "
            "LOW means no industry override (falls back to other rules). "
            "Changes apply to all LU tabs after saving."
        ),
        font=F(8),
        fg=_TXT_SOFT,
        bg=_NAVY_MIST,
        anchor="w",
        justify="left"
    ).pack(fill="x", padx=10, pady=8)

    # Search
    search_row = tk.Frame(dialog, bg=_CARD_WHITE)
    search_row.pack(fill="x", padx=16, pady=(4, 6))
    tk.Label(search_row, text="🔍", font=F(10), fg=_TXT_SOFT, bg=_CARD_WHITE).pack(side="left")
    search_var = tk.StringVar(value="")
    search_entry = ctk.CTkEntry(
        search_row,
        textvariable=search_var,
        width=420,
        height=28,
        corner_radius=4,
        fg_color=_WHITE,
        text_color=_TXT_NAVY,
        border_color=_BORDER_MID,
        font=FF(9),
        placeholder_text="Search industry..."
    )
    search_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

    # Column header
    col_hdr = tk.Frame(dialog, bg=_NAVY_MID, height=30)
    col_hdr.pack(fill="x", padx=16)
    col_hdr.pack_propagate(False)
    tk.Label(col_hdr, text="Industry", font=F(8, "bold"), fg=_WHITE, bg=_NAVY_MID).pack(side="left", padx=10, pady=6)
    tk.Label(col_hdr, text="Risk Level", font=F(8, "bold"), fg=_WHITE, bg=_NAVY_MID).pack(side="right", padx=10, pady=6)

    # Scrollable list
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
    canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", lambda ev: canvas.yview_scroll(int(-1*(ev.delta/120)), "units")))
    canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

    row_state = {}
    row_widgets = []

    def _paint_buttons(var, high_btn, low_btn):
        if var.get() == "HIGH":
            high_btn.configure(bg="#FFB3B3", fg=_ACCENT_RED, relief="sunken", font=F(8, "bold"))
            low_btn.configure(bg="#F0FBE8", fg=_TXT_MUTED, relief="flat", font=F(8))
        else:
            low_btn.configure(bg="#B7E8A0", fg=_ACCENT_SUCCESS, relief="sunken", font=F(8, "bold"))
            high_btn.configure(bg="#FFE8E8", fg=_TXT_MUTED, relief="flat", font=F(8))

    def _build_rows():
        for w in rows_frame.winfo_children():
            w.destroy()
        row_widgets.clear()
        for idx, industry in enumerate(industries):
            risk_var = row_state.get(industry)
            if risk_var is None:
                risk_var = tk.StringVar(value="HIGH" if industry.strip().lower() in high_set else "LOW")
                row_state[industry] = risk_var

            row_bg = _WHITE if idx % 2 == 0 else _OFF_WHITE
            row = tk.Frame(rows_frame, bg=row_bg)
            row.pack(fill="x")
            tk.Frame(row, bg=_BORDER_LIGHT, height=1).pack(fill="x")

            inner = tk.Frame(row, bg=row_bg)
            inner.pack(fill="x", padx=8, pady=4)
            tk.Label(
                inner,
                text=industry,
                font=F(9),
                fg=_TXT_NAVY,
                bg=row_bg,
                anchor="w",
                justify="left",
                wraplength=420
            ).pack(side="left", fill="x", expand=True, padx=(4, 0))

            btn_wrap = tk.Frame(inner, bg=row_bg)
            btn_wrap.pack(side="right", padx=6)
            high_btn = tk.Button(
                btn_wrap, text="🟠 HIGH", font=F(8), fg=_TXT_MUTED, bg="#FFE8E8",
                relief="flat", bd=1, padx=8, pady=3, cursor="hand2", activebackground="#FFB3B3",
                command=lambda v=risk_var: v.set("HIGH")
            )
            low_btn = tk.Button(
                btn_wrap, text="🟢 LOW", font=F(8), fg=_TXT_MUTED, bg="#F0FBE8",
                relief="flat", bd=1, padx=8, pady=3, cursor="hand2", activebackground="#B7E8A0",
                command=lambda v=risk_var: v.set("LOW")
            )
            high_btn.pack(side="left", padx=(0, 4))
            low_btn.pack(side="left")
            risk_var.trace_add("write", lambda *_a, v=risk_var, hb=high_btn, lb=low_btn: _paint_buttons(v, hb, lb))
            _paint_buttons(risk_var, high_btn, low_btn)
            row_widgets.append((row, industry))

    def _apply_filter(*_args):
        query = search_var.get().strip().lower()
        for row, industry in row_widgets:
            if not query or query in industry.lower():
                if not row.winfo_ismapped():
                    row.pack(fill="x")
            else:
                if row.winfo_ismapped():
                    row.pack_forget()

    search_var.trace_add("write", _apply_filter)

    def _set_all(val: str):
        for v in row_state.values():
            v.set(val)

    def save():
        new_high = [industry for industry, var in row_state.items() if var.get() == "HIGH"]
        set_high_risk_industries(new_high)
        _lu_rescore_all(self)
        _lu_render_results(self, self._lu_all_data.get("general", []))
        dialog.destroy()

    _build_rows()

    footer = tk.Frame(dialog, bg=_OFF_WHITE, highlightbackground=_BORDER_MID, highlightthickness=1)
    footer.pack(fill="x", padx=16, pady=(2, 14))
    tk.Button(
        footer, text="Set ALL -> HIGH", font=F(8, "bold"), fg=_ACCENT_RED, bg="#FFE8E8",
        relief="flat", bd=0, padx=10, pady=6, cursor="hand2", command=lambda: _set_all("HIGH")
    ).pack(side="left", padx=(12, 4), pady=8)
    tk.Button(
        footer, text="Set ALL -> LOW", font=F(8, "bold"), fg=_ACCENT_SUCCESS, bg="#DCEDC8",
        relief="flat", bd=0, padx=10, pady=6, cursor="hand2", command=lambda: _set_all("LOW")
    ).pack(side="left", padx=4, pady=8)
    tk.Button(
        footer, text="Cancel", font=F(9), fg=_TXT_SOFT, bg=_OFF_WHITE,
        relief="flat", bd=0, padx=10, pady=8, cursor="hand2", command=dialog.destroy
    ).pack(side="right", padx=(0, 4), pady=8)
    tk.Button(
        footer, text="  ✔  Apply & Close  ", font=F(9, "bold"), fg=_WHITE, bg=_NAVY_MID,
        activebackground=_NAVY_LIGHT, activeforeground=_WHITE,
        relief="flat", bd=0, padx=14, pady=8, cursor="hand2", command=save
    ).pack(side="right", padx=12, pady=8)


def _open_product_risk_dialog(self):
    if not self._lu_all_data or not self._lu_all_data.get("unique_product_names"):
        messagebox.showwarning(
            "No Data",
            "Load an Excel file first. Product Name column must contain values.",
        )
        return

    _detected = (self._lu_all_data or {}).get("unique_product_names", [])
    current = get_product_risk_overrides()
    products = sorted({str(x).strip() for x in _detected if str(x).strip()}, key=str.lower)

    dialog = ctk.CTkToplevel(self)
    dialog.title("Product Name Risk Settings")
    dialog.geometry("700x620")
    dialog.minsize(620, 520)
    dialog.transient(self)
    dialog.grab_set()
    dialog.configure(fg_color=_CARD_WHITE)

    hdr = tk.Frame(dialog, bg=_NAVY_MID, height=52)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)
    tk.Label(
        hdr,
        text="⚙  Product Name Risk Settings",
        font=F(11, "bold"),
        fg=_WHITE,
        bg=_NAVY_MID,
    ).pack(side="left", padx=16, pady=12)

    note = tk.Frame(dialog, bg=_NAVY_MIST, highlightbackground=_BORDER_MID, highlightthickness=1)
    note.pack(fill="x", padx=16, pady=(10, 6))
    tk.Label(
        note,
        text=(
            "Set overrides per atomic product (comma-separated names in Excel are split). "
            "HIGH forces HIGH risk if any listed product appears in a client's Product Name cell. "
            "LOW means no product override (falls back to Expense/Industry rules). "
            "Applies to all LU tabs after saving (re-scan)."
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
        width=420,
        height=28,
        corner_radius=4,
        fg_color=_WHITE,
        text_color=_TXT_NAVY,
        border_color=_BORDER_MID,
        font=FF(9),
        placeholder_text="Search product name...",
    )
    search_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

    col_hdr = tk.Frame(dialog, bg=_NAVY_MID, height=30)
    col_hdr.pack(fill="x", padx=16)
    col_hdr.pack_propagate(False)
    tk.Label(col_hdr, text="Product Name", font=F(8, "bold"), fg=_WHITE, bg=_NAVY_MID).pack(
        side="left", padx=10, pady=6
    )
    tk.Label(col_hdr, text="Risk Level", font=F(8, "bold"), fg=_WHITE, bg=_NAVY_MID).pack(
        side="right", padx=10, pady=6
    )

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
    canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", lambda ev: canvas.yview_scroll(int(-1*(ev.delta/120)), "units")))
    canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

    row_state = {}
    row_widgets = []

    def _paint_buttons(var, high_btn, low_btn):
        if var.get() == "HIGH":
            high_btn.configure(bg="#FFB3B3", fg=_ACCENT_RED, relief="sunken", font=F(8, "bold"))
            low_btn.configure(bg="#F0FBE8", fg=_TXT_MUTED, relief="flat", font=F(8))
        else:
            low_btn.configure(bg="#B7E8A0", fg=_ACCENT_SUCCESS, relief="sunken", font=F(8, "bold"))
            high_btn.configure(bg="#FFE8E8", fg=_TXT_MUTED, relief="flat", font=F(8))

    def _build_rows():
        for w in rows_frame.winfo_children():
            w.destroy()
        row_widgets.clear()
        for idx, product in enumerate(products):
            pk = product.strip().lower()
            lvl = current.get(pk, "LOW")
            risk_var = tk.StringVar(value="HIGH" if lvl == "HIGH" else "LOW")
            row_state[product] = risk_var

            row_bg = _WHITE if idx % 2 == 0 else _OFF_WHITE
            row = tk.Frame(rows_frame, bg=row_bg)
            row.pack(fill="x")
            tk.Frame(row, bg=_BORDER_LIGHT, height=1).pack(fill="x")

            inner = tk.Frame(row, bg=row_bg)
            inner.pack(fill="x", padx=8, pady=4)
            tk.Label(
                inner,
                text=product,
                font=F(9),
                fg=_TXT_NAVY,
                bg=row_bg,
                anchor="w",
                justify="left",
                wraplength=420,
            ).pack(side="left", fill="x", expand=True, padx=(4, 0))

            btn_wrap = tk.Frame(inner, bg=row_bg)
            btn_wrap.pack(side="right", padx=6)
            high_btn = tk.Button(
                btn_wrap, text="🟠 HIGH", font=F(8), fg=_TXT_MUTED, bg="#FFE8E8",
                relief="flat", bd=1, padx=8, pady=3, cursor="hand2", activebackground="#FFB3B3",
                command=lambda v=risk_var: v.set("HIGH"),
            )
            low_btn = tk.Button(
                btn_wrap, text="🟢 LOW", font=F(8), fg=_TXT_MUTED, bg="#F0FBE8",
                relief="flat", bd=1, padx=8, pady=3, cursor="hand2", activebackground="#B7E8A0",
                command=lambda v=risk_var: v.set("LOW"),
            )
            high_btn.pack(side="left", padx=(0, 4))
            low_btn.pack(side="left")
            risk_var.trace_add("write", lambda *_a, v=risk_var, hb=high_btn, lb=low_btn: _paint_buttons(v, hb, lb))
            _paint_buttons(risk_var, high_btn, low_btn)
            row_widgets.append((row, product))

    def _apply_filter(*_args):
        q = search_var.get().strip().lower()
        for row, product in row_widgets:
            if not q or q in product.lower():
                if not row.winfo_ismapped():
                    row.pack(fill="x")
            else:
                if row.winfo_ismapped():
                    row.pack_forget()

    search_var.trace_add("write", _apply_filter)

    def _set_all(val: str):
        for v in row_state.values():
            v.set(val)

    def save():
        # Store HIGH-only product overrides so default LOW state does not
        # unintentionally force all clients to LOW and block other rules.
        mapping = {name: row_state[name].get() for name in row_state if row_state[name].get() == "HIGH"}
        set_product_risk_overrides(mapping)
        _lu_rescore_all(self)
        _lu_render_results(self, self._lu_all_data.get("general", []))
        dialog.destroy()

    _build_rows()

    footer = tk.Frame(dialog, bg=_OFF_WHITE, highlightbackground=_BORDER_MID, highlightthickness=1)
    footer.pack(fill="x", padx=16, pady=(2, 14))
    tk.Button(
        footer, text="Set ALL -> HIGH", font=F(8, "bold"), fg=_ACCENT_RED, bg="#FFE8E8",
        relief="flat", bd=0, padx=10, pady=6, cursor="hand2", command=lambda: _set_all("HIGH"),
    ).pack(side="left", padx=(12, 4), pady=8)
    tk.Button(
        footer, text="Set ALL -> LOW", font=F(8, "bold"), fg=_ACCENT_SUCCESS, bg="#DCEDC8",
        relief="flat", bd=0, padx=10, pady=6, cursor="hand2", command=lambda: _set_all("LOW"),
    ).pack(side="left", padx=4, pady=8)
    tk.Button(
        footer, text="Cancel", font=F(9), fg=_TXT_SOFT, bg=_OFF_WHITE,
        relief="flat", bd=0, padx=10, pady=8, cursor="hand2", command=dialog.destroy,
    ).pack(side="right", padx=(0, 4), pady=8)
    tk.Button(
        footer, text="  ✔  Apply & Close  ", font=F(9, "bold"), fg=_WHITE, bg=_NAVY_MID,
        activebackground=_NAVY_LIGHT, activeforeground=_WHITE,
        relief="flat", bd=0, padx=14, pady=8, cursor="hand2", command=save,
    ).pack(side="right", padx=12, pady=8)


def _open_expense_risk_dialog(self):
    # ── Canonical expense categories from expense_categories.pdf ──────
    # These are the fixed canonical titles (one per category) used as the
    # basis for risk tagging.  The list is always shown in full regardless
    # of what is present in the loaded Excel file.
    _PDF_EXPENSE_SECTIONS = [
        (
            "🏠  Household / Personal Expenses",
            "#E8F0FE",  # section header bg
            "#3A5BA0",  # section header fg
            [
                "Food & Groceries",
                "Utilities – Electricity",
                "Utilities – Water",
                "Utilities – Gas / LPG",
                "Utilities – Internet & Cable",
                "Communication – Phone Load",
                "Fuel & Transportation",
                "Vehicle Maintenance",
                "Education",
                "Loans & Amortizations",
                "Credit Card Payments",
                "Insurance & Gov't Contributions",
                "Housing & Rent",
                "Household Help",
                "Health & Medicine",
                "Personal & Family Support",
                "Savings & Investments",
            ],
        ),
        (
            "🏢  Business Expenses",
            "#E8F5E9",  # section header bg
            "#2E7D32",  # section header fg
            [
                "Cost of Goods / Purchases",
                "Salaries & Labor",
                "Rent & Space",
                "Utilities – Electricity & Water",
                "Fuel & Transportation",
                "Vehicle & Equipment Maintenance",
                "Permits, Taxes & Licenses",
                "Operating Expenses (General)",
                "Loans & Amortizations",
                "Supplies & Materials",
                "Bookkeeping & Professional Fees",
                "Agriculture & Farming",
                "Insurance (Business)",
                "Other / Miscellaneous",
            ],
        ),
    ]

    current = get_expense_risk_overrides()

    def _norm(s: str) -> str:
        s = str(s or "").lower()
        s = re.sub(r"[^a-z0-9\s/]", " ", s)
        return re.sub(r"\s+", " ", s).strip()

    dialog = ctk.CTkToplevel(self)
    dialog.title("Expense Risk Settings")
    dialog.geometry("780x660")
    dialog.minsize(680, 560)
    dialog.transient(self)
    dialog.grab_set()
    dialog.configure(fg_color=_CARD_WHITE)

    # ── Header ─────────────────────────────────────────────────────────
    hdr = tk.Frame(dialog, bg=_NAVY_MID, height=52)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)
    tk.Label(
        hdr,
        text="⚙  Expense Risk Settings",
        font=F(11, "bold"),
        fg=_WHITE,
        bg=_NAVY_MID,
    ).pack(side="left", padx=16, pady=12)
    tk.Label(
        hdr,
        text=(
            "Based on expense_categories.pdf  •  17 Household + 14 Business categories"
        ),
        font=F(7),
        fg=_NAVY_PALE,
        bg=_NAVY_MID,
    ).pack(side="left", padx=(0, 16), pady=12)

    # ── Note strip ─────────────────────────────────────────────────────
    note = tk.Frame(dialog, bg=_NAVY_MIST, highlightbackground=_BORDER_MID, highlightthickness=1)
    note.pack(fill="x", padx=16, pady=(10, 4))
    tk.Label(
        note,
        text=(
            "Each row represents a canonical expense category from the PDF reference. "
            "Set to HIGH or LOW.  Precedence: Product Risk > Expense Risk > Industry Risk.  "
            "Changes apply to all LU tabs after saving."
        ),
        font=F(8),
        fg=_TXT_SOFT,
        bg=_NAVY_MIST,
        anchor="w",
        justify="left",
        wraplength=720,
    ).pack(fill="x", padx=10, pady=8)

    # ── Search bar ─────────────────────────────────────────────────────
    search_row = tk.Frame(dialog, bg=_CARD_WHITE)
    search_row.pack(fill="x", padx=16, pady=(4, 4))
    tk.Label(search_row, text="🔍", font=F(10), fg=_TXT_SOFT, bg=_CARD_WHITE).pack(side="left")
    search_var = tk.StringVar(value="")
    ctk.CTkEntry(
        search_row,
        textvariable=search_var,
        width=460,
        height=28,
        corner_radius=4,
        fg_color=_WHITE,
        text_color=_TXT_NAVY,
        border_color=_BORDER_MID,
        font=FF(9),
        placeholder_text="Search expense category…",
    ).pack(side="left", fill="x", expand=True, padx=(6, 0))

    # ── Column header ──────────────────────────────────────────────────
    col_hdr = tk.Frame(dialog, bg=_NAVY_MID, height=30)
    col_hdr.pack(fill="x", padx=16)
    col_hdr.pack_propagate(False)
    tk.Label(col_hdr, text="Expense Category", font=F(8, "bold"), fg=_WHITE, bg=_NAVY_MID).pack(
        side="left", padx=10, pady=6
    )
    tk.Label(col_hdr, text="Risk Level", font=F(8, "bold"), fg=_WHITE, bg=_NAVY_MID).pack(
        side="right", padx=10, pady=6
    )

    # ── Scrollable list ────────────────────────────────────────────────
    list_wrap = tk.Frame(dialog, bg=_CARD_WHITE)
    list_wrap.pack(fill="both", expand=True, padx=16, pady=(0, 6))
    sb = tk.Scrollbar(list_wrap, relief="flat", troughcolor=_OFF_WHITE,
                      bg=_BORDER_LIGHT, width=8, bd=0)
    sb.pack(side="right", fill="y")
    canvas = tk.Canvas(list_wrap, bg=_CARD_WHITE, highlightthickness=0, yscrollcommand=sb.set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.config(command=canvas.yview)
    rows_frame = tk.Frame(canvas, bg=_CARD_WHITE)
    win = canvas.create_window((0, 0), window=rows_frame, anchor="nw")
    rows_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
    canvas.bind("<Enter>", lambda _e: canvas.bind_all(
        "<MouseWheel>", lambda ev: canvas.yview_scroll(int(-1 * (ev.delta / 120)), "units")))
    canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

    row_state   = {}   # row_id -> {"name": category_name, "var": tk.StringVar}
    row_widgets = []   # list of (frame_widget, exp_name, section_title)

    def _paint_buttons(var, high_btn, low_btn):
        if var.get() == "HIGH":
            high_btn.configure(bg="#FFB3B3", fg=_ACCENT_RED, relief="sunken", font=F(8, "bold"))
            low_btn.configure( bg="#F0FBE8", fg=_TXT_MUTED,  relief="flat",   font=F(8))
        else:
            low_btn.configure( bg="#B7E8A0", fg=_ACCENT_SUCCESS, relief="sunken", font=F(8, "bold"))
            high_btn.configure(bg="#FFE8E8", fg=_TXT_MUTED,      relief="flat",   font=F(8))

    def _build_rows():
        for w in rows_frame.winfo_children():
            w.destroy()
        row_widgets.clear()

        row_idx = 0  # alternating bg counter across all sections
        for section_title, sec_hdr_bg, sec_hdr_fg, categories in _PDF_EXPENSE_SECTIONS:
            # Section divider label
            sec_frame = tk.Frame(rows_frame, bg=sec_hdr_bg)
            sec_frame.pack(fill="x")
            sec_lbl = tk.Label(
                sec_frame,
                text=f"  {section_title}",
                font=F(8, "bold"),
                fg=sec_hdr_fg,
                bg=sec_hdr_bg,
                anchor="w",
            )
            sec_lbl.pack(fill="x", padx=6, pady=5)
            row_widgets.append((sec_frame, None, section_title))  # section sentinel

            for cat_name in categories:
                ek  = _lu_core._expense_override_key(cat_name)
                lvl = current.get(ek, "LOW")
                risk_var = tk.StringVar(value="HIGH" if lvl == "HIGH" else "LOW")
                row_id = f"{section_title}::{cat_name}"
                row_state[row_id] = {"name": cat_name, "var": risk_var}

                row_bg = _WHITE if row_idx % 2 == 0 else _OFF_WHITE
                row_idx += 1
                row = tk.Frame(rows_frame, bg=row_bg)
                row.pack(fill="x")
                tk.Frame(row, bg=_BORDER_LIGHT, height=1).pack(fill="x")

                inner = tk.Frame(row, bg=row_bg)
                inner.pack(fill="x", padx=8, pady=5)
                tk.Label(
                    inner,
                    text=cat_name,
                    font=F(9),
                    fg=_TXT_NAVY,
                    bg=row_bg,
                    anchor="w",
                    justify="left",
                    wraplength=500,
                ).pack(side="left", fill="x", expand=True, padx=(4, 0))

                btn_wrap = tk.Frame(inner, bg=row_bg)
                btn_wrap.pack(side="right", padx=6)
                high_btn = tk.Button(
                    btn_wrap, text="🟠 HIGH", font=F(8), fg=_TXT_MUTED, bg="#FFE8E8",
                    relief="flat", bd=1, padx=8, pady=3, cursor="hand2",
                    activebackground="#FFB3B3",
                    command=lambda v=risk_var: v.set("HIGH"),
                )
                low_btn = tk.Button(
                    btn_wrap, text="🟢 LOW", font=F(8), fg=_TXT_MUTED, bg="#F0FBE8",
                    relief="flat", bd=1, padx=8, pady=3, cursor="hand2",
                    activebackground="#B7E8A0",
                    command=lambda v=risk_var: v.set("LOW"),
                )
                high_btn.pack(side="left", padx=(0, 4))
                low_btn.pack(side="left")
                risk_var.trace_add(
                    "write",
                    lambda *_a, v=risk_var, hb=high_btn, lb=low_btn: _paint_buttons(v, hb, lb),
                )
                _paint_buttons(risk_var, high_btn, low_btn)
                # Store (row_frame, cat_name, section_title) for filtering
                row_widgets.append((row, cat_name, section_title))

    def _apply_filter(*_args):
        q = _norm(search_var.get())
        # Track which section headers should be visible
        section_has_match: dict[str, bool] = {}
        for entry in row_widgets:
            frame, cat_name, section_title = entry
            if cat_name is None:
                # This is a section header sentinel — decide later
                continue
            match = not q or q in _norm(cat_name) or q in _norm(section_title)
            section_has_match[section_title] = section_has_match.get(section_title, False) or match
            if match:
                if not frame.winfo_ismapped():
                    frame.pack(fill="x")
            else:
                if frame.winfo_ismapped():
                    frame.pack_forget()

        # Show/hide section header rows based on whether any child matches
        for entry in row_widgets:
            frame, cat_name, section_title = entry
            if cat_name is None:  # section header sentinel
                visible = section_has_match.get(section_title, False) or not q
                if visible:
                    if not frame.winfo_ismapped():
                        frame.pack(fill="x")
                else:
                    if frame.winfo_ismapped():
                        frame.pack_forget()

        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    search_var.trace_add("write", _apply_filter)

    def _set_all(val: str):
        for row in row_state.values():
            row["var"].set(val)

    def save():
        # Save HIGH-only flags (LOW = not set)
        mapping = {
            row["name"]: row["var"].get()
            for row in row_state.values()
            if row["var"].get() == "HIGH"
        }
        set_expense_risk_overrides(mapping)
        _lu_rescore_all(self)
        _lu_render_results(self, self._lu_all_data.get("general", []))
        dialog.destroy()

    _build_rows()
    _apply_filter()

    # ── Footer ─────────────────────────────────────────────────────────
    footer = tk.Frame(dialog, bg=_OFF_WHITE, highlightbackground=_BORDER_MID, highlightthickness=1)
    footer.pack(fill="x", padx=16, pady=(2, 14))
    tk.Button(
        footer, text="Set ALL → HIGH", font=F(8, "bold"), fg=_ACCENT_RED, bg="#FFE8E8",
        relief="flat", bd=0, padx=10, pady=6, cursor="hand2", command=lambda: _set_all("HIGH"),
    ).pack(side="left", padx=(12, 4), pady=8)
    tk.Button(
        footer, text="Set ALL → LOW", font=F(8, "bold"), fg=_ACCENT_SUCCESS, bg="#DCEDC8",
        relief="flat", bd=0, padx=10, pady=6, cursor="hand2", command=lambda: _set_all("LOW"),
    ).pack(side="left", padx=4, pady=8)
    tk.Button(
        footer, text="Cancel", font=F(9), fg=_TXT_SOFT, bg=_OFF_WHITE,
        relief="flat", bd=0, padx=10, pady=8, cursor="hand2", command=dialog.destroy,
    ).pack(side="right", padx=(0, 4), pady=8)
    tk.Button(
        footer, text="  ✔  Apply & Close  ", font=F(9, "bold"), fg=_WHITE, bg=_NAVY_MID,
        activebackground=_NAVY_LIGHT, activeforeground=_WHITE,
        relief="flat", bd=0, padx=14, pady=8, cursor="hand2", command=save,
    ).pack(side="right", padx=12, pady=8)

    # ── No-data guard removed: dialog is always openable since categories ──
    # ── are PDF-based and do not depend on a loaded Excel file.         ──


# ══════════════════════════════════════════════════════════════════════
#  ATTACH
# ══════════════════════════════════════════════════════════════════════

def attach(cls):
    """Attach ALL LU Analysis methods to the app class."""
    # Core panel + file loading
    cls._build_lu_analysis_panel = _build_lu_analysis_panel
    cls._lu_switch_view          = _lu_switch_view
    cls._lu_browse_file          = _lu_browse_file
    cls._lu_run_analysis         = _lu_run_analysis
    cls._lu_rescore_all          = _lu_rescore_all
    cls._open_industry_risk_dialog = _open_industry_risk_dialog
    cls._open_product_risk_dialog = _open_product_risk_dialog
    cls._open_expense_risk_dialog = _open_expense_risk_dialog

    # Each tab module attaches its own methods
    lu_tab_analysis.attach(cls)
    lu_tab_charts.attach(cls)
    lu_simulator_patch.attach(cls)
    lu_loanbal_export_patch.attach(cls)
    lu_tab_report.attach(cls)

    # Placeholder attributes set by builder
    cls._lu_export_btn    = None
    cls._loanbal_export_btn = None
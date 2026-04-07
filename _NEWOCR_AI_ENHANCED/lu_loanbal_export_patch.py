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


def F(size, weight="normal"):
    return ("Segoe UI", size, weight)

def FF(size, weight="normal"):
    return ctk.CTkFont(family="Segoe UI", size=size, weight=weight)


# ── Scrollable helper (self-contained copy so we don't import from lu_ui) ────

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
#  Identical to the original except we add the Export button in the header.
# ══════════════════════════════════════════════════════════════════════════════

def _build_loanbal_panel_patched(self, parent):
    # ── Header bar ────────────────────────────────────────────────────────────
    hdr = tk.Frame(parent, bg=_NAVY_MID, height=38)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)

    tk.Label(
        hdr,
        text="📊  Sector vs Total Loan Balance  —  Exposure Analysis",
        font=F(10, "bold"), fg=_WHITE, bg=_NAVY_MID,
    ).pack(side="left", padx=20, pady=8)

    # ── Export button ─────────────────────────────────────────────────────────
    self._loanbal_export_btn = ctk.CTkButton(
        hdr,
        text="💾  Export",
        command=lambda: _loanbal_show_export_menu(self),
        width=100, height=26, corner_radius=4,
        fg_color=_LIME_DARK, hover_color=_LIME_MID,
        text_color=_TXT_ON_LIME, font=FF(8, "bold"),
        state="disabled",
    )
    self._loanbal_export_btn.pack(side="right", padx=12, pady=6)

    # ── Body ──────────────────────────────────────────────────────────────────
    self._loanbal_body = tk.Frame(parent, bg=_CARD_WHITE)
    self._loanbal_body.pack(fill="both", expand=True)
    tk.Label(
        self._loanbal_body,
        text="Run an analysis first to view loan balance exposure.",
        font=F(10), fg=_TXT_MUTED, bg=_CARD_WHITE,
    ).pack(pady=60)


# ══════════════════════════════════════════════════════════════════════════════
#  PATCHED _loanbal_render
#  Identical to original plus it enables the Export button when data is present.
# ══════════════════════════════════════════════════════════════════════════════

def _loanbal_render_patched(self):
    for w in self._loanbal_body.winfo_children():
        w.destroy()
    plt.close("all")

    all_data   = self._lu_all_data
    general    = all_data.get("general", [])
    sector_map = all_data.get("sector_map", {})
    totals     = all_data.get("totals", {})
    grand_lb   = totals.get("loan_balance", 0) or 0

    # Enable / disable the export button based on data availability
    btn = getattr(self, "_loanbal_export_btn", None)
    if btn:
        btn.configure(state="normal" if general else "disabled")

    if not general:
        tk.Label(
            self._loanbal_body,
            text="No data available.",
            font=F(10), fg=_TXT_MUTED, bg=_CARD_WHITE,
        ).pack(pady=60)
        return

    _, inner, _ = _make_scrollable(self._loanbal_body, _CARD_WHITE)
    pad = tk.Frame(inner, bg=_CARD_WHITE)
    pad.pack(fill="both", expand=True, padx=24, pady=16)

    # ── Grand total card ──────────────────────────────────────────────────────
    grand_card = tk.Frame(pad, bg=_NAVY_DEEP,
                          highlightbackground=_NAVY_MID, highlightthickness=1)
    grand_card.pack(fill="x", pady=(0, 16))
    gc_inner = tk.Frame(grand_card, bg=_NAVY_DEEP)
    gc_inner.pack(fill="x", padx=24, pady=16)
    tk.Label(gc_inner, text="💰  GRAND TOTAL LOAN BALANCE",
             font=F(10, "bold"), fg=_TXT_MUTED, bg=_NAVY_DEEP).pack(anchor="w")
    tk.Label(gc_inner, text=f"₱{grand_lb:,.2f}",
             font=F(22, "bold"), fg=_LIME_MID, bg=_NAVY_DEEP).pack(anchor="w")
    tk.Label(gc_inner,
             text=f"{len(general)} clients  ·  {len(sector_map)} sectors detected",
             font=F(9), fg=_TXT_SOFT, bg=_NAVY_DEEP).pack(anchor="w")

    # ── Per-sector breakdown table ────────────────────────────────────────────
    tk.Label(pad, text="Sector Loan Balance Breakdown",
             font=F(11, "bold"), fg=_TXT_NAVY, bg=_CARD_WHITE).pack(anchor="w", pady=(0, 8))

    tbl_hdr = tk.Frame(pad, bg=_NAVY_MID)
    tbl_hdr.pack(fill="x")
    for col, lbl, w in [
        (0, "Sector",              200),
        (1, "# Clients",            80),
        (2, "Total Loan Balance",  180),
        (3, "% of Total",          100),
        (4, "Avg Loan per Client", 180),
        (5, "Avg Net Income",      160),
        (6, "Risk Profile",        100),
    ]:
        tbl_hdr.columnconfigure(col, minsize=w)
        tk.Label(tbl_hdr, text=lbl, font=F(8, "bold"), fg=_WHITE,
                 bg=_NAVY_MID, padx=10, pady=8, anchor="w"
                 ).grid(row=0, column=col, sticky="ew")

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

    for idx, (sector, n, s_lb, pct, avg_lb, avg_net, risk_label) in enumerate(sector_rows_data):
        row_bg    = _CARD_WHITE if idx % 2 == 0 else _OFF_WHITE
        col_color = _SECTOR_COLORS.get(sector, _NAVY_MID)
        icon      = _SECTOR_ICON.get(sector, "📋")
        row       = tk.Frame(pad, bg=row_bg,
                             highlightbackground=_BORDER_LIGHT, highlightthickness=1)
        row.pack(fill="x")
        for col in range(7):
            row.columnconfigure(col, minsize=[200, 80, 180, 100, 180, 160, 100][col])

        tk.Label(row, text=f"{icon}  {sector}", font=F(9, "bold"),
                 fg=col_color, bg=row_bg, padx=10, pady=10, anchor="w"
                 ).grid(row=0, column=0, sticky="ew")
        tk.Label(row, text=str(n),
                 font=F(9), fg=_TXT_NAVY, bg=row_bg, padx=10, pady=10, anchor="center"
                 ).grid(row=0, column=1, sticky="ew")
        tk.Label(row, text=f"₱{s_lb:,.2f}",
                 font=F(9, "bold"), fg=_TXT_NAVY, bg=row_bg, padx=10, pady=10, anchor="e"
                 ).grid(row=0, column=2, sticky="ew")

        pct_cell = tk.Frame(row, bg=row_bg)
        pct_cell.grid(row=0, column=3, sticky="ew", padx=4, pady=4)
        tk.Label(pct_cell, text=f"{pct:.1f}%", font=F(9, "bold"),
                 fg=col_color, bg=row_bg).pack(anchor="w", padx=6)
        bar_frame = tk.Frame(pct_cell, bg=row_bg)
        bar_frame.pack(fill="x", padx=6)
        bar_w = max(4, int(80 * pct / 100))
        tk.Frame(bar_frame, bg=col_color, height=6, width=bar_w).pack(side="left")
        tk.Frame(bar_frame, bg=_BORDER_LIGHT, height=6,
                 width=80 - bar_w).pack(side="left")

        tk.Label(row, text=f"₱{avg_lb:,.2f}",
                 font=F(9), fg=_TXT_NAVY, bg=row_bg, padx=10, pady=10, anchor="e"
                 ).grid(row=0, column=4, sticky="ew")
        tk.Label(row, text=f"₱{avg_net:,.2f}" if avg_net else "—",
                 font=F(9), fg=_TXT_NAVY, bg=row_bg, padx=10, pady=10, anchor="e"
                 ).grid(row=0, column=5, sticky="ew")

        risk_fg = _RISK_COLOR.get(risk_label, _TXT_SOFT)
        risk_bg = _RISK_BADGE_BG.get(risk_label, _OFF_WHITE)
        tk.Label(row, text=risk_label, font=F(8, "bold"),
                 fg=risk_fg, bg=risk_bg, padx=8, pady=4
                 ).grid(row=0, column=6, sticky="w", padx=8, pady=8)

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
                labels  = [f"{_SECTOR_ICON.get(s, '')}\n{s}\n{p:.1f}%"
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

                short_names = [f"{_SECTOR_ICON.get(s, '')} {s[:22]}" for s in snames]
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

    # ── Per-client table ──────────────────────────────────────────────────────
    tk.Label(pad, text="Individual Client Loan Balance",
             font=F(11, "bold"), fg=_TXT_NAVY, bg=_CARD_WHITE
             ).pack(anchor="w", pady=(16, 8))

    cl_hdr = tk.Frame(pad, bg=_NAVY_MID)
    cl_hdr.pack(fill="x")
    for col, lbl, w in [
        (0, "Client",         180),
        (1, "ID",              60),
        (2, "Sector",         160),
        (3, "Loan Balance",   130),
        (4, "% of Total",      90),
        (5, "Net Income",     120),
        (6, "Current Amort",  130),
        (7, "Risk",            80),
    ]:
        cl_hdr.columnconfigure(col, minsize=w)
        tk.Label(cl_hdr, text=lbl, font=F(8, "bold"), fg=_WHITE,
                 bg=_NAVY_MID, padx=8, pady=8, anchor="w"
                 ).grid(row=0, column=col, sticky="ew")

    clients_sorted = sorted(general, key=lambda r: -(r.get("loan_balance") or 0))
    for idx, rec in enumerate(clients_sorted):
        row_bg    = _CARD_WHITE if idx % 2 == 0 else _OFF_WHITE
        lb        = rec.get("loan_balance") or 0
        net       = rec.get("net_income") or 0
        amrt      = rec.get("current_amort") or 0
        pct       = (lb / grand_lb * 100) if grand_lb > 0 else 0.0
        rl        = rec.get("score_label", "N/A")
        sec       = rec.get("sector", "—")
        col_color = _SECTOR_COLORS.get(sec, _NAVY_MID)
        row       = tk.Frame(pad, bg=row_bg,
                             highlightbackground=_BORDER_LIGHT, highlightthickness=1)
        row.pack(fill="x")
        for c in range(8):
            row.columnconfigure(c, minsize=[180, 60, 160, 130, 90, 120, 130, 80][c])

        tk.Label(row, text=rec["client"][:28], font=F(9, "bold"),
                 fg=_TXT_NAVY, bg=row_bg, padx=8, pady=8, anchor="w"
                 ).grid(row=0, column=0, sticky="ew")
        tk.Label(row, text=rec.get("client_id", "—"), font=F(8),
                 fg=_TXT_SOFT, bg=row_bg, padx=8, pady=8, anchor="center"
                 ).grid(row=0, column=1, sticky="ew")
        tk.Label(row, text=f"{_SECTOR_ICON.get(sec, '')} {sec[:22]}", font=F(8),
                 fg=col_color, bg=row_bg, padx=8, pady=8, anchor="w"
                 ).grid(row=0, column=2, sticky="ew")
        tk.Label(row, text=f"₱{lb:,.2f}", font=F(9, "bold"),
                 fg=_TXT_NAVY, bg=row_bg, padx=8, pady=8, anchor="e"
                 ).grid(row=0, column=3, sticky="ew")

        pct_f = tk.Frame(row, bg=row_bg)
        pct_f.grid(row=0, column=4, sticky="ew", padx=4, pady=4)
        tk.Label(pct_f, text=f"{pct:.2f}%", font=F(8, "bold"),
                 fg=col_color, bg=row_bg).pack(anchor="w", padx=4)
        bw = max(2, int(70 * pct / 100)) if grand_lb > 0 else 2
        bf = tk.Frame(pct_f, bg=row_bg)
        bf.pack(fill="x", padx=4)
        tk.Frame(bf, bg=col_color, height=4, width=bw).pack(side="left")
        tk.Frame(bf, bg=_BORDER_LIGHT, height=4,
                 width=max(0, 70 - bw)).pack(side="left")

        tk.Label(row, text=f"₱{net:,.2f}" if net else "—", font=F(9),
                 fg=_TXT_NAVY, bg=row_bg, padx=8, pady=8, anchor="e"
                 ).grid(row=0, column=5, sticky="ew")
        tk.Label(row, text=f"₱{amrt:,.2f}" if amrt else "—", font=F(9),
                 fg=_TXT_NAVY, bg=row_bg, padx=8, pady=8, anchor="e"
                 ).grid(row=0, column=6, sticky="ew")

        risk_fg = _RISK_COLOR.get(rl, _TXT_SOFT)
        risk_bg = _RISK_BADGE_BG.get(rl, _OFF_WHITE)
        tk.Label(row, text=rl, font=F(7, "bold"),
                 fg=risk_fg, bg=risk_bg, padx=6, pady=3
                 ).grid(row=0, column=7, sticky="w", padx=6, pady=8)

    plt.close("all")


# ══════════════════════════════════════════════════════════════════════════════
#  EXPORT MENU POPUP
# ══════════════════════════════════════════════════════════════════════════════

def _loanbal_show_export_menu(self):
    """Popup menu anchored to the Export button in the loanbal header."""
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
    menu.add_command(
        label="📊  Export Excel — Loan Balance Workbook",
        command=lambda: _loanbal_export_excel(self),
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
#  PDF EXPORT  (Loan Balance focused)
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

    # ── Grand total ───────────────────────────────────────────────────────────
    grand_str = f"₱{grand_lb:,.2f}"
    story.append(Paragraph(
        f"<b>Grand Total Loan Balance: {grand_str}</b>   |   "
        f"{len(general)} clients  ·  {len(sector_map)} sectors",
        bold_s))
    story.append(Spacer(1, 0.4 * cm))

    # ── Sector summary table ──────────────────────────────────────────────────
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

    # ── Per-client table (new page, landscape keeps it wide enough) ───────────
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
    """Convert a reportlab HexColor back to a 6-char hex string."""
    try:
        return f"{int(color.red*255):02X}{int(color.green*255):02X}{int(color.blue*255):02X}"
    except Exception:
        return "000000"


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL EXPORT  (Loan Balance focused)
# ══════════════════════════════════════════════════════════════════════════════

def _loanbal_export_excel(self):
    if not self._lu_all_data:
        messagebox.showwarning("No Data", "Run an analysis first.")
        return
    if not _HAS_OPENPYXL:
        messagebox.showerror(
            "Missing Library",
            "openpyxl is not installed.\nRun:  pip install openpyxl")
        return

    default_name = (
        f"LU_LoanBalance_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
    path = filedialog.asksaveasfilename(
        title="Save Loan Balance Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=default_name,
    )
    if not path:
        return

    try:
        _generate_loanbal_excel(self._lu_all_data,
                                path,
                                filepath=self._lu_filepath or "")
        messagebox.showinfo("Export Complete", f"Excel saved to:\n{path}")
    except Exception as ex:
        messagebox.showerror("Excel Export Error", str(ex))


def _generate_loanbal_excel(all_data, out_path, filepath=""):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    general    = all_data.get("general", [])
    sector_map = all_data.get("sector_map", {})
    totals     = all_data.get("totals", {})
    grand_lb   = totals.get("loan_balance", 0) or 0
    fname      = Path(filepath).name if filepath else "—"
    now        = datetime.now().strftime("%Y-%m-%d %H:%M")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def fill(hex_col):
        return PatternFill("solid", fgColor=hex_col.lstrip("#"))

    def thin_border():
        s = Side(style="thin", color="C5D0E8")
        return Border(left=s, right=s, top=s, bottom=s)

    # ── Exact fills from reference Excel ─────────────────────────────────────
    FILLS = {
        "navy":     fill("#1A3A6B"),   # header bg
        "mist":     fill("#EEF3FB"),   # title row + alternating odd data rows
        "white":    fill("#FFFFFF"),   # alternating even data rows
        # Client Detail row fills (match reference: FFFBF0 for all data rows)
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

    # ══════════════════════════════════════════════════════════════════════════
    #  SHEET 1 — Sector Summary
    #  Matches: A1:G1 merged title (mist bg, navy text, bold, sz13)
    #           A2:G2 merged subtitle (no bg, soft grey text, sz8)
    #           Row 3: navy header, white text, bold, sz9, centered, thin border
    #           Data rows: alternating mist/white, sz9, thin border
    #           Freeze at A2  (reference shows A2, not A4)
    # ══════════════════════════════════════════════════════════════════════════
    ws1 = wb.create_sheet("Sector Summary")

    # Row 1 — merged title
    ws1.merge_cells("A1:G1")
    c = ws1["A1"]
    c.value     = "LU Analysis — Sector vs Loan Balance"
    c.font      = Font(bold=True, size=13, color="0A1628")
    c.fill      = FILLS["mist"]
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[1].height = 22.05

    # Row 2 — merged subtitle
    ws1.merge_cells("A2:G2")
    c = ws1["A2"]
    c.value     = (f"File: {fname}    Generated: {now}    "
                   f"Grand Total Loan Balance: ₱{grand_lb:,.2f}    "
                   f"Clients: {len(general)}")
    c.font      = Font(size=8, color="6B7FA3")
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws1.row_dimensions[2].height = 16.05

    # Row 3 — column headers  (navy bg, white bold sz9, centered, thin border)
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

    # Freeze below header row (reference: A2 — keeps title visible, header scrolls)
    ws1.freeze_panes = "A2"

    # Build sector rows sorted by loan balance descending
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

    # Data rows start at row 4
    # Alternating: even index (0,2,4…) → mist; odd index (1,3,5…) → white
    # (matches reference: row4=mist, row5=white, row6=mist, …)
    for idx, (sector, n, s_lb, pct, avg_lb, avg_net, risk_label) in \
            enumerate(sector_rows_data):
        ri      = 4 + idx
        icon    = _SECTOR_ICON.get(sector, "")
        sec_fc  = SEC_FC.get(sector, "1A2B4A")
        risk_fc = RISK_FC.get(risk_label, "9AAACE")
        row_fill = FILLS["mist"] if idx % 2 == 0 else FILLS["white"]

        # Col A: sector name — sector color, bold, left
        # Col B: # clients  — navy text, normal, center
        # Col C: total lb   — navy text, bold, right, number format
        # Col D: % of total — sector color, bold, center (stored as string "X.X%")
        # Col E: avg lb     — navy text, normal, right, number format
        # Col F: avg net    — navy text, normal, right, number format
        # Col G: risk label — risk color, bold, center
        row_def = [
            (f"{icon} {sector}",            sec_fc,  True,  "left",   None),
            (n,                             "1A2B4A", False, "center", "0"),
            (s_lb,                          "1A2B4A", True,  "right",  NUM_FMT),
            (f"{pct:.1f}%",                 sec_fc,  True,  "center", None),
            (avg_lb,                        "1A2B4A", False, "right",  NUM_FMT),
            (avg_net if avg_net else "—",   "1A2B4A", False, "right",
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

    # ══════════════════════════════════════════════════════════════════════════
    #  SHEET 2 — Client Detail
    #  Matches: Row 1 navy header, white bold sz9, centered, thin border
    #           Data rows: FFFBF0 fill (warm cream — matches reference exactly)
    #           Freeze at A2
    #           Col widths from reference: A=28, B=14, C=12, D=22, E=20,
    #                                      F=14, G=18, H=18(curr amort),
    #                                      I=18(amort hist), J=18(total src),
    #                                      K=14, L=12
    # ══════════════════════════════════════════════════════════════════════════
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

    clients_sorted = sorted(general, key=lambda r: -(r.get("loan_balance") or 0))
    for ri, rec in enumerate(clients_sorted, start=2):
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

        # Reference uses FFFBF0 (warm cream) for ALL data rows uniformly
        row_fill = FILLS["cd_row"]

        # Col A: client name   — navy text, normal, left
        # Col B: client ID     — soft grey, normal, left
        # Col C: PN            — soft grey, normal, left
        # Col D: sector+icon   — sector color, normal, left
        # Col E: loan balance  — navy text, bold, right, number
        # Col F: % of total    — sector color, bold, center (string)
        # Col G: net income    — navy text, normal, right, number / "—"
        # Col H: current amort — navy text, normal, right, number / "—"
        # Col I: amort history — navy text, normal, right, number / "—"
        # Col J: total source  — navy text, normal, right, number / "—"
        # Col K: risk label    — risk color, bold, center
        # Col L: risk score    — navy text, normal, right, number
        row_def = [
            (rec.get("client", ""),          "1A2B4A", False, "left",   None),
            (str(rec.get("client_id", "")),  "6B7FA3", False, "left",   None),
            (str(rec.get("pn", "")),         "6B7FA3", False, "left",   None),
            (f"{icon} {sec}",                sec_fc,  False, "left",   None),
            (lb,                             "1A2B4A", True,  "right",  NUM_FMT),
            (f"{pct:.2f}%",                  sec_fc,  True,  "center", None),
            (net  if net  else "—",          "1A2B4A", False, "right",
             NUM_FMT if net  else None),
            (amrt if amrt else "—",          "1A2B4A", False, "right",
             NUM_FMT if amrt else None),
            (hist if hist else "—",          "1A2B4A", False, "right",
             NUM_FMT if hist else None),
            (src  if src  else "—",          "1A2B4A", False, "right",
             NUM_FMT if src  else None),
            (rl,                             risk_fc,  True,  "center", None),
            (rec.get("score", 0),            "1A2B4A", False, "right",  "0.0"),
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
    cls._build_loanbal_panel     = _build_loanbal_panel_patched
    cls._loanbal_render          = _loanbal_render_patched
    cls._loanbal_show_export_menu = _loanbal_show_export_menu
    cls._loanbal_export_pdf      = _loanbal_export_pdf
    cls._loanbal_export_excel    = _loanbal_export_excel
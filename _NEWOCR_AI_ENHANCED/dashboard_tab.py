"""
dashboard_tab.py — Dashboard panel (LU-driven)
==============================================
Isolated dashboard module so it is easier to edit independently from ui_panels.py.
"""

import json
import tkinter as tk
from pathlib import Path

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import FuncFormatter
import numpy as np

from app_constants import *

_DASH_CACHE_PATH = Path(__file__).with_name(".cache") / "dashboard_lu_snapshot.json"

# ── Design tokens ─────────────────────────────────────────────────────────────
_BG       = "#F0F3FA"
_CARD_BG  = "#FFFFFF"
_NAVY     = "#1A2B5F"
_NAVY_MID = "#2E4A9E"
_ACCENT   = "#3D6EE8"
_RED      = "#E84040"
_GREEN    = "#22C87A"
_GOLD     = "#F5A623"
_SOFT     = "#8A97B5"
_BORDER   = "#DDE3F0"

# Original bar chart colours (unchanged)
_INDUSTRY_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b",
    "#e377c2", "#7f7f7f", "#bcbd22", "#17becf", "#3F51B5", "#009688"
]

plt.rcParams.update({
    "font.family":       "Segoe UI",
    "axes.spines.top":   False,
    "axes.spines.right": False,
    "axes.grid":         True,
    "grid.color":        "#E8EDF8",
    "grid.linewidth":    0.8,
    "axes.facecolor":    _CARD_BG,
    "figure.facecolor":  _CARD_BG,
    "text.color":        _NAVY,
    "axes.labelcolor":   _SOFT,
    "xtick.color":       _SOFT,
    "ytick.color":       _SOFT,
    "xtick.labelsize":   8,
    "ytick.labelsize":   8,
})

# ── Helpers ───────────────────────────────────────────────────────────────────

def _F(size=10, weight="normal"):
    return ("Segoe UI", size, weight)


_KPI_ICONS = {
    "clients":      "👥",
    "high_risk":    "⚠️",
    "industries":   "🏭",
    "loan_balance": "💰",
}

_DETAIL_ICONS = {
    "low_risk":   "✅",
    "high_ratio": "📊",
    "avg_loan":   "🏦",
    "avg_net":    "💵",
    "top_sector": "🏆",
}

_KPI_ACCENTS = {
    "clients":      _ACCENT,
    "high_risk":    _RED,
    "industries":   _GOLD,
    "loan_balance": _GREEN,
}


def _make_kpi_card(parent, key, label, value, delta):
    accent = _KPI_ACCENTS.get(key, _ACCENT)
    outer  = tk.Frame(parent, bg=_CARD_BG, highlightbackground=_BORDER, highlightthickness=1)
    # Coloured top accent bar
    tk.Frame(outer, bg=accent, height=4).pack(fill="x")
    body = tk.Frame(outer, bg=_CARD_BG)
    body.pack(fill="both", expand=True, padx=14, pady=10)
    # Icon + label row
    top_row = tk.Frame(body, bg=_CARD_BG)
    top_row.pack(fill="x")
    tk.Label(top_row, text=_KPI_ICONS.get(key, ""), font=_F(14), bg=_CARD_BG).pack(side="left", padx=(0, 6))
    tk.Label(top_row, text=label.upper(), font=_F(7, "bold"), fg=_SOFT, bg=_CARD_BG).pack(side="left", anchor="s", pady=(4, 0))
    # Large value
    value_lbl = tk.Label(body, text=value, font=("Segoe UI", 22, "bold"), fg=_NAVY, bg=_CARD_BG)
    value_lbl.pack(anchor="w", pady=(4, 0))
    # Divider
    tk.Frame(body, bg=_BORDER, height=1).pack(fill="x", pady=6)
    # Delta / status
    delta_lbl = tk.Label(body, text=delta, font=_F(8, "bold"), fg=_SOFT, bg=_CARD_BG)
    delta_lbl.pack(anchor="w")
    return outer, value_lbl, delta_lbl


def _make_detail_card(parent, key, label, value):
    outer = tk.Frame(parent, bg=_CARD_BG, highlightbackground=_BORDER, highlightthickness=1)
    tk.Frame(outer, bg=_SOFT, height=3).pack(fill="x")
    body = tk.Frame(outer, bg=_CARD_BG)
    body.pack(fill="both", expand=True, padx=10, pady=8)
    tk.Label(body, text=f"{_DETAIL_ICONS.get(key, '')}  {label}",
             font=_F(8), fg=_SOFT, bg=_CARD_BG).pack(anchor="w")
    v = tk.Label(body, text=value, font=("Segoe UI", 13, "bold"), fg=_NAVY, bg=_CARD_BG)
    v.pack(anchor="w", pady=(2, 0))
    return outer, v


# ── Matplotlib line chart ─────────────────────────────────────────────────────

def _build_line_chart_widget(parent):
    fig, ax = plt.subplots(figsize=(5.4, 2.8), dpi=96)
    fig.subplots_adjust(left=0.10, right=0.97, top=0.82, bottom=0.14)
    ax.set_facecolor(_CARD_BG)
    fig.set_facecolor(_CARD_BG)
    canvas = FigureCanvasTkAgg(fig, master=parent)
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=12, pady=(6, 4))
    return fig, ax, canvas


def _redraw_line_chart(fig, ax, series_loan, series_net, industry_names=None):
    ax.clear()
    n = max(2, len(series_loan), len(series_net))

    def _pad(s):
        s = list(s)
        s.extend([0.0] * (n - len(s)))
        return np.array(s[:n], dtype=float)

    loan = _pad(series_loan)
    net  = _pad(series_net)
    xs   = np.arange(n)

    ax.fill_between(xs, loan, alpha=0.12, color=_ACCENT)
    ax.fill_between(xs, net,  alpha=0.10, color=_GREEN)

    ax.plot(xs, loan, color=_ACCENT, linewidth=2.2, marker="o", markersize=4,
            markerfacecolor=_CARD_BG, markeredgewidth=1.8, label="Loan Exposure", zorder=3)
    ax.plot(xs, net,  color=_GREEN,  linewidth=2.0, marker="s", markersize=3.5,
            markerfacecolor=_CARD_BG, markeredgewidth=1.6, linestyle="--",
            label="Avg Net Income", zorder=3)

    ax.set_xlim(-0.4, n - 0.6)
    ax.yaxis.set_major_formatter(
        FuncFormatter(lambda v, _: f"₱{v/1e6:.1f}M" if v >= 1e6 else f"₱{v/1e3:.0f}K")
    )
    ax.set_xticks(xs)
    if industry_names and len(industry_names) >= n:
        short = [nm[:12] + "…" if len(nm) > 13 else nm for nm in industry_names[:n]]
        ax.set_xticklabels(short, rotation=25, ha="right", fontsize=7)
    else:
        ax.set_xticklabels([str(i + 1) for i in xs], fontsize=8)

    ax.spines["left"].set_color(_BORDER)
    ax.spines["bottom"].set_color(_BORDER)
    ax.tick_params(axis="both", which="both", length=0)
    ax.legend(fontsize=7.5, frameon=False, loc="upper right", labelcolor=_SOFT, ncol=2)
    try:
        widget = fig.canvas.get_tk_widget()
        if widget.winfo_exists():
            fig.canvas.draw_idle()
    except Exception:
        pass


# ── Main panel builder ────────────────────────────────────────────────────────

def _build_dashboard_panel(self, parent):
    self._dashboard_frame = tk.Frame(parent, bg=_BG)

    canvas_outer = tk.Frame(self._dashboard_frame, bg=_BG)
    canvas_outer.pack(fill="both", expand=True)
    dash_canvas = tk.Canvas(canvas_outer, bg=_BG, highlightthickness=0)
    dash_scroll = tk.Scrollbar(canvas_outer, orient="vertical", command=dash_canvas.yview)
    dash_canvas.configure(yscrollcommand=dash_scroll.set)
    dash_scroll.pack(side="right", fill="y")
    dash_canvas.pack(side="left", fill="both", expand=True)

    body = tk.Frame(dash_canvas, bg=_BG)
    body_id = dash_canvas.create_window((0, 0), window=body, anchor="nw")
    body.bind("<Configure>", lambda _e=None: dash_canvas.configure(scrollregion=dash_canvas.bbox("all")))
    dash_canvas.bind("<Configure>", lambda e: dash_canvas.itemconfigure(body_id, width=e.width))

    # ── Header ────────────────────────────────────────────────────────────────
    head = tk.Frame(body, bg=_BG)
    head.pack(fill="x", padx=20, pady=(16, 10))
    tk.Label(head, text="Performance Overview", font=("Segoe UI", 17, "bold"),
             fg=_NAVY, bg=_BG).pack(side="left")
    self._dash_subtitle_lbl = tk.Label(head, text="Waiting for LU Analysis data…",
                                        font=_F(9), fg=_SOFT, bg=_BG)
    self._dash_subtitle_lbl.pack(side="right", anchor="s", pady=4)

    # ── KPI row 1 ─────────────────────────────────────────────────────────────
    kpi_row = tk.Frame(body, bg=_BG)
    kpi_row.pack(fill="x", padx=20, pady=(0, 8))
    kpi_defs = [
        ("clients",      "Total Clients",      "0",  "No LU file loaded"),
        ("high_risk",    "High Risk Clients",  "0",  "No LU file loaded"),
        ("industries",   "Industries",         "0",  "No LU file loaded"),
        ("loan_balance", "Total Loan Balance", "₱0", "No LU file loaded"),
    ]
    self._dash_value_lbls = {}
    self._dash_delta_lbls = {}
    for idx, (key, label, value, delta) in enumerate(kpi_defs):
        outer, v_lbl, d_lbl = _make_kpi_card(kpi_row, key, label, value, delta)
        outer.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 9, 0))
        kpi_row.grid_columnconfigure(idx, weight=1)
        self._dash_value_lbls[key] = v_lbl
        self._dash_delta_lbls[key] = d_lbl

    # ── KPI row 2 ─────────────────────────────────────────────────────────────
    detail_row = tk.Frame(body, bg=_BG)
    detail_row.pack(fill="x", padx=20, pady=(0, 12))
    detail_defs = [
        ("low_risk",   "Low Risk Clients", "0"),
        ("high_ratio", "High Risk Ratio",  "0.0%"),
        ("avg_loan",   "Avg Loan / Client","₱0"),
        ("avg_net",    "Avg Net Income",   "₱0"),
        ("top_sector", "Top Sector",       "—"),
    ]
    self._dash_detail_lbls = {}
    for idx, (key, label, value) in enumerate(detail_defs):
        outer, v_lbl = _make_detail_card(detail_row, key, label, value)
        outer.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 8, 0))
        detail_row.grid_columnconfigure(idx, weight=1)
        self._dash_detail_lbls[key] = v_lbl

    # ── Chart row ─────────────────────────────────────────────────────────────
    charts = tk.Frame(body, bg=_BG)
    charts.pack(fill="both", expand=True, padx=20, pady=(0, 12))
    charts.grid_columnconfigure(0, weight=2)
    charts.grid_columnconfigure(1, weight=1)

    # Left — matplotlib line chart
    left_card = tk.Frame(charts, bg=_CARD_BG, highlightbackground=_BORDER, highlightthickness=1)
    left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    tk.Frame(left_card, bg=_ACCENT, height=4).pack(fill="x")
    top_left = tk.Frame(left_card, bg=_CARD_BG)
    top_left.pack(fill="x", padx=14, pady=(12, 0))
    tk.Label(top_left, text="At-Risk Industries", font=_F(11, "bold"), fg=_NAVY_MID, bg=_CARD_BG).pack(side="left")
    self._dash_left_note_lbl = tk.Label(top_left, text="Loan Exposure vs Avg Net Income",
                                         font=_F(8), fg=_SOFT, bg=_CARD_BG)
    self._dash_left_note_lbl.pack(side="right")

    self._dash_line_fig, self._dash_line_ax, _ = _build_line_chart_widget(left_card)

    self._dash_at_risk_names_lbl = tk.Label(
        left_card, text="At-risk industries: —",
        font=_F(8), fg=_SOFT, bg=_CARD_BG, justify="left", wraplength=700
    )
    self._dash_at_risk_names_lbl.pack(fill="x", padx=12, pady=(0, 10))

    # Right — ORIGINAL tkinter Canvas bar chart (completely unchanged)
    right_card = tk.Frame(charts, bg=_CARD_BG, highlightbackground=_BORDER, highlightthickness=1)
    right_card.grid(row=0, column=1, sticky="nsew")
    tk.Frame(right_card, bg=_NAVY_MID, height=4).pack(fill="x")
    top_right = tk.Frame(right_card, bg=_CARD_BG)
    top_right.pack(fill="x", padx=14, pady=(12, 0))
    tk.Label(top_right, text="Risk by Industry", font=_F(11, "bold"), fg=_NAVY, bg=_CARD_BG).pack(side="left")
    self._dash_right_note_lbl = tk.Label(top_right, text="High vs Low Clients",
                                          font=_F(8), fg=_SOFT, bg=_CARD_BG)
    self._dash_right_note_lbl.pack(side="right")

    rchart = tk.Canvas(right_card, bg=_CARD_BG, highlightthickness=0, height=250)
    rchart.pack(fill="both", expand=True, padx=12, pady=(6, 10))

    def _draw_bar_chart(_e=None, cv=rchart):
        cv.delete("all")
        w, h = max(cv.winfo_width(), 260), max(cv.winfo_height(), 200)
        pad_l, pad_r, pad_t, pad_b = 26, 16, 12, 12
        plot_w = w - pad_l - pad_r
        plot_h = h - pad_t - pad_b
        lows  = list(getattr(self, "_dash_industry_low",   []))
        highs = list(getattr(self, "_dash_industry_high",  []))
        names = list(getattr(self, "_dash_industry_names", []))
        n = max(1, len(lows), len(highs))
        if len(lows)  < n: lows.extend([0]    * (n - len(lows)))
        if len(highs) < n: highs.extend([0]   * (n - len(highs)))
        if len(names) < n: names.extend(["—"] * (n - len(names)))
        total_max = max(1, max((l + h) for l, h in zip(lows, highs)))
        row_h = max(18, plot_h / max(1, n))
        bar_h = max(8, row_h * 0.58)
        y = pad_t + (row_h - bar_h) / 2
        for idx, (lo, hi, _nm) in enumerate(zip(lows, highs, names), 1):
            cv.create_text(8, y + (bar_h / 2), text=str(idx), fill=_SOFT,
                           anchor="w", font=("Segoe UI", 7, "bold"))
            lw = (lo / total_max) * plot_w
            hw = (hi / total_max) * plot_w
            x0 = pad_l
            base_color = _INDUSTRY_COLORS[idx % len(_INDUSTRY_COLORS)]
            cv.create_rectangle(x0, y, x0 + lw, y + bar_h, fill=base_color, outline="")
            cv.create_rectangle(x0 + lw, y, x0 + lw + hw, y + bar_h, fill="#A7DCE8", outline="")
            cv.create_line(pad_l, y + bar_h + 1, pad_l + plot_w, y + bar_h + 1, fill="#EDF2F8")
            y += row_h

    rchart.bind("<Configure>", _draw_bar_chart)
    self._dash_draw_bar = _draw_bar_chart

    self._dash_right_legend = tk.Frame(right_card, bg=_CARD_BG)
    self._dash_right_legend.pack(fill="x", padx=12, pady=(0, 10))

    # ── High-risk table ───────────────────────────────────────────────────────
    table_wrap = tk.Frame(body, bg=_CARD_BG, highlightbackground=_BORDER, highlightthickness=1)
    table_wrap.pack(fill="both", expand=False, padx=20, pady=(0, 18))
    tk.Frame(table_wrap, bg=_RED, height=4).pack(fill="x")
    thdr = tk.Frame(table_wrap, bg=_CARD_BG)
    thdr.pack(fill="x", padx=14, pady=(10, 4))
    tk.Label(thdr, text="⚠️  Highest-Risk Clients", font=_F(11, "bold"), fg=_NAVY, bg=_CARD_BG).pack(side="left")
    self._dash_table_count_lbl = tk.Label(thdr, text="", font=_F(8), fg=_SOFT, bg=_CARD_BG)
    self._dash_table_count_lbl.pack(side="right")
    tk.Label(table_wrap, text="Basis: only HIGH-risk clients, ranked by lowest net income.",
             font=_F(8), fg=_SOFT, bg=_CARD_BG).pack(anchor="w", padx=14, pady=(0, 6))

    cols = tk.Frame(table_wrap, bg=_NAVY, height=30)
    cols.pack(fill="x", padx=12)
    cols.pack_propagate(False)
    self._dash_table_specs = [
        ("Client",       "w"),
        ("Industry",     "w"),
        ("Risk",         "center"),
        ("Loan Balance", "center"),
        ("Net Income",   "center"),
    ]
    self._dash_table_col_weights = (32, 34, 10, 12, 12)
    for i, w in enumerate(self._dash_table_col_weights):
        cols.grid_columnconfigure(i * 2, weight=w, uniform="dash_tbl")
        if i < len(self._dash_table_col_weights) - 1:
            cols.grid_columnconfigure(i * 2 + 1, weight=0)
    for idx, (text, anchor) in enumerate(self._dash_table_specs):
        tk.Label(cols, text=text, anchor=anchor, font=_F(8, "bold"),
                 fg="#FFFFFF", bg=_NAVY, padx=8
                 ).grid(row=0, column=idx * 2, sticky="nsew", pady=7)
        if idx < len(self._dash_table_specs) - 1:
            tk.Frame(cols, bg="#3A4B7A", width=1).grid(row=0, column=idx * 2 + 1, sticky="ns", pady=5)

    self._dash_table_rows = tk.Frame(table_wrap, bg=_CARD_BG)
    self._dash_table_rows.pack(fill="x", padx=12, pady=(0, 12))
    for i, w in enumerate(self._dash_table_col_weights):
        self._dash_table_rows.grid_columnconfigure(i * 2, weight=w, uniform="dash_tbl")
        if i < len(self._dash_table_col_weights) - 1:
            self._dash_table_rows.grid_columnconfigure(i * 2 + 1, weight=0)

    def _wheel(e):
        if dash_canvas.winfo_exists():
            dash_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
    dash_canvas.bind("<MouseWheel>", _wheel)

    self._dash_cache = _load_dashboard_cache()
    self._refresh_dashboard_from_lu()


# ── Refresh ───────────────────────────────────────────────────────────────────

def _refresh_dashboard_from_lu(self):
    all_data = getattr(self, "_lu_all_data", None) or {}
    general  = list(all_data.get("general") or [])
    totals   = dict(all_data.get("totals")  or {})
    cache    = getattr(self, "_dash_cache",  None) or {}

    if not general and cache:
        clients    = int(cache.get("clients")    or 0)
        high       = int(cache.get("high")       or 0)
        industries = int(cache.get("industries") or 0)
        total_loan = float(cache.get("total_loan") or 0.0)
        total_net  = float(cache.get("total_net")  or 0.0)
        has_data   = clients > 0
        self._dash_series_loan    = list(cache.get("series_loan")    or [0.0, 0.0])
        self._dash_series_net     = list(cache.get("series_net")     or [0.0, 0.0])
        self._dash_industry_low   = list(cache.get("industry_low")   or [0])
        self._dash_industry_high  = list(cache.get("industry_high")  or [0])
        self._dash_industry_names = list(cache.get("industry_names") or ["—"])
        at_risk_names = list(cache.get("at_risk_names") or [])
        high_rows     = list(cache.get("high_rows")     or [])
        top_sector    = cache.get("top_sector") or "—"
    else:
        clients    = len(general)
        high       = sum(1 for r in general if str(r.get("score_label") or "").upper() == "HIGH")
        industries = len(all_data.get("unique_industries") or [])
        total_loan = float(totals.get("loan_balance") or 0.0)
        total_net  = float(totals.get("total_net")    or 0.0)
        has_data   = clients > 0

        industry_risk_counts = {}
        industry_finance     = {}
        for rec in general:
            ind = str(rec.get("industry") or "Unknown Industry").strip() or "Unknown Industry"
            industry_risk_counts.setdefault(ind, {"high": 0, "low": 0})
            industry_finance.setdefault(ind, {"loan": 0.0, "net": 0.0, "count": 0})
            if str(rec.get("score_label") or "").upper() == "HIGH":
                industry_risk_counts[ind]["high"] += 1
            else:
                industry_risk_counts[ind]["low"] += 1
            industry_finance[ind]["loan"]  += float(rec.get("loan_balance") or 0.0)
            industry_finance[ind]["net"]   += float(rec.get("net_income")   or 0.0)
            industry_finance[ind]["count"] += 1

        top_industries = sorted(
            industry_risk_counts.items(),
            key=lambda kv: (kv[1]["high"], kv[1]["high"] + kv[1]["low"]),
            reverse=True
        )[:12]

        at_risk_for_line = [(n, v) for (n, v) in top_industries if int(v.get("high", 0)) > 0] or top_industries
        self._dash_series_loan = [industry_finance[n]["loan"] for n, _ in at_risk_for_line] or [0.0, 0.0]
        self._dash_series_net  = [
            (industry_finance[n]["net"] / max(1, industry_finance[n]["count"]))
            for n, _ in at_risk_for_line
        ] or [0.0, 0.0]

        if top_industries:
            nm = top_industries[0][0]
            self._dash_left_note_lbl.config(text=f"Top risk: {nm[:32]}{'…' if len(nm) > 32 else ''}")

        sector_counts = {}
        for rec in general:
            sec = str(rec.get("sector") or "Other")
            sector_counts.setdefault(sec, {"high": 0, "low": 0})
            if str(rec.get("score_label") or "").upper() == "HIGH":
                sector_counts[sec]["high"] += 1
            else:
                sector_counts[sec]["low"] += 1
        top_items = sorted(sector_counts.items(), key=lambda kv: kv[1]["high"] + kv[1]["low"], reverse=True)[:7]

        self._dash_industry_low   = [v["low"]  for _, v in top_industries[:7]] or [0]
        self._dash_industry_high  = [v["high"] for _, v in top_industries[:7]] or [0]
        self._dash_industry_names = [name      for name, _ in top_industries[:7]] or ["—"]
        at_risk_names = [name for name, v in top_industries if int(v.get("high", 0)) > 0]
        top_sector    = top_items[0][0] if top_items else "—"
        high_rows = sorted(
            [r for r in general if str(r.get("score_label") or "").upper() == "HIGH"],
            key=lambda r: (float(r.get("net_income") or 0.0), float(r.get("loan_balance") or 0.0))
        )[:8]

    # Subtitle
    if has_data:
        cur = Path(getattr(self, "_lu_filepath", "")).name if getattr(self, "_lu_filepath", None) else ""
        self._dash_subtitle_lbl.config(text=(f"From LU file: {cur}" if cur else "From saved LU snapshot"))
    else:
        self._dash_subtitle_lbl.config(text="Waiting for LU Analysis data…")

    # KPI card values
    self._dash_value_lbls["clients"].config(text=f"{clients:,}")
    self._dash_value_lbls["high_risk"].config(text=f"{high:,}")
    self._dash_value_lbls["industries"].config(text=f"{industries:,}")
    self._dash_value_lbls["loan_balance"].config(text=f"₱{total_loan:,.0f}")

    high_ratio = (high / clients * 100.0) if clients else 0.0
    low        = max(0, clients - high)
    net_ratio  = (total_net / total_loan * 100.0) if total_loan else 0.0
    avg_loan   = (total_loan / clients) if clients else 0.0
    avg_net    = (total_net  / clients) if clients else 0.0

    self._dash_delta_lbls["clients"].config(
        text=(f"{clients:,} loaded" if has_data else "No LU file loaded"),
        fg=(_GREEN if has_data else _SOFT))
    self._dash_delta_lbls["high_risk"].config(
        text=(f"{high_ratio:.1f}% of clients" if has_data else "No LU file loaded"),
        fg=(_RED if high > 0 else _GREEN if has_data else _SOFT))
    self._dash_delta_lbls["industries"].config(
        text=(f"{industries:,} active tags" if has_data else "No LU file loaded"),
        fg=(_GOLD if has_data else _SOFT))
    self._dash_delta_lbls["loan_balance"].config(
        text=(f"Net income {net_ratio:.1f}%" if has_data else "No LU file loaded"),
        fg=(_GREEN if has_data else _SOFT))

    # Detail cards
    self._dash_detail_lbls["low_risk"].config(text=f"{low:,}")
    self._dash_detail_lbls["high_ratio"].config(text=f"{high_ratio:.1f}%")
    self._dash_detail_lbls["avg_loan"].config(text=f"₱{avg_loan:,.0f}")
    self._dash_detail_lbls["avg_net"].config(text=f"₱{avg_net:,.0f}")
    self._dash_detail_lbls["top_sector"].config(text=top_sector)

    if hasattr(self, "_dash_at_risk_names_lbl"):
        self._dash_at_risk_names_lbl.config(
            text=("At-risk industries: " + ", ".join(at_risk_names)) if at_risk_names else "At-risk industries: none"
        )

    # Legend (right card — original logic unchanged)
    if hasattr(self, "_dash_right_legend"):
        for w in self._dash_right_legend.winfo_children():
            w.destroy()
        if getattr(self, "_dash_industry_names", None):
            total_clients = max(1, sum(self._dash_industry_high) + sum(self._dash_industry_low))
            for idx, name in enumerate(self._dash_industry_names):
                hi  = self._dash_industry_high[idx] if idx < len(self._dash_industry_high) else 0
                lo  = self._dash_industry_low[idx]  if idx < len(self._dash_industry_low)  else 0
                pct = ((hi + lo) / total_clients) * 100.0
                item = tk.Frame(self._dash_right_legend, bg=_CARD_BG)
                item.grid(row=idx // 2, column=idx % 2, sticky="w", padx=(0, 10), pady=2)
                sw = tk.Canvas(item, width=14, height=10, bg=_CARD_BG, highlightthickness=0)
                sw.pack(side="left", padx=(0, 5))
                sw.create_rectangle(0, 2, 14, 8,
                                    fill=_INDUSTRY_COLORS[idx % len(_INDUSTRY_COLORS)], outline="")
                short = name if len(name) <= 34 else name[:33] + "…"
                tk.Label(item, text=f"{short}  {pct:.1f}%",
                         font=_F(8), fg=_NAVY, bg=_CARD_BG).pack(side="left")
        else:
            tk.Label(self._dash_right_legend, text="No industries to show",
                     font=_F(8), fg=_SOFT, bg=_CARD_BG).pack(anchor="w")

    # Redraw matplotlib line chart
    _redraw_line_chart(
        self._dash_line_fig, self._dash_line_ax,
        getattr(self, "_dash_series_loan", [0.0, 0.0]),
        getattr(self, "_dash_series_net",  [0.0, 0.0]),
        getattr(self, "_dash_industry_names", None),
    )

    # Redraw original tkinter bar chart
    if hasattr(self, "_dash_draw_bar"):
        self._dash_draw_bar()

    _render_high_risk_table(self, high_rows)


# ── High-risk table renderer ──────────────────────────────────────────────────

def _render_high_risk_table(self, rows):
    for w in self._dash_table_rows.winfo_children():
        w.destroy()
    self._dash_table_count_lbl.config(text=f"{len(rows)} shown")
    if not rows:
        tk.Label(self._dash_table_rows, text="No high-risk clients to display.",
                 font=_F(9), fg=_SOFT, bg=_CARD_BG
                 ).grid(row=0, column=0, columnspan=9, sticky="w", padx=6, pady=10)
        return
    for r, rec in enumerate(rows):
        bg = _CARD_BG if r % 2 == 0 else "#F5F7FC"
        vals = [
            (str(rec.get("client")   or "—"), "w"),
            (str(rec.get("industry") or "—"), "w"),
            ("HIGH",                           "center"),
            (f"₱{float(rec.get('loan_balance') or 0.0):,.0f}", "center"),
            (f"₱{float(rec.get('net_income')   or 0.0):,.0f}", "center"),
        ]
        for c, (text, anchor) in enumerate(vals):
            fg   = _RED if c == 2 else _NAVY
            font = _F(8, "bold") if c == 2 else _F(8)
            tk.Label(self._dash_table_rows, text=text, anchor=anchor, font=font,
                     fg=fg, bg=bg, padx=8, pady=5
                     ).grid(row=r * 2, column=c * 2, sticky="nsew")
            if c < len(vals) - 1:
                tk.Frame(self._dash_table_rows, bg=_BORDER, width=1
                         ).grid(row=r * 2, column=c * 2 + 1, sticky="ns")
        tk.Frame(self._dash_table_rows, bg=_BORDER, height=1
                 ).grid(row=r * 2 + 1, column=0, columnspan=9, sticky="ew")


# ── Snapshot persist / load ───────────────────────────────────────────────────

def _persist_dashboard_snapshot_from_lu(self):
    all_data = getattr(self, "_lu_all_data", None) or {}
    general  = list(all_data.get("general") or [])
    if not general:
        return
    totals = dict(all_data.get("totals") or {})

    industry_risk_counts = {}
    industry_finance     = {}
    for rec in general:
        ind = str(rec.get("industry") or "Unknown Industry").strip() or "Unknown Industry"
        industry_risk_counts.setdefault(ind, {"high": 0, "low": 0})
        industry_finance.setdefault(ind, {"loan": 0.0, "net": 0.0, "count": 0})
        if str(rec.get("score_label") or "").upper() == "HIGH":
            industry_risk_counts[ind]["high"] += 1
        else:
            industry_risk_counts[ind]["low"] += 1
        industry_finance[ind]["loan"]  += float(rec.get("loan_balance") or 0.0)
        industry_finance[ind]["net"]   += float(rec.get("net_income")   or 0.0)
        industry_finance[ind]["count"] += 1

    top_industries = sorted(
        industry_risk_counts.items(),
        key=lambda kv: (kv[1]["high"], kv[1]["high"] + kv[1]["low"]),
        reverse=True
    )[:12]

    sector_counts = {}
    for rec in general:
        sec = str(rec.get("sector") or "Other")
        sector_counts.setdefault(sec, {"high": 0, "low": 0})
        if str(rec.get("score_label") or "").upper() == "HIGH":
            sector_counts[sec]["high"] += 1
        else:
            sector_counts[sec]["low"] += 1
    top_sectors = sorted(sector_counts.items(), key=lambda kv: kv[1]["high"] + kv[1]["low"], reverse=True)[:7]

    at_risk_for_line = [(n, v) for (n, v) in top_industries if int(v.get("high", 0)) > 0] or top_industries
    at_risk_names    = [name for name, v in top_industries if int(v.get("high", 0)) > 0]
    high_rows = sorted(
        [r for r in general if str(r.get("score_label") or "").upper() == "HIGH"],
        key=lambda r: (float(r.get("net_income") or 0.0), float(r.get("loan_balance") or 0.0))
    )[:8]

    payload = {
        "clients":        len(general),
        "high":           sum(1 for r in general if str(r.get("score_label") or "").upper() == "HIGH"),
        "industries":     len(all_data.get("unique_industries") or []),
        "total_loan":     float(totals.get("loan_balance") or 0.0),
        "total_net":      float(totals.get("total_net")    or 0.0),
        "series_loan":    [industry_finance[n]["loan"] for n, _ in at_risk_for_line] or [0.0, 0.0],
        "series_net":     [(industry_finance[n]["net"] / max(1, industry_finance[n]["count"])) for n, _ in at_risk_for_line] or [0.0, 0.0],
        "industry_low":   [v["low"]  for _, v in top_industries[:7]] or [0],
        "industry_high":  [v["high"] for _, v in top_industries[:7]] or [0],
        "industry_names": [name      for name, _ in top_industries[:7]] or ["—"],
        "at_risk_names":  at_risk_names,
        "top_sector":     top_sectors[0][0] if top_sectors else "—",
        "high_rows": [
            {
                "client":       str(r.get("client")       or ""),
                "industry":     str(r.get("industry")     or ""),
                "loan_balance": float(r.get("loan_balance") or 0.0),
                "net_income":   float(r.get("net_income")   or 0.0),
            }
            for r in high_rows
        ],
    }
    _save_dashboard_cache(payload)
    self._dash_cache = payload


def _load_dashboard_cache():
    try:
        if not _DASH_CACHE_PATH.exists():
            return {}
        return json.loads(_DASH_CACHE_PATH.read_text(encoding="utf-8") or "{}")
    except Exception:
        return {}


def _save_dashboard_cache(payload: dict):
    try:
        _DASH_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        _DASH_CACHE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def attach(cls):
    cls._build_dashboard_panel              = _build_dashboard_panel
    cls._refresh_dashboard_from_lu          = _refresh_dashboard_from_lu
    cls._persist_dashboard_snapshot_from_lu = _persist_dashboard_snapshot_from_lu
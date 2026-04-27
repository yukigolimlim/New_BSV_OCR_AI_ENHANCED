"""
dashboard_tab.py — Dashboard panel (LU-driven)
==============================================
Isolated dashboard module so it is easier to edit independently from ui_panels.py.
"""

import json
import tkinter as tk
from pathlib import Path

from app_constants import *

_DASH_CACHE_PATH = Path(__file__).with_name(".cache") / "dashboard_lu_snapshot.json"
_INDUSTRY_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b",
    "#e377c2", "#7f7f7f", "#bcbd22", "#17becf", "#3F51B5", "#009688"
]

def _add_container_header_bar(widget, height=3):
    bar = tk.Frame(widget, bg=NAVY_DEEP, height=height)
    bar.pack(fill="x", side="top")
    bar.pack_propagate(False)


def _build_dashboard_panel(self, parent):
    self._dashboard_frame = tk.Frame(parent, bg=WHITE)

    canvas_outer = tk.Frame(self._dashboard_frame, bg=WHITE)
    canvas_outer.pack(fill="both", expand=True)
    dash_canvas = tk.Canvas(canvas_outer, bg=WHITE, highlightthickness=0)
    dash_scroll = tk.Scrollbar(canvas_outer, orient="vertical", command=dash_canvas.yview)
    dash_canvas.configure(yscrollcommand=dash_scroll.set)
    dash_scroll.pack(side="right", fill="y")
    dash_canvas.pack(side="left", fill="both", expand=True)

    body = tk.Frame(dash_canvas, bg="#F4F6FA")
    body_id = dash_canvas.create_window((0, 0), window=body, anchor="nw")
    body.bind("<Configure>", lambda _e=None: dash_canvas.configure(scrollregion=dash_canvas.bbox("all")))
    dash_canvas.bind("<Configure>", lambda e: dash_canvas.itemconfigure(body_id, width=e.width))

    head = tk.Frame(body, bg="#F4F6FA")
    head.pack(fill="x", padx=18, pady=(14, 10))
    tk.Label(head, text="Performance Overview", font=F(16, "bold"),
             fg=NAVY_DEEP, bg="#F4F6FA").pack(side="left")
    self._dash_subtitle_lbl = tk.Label(
        head, text="Waiting for LU Analysis data", font=F(9),
        fg=TXT_SOFT, bg="#F4F6FA"
    )
    self._dash_subtitle_lbl.pack(side="right")

    # KPI cards (row 1)
    cards = tk.Frame(body, bg="#F4F6FA")
    cards.pack(fill="x", padx=18, pady=(0, 8))
    card_data = [
        ("clients", "Total Clients", "0", "No LU file loaded", TXT_SOFT),
        ("high_risk", "High Risk Clients", "0", "No LU file loaded", TXT_SOFT),
        ("industries", "Industries", "0", "No LU file loaded", TXT_SOFT),
        ("loan_balance", "Total Loan Balance", "P0", "No LU file loaded", TXT_SOFT),
    ]
    self._dash_value_lbls = {}
    self._dash_delta_lbls = {}
    for idx, (key, label, value, delta, dcolor) in enumerate(card_data):
        card = tk.Frame(cards, bg=WHITE, highlightbackground=BORDER_LIGHT, highlightthickness=1)
        card.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 8, 0))
        cards.grid_columnconfigure(idx, weight=1)
        _add_container_header_bar(card)
        hdr = tk.Frame(card, bg=WHITE)
        hdr.pack(fill="x", padx=14, pady=(10, 4))
        tk.Label(hdr, text=label.upper(), font=F(7, "bold"), fg=TXT_SOFT, bg=WHITE).pack(side="left")
        value_lbl = tk.Label(card, text=value, font=F(16, "bold"), fg=TXT_NAVY, bg=WHITE)
        value_lbl.pack(anchor="w", padx=14)
        self._dash_value_lbls[key] = value_lbl
        foot = tk.Frame(card, bg=WHITE)
        foot.pack(fill="x", padx=14, pady=(3, 10))
        tk.Label(foot, text=label, font=F(8), fg=TXT_SOFT, bg=WHITE).pack(side="left")
        delta_lbl = tk.Label(foot, text=delta, font=F(9, "bold"), fg=dcolor, bg=WHITE)
        delta_lbl.pack(side="right")
        self._dash_delta_lbls[key] = delta_lbl

    # KPI cards (row 2 - more detail)
    details = tk.Frame(body, bg="#F4F6FA")
    details.pack(fill="x", padx=18, pady=(0, 12))
    detail_cards = [
        ("low_risk", "Low Risk Clients", "0"),
        ("high_ratio", "High Risk Ratio", "0.0%"),
        ("avg_loan", "Avg Loan / Client", "P0"),
        ("avg_net", "Avg Net Income", "P0"),
        ("top_sector", "Top Sector", "—"),
    ]
    self._dash_detail_lbls = {}
    for idx, (key, label, value) in enumerate(detail_cards):
        c = tk.Frame(details, bg=WHITE, highlightbackground=BORDER_LIGHT, highlightthickness=1)
        c.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 8, 0))
        details.grid_columnconfigure(idx, weight=1)
        _add_container_header_bar(c, height=2)
        tk.Label(c, text=label, font=F(8), fg=TXT_SOFT, bg=WHITE).pack(anchor="w", padx=10, pady=(9, 2))
        v = tk.Label(c, text=value, font=F(11, "bold"), fg=TXT_NAVY, bg=WHITE)
        v.pack(anchor="w", padx=10, pady=(0, 9))
        self._dash_detail_lbls[key] = v

    charts = tk.Frame(body, bg="#F4F6FA")
    charts.pack(fill="both", expand=True, padx=18, pady=(0, 12))
    charts.grid_columnconfigure(0, weight=2)
    charts.grid_columnconfigure(1, weight=1)

    left_card = tk.Frame(charts, bg=WHITE, highlightbackground=BORDER_LIGHT, highlightthickness=1)
    left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    _add_container_header_bar(left_card)
    top_left = tk.Frame(left_card, bg=WHITE)
    top_left.pack(fill="x", padx=14, pady=(12, 0))
    tk.Label(top_left, text="At-Risk Industries", font=F(11, "bold"), fg=NAVY_MID, bg=WHITE).pack(side="left")
    self._dash_left_note_lbl = tk.Label(top_left, text="Loan Exposure vs Avg Net Income", font=F(8), fg=TXT_SOFT, bg=WHITE)
    self._dash_left_note_lbl.pack(side="right")

    lchart = tk.Canvas(left_card, bg=WHITE, highlightthickness=0, height=250)
    lchart.pack(fill="both", expand=True, padx=12, pady=(6, 10))

    def _draw_line_chart(_e=None, cv=lchart):
        cv.delete("all")
        w, h = max(cv.winfo_width(), 360), max(cv.winfo_height(), 180)
        pad_l, pad_r, pad_t, pad_b = 30, 18, 16, 30
        plot_w = w - pad_l - pad_r
        plot_h = h - pad_t - pad_b
        bottom_y = pad_t + plot_h
        for i in range(6):
            y = pad_t + (plot_h * i / 5)
            cv.create_line(pad_l, y, w - pad_r, y, fill="#EEF2F8")
        series_a = list(getattr(self, "_dash_series_loan", []))
        series_b = list(getattr(self, "_dash_series_net", []))
        n = max(2, len(series_a), len(series_b))
        if len(series_a) < n:
            series_a.extend([0.0] * (n - len(series_a)))
        if len(series_b) < n:
            series_b.extend([0.0] * (n - len(series_b)))
        xs = [pad_l + i * (plot_w / max(1, n - 1)) for i in range(n)]
        vmax = max(1.0, max(series_a + series_b))
        def _to_y(v):
            return pad_t + plot_h - (v / vmax) * plot_h
        pts_a, pts_b = [], []
        for x, va, vb in zip(xs, series_a, series_b):
            pts_a.extend((x, _to_y(va)))
            pts_b.extend((x, _to_y(vb)))

        # Shaded fill under series_b (net income - light blue)
        fill_b = [pad_l, bottom_y] + pts_b + [xs[-1], bottom_y]
        cv.create_polygon(fill_b, fill="#D6F0F8", outline="", smooth=True)

        # Shaded fill under series_a (loan - darker blue)
        fill_a = [pad_l, bottom_y] + pts_a + [xs[-1], bottom_y]
        cv.create_polygon(fill_a, fill="#C8D0F5", outline="", smooth=True)

        cv.create_line(*pts_b, fill="#89D6E9", width=2, smooth=True)
        cv.create_line(*pts_a, fill="#4A63D9", width=2, smooth=True)
        for x, va in zip(xs, series_a):
            y = _to_y(va)
            cv.create_oval(x - 2, y - 2, x + 2, y + 2, fill="#4A63D9", outline="#4A63D9")
        for x, vb in zip(xs, series_b):
            y = _to_y(vb)
            cv.create_oval(x - 2, y - 2, x + 2, y + 2, fill="#89D6E9", outline="#89D6E9")

    lchart.bind("<Configure>", _draw_line_chart)
    self._dash_draw_line = _draw_line_chart
    self._dash_at_risk_names_lbl = tk.Label(
        left_card, text="At-risk industries: —",
        font=F(8), fg=TXT_SOFT, bg=WHITE, justify="left", wraplength=700
    )
    self._dash_at_risk_names_lbl.pack(fill="x", padx=12, pady=(0, 10))

    right_card = tk.Frame(charts, bg=WHITE, highlightbackground=BORDER_LIGHT, highlightthickness=1)
    right_card.grid(row=0, column=1, sticky="nsew")
    _add_container_header_bar(right_card)
    top_right = tk.Frame(right_card, bg=WHITE)
    top_right.pack(fill="x", padx=14, pady=(12, 0))
    tk.Label(top_right, text="Risk by Industry", font=F(11, "bold"), fg=TXT_NAVY, bg=WHITE).pack(side="left")
    self._dash_right_note_lbl = tk.Label(top_right, text="High vs Low Clients", font=F(8), fg=TXT_SOFT, bg=WHITE)
    self._dash_right_note_lbl.pack(side="right")

    rchart = tk.Canvas(right_card, bg=WHITE, highlightthickness=0, height=250)
    rchart.pack(fill="both", expand=True, padx=12, pady=(6, 10))

    def _draw_bar_chart(_e=None, cv=rchart):
        cv.delete("all")
        w, h = max(cv.winfo_width(), 260), max(cv.winfo_height(), 200)
        pad_l, pad_r, pad_t, pad_b = 26, 16, 12, 12
        plot_w = w - pad_l - pad_r
        plot_h = h - pad_t - pad_b
        lows = list(getattr(self, "_dash_industry_low", []))
        highs = list(getattr(self, "_dash_industry_high", []))
        names = list(getattr(self, "_dash_industry_names", []))
        n = max(1, len(lows), len(highs))
        if len(lows) < n:
            lows.extend([0] * (n - len(lows)))
        if len(highs) < n:
            highs.extend([0] * (n - len(highs)))
        if len(names) < n:
            names.extend(["—"] * (n - len(names)))
        total_max = max(1, max((l + h) for l, h in zip(lows, highs)))
        row_h = max(18, plot_h / max(1, n))
        bar_h = max(8, row_h * 0.58)
        y = pad_t + (row_h - bar_h) / 2
        for idx, (lo, hi, _nm) in enumerate(zip(lows, highs, names), 1):
            cv.create_text(8, y + (bar_h / 2), text=str(idx), fill=TXT_SOFT, anchor="w", font=("Segoe UI", 7, "bold"))
            lw = (lo / total_max) * plot_w
            hw = (hi / total_max) * plot_w
            x0 = pad_l
            base_color = _INDUSTRY_COLORS[idx % len(_INDUSTRY_COLORS)]
            cv.create_rectangle(x0, y, x0 + lw, y + bar_h, fill=base_color, outline="")
            cv.create_rectangle(x0 + lw, y, x0 + lw + hw, y + bar_h, fill="#A7DCE8", outline="")
            # subtle baseline
            cv.create_line(pad_l, y + bar_h + 1, pad_l + plot_w, y + bar_h + 1, fill="#EDF2F8")
            y += row_h

    rchart.bind("<Configure>", _draw_bar_chart)
    self._dash_draw_bar = _draw_bar_chart
    self._dash_right_legend = tk.Frame(right_card, bg=WHITE)
    self._dash_right_legend.pack(fill="x", padx=12, pady=(0, 10))

    # High-risk details table
    table_wrap = tk.Frame(body, bg=WHITE, highlightbackground=BORDER_LIGHT, highlightthickness=1)
    table_wrap.pack(fill="both", expand=False, padx=18, pady=(0, 16))
    _add_container_header_bar(table_wrap)
    header = tk.Frame(table_wrap, bg=WHITE)
    header.pack(fill="x", padx=12, pady=(10, 6))
    tk.Label(header, text="Highest-Risk People", font=F(11, "bold"), fg=TXT_NAVY, bg=WHITE).pack(side="left")
    self._dash_table_count_lbl = tk.Label(header, text="", font=F(8), fg=TXT_SOFT, bg=WHITE)
    self._dash_table_count_lbl.pack(side="right")
    tk.Label(
        table_wrap,
        text="Basis: only HIGH-risk clients, ranked by lowest net income.",
        font=F(8), fg=TXT_SOFT, bg=WHITE
    ).pack(anchor="w", padx=12, pady=(0, 6))

    cols = tk.Frame(table_wrap, bg=NAVY_DEEP, height=30)
    cols.pack(fill="x", padx=10)
    cols.pack_propagate(False)
    self._dash_table_specs = [
        ("Client", "w"),
        ("Industry", "w"),
        ("Risk", "center"),
        ("Loan Balance", "center"),
        ("Net Income", "center"),
    ]
    self._dash_table_col_weights = (32, 34, 10, 12, 12)
    for i, w in enumerate(self._dash_table_col_weights):
        cols.grid_columnconfigure(i * 2, weight=w, uniform="dash_tbl")
        if i < len(self._dash_table_col_weights) - 1:
            cols.grid_columnconfigure(i * 2 + 1, weight=0)

    for idx, (text, anchor) in enumerate(self._dash_table_specs):
        tk.Label(
            cols, text=text, anchor=anchor, font=F(8, "bold"),
            fg=WHITE, bg=NAVY_DEEP, padx=8
        ).grid(row=0, column=idx * 2, sticky="nsew", pady=6)
        if idx < len(self._dash_table_specs) - 1:
            tk.Frame(cols, bg=NAVY_LIGHT, width=1).grid(row=0, column=idx * 2 + 1, sticky="ns", pady=5)

    self._dash_table_rows = tk.Frame(table_wrap, bg=WHITE)
    self._dash_table_rows.pack(fill="x", padx=10, pady=(0, 10))
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


def _refresh_dashboard_from_lu(self):
    all_data = getattr(self, "_lu_all_data", None) or {}
    general = list(all_data.get("general") or [])
    totals = dict(all_data.get("totals") or {})
    cache = getattr(self, "_dash_cache", None) or {}

    if not general and cache:
        clients = int(cache.get("clients") or 0)
        high = int(cache.get("high") or 0)
        industries = int(cache.get("industries") or 0)
        total_loan = float(cache.get("total_loan") or 0.0)
        total_net = float(cache.get("total_net") or 0.0)
        has_data = clients > 0
        self._dash_series_loan = list(cache.get("series_loan") or [0.0, 0.0])
        self._dash_series_net = list(cache.get("series_net") or [0.0, 0.0])
        self._dash_industry_low = list(cache.get("industry_low") or [0])
        self._dash_industry_high = list(cache.get("industry_high") or [0])
        self._dash_industry_names = list(cache.get("industry_names") or ["—"])
        at_risk_names = list(cache.get("at_risk_names") or [])
        high_rows = list(cache.get("high_rows") or [])
        top_sector = cache.get("top_sector") or "—"
    else:
        clients = len(general)
        high = sum(1 for r in general if str(r.get("score_label") or "").upper() == "HIGH")
        industries = len(all_data.get("unique_industries") or [])
        total_loan = float(totals.get("loan_balance") or 0.0)
        total_net = float(totals.get("total_net") or 0.0)
        has_data = clients > 0
        industry_risk_counts = {}
        industry_finance = {}
        for rec in general:
            industry = str(rec.get("industry") or "Unknown Industry").strip() or "Unknown Industry"
            if industry not in industry_risk_counts:
                industry_risk_counts[industry] = {"high": 0, "low": 0}
            if industry not in industry_finance:
                industry_finance[industry] = {"loan": 0.0, "net": 0.0, "count": 0}
            if str(rec.get("score_label") or "").upper() == "HIGH":
                industry_risk_counts[industry]["high"] += 1
            else:
                industry_risk_counts[industry]["low"] += 1
            industry_finance[industry]["loan"] += float(rec.get("loan_balance") or 0.0)
            industry_finance[industry]["net"] += float(rec.get("net_income") or 0.0)
            industry_finance[industry]["count"] += 1

        top_industries = sorted(
            industry_risk_counts.items(),
            key=lambda kv: (kv[1]["high"], kv[1]["high"] + kv[1]["low"]),
            reverse=True
        )[:12]
        # Line chart metric (different from bar chart):
        #   - series_loan: total loan exposure by at-risk industry
        #   - series_net : average net income by at-risk industry
        at_risk_for_line = [(n, v) for (n, v) in top_industries if int(v.get("high", 0)) > 0]
        if not at_risk_for_line:
            at_risk_for_line = top_industries
        self._dash_series_loan = [industry_finance[n]["loan"] for n, _ in at_risk_for_line] or [0.0, 0.0]
        self._dash_series_net = [
            (industry_finance[n]["net"] / max(1, industry_finance[n]["count"]))
            for n, _ in at_risk_for_line
        ] or [0.0, 0.0]
        if top_industries:
            top_industry_name = top_industries[0][0]
            if len(top_industry_name) > 32:
                top_industry_name = top_industry_name[:31] + "…"
            self._dash_left_note_lbl.config(text=f"Top risk: {top_industry_name}")

        sector_counts = {}
        for rec in general:
            sector = str(rec.get("sector") or "Other")
            if sector not in sector_counts:
                sector_counts[sector] = {"high": 0, "low": 0}
            if str(rec.get("score_label") or "").upper() == "HIGH":
                sector_counts[sector]["high"] += 1
            else:
                sector_counts[sector]["low"] += 1
        top_items = sorted(sector_counts.items(), key=lambda kv: kv[1]["high"] + kv[1]["low"], reverse=True)[:7]
        self._dash_industry_low = [v["low"] for _, v in top_industries[:7]] or [0]
        self._dash_industry_high = [v["high"] for _, v in top_industries[:7]] or [0]
        self._dash_industry_names = [name for name, _ in top_industries[:7]] or ["—"]
        at_risk_names = [name for name, v in top_industries if int(v.get("high", 0)) > 0]
        top_sector = top_items[0][0] if top_items else "—"
        high_rows = sorted(
            [r for r in general if str(r.get("score_label") or "").upper() == "HIGH"],
            key=lambda r: (float(r.get("net_income") or 0.0), float(r.get("loan_balance") or 0.0))
        )[:8]

    if has_data:
        current_file = Path(getattr(self, "_lu_filepath", "")).name if getattr(self, "_lu_filepath", None) else ""
        self._dash_subtitle_lbl.config(text=(f"From LU file: {current_file}" if current_file else "From saved LU snapshot"))
    else:
        self._dash_subtitle_lbl.config(text="Waiting for LU Analysis data")

    self._dash_value_lbls["clients"].config(text=f"{clients:,}")
    self._dash_value_lbls["high_risk"].config(text=f"{high:,}")
    self._dash_value_lbls["industries"].config(text=f"{industries:,}")
    self._dash_value_lbls["loan_balance"].config(text=f"P{total_loan:,.0f}")

    high_ratio = (high / clients * 100.0) if clients else 0.0
    low = max(0, clients - high)
    net_ratio = (total_net / total_loan * 100.0) if total_loan else 0.0
    avg_loan = (total_loan / clients) if clients else 0.0
    avg_net = (total_net / clients) if clients else 0.0

    self._dash_delta_lbls["clients"].config(text=(f"{clients:,} loaded" if has_data else "No LU file loaded"),
                                            fg=(ACCENT_SUCCESS if has_data else TXT_SOFT))
    self._dash_delta_lbls["high_risk"].config(text=(f"{high_ratio:.1f}% of clients" if has_data else "No LU file loaded"),
                                              fg=(ACCENT_RED if high > 0 else ACCENT_SUCCESS if has_data else TXT_SOFT))
    self._dash_delta_lbls["industries"].config(text=(f"{industries:,} active tags" if has_data else "No LU file loaded"),
                                               fg=(ACCENT_SUCCESS if has_data else TXT_SOFT))
    self._dash_delta_lbls["loan_balance"].config(text=(f"Net income {net_ratio:.1f}%" if has_data else "No LU file loaded"),
                                                 fg=(ACCENT_SUCCESS if has_data else TXT_SOFT))

    self._dash_detail_lbls["low_risk"].config(text=f"{low:,}")
    self._dash_detail_lbls["high_ratio"].config(text=f"{high_ratio:.1f}%")
    self._dash_detail_lbls["avg_loan"].config(text=f"P{avg_loan:,.0f}")
    self._dash_detail_lbls["avg_net"].config(text=f"P{avg_net:,.0f}")
    self._dash_detail_lbls["top_sector"].config(text=top_sector)
    if hasattr(self, "_dash_at_risk_names_lbl"):
        if at_risk_names:
            self._dash_at_risk_names_lbl.config(text="At-risk industries: " + ", ".join(at_risk_names))
        else:
            self._dash_at_risk_names_lbl.config(text="At-risk industries: none")
    if hasattr(self, "_dash_right_legend"):
        for w in self._dash_right_legend.winfo_children():
            w.destroy()
        if getattr(self, "_dash_industry_names", None):
            total_clients = max(1, sum(self._dash_industry_high) + sum(self._dash_industry_low))
            for idx, name in enumerate(self._dash_industry_names):
                hi = self._dash_industry_high[idx] if idx < len(self._dash_industry_high) else 0
                lo = self._dash_industry_low[idx] if idx < len(self._dash_industry_low) else 0
                pct = ((hi + lo) / total_clients) * 100.0
                item = tk.Frame(self._dash_right_legend, bg=WHITE)
                item.grid(row=idx // 2, column=idx % 2, sticky="w", padx=(0, 10), pady=2)
                sw = tk.Canvas(item, width=14, height=10, bg=WHITE, highlightthickness=0)
                sw.pack(side="left", padx=(0, 5))
                sw.create_rectangle(0, 2, 14, 8, fill=_INDUSTRY_COLORS[idx % len(_INDUSTRY_COLORS)], outline="")
                short = name if len(name) <= 34 else name[:33] + "…"
                tk.Label(item, text=f"{short}  {pct:.1f}%", font=F(8), fg=TXT_NAVY, bg=WHITE).pack(side="left")
        else:
            tk.Label(self._dash_right_legend, text="No industries to show", font=F(8), fg=TXT_SOFT, bg=WHITE).pack(anchor="w")

    if hasattr(self, "_dash_draw_line"):
        self._dash_draw_line()
    if hasattr(self, "_dash_draw_bar"):
        self._dash_draw_bar()

    _render_high_risk_table(self, high_rows)


def _render_high_risk_table(self, rows):
    for w in self._dash_table_rows.winfo_children():
        w.destroy()
    self._dash_table_count_lbl.config(text=f"{len(rows)} shown")
    if not rows:
        tk.Label(self._dash_table_rows, text="No high-risk clients to display.",
                 font=F(9), fg=TXT_SOFT, bg=WHITE).grid(row=0, column=0, columnspan=9, sticky="w", padx=6, pady=8)
        return
    for r, rec in enumerate(rows):
        bg = WHITE if r % 2 == 0 else "#F8FAFD"
        vals = [
            (str(rec.get("client") or "—"), "w"),
            (str(rec.get("industry") or "—"), "w"),
            ("HIGH", "center"),
            (f"P{float(rec.get('loan_balance') or 0.0):,.0f}", "center"),
            (f"P{float(rec.get('net_income') or 0.0):,.0f}", "center"),
        ]
        for c, (text, anchor) in enumerate(vals):
            fg = ACCENT_RED if c == 2 else TXT_NAVY
            tk.Label(
                self._dash_table_rows, text=text, anchor=anchor, font=F(8),
                fg=fg, bg=bg, padx=8, pady=4
            ).grid(row=r * 2, column=c * 2, sticky="nsew")
            if c < len(vals) - 1:
                tk.Frame(self._dash_table_rows, bg=BORDER_LIGHT, width=1).grid(row=r * 2, column=c * 2 + 1, sticky="ns")
        tk.Frame(self._dash_table_rows, bg=BORDER_LIGHT, height=1).grid(
            row=r * 2 + 1, column=0, columnspan=9, sticky="ew"
        )


def _persist_dashboard_snapshot_from_lu(self):
    all_data = getattr(self, "_lu_all_data", None) or {}
    general = list(all_data.get("general") or [])
    if not general:
        return
    totals = dict(all_data.get("totals") or {})
    industry_risk_counts = {}
    for rec in general:
        industry = str(rec.get("industry") or "Unknown Industry").strip() or "Unknown Industry"
        if industry not in industry_risk_counts:
            industry_risk_counts[industry] = {"high": 0, "low": 0}
        if str(rec.get("score_label") or "").upper() == "HIGH":
            industry_risk_counts[industry]["high"] += 1
        else:
            industry_risk_counts[industry]["low"] += 1
    top_industries = sorted(
        industry_risk_counts.items(),
        key=lambda kv: (kv[1]["high"], kv[1]["high"] + kv[1]["low"]),
        reverse=True
    )[:12]
    sector_counts = {}
    for rec in general:
        sector = str(rec.get("sector") or "Other")
        if sector not in sector_counts:
            sector_counts[sector] = {"high": 0, "low": 0}
        if str(rec.get("score_label") or "").upper() == "HIGH":
            sector_counts[sector]["high"] += 1
        else:
            sector_counts[sector]["low"] += 1
    top_sectors = sorted(sector_counts.items(), key=lambda kv: kv[1]["high"] + kv[1]["low"], reverse=True)[:7]
    industry_risk_counts = {}
    for rec in general:
        industry = str(rec.get("industry") or "Unknown Industry").strip() or "Unknown Industry"
        if industry not in industry_risk_counts:
            industry_risk_counts[industry] = {"high": 0, "low": 0}
        if str(rec.get("score_label") or "").upper() == "HIGH":
            industry_risk_counts[industry]["high"] += 1
        else:
            industry_risk_counts[industry]["low"] += 1
    top_industries = sorted(
        industry_risk_counts.items(),
        key=lambda kv: (kv[1]["high"], kv[1]["high"] + kv[1]["low"]),
        reverse=True
    )[:12]
    industry_finance = {}
    for rec in general:
        industry = str(rec.get("industry") or "Unknown Industry").strip() or "Unknown Industry"
        if industry not in industry_finance:
            industry_finance[industry] = {"loan": 0.0, "net": 0.0, "count": 0}
        industry_finance[industry]["loan"] += float(rec.get("loan_balance") or 0.0)
        industry_finance[industry]["net"] += float(rec.get("net_income") or 0.0)
        industry_finance[industry]["count"] += 1
    at_risk_for_line = [(n, v) for (n, v) in top_industries if int(v.get("high", 0)) > 0]
    if not at_risk_for_line:
        at_risk_for_line = top_industries
    at_risk_names = [name for name, v in top_industries if int(v.get("high", 0)) > 0]
    high_rows = sorted(
        [r for r in general if str(r.get("score_label") or "").upper() == "HIGH"],
        key=lambda r: (float(r.get("net_income") or 0.0), float(r.get("loan_balance") or 0.0))
    )[:8]
    payload = {
        "clients": len(general),
        "high": sum(1 for r in general if str(r.get("score_label") or "").upper() == "HIGH"),
        "industries": len(all_data.get("unique_industries") or []),
        "total_loan": float(totals.get("loan_balance") or 0.0),
        "total_net": float(totals.get("total_net") or 0.0),
        "series_loan": [industry_finance[n]["loan"] for n, _ in at_risk_for_line] or [0.0, 0.0],
        "series_net": [
            (industry_finance[n]["net"] / max(1, industry_finance[n]["count"]))
            for n, _ in at_risk_for_line
        ] or [0.0, 0.0],
        "industry_low": [v["low"] for _, v in top_industries[:7]] or [0],
        "industry_high": [v["high"] for _, v in top_industries[:7]] or [0],
        "industry_names": [name for name, _ in top_industries[:7]] or ["—"],
        "at_risk_names": at_risk_names,
        "top_sector": top_sectors[0][0] if top_sectors else "—",
        "high_rows": [
            {
                "client": str(r.get("client") or ""),
                "industry": str(r.get("industry") or ""),
                "loan_balance": float(r.get("loan_balance") or 0.0),
                "net_income": float(r.get("net_income") or 0.0),
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
    cls._build_dashboard_panel = _build_dashboard_panel
    cls._refresh_dashboard_from_lu = _refresh_dashboard_from_lu
    cls._persist_dashboard_snapshot_from_lu = _persist_dashboard_snapshot_from_lu

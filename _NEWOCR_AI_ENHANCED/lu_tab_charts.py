"""
lu_tab_charts.py — Charts Tab
================================
Renders matplotlib charts:
  1. Client risk mix — HIGH / MEDIUM / LOW / other (donut, % of clients in view)
  2. Client count per sector (bar)
  3. Risk distribution per sector (stacked horizontal bar)
  4. Total loan balance per sector (bar)

Also shows a detailed expense breakdown bar chart for per‑client view.

Standalone: imports only lu_core and lu_shared.
Attached to app class via attach(cls).

Public surface
--------------
  attach(cls)
  _build_charts_panel(self, parent)
  _charts_show_placeholder(self)
  _charts_render(self)
"""

import tkinter as tk
import re
import textwrap

from lu_shared import (
    F,
    _NAVY_MID, _NAVY_LIGHT, _WHITE, _CARD_WHITE, _BORDER_MID, _TXT_MUTED,
    _LIME_MID, _ACCENT_RED, _ACCENT_GOLD, _ACCENT_SUCCESS,
    _MPL_BG, _MPL_NAVY, _MPL_HIGH, _MPL_MOD, _MPL_LOW,
    _RISK_COLOR,
    _make_scrollable,
    _lu_filter_data_by_query,
    _lu_get_active_sectors, _lu_get_filtered_all_data,
)

# ── Tuneable constant ──────────────────────────────────────────────
CHART_MAX_ITEMS = 30   # max bars in per‑client expense chart
CHART_MAX_INDUSTRIES = 18  # max x-axis categories before aggregating to "Others"

try:
    import matplotlib
    matplotlib.use("TkAgg")
    import matplotlib.pyplot as plt
    import matplotlib.ticker
    import matplotlib.patches as mpatches
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    _HAS_MPL = True
except ImportError:
    _HAS_MPL = False


_INDUSTRY_SPLIT_RE = re.compile(r"\s*(?:,|/|;|&|\band\b)\s*", re.I)


def _extract_industry_tags(raw_industry: str) -> list[str]:
    """
    Normalize a single industry cell into unique tags.
    Example: "Construction, Transportation" -> ["Construction", "Transportation"].
    """
    raw = (raw_industry or "").strip()
    if not raw:
        return []
    parts = [p.strip() for p in _INDUSTRY_SPLIT_RE.split(raw) if p and p.strip()]
    tags = []
    seen = set()
    for p in parts:
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        tags.append(p)
    return tags


def _build_industry_map(records: list[dict]) -> dict[str, list[dict]]:
    """
    Build industry->records map using split tags from Industry Name.
    A client with combined industries contributes to each component tag.
    """
    industry_map: dict[str, list[dict]] = {}
    for rec in records:
        # Prefer normalized tags from lu_core so this stays in sync with
        # Risk Settings' unique_industries list.
        tags = rec.get("industry_tags") or _extract_industry_tags(rec.get("industry", ""))
        if not tags:
            continue
        for tag in tags:
            industry_map.setdefault(tag, []).append(rec)
    return industry_map


# ══════════════════════════════════════════════════════════════════════
#  PANEL BUILDER
# ══════════════════════════════════════════════════════════════════════

def _build_charts_panel(self, parent):
    hdr = tk.Frame(parent, bg=_NAVY_MID, height=46)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)
    self._charts_hdr_lbl = tk.Label(
        hdr,
        text="📊  Industry Charts — Based on Industry Name column",
        font=F(10, "bold"), fg=_WHITE, bg=_NAVY_MID)
    self._charts_hdr_lbl.pack(side="left", padx=20, pady=8)
    tk.Label(hdr, text="🔎", font=F(9), fg=_WHITE, bg=_NAVY_MID).pack(side="left", padx=(12, 4))
    self._charts_search_var = tk.StringVar()
    search = tk.Entry(
        hdr, textvariable=self._charts_search_var,
        font=F(8), relief="flat", bg=_WHITE, fg="#1A2B4A",
        insertbackground="#1A2B4A", highlightbackground=_NAVY_LIGHT, highlightthickness=1)
    search.pack(side="left", padx=(0, 8), ipady=3)
    self._charts_match_lbl = tk.Label(
        hdr, text="", font=F(8, "bold"), fg=_WHITE, bg=_NAVY_MID, padx=8, pady=3)
    self._charts_match_lbl.pack(side="left", padx=(0, 8), pady=8)
    search.insert(0, "")
    self._charts_search_var.trace_add(
        "write", lambda *_: _charts_render(self) if getattr(self, "_lu_all_data", None) else None)
    self._charts_body = tk.Frame(parent, bg=_CARD_WHITE)
    self._charts_body.pack(fill="both", expand=True)
    _charts_show_placeholder(self)


def _charts_show_placeholder(self):
    for w in self._charts_body.winfo_children():
        w.destroy()
    tk.Label(self._charts_body,
             text="Run an analysis first to view industry charts.",
             font=F(10), fg=_TXT_MUTED, bg=_CARD_WHITE).pack(pady=60)


# ══════════════════════════════════════════════════════════════════════
#  RENDERER (with figure cleanup and item cap)
# ══════════════════════════════════════════════════════════════════════

def _charts_render(self):
    # Close all figures before starting (prevents memory leak)
    plt.close("all")

    for w in self._charts_body.winfo_children():
        w.destroy()

    if not _HAS_MPL:
        tk.Label(self._charts_body,
                 text="matplotlib is not installed.\nRun:  pip install matplotlib",
                 font=F(10), fg=_ACCENT_RED, bg=_CARD_WHITE).pack(pady=40)
        return

    all_data   = _lu_get_filtered_all_data(self)
    q          = getattr(self, "_charts_search_var", tk.StringVar(value="")).get().strip()
    all_data   = _lu_filter_data_by_query(all_data, q)
    general    = all_data.get("general", [])
    match_lbl = getattr(self, "_charts_match_lbl", None)
    if match_lbl is not None:
        if q:
            client_names = sorted({(r.get("client") or "").strip() for r in general if r.get("client")})
            if len(client_names) == 1:
                match_lbl.config(text=client_names[0][:28], bg="#4A6FA5")
            else:
                match_lbl.config(text=f"{len(general)} CLIENTS MATCHED", bg="#4A6FA5")
        else:
            match_lbl.config(text="", bg=_NAVY_MID)

    if not general:
        _charts_show_placeholder(self)
        return

    active_sectors = _lu_get_active_sectors(self)
    if active_sectors:
        self._charts_hdr_lbl.config(
            text=f"📊  Charts — Filtered: {' · '.join(active_sectors)}",
            fg=_LIME_MID)
    else:
        self._charts_hdr_lbl.config(
            text="📊  Industry Charts — Based on Industry Name column",
            fg=_WHITE)
    if q:
        self._charts_hdr_lbl.config(
            text=f"📊  Charts — Search: {q[:30]}",
            fg=_LIME_MID)

    _, inner, _ = _make_scrollable(self._charts_body, _CARD_WHITE)
    pad = tk.Frame(inner, bg=_CARD_WHITE)
    pad.pack(fill="both", expand=True, padx=24, pady=16)

    industry_map = _build_industry_map(general)
    # Use lu_core output as source-of-truth to match Risk Settings list.
    all_unique_industries = all_data.get("unique_industries", [])
    industry_names = [ind for ind in all_unique_industries if ind in industry_map]
    if not industry_names:
        industry_names = sorted(
            industry_map.keys(),
            key=lambda ind: (-len(industry_map[ind]), ind.lower())
        )

    def _wrap_label(s: str, width: int = 14) -> str:
        s = str(s or "").strip()
        if not s:
            return "—"
        s = re.sub(r"\s+", " ", s).strip()
        return textwrap.fill(s, width=width)

    def _cap_and_aggregate(names: list[str], values: list[float], max_items: int):
        if len(names) <= max_items:
            return names, values
        pairs = list(zip(names, values))
        pairs.sort(key=lambda x: -float(x[1] or 0))
        keep = pairs[:max_items]
        rest = pairs[max_items:]
        other_sum = sum(float(v or 0) for _n, v in rest)
        out_n = [n for n, _v in keep] + (["Others"] if other_sum else [])
        out_v = [v for _n, v in keep] + ([other_sum] if other_sum else [])
        return out_n, out_v

    def _embed(fig, frame):
        FigureCanvasTkAgg(fig, master=frame).get_tk_widget().pack(
            fill="both", expand=True, padx=4, pady=4)

    tk.Label(
        pad,
        text=("Note: combined industries (e.g. 'Construction, Transportation') "
              "are split and counted under each individual industry."),
        font=F(8), fg=_TXT_MUTED, bg=_CARD_WHITE
    ).pack(anchor="w", pady=(0, 10))

    # ── Chart 0: Client risk mix (same filtered dataset as search bar) ─
    tk.Label(
        pad,
        text="Client risk mix (% of clients in this view)",
        font=F(11, "bold"),
        fg="#1A2B4A",
        bg=_CARD_WHITE,
    ).pack(anchor="w", pady=(0, 8))

    risk_counts = {"HIGH": 0, "MEDIUM": 0, "LOW": 0, "OTHER": 0}
    for r in general:
        lab = str(r.get("score_label") or "N/A").strip().upper()
        if lab == "HIGH":
            risk_counts["HIGH"] += 1
        elif lab == "LOW":
            risk_counts["LOW"] += 1
        elif lab in ("MODERATE", "MEDIUM"):
            risk_counts["MEDIUM"] += 1
        else:
            risk_counts["OTHER"] += 1
    n_clients_view = len(general)

    _risk_order = (
        ("HIGH", _ACCENT_RED),
        ("MEDIUM", _ACCENT_GOLD),
        ("LOW", _ACCENT_SUCCESS),
        ("OTHER", "#9AAACE"),
    )
    fig0 = None
    c0 = tk.Frame(pad, bg=_WHITE, highlightbackground=_BORDER_MID, highlightthickness=1)
    c0.pack(fill="x", pady=(0, 16))
    try:
        xs = []
        ys = []
        tick_labels = []
        point_colors = []
        for key, col in _risk_order:
            c = risk_counts[key]
            pct = 100.0 * c / n_clients_view if n_clients_view else 0.0
            xs.append(len(xs))
            ys.append(pct)
            tick_labels.append(key)
            point_colors.append(col)

        if any(v > 0 for v in ys):
            fig0, ax0 = plt.subplots(figsize=(10.5, 3.5))
            fig0.patch.set_facecolor(_MPL_BG)
            ax0.set_facecolor(_MPL_BG)

            # Bar chart per risk bucket
            bars = ax0.bar(
                xs,
                ys,
                color=point_colors,
                edgecolor=_MPL_BG,
                linewidth=1.5,
                width=0.58,
            )
            for bar, y, key in zip(bars, ys, tick_labels):
                c = risk_counts[key]
                ax0.text(
                    bar.get_x() + bar.get_width() / 2.0,
                    y + (2.0 if y < 96 else -4.0),
                    f"{c} ({y:.1f}%)",
                    ha="center",
                    va="bottom" if y < 96 else "top",
                    fontsize=9,
                    fontweight="bold",
                    color=_MPL_NAVY,
                )

            ax0.set_xticks(xs, tick_labels, fontsize=9)
            ax0.set_ylim(0, 100)
            ax0.set_ylabel("% of clients", fontsize=9, color=_MPL_NAVY)
            ax0.grid(axis="y", linestyle="--", alpha=0.25)
            ax0.spines[["top", "right"]].set_visible(False)

            ax0.set_title(
                "Risk mix — share of clients",
                fontsize=10,
                color="#1A2B4A",
                pad=10,
            )
            fig0.tight_layout(pad=1.1)
            _embed(fig0, c0)
        else:
            tk.Label(
                c0,
                text="No risk labels to chart.",
                font=F(9),
                fg=_TXT_MUTED,
                bg=_WHITE,
            ).pack(pady=18)
    except Exception:
        parts = []
        for key, _col in _risk_order:
            c = risk_counts[key]
            if c > 0:
                pct = 100.0 * c / n_clients_view if n_clients_view else 0.0
                parts.append(f"{key} {c} ({pct:.1f}%)")
        tk.Label(
            c0,
            text="Chart unavailable — " + ("  ·  ".join(parts) if parts else "no data"),
            font=F(9),
            fg=_TXT_MUTED,
            bg=_WHITE,
            wraplength=640,
            justify="left",
        ).pack(padx=12, pady=18)
    finally:
        if fig0:
            plt.close(fig0)

    # ── Chart 1: Client count per industry ───────────────────────────
    tk.Label(pad, text="Client Distribution by Industry",
             font=F(11, "bold"), fg="#1A2B4A", bg=_CARD_WHITE).pack(anchor="w", pady=(0, 8))
    fig1 = None
    try:
        fig1, ax1 = plt.subplots(figsize=(10.5, 4.4))
        fig1.patch.set_facecolor(_MPL_BG)
        ax1.set_facecolor(_MPL_BG)
        ind_names  = industry_names
        ind_counts = [len(industry_map[ind]) for ind in ind_names]
        ind_names, ind_counts = _cap_and_aggregate(ind_names, ind_counts, CHART_MAX_INDUSTRIES)
        colors = [plt.cm.tab20(i % 20) for i in range(len(ind_names))]
        if ind_names:
            xlabels = [_wrap_label(n, width=16) for n in ind_names]
            bars = ax1.bar(xlabels, ind_counts, color=colors,
                           edgecolor=_MPL_BG, linewidth=1.5, width=0.55)
            ax1.set_ylabel("Number of Clients", fontsize=9, color=_MPL_NAVY)
            ax1.tick_params(axis="x", labelsize=7)
            plt.setp(ax1.get_xticklabels(), rotation=35, ha="right")
            ax1.tick_params(axis="y", labelsize=8)
            ax1.spines[["top", "right"]].set_visible(False)
            for bar, val in zip(bars, ind_counts):
                ax1.text(bar.get_x() + bar.get_width()/2,
                         bar.get_height() + 0.05, str(val),
                         ha="center", va="bottom",
                         fontsize=9, fontweight="bold", color=_MPL_NAVY)
        else:
            ax1.text(0.5, 0.5, "No industry data", ha="center", va="center",
                     transform=ax1.transAxes, color=_TXT_MUTED)
        fig1.subplots_adjust(bottom=0.32)
        fig1.tight_layout(pad=1.2)
        c1 = tk.Frame(pad, bg=_WHITE,
                      highlightbackground=_BORDER_MID, highlightthickness=1)
        c1.pack(fill="x", pady=(0, 16))
        _embed(fig1, c1)
    except Exception:
        pass
    finally:
        if fig1:
            plt.close(fig1)

    # ── Chart 2: Risk distribution by industry ───────────────────────
    tk.Label(pad, text="Risk Distribution by Industry",
             font=F(11, "bold"), fg="#1A2B4A", bg=_CARD_WHITE).pack(anchor="w", pady=(0, 8))
    fig2 = None
    try:
        fig2_h = max(3.2, len(industry_names) * 0.35 + 0.8)
        fig2, ax2 = plt.subplots(figsize=(9, fig2_h))
        fig2.patch.set_facecolor(_MPL_BG)
        ax2.set_facecolor(_MPL_BG)

        if industry_names:
            highs = []
            mods = []
            lows = []
            for ind in industry_names:
                recs = industry_map[ind]
                highs.append(sum(1 for r in recs if r.get("score_label") == "HIGH"))
                mods.append(sum(1 for r in recs if str(r.get("score_label") or "").strip().upper() in ("MEDIUM", "MODERATE")))
                lows.append(sum(1 for r in recs if str(r.get("score_label") or "").strip().upper() not in ("HIGH", "MEDIUM", "MODERATE")))
            y = list(range(len(industry_names)))
            ax2.barh(y, highs, color=_MPL_HIGH, label="HIGH")
            ax2.barh(y, mods, left=highs, color=_MPL_MOD, label="MEDIUM")
            ax2.barh(y, lows, left=[h + m for h, m in zip(highs, mods)], color=_MPL_LOW, label="LOW")
            ax2.set_yticks(y, industry_names, fontsize=8)
            ax2.tick_params(axis="x", labelsize=8)
            ax2.spines[["top", "right"]].set_visible(False)
            ax2.legend(fontsize=8, frameon=False, loc="lower right")
            ax2.set_xlabel("Client Count", fontsize=9, color=_MPL_NAVY)
            ax2.invert_yaxis()
        else:
            ax2.text(0.5, 0.5, "No data", ha="center", va="center",
                     transform=ax2.transAxes, color=_TXT_MUTED, fontsize=8)

        fig2.tight_layout(pad=1.2)
        c2 = tk.Frame(pad, bg=_WHITE,
                      highlightbackground=_BORDER_MID, highlightthickness=1)
        c2.pack(fill="x", pady=(0, 16))
        _embed(fig2, c2)
    except Exception:
        pass
    finally:
        if fig2:
            plt.close(fig2)

    # ── Chart 3: Total Loan Balance by industry ───────────────────────
    tk.Label(pad, text="Total Loan Balance by Industry",
             font=F(11, "bold"), fg="#1A2B4A", bg=_CARD_WHITE).pack(anchor="w", pady=(0, 8))
    fig3 = None
    try:
        fig3, ax3 = plt.subplots(figsize=(10.5, 4.4))
        fig3.patch.set_facecolor(_MPL_BG)
        ax3.set_facecolor(_MPL_BG)
        ind_names = industry_names
        ind_vals  = [sum(r.get("loan_balance") or 0 for r in industry_map[ind]) for ind in ind_names]
        ind_names, ind_vals = _cap_and_aggregate(ind_names, ind_vals, CHART_MAX_INDUSTRIES)
        colors = [plt.cm.tab20(i % 20) for i in range(len(ind_names))]
        if ind_names and any(v > 0 for v in ind_vals):
            xlabels = [_wrap_label(n, width=16) for n in ind_names]
            bars = ax3.bar(xlabels, ind_vals, color=colors,
                           edgecolor=_MPL_BG, linewidth=1.5, width=0.55)
            ax3.yaxis.set_major_formatter(
                matplotlib.ticker.FuncFormatter(
                    lambda x, _: f"₱{x/1e6:.1f}M" if x >= 1e6 else f"₱{x:,.0f}"))
            ax3.tick_params(axis="x", labelsize=7)
            plt.setp(ax3.get_xticklabels(), rotation=35, ha="right")
            ax3.tick_params(axis="y", labelsize=8)
            ax3.spines[["top", "right"]].set_visible(False)
            max_v = max(ind_vals) if ind_vals else 1
            for bar, val in zip(bars, ind_vals):
                lbl = f"₱{val/1e6:.2f}M" if val >= 1e6 else f"₱{val:,.0f}"
                ax3.text(bar.get_x() + bar.get_width()/2,
                         val + max_v * 0.01, lbl,
                         ha="center", va="bottom", fontsize=7, color=_MPL_NAVY)
        else:
            ax3.text(0.5, 0.5, "No loan balance data", ha="center", va="center",
                     transform=ax3.transAxes, color=_TXT_MUTED)
        fig3.subplots_adjust(bottom=0.32)
        fig3.tight_layout(pad=1.2)
        c3 = tk.Frame(pad, bg=_WHITE,
                      highlightbackground=_BORDER_MID, highlightthickness=1)
        c3.pack(fill="x", pady=(0, 16))
        _embed(fig3, c3)
    except Exception:
        pass
    finally:
        if fig3:
            plt.close(fig3)

    # ── Per‑client horizontal bar chart (if not in general view) ──────
    # Only show if a specific client is selected and we have expenses
    if hasattr(self, "_lu_active_client") and self._lu_active_client != "📊  General (All Clients)":
        # Get the selected client's data
        clients = all_data.get("clients", {})
        rec = clients.get(self._lu_active_client)
        if rec:
            expenses = [e for e in rec.get("expenses", []) if e["total"] > 0]
            if expenses:
                # Sort and cap
                expenses_sorted = sorted(expenses, key=lambda x: x["total"], reverse=True)
                expenses_capped = expenses_sorted[:CHART_MAX_ITEMS]

                tk.Label(pad, text=f"Expense Breakdown — {self._lu_active_client}",
                         font=F(11, "bold"), fg="#1A2B4A", bg=_CARD_WHITE).pack(anchor="w", pady=(16, 6))

                fig4 = None
                try:
                    fig4, ax4 = plt.subplots(figsize=(9, max(2.8, len(expenses_capped) * 0.45 + 0.8)))
                    fig4.patch.set_facecolor(_MPL_BG)
                    ax4.set_facecolor(_MPL_BG)
                    names = [e["name"][:28] + "…" if len(e["name"]) > 28 else e["name"] for e in expenses_capped]
                    values = [e["total"] for e in expenses_capped]
                    colors = [_RISK_COLOR.get(e["risk"], _MPL_LOW) for e in expenses_capped]

                    bars = ax4.barh(names, values, color=colors, edgecolor=_MPL_BG, linewidth=1.2, height=0.6)
                    ax4.invert_yaxis()
                    ax4.xaxis.set_major_formatter(
                        matplotlib.ticker.FuncFormatter(lambda x, _: f"₱{x:,.0f}"))
                    ax4.tick_params(axis="x", labelsize=8)
                    ax4.tick_params(axis="y", labelsize=9)
                    ax4.spines[["top", "right"]].set_visible(False)
                    max_val = max(values) if values else 1
                    for bar, val in zip(bars, values):
                        ax4.text(val + max_val * 0.01, bar.get_y() + bar.get_height() / 2,
                                 f"₱{val:,.2f}", va="center", fontsize=8, color=_MPL_NAVY)

                    legend_patches = [
                        mpatches.Patch(color=_MPL_HIGH, label="HIGH risk"),
                        mpatches.Patch(color=_MPL_MOD,  label="MEDIUM risk"),
                        mpatches.Patch(color=_MPL_LOW,  label="LOW risk"),
                    ]
                    ax4.legend(handles=legend_patches, fontsize=8, frameon=False, loc="lower right")
                    fig4.tight_layout(pad=1.2)

                    c4 = tk.Frame(pad, bg=_WHITE,
                                  highlightbackground=_BORDER_MID, highlightthickness=1)
                    c4.pack(fill="x", pady=(0, 16))
                    _embed(fig4, c4)
                except Exception:
                    pass
                finally:
                    if fig4:
                        plt.close(fig4)

    # Final safety close
    plt.close("all")


# ══════════════════════════════════════════════════════════════════════
#  ATTACH
# ══════════════════════════════════════════════════════════════════════

def attach(cls):
    """Attach Charts-tab methods to the app class."""
    cls._build_charts_panel      = _build_charts_panel
    cls._charts_show_placeholder = _charts_show_placeholder
    cls._charts_render           = _charts_render
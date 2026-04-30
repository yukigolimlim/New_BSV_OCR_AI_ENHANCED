"""
lu_shared.py — Shared UI constants, helpers, and filter utilities
==================================================================
Imported by every tab file.  No tab-specific logic lives here.

Provides
--------
  - Colour / font constants
  - F(), FF() font helpers
  - _bind_mousewheel(), _make_scrollable()
  - Filter-state helpers (_lu_get_active_sectors, etc.)
  - Sector icon / colour maps
  - _SECTOR_COLS / LU_CLIENT_TREE_SPEC (loan balance + analysis grids)
"""

import re
import tkinter as tk
import customtkinter as ctk

from lu_core import (
    GENERAL_CLIENT, _RISK_ORDER, _SCORE_BANDS,
    SECTOR_WHOLESALE, SECTOR_AGRICULTURE,
    SECTOR_TRANSPORT, SECTOR_REMITTANCE,
    SECTOR_CONSUMER, SECTOR_OTHER,
    _compute_risk_score, _fmt_value,
    split_product_name_tokens,
)

# ══════════════════════════════════════════════════════════════════════
#  COLOUR PALETTE
# ══════════════════════════════════════════════════════════════════════

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

_RISK_COLOR    = {"HIGH": _ACCENT_RED, "MODERATE": _ACCENT_GOLD, "MEDIUM": _ACCENT_GOLD, "LOW": _ACCENT_SUCCESS}
_RISK_BG       = {"HIGH": "#FFF5F5",   "MODERATE": "#FFFBF0",    "MEDIUM": "#FFFBF0",    "LOW": "#F0FBE8"}
_RISK_BADGE_BG = {"HIGH": "#FFE8E8",   "MODERATE": "#FFF3CD",    "MEDIUM": "#FFF3CD",    "LOW": "#DCEDC8"}

_CLIENT_HERO_BG = {
    "CRITICAL": _NAVY_DEEP, "HIGH": _NAVY_DEEP,
    "MODERATE": _NAVY_DEEP, "MEDIUM": _NAVY_DEEP, "LOW": _NAVY_DEEP, "N/A": _NAVY_DEEP,
}
_CLIENT_HERO_ACCENT = {
    "CRITICAL": "#FF4444", "HIGH": "#E53E3E",
    "MODERATE": "#D4A017", "MEDIUM": "#D4A017", "LOW": "#2E7D32", "N/A": "#4A6FA5",
}

_MPL_HIGH = "#E53E3E"
_MPL_MOD  = "#D4A017"
_MPL_LOW  = "#2E7D32"
_MPL_NAVY = "#1A3A6B"
_MPL_BG   = "#FAFBFD"

_SIM_BAR_BASE = "#4A6FA5"
_SIM_BAR_SIM  = "#E53E3E"

# ── Sector colours & icons ──────────────────────────────────────────
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
_ALL_SECTORS = _CHART_SECTORS + [SECTOR_OTHER]

# Loan-balance / industry breakdown when industry tags are missing (not in reference list).
LU_SECTOR_UNSPECIFIED_LABEL = "Unspecified/Not in Reference File"

# ── Loan-balance table column specs ────────────────────────────────
_SECTOR_COLS = [
    ("Sector",              2, 160, "w"),
    ("# Clients",           1,  70, "center"),
    ("Total Loan Balance",  2, 150, "center"),
    ("% of Total",          2, 110, "center"),
    ("Avg Loan per Client", 2, 140, "center"),
    ("Avg Net Income",      2, 130, "center"),
    ("Risk Profile",        2,  90, "center"),
]

# Excel-aligned columns (same order as typical Look-Up / portfolio export).
# (tree_column_id, heading, field_key, min_width_px, anchor, kind)
# kind: plain | text | asset_text | money | risk
# asset_text = multi-line inventory-style notes (shown one line in tree with · separators)
LU_CLIENT_TREE_SPEC = (
    ("client_id", "Client ID", "client_id", 78, "center", "plain"),
    ("pn", "PN", "pn", 68, "center", "plain"),
    ("client", "Applicant", "client", 148, "w", "text"),
    ("residence", "Residence Address", "residence", 128, "w", "text"),
    ("office", "Office Address", "office", 128, "w", "text"),
    ("industry", "Industry Name", "industry", 132, "w", "text"),
    ("spouse_info", "Spouse Info", "spouse_info", 100, "w", "text"),
    ("personal_assets", "Personal Assets", "personal_assets", 200, "w", "asset_text"),
    ("business_assets", "Business Assets", "business_assets", 200, "w", "asset_text"),
    ("business_inventory", "Business Inventory", "business_inventory", 200, "w", "asset_text"),
    ("source_income", "Source of Income", "source_income", 140, "w", "text"),
    ("total_source", "Total Source Of Income", "total_source", 118, "e", "money"),
    ("biz_exp_detail", "Business Expenses", "biz_exp_detail", 120, "w", "text"),
    ("total_biz_exp", "Total Business Expenses", "total_biz_exp", 120, "e", "money"),
    ("hhld_exp_detail", "Household / Personal Expenses", "hhld_exp_detail", 130, "w", "text"),
    ("total_hhld_exp", "Total Household / Personal Expenses", "total_hhld_exp", 130, "e", "money"),
    ("net_income", "Total Net Income", "net_income", 108, "e", "money"),
    ("amort_history", "Total Amortization History", "amort_history", 118, "e", "money"),
    ("current_amort", "Total Current Amortization", "current_amort", 118, "e", "money"),
    ("loan_balance", "Loan Balance", "loan_balance", 112, "e", "money"),
    ("total_amortized_cost", "Total Amortized Cost", "total_amortized_cost", 118, "e", "money"),
    ("principal_loan", "Principal Loan", "principal_loan", 112, "e", "money"),
    ("maturity", "Maturity", "maturity", 88, "center", "plain"),
    ("interest_rate", "Interest Rate", "interest_rate", 88, "center", "plain"),
    ("branch", "Branch", "branch", 88, "center", "plain"),
    ("loan_class", "Loan Class", "loan_class", 88, "center", "plain"),
    ("product_name", "Product Name", "product_name", 110, "w", "text"),
    ("loan_date", "Loan Date", "loan_date", 88, "center", "plain"),
    ("term_unit", "Term Unit", "term_unit", 72, "center", "plain"),
    ("term", "Term", "term", 56, "center", "plain"),
    ("security", "Security", "security", 88, "w", "text"),
    ("release_tag", "Release Tag", "release_tag", 88, "center", "plain"),
    ("loan_amount", "Loan Amount", "loan_amount", 104, "e", "money"),
    ("loan_status", "Loan Status", "loan_status", 88, "center", "plain"),
    ("ao_name", "AO Name", "ao_name", 100, "w", "text"),
    ("score_label", "Risk", "score_label", 80, "center", "risk"),
    ("risk_reasoning", "Risk Reasoning", "risk_reasoning", 320, "w", "text"),
)


def lu_format_lu_cell(rec: dict, field: str, kind: str, text_limit: int = 56) -> str:
    """Format one cell for LU analysis / loan-balance tree views."""
    raw = rec.get(field)
    if kind == "asset_text":
        s = str(raw or "").strip()
        if not s:
            return "—"
        s = s.replace("\r\n", "\n").replace("\r", "\n")
        s = re.sub(r"\n+", " · ", s)
        s = re.sub(r"[ \t]+", " ", s).strip()
        cap = 120
        if len(s) > cap:
            return s[: cap - 1] + "…"
        return s
    if kind == "money":
        if raw is None or raw == "":
            return "—"
        try:
            return f"₱{float(raw):,.2f}"
        except (TypeError, ValueError):
            return "—"
    if kind == "risk":
        rl = str(rec.get("score_label") or "").strip()
        if not rl:
            return "⚪ —"
        icon = {"HIGH": "🟠", "LOW": "🟢", "MODERATE": "🟡", "MEDIUM": "🟡", "N/A": "⚪"}.get(rl.upper(), "⚪")
        return f"{icon} {rl}"
    s = str(raw or "").strip()
    if not s:
        return "—"
    if kind == "text" and len(s) > text_limit:
        return s[: text_limit - 1] + "…"
    return s


def lu_client_row_tuple(rec: dict) -> tuple:
    return tuple(
        lu_format_lu_cell(rec, field, kind)
        for (_cid, _hdr, field, _w, _a, kind) in LU_CLIENT_TREE_SPEC
    )


# ══════════════════════════════════════════════════════════════════════
#  FONT HELPERS
# ══════════════════════════════════════════════════════════════════════

def F(size, weight="normal"):
    z = 1.0
    try:
        from app_constants import get_ui_zoom
        z = float(get_ui_zoom())
    except Exception:
        pass
    return ("Segoe UI", max(6, int(round(size * z))), weight)

def FF(size, weight="normal"):
    z = 1.0
    try:
        from app_constants import get_ui_zoom
        z = float(get_ui_zoom())
    except Exception:
        pass
    return ctk.CTkFont(family="Segoe UI", size=max(6, int(round(size * z))), weight=weight)


# ══════════════════════════════════════════════════════════════════════
#  MOUSEWHEEL + SCROLLABLE FRAME
# ══════════════════════════════════════════════════════════════════════

def _bind_mousewheel(canvas: tk.Canvas):
    def _on_enter(e):
        canvas.bind_all("<MouseWheel>",
                        lambda ev: canvas.yview_scroll(int(-1*(ev.delta/120)), "units"))
    def _on_leave(e):
        canvas.unbind_all("<MouseWheel>")
    canvas.bind("<Enter>", _on_enter)
    canvas.bind("<Leave>", _on_leave)


def _make_scrollable(parent, bg=None):
    """Return (outer_frame, inner_frame, canvas) with scrollbar."""
    bg    = bg or _CARD_WHITE
    outer = tk.Frame(parent, bg=bg)
    outer.pack(fill="both", expand=True)
    sb    = tk.Scrollbar(outer, relief="flat", troughcolor=_OFF_WHITE,
                         bg=_BORDER_LIGHT, width=8, bd=0)
    sb.pack(side="right", fill="y")
    def _safe_scroll_set(first, last):
        try:
            if sb.winfo_exists():
                sb.set(first, last)
        except tk.TclError:
            # Widget may be destroyed while async redraw still runs.
            return
    canvas = tk.Canvas(outer, bg=bg, highlightthickness=0, yscrollcommand=_safe_scroll_set)
    canvas.pack(side="left", fill="both", expand=True)
    sb.config(command=canvas.yview)
    inner = tk.Frame(canvas, bg=bg)
    win   = canvas.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))
    _bind_mousewheel(canvas)
    return outer, inner, canvas


# ══════════════════════════════════════════════════════════════════════
#  TABLE HELPERS  (used by loanbal tab)
# ══════════════════════════════════════════════════════════════════════

def _make_table_frame(parent, col_specs):
    tf = tk.Frame(parent, bg=_WHITE)
    tf.pack(fill="x", padx=0, pady=0)
    for ci, (_, weight, min_px, _anchor) in enumerate(col_specs):
        tf.columnconfigure(ci, weight=weight, minsize=min_px)
    return tf


def _table_header(table_frame, col_specs):
    hdr_bg   = _NAVY_MID
    bg_strip = tk.Frame(table_frame, bg=hdr_bg)
    bg_strip.grid(row=0, column=0, columnspan=len(col_specs), sticky="nsew")
    bg_strip.lower()
    for ci, (label, _w, _min, anchor) in enumerate(col_specs):
        tk.Label(table_frame, text=label,
                 font=F(8, "bold"), fg=_WHITE, bg=hdr_bg,
                 anchor=anchor, padx=10, pady=9
                 ).grid(row=0, column=ci, sticky="nsew", ipadx=0)


def _table_divider(table_frame, row_idx, n_cols, color=None):
    color = color or _BORDER_LIGHT
    tk.Frame(table_frame, bg=color, height=1
             ).grid(row=row_idx, column=0, columnspan=n_cols, sticky="ew")


# ══════════════════════════════════════════════════════════════════════
#  FILTER STATE HELPERS  (shared across all tabs)
# ══════════════════════════════════════════════════════════════════════

def _lu_get_active_sectors(self) -> list:
    """Return list of active sector-filter names, or None."""
    return getattr(self, "_lu_filtered_sectors", None)


def _lu_get_filtered_all_data(self) -> dict:
    """Return all_data filtered to active sectors (or full data if no filter)."""
    all_data       = self._lu_all_data
    active_sectors = _lu_get_active_sectors(self)
    if not active_sectors:
        return all_data

    # Backward-compatible filtering:
    # - older UI paths store industry names in _lu_filtered_sectors
    # - newer tab views filter by canonical sector names
    active_set = set(active_sectors)
    filtered_general = [
        r for r in all_data.get("general", [])
        if (r.get("sector") in active_set) or (r.get("industry") in active_set)
    ]
    filtered_sector_map = {s: v for s, v in all_data.get("sector_map", {}).items()
                           if s in active_sectors}
    filtered_totals = {
        "loan_balance": sum(r.get("loan_balance") or 0 for r in filtered_general),
        "total_net":    sum(r.get("net_income")   or 0 for r in filtered_general),
    }
    filtered_clients = {r["client"]: r for r in filtered_general}
    return {
        "general":    filtered_general,
        "sector_map": filtered_sector_map,
        "totals":     filtered_totals,
        "clients":    filtered_clients,
    }


def _lu_filter_data_by_query(all_data: dict, query: str) -> dict:
    """Return all_data narrowed by free-text query across core client fields."""
    if not all_data:
        return {
            "general": [], "sector_map": {}, "totals": {}, "clients": {},
            "unique_industries": [], "unique_product_names": [],
        }

    q = (query or "").strip().lower()
    if not q:
        return all_data

    def _match(rec: dict) -> bool:
        industry_tags = rec.get("industry_tags") or []
        haystack = [
            rec.get("client", ""),
            rec.get("industry", ""),
            rec.get("sector", ""),
            str(rec.get("client_id", "")),
            str(rec.get("pn", "")),
            str(rec.get("product_name", "")),
            " ".join(industry_tags),
        ]
        return any(q in str(v).lower() for v in haystack if v is not None)

    general = all_data.get("general", [])
    filtered_general = [r for r in general if _match(r)]
    filtered_clients = {r.get("client", ""): r for r in filtered_general if r.get("client")}

    filtered_sector_map = {}
    for rec in filtered_general:
        sec = rec.get("sector")
        if sec:
            filtered_sector_map.setdefault(sec, []).append(rec)

    filtered_totals = {
        "loan_balance": sum(r.get("loan_balance") or 0 for r in filtered_general),
        "total_net":    sum(r.get("net_income") or 0 for r in filtered_general),
    }

    unique_industries = sorted({
        (tag or "").strip()
        for rec in filtered_general
        for tag in (rec.get("industry_tags") or [rec.get("industry", "")])
        if (tag or "").strip()
    }, key=str.lower)

    _prod_first: dict[str, str] = {}
    for rec in filtered_general:
        pn = (rec.get("product_name") or "").strip()
        if not pn:
            continue
        for tok in split_product_name_tokens(pn):
            lk = tok.strip().lower()
            if lk and lk not in _prod_first:
                _prod_first[lk] = tok.strip()
    unique_product_names = sorted(_prod_first.values(), key=str.lower)

    return {
        "general": filtered_general,
        "clients": filtered_clients,
        "sector_map": filtered_sector_map,
        "totals": filtered_totals,
        "unique_industries": unique_industries,
        "unique_product_names": unique_product_names,
    }
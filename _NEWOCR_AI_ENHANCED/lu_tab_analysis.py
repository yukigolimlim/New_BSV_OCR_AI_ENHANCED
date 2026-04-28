"""
lu_tab_analysis.py — Analysis Tab
===================================
Handles: client search, general treeview, per-client risk display,
and the Industry Risk Settings dialog.

Industry‑based risk: risk = HIGH if industry is in the user‑defined
override table (_INDUSTRY_RISK_OVERRIDES), else LOW.
No expense breakdown is shown.

Standalone: imports lu_core and lu_shared; pulls lu_loanbal_export_patch only for Excel export.
Attached to app class via attach(cls).

Public surface
--------------
  attach(cls)
  _build_analysis_view(self, parent)
  _lu_render_results(self, results)
  _lu_render_general_view(self, results, parent)
  _lu_render_client_view(self, results)
  _lu_on_client_change(self, value)
  _lu_filter_by_search(self)
  _lu_populate_client_dropdown(self)
  _lu_update_filter_pill(self)
  _lu_clear_sector_filter(self)
  _lu_open_risk_settings(self)        ← NEW
  _lu_apply_risk_overrides(self)      ← NEW
"""

import tkinter as tk
import tkinter.ttk as ttk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
from datetime import datetime

import customtkinter as ctk

import lu_core
from lu_core import GENERAL_CLIENT
from lu_shared import (
    F, FF,
    LU_CLIENT_TREE_SPEC, lu_client_row_tuple, lu_format_lu_cell,
    _NAVY_DEEP, _NAVY_MID, _NAVY_LIGHT, _NAVY_MIST, _NAVY_GHOST, _NAVY_PALE,
    _WHITE, _CARD_WHITE, _OFF_WHITE, _BORDER_LIGHT, _BORDER_MID,
    _TXT_NAVY, _TXT_SOFT, _TXT_MUTED, _TXT_ON_LIME,
    _LIME_MID, _LIME_DARK, _LIME_PALE,
    _ACCENT_RED, _ACCENT_GOLD, _ACCENT_SUCCESS,
    _RISK_COLOR, _RISK_BG, _RISK_BADGE_BG,
    _CLIENT_HERO_BG, _CLIENT_HERO_ACCENT,
    _make_scrollable,
)

# Reserved Analysis search phrases → filter all HIGH or all LOW risk (substring-safe).
_LU_SEARCH_HIGH_RISK = frozenset({
    "high risk", "all high risk", "risk:high", "risk high", "all high",
    "#high", "!high", "clients high", "only high risk", "high risk only",
})
_LU_SEARCH_LOW_RISK = frozenset({
    "low risk", "all low risk", "risk:low", "risk low", "all low",
    "#low", "!low", "clients low", "only low risk", "low risk only",
})
_LU_SEARCH_MEDIUM_RISK = frozenset({
    "medium risk", "all medium risk", "risk:medium", "risk medium",
    "#medium", "!medium", "clients medium", "only medium risk", "medium risk only",
})


# ══════════════════════════════════════════════════════════════════════
#  INDUSTRY RISK OVERRIDES
#  Persists across re-renders.  Keys = exact industry strings from data.
#  Values = "HIGH" | "LOW"
# ══════════════════════════════════════════════════════════════════════

_INDUSTRY_RISK_OVERRIDES: dict = {}

# ── Monkey-patch lu_core._compute_risk_score ──────────────────────────
# The original function scores based on expenses. We wrap it so that a
# user-set override for the industry always wins.

_original_compute_risk_score = lu_core._compute_risk_score


def _patched_compute_risk_score(expenses: list, industry: str = "", product_name: str = ""):
    # Support both old tuple return and newer dict return.
    try:
        result = _original_compute_risk_score(expenses, industry, product_name)
    except TypeError:
        try:
            result = _original_compute_risk_score(expenses, industry)
        except TypeError:
            result = _original_compute_risk_score(expenses)

    is_tuple = isinstance(result, tuple)
    if is_tuple:
        score = float(result[0]) if len(result) > 0 else 0.0
        label = (result[1] if len(result) > 1 else "LOW")
    else:
        score = float(result.get("score", 0.0))
        label = str(result.get("score_label") or result.get("label") or "LOW")

    # Keep color/label in sync with core settings:
    # - canonical industry HIGH list from lu_core
    # - legacy local dialog overrides (if used)
    high_ind_set = {str(x).strip().lower() for x in lu_core.get_high_risk_industries()}
    medium_ind_set = {str(x).strip().lower() for x in lu_core.get_medium_risk_industries()}
    tags = lu_core._extract_industry_tags(industry) if hasattr(lu_core, "_extract_industry_tags") else [industry]
    industry_is_high = any(str(t or "").strip().lower() in high_ind_set for t in tags)
    industry_is_medium = (not industry_is_high) and any(str(t or "").strip().lower() in medium_ind_set for t in tags)

    override = _INDUSTRY_RISK_OVERRIDES.get(industry, "").upper()
    if override == "HIGH" or industry_is_high:
        label = "HIGH"
        if score < 0.6:
            score = 0.75
    elif override == "MEDIUM" or industry_is_medium:
        label = "MEDIUM"
        if score < 0.3:
            score = 0.45

    # Product Name risk (lu_core registry) wins over expense/industry when set.
    pr, _pr_tok = lu_core.lookup_product_risk_override(str(product_name or ""))
    if pr == "HIGH":
        label = "HIGH"
        if score < 0.6:
            score = 0.75
    elif pr == "MEDIUM":
        if label != "HIGH":
            label = "MEDIUM"
        if score < 0.3:
            score = 0.45
    else:
        # Expense risk overrides industry when product risk is not set.
        er = lu_core.get_expense_risk_overrides()
        has_low = False
        for e in expenses or []:
            nm = str((e or {}).get("name") or "").strip()
            if not nm:
                continue
            lvl = next(
                (er.get(k) for k in lu_core._expense_override_lookup_keys(nm) if er.get(k) in ("HIGH", "MEDIUM", "MODERATE")),
                None,
            )
            if lvl == "HIGH":
                label = "HIGH"
                if score < 0.6:
                    score = 0.75
                break
            if lvl in ("MEDIUM", "MODERATE"):
                if label != "HIGH":
                    label = "MEDIUM"
                if score < 0.3:
                    score = 0.45
            if lvl == "LOW":
                has_low = True
        else:
            if has_low:
                label = "LOW"
                if score > 0.4:
                    score = 0.20

    if is_tuple:
        fg = "#E53E3E" if label == "HIGH" else ("#D4A017" if label in ("MEDIUM", "MODERATE") else "#2E7D32")
        bg = "#FFF5F5" if label == "HIGH" else ("#FFFBF0" if label in ("MEDIUM", "MODERATE") else "#F0FBE8")
        return (score, label, fg, bg)

    result["label"] = label
    result["score_label"] = label
    result["score"] = score
    return result


lu_core._compute_risk_score = _patched_compute_risk_score


# ══════════════════════════════════════════════════════════════════════
#  RE-SCORE HELPER
# ══════════════════════════════════════════════════════════════════════

def _lu_apply_risk_overrides(self):
    """Re-score every client using current _INDUSTRY_RISK_OVERRIDES,
    then refresh the visible view."""
    all_data = getattr(self, "_lu_all_data", None)
    if not all_data:
        return

    # Sync local dialog overrides (_INDUSTRY_RISK_OVERRIDES) into lu_core so
    # that industries toggled HIGH here are treated as HIGH everywhere
    # (including on a fresh re-render from the raw data).
    local_high = {
        ind for ind, lvl in _INDUSTRY_RISK_OVERRIDES.items() if lvl == "HIGH"
    }
    # Replace with local HIGH set so switching industries to LOW truly clears them.
    lu_core.set_high_risk_industries(local_high)

    for rec in all_data.get("general", []):
        industry = rec.get("industry", "")
        expenses = rec.get("expenses", [])
        product_name = str(rec.get("product_name") or "")
        scored = lu_core._compute_risk_score(
            expenses, industry, product_name
        )
        if isinstance(scored, tuple):
            score_val = scored[0] if len(scored) > 0 else 0.0
            label_val = scored[1] if len(scored) > 1 else "LOW"
        else:
            score_val = scored.get("score", 0.0)
            label_val = scored.get("score_label", "LOW")
        rec["score"] = score_val
        rec["score_label"] = label_val
        rec["risk"] = label_val
        # Keep color fields in sync so the treeview tag picks up the right color.
        rec["score_fg"] = "#E53E3E" if label_val == "HIGH" else ("#D4A017" if label_val in ("MEDIUM", "MODERATE") else "#2E7D32")
        rec["score_bg"] = "#FFF5F5" if label_val == "HIGH" else ("#FFFBF0" if label_val in ("MEDIUM", "MODERATE") else "#F0FBE8")
        pr, pr_matched = lu_core.lookup_product_risk_override(product_name)
        high_exp = ""
        medium_exp = ""
        exp_overrides = lu_core.get_expense_risk_overrides()
        for e in expenses or []:
            nm = str((e or {}).get("name") or "").strip()
            if nm and any(
                exp_overrides.get(k) == "HIGH"
                for k in lu_core._expense_override_lookup_keys(nm)
            ):
                high_exp = nm
                break
            if nm and not medium_exp and any(
                str(exp_overrides.get(k) or "").upper() in ("MEDIUM", "MODERATE")
                for k in lu_core._expense_override_lookup_keys(nm)
            ):
                medium_exp = nm
        high_ind_set = {str(x).strip().lower() for x in lu_core.get_high_risk_industries()}
        medium_ind_set = {str(x).strip().lower() for x in lu_core.get_medium_risk_industries()}
        tags = rec.get("industry_tags") or [industry]
        high_ind = any(str(t or "").strip().lower() in high_ind_set for t in tags)
        medium_ind = (not high_ind) and any(str(t or "").strip().lower() in medium_ind_set for t in tags)
        rec["risk_reasoning"] = lu_core._compute_risk_reasoning(
            industry=industry,
            product_name=product_name,
            product_override=pr,
            expense_high_name=high_exp,
            expense_medium_name=medium_exp,
            is_high_industry=high_ind,
            is_medium_industry=medium_ind,
            product_matched_token=pr_matched,
        )
    all_data["clients"] = {r["client"]: r for r in all_data.get("general", [])}
    _lu_populate_client_dropdown(self)
    _lu_on_client_change(self, getattr(self, "_lu_analysis_active_client", GENERAL_CLIENT))


# ══════════════════════════════════════════════════════════════════════
#  RISK SETTINGS DIALOG
# ══════════════════════════════════════════════════════════════════════

class _RiskSettingsDialog(tk.Toplevel):
    """
    Modal dialog listing every unique industry from the loaded data.
    Each row has  [🟠 HIGH]  [🟢 LOW]  toggle buttons.
    Clicking Apply & Close writes to _INDUSTRY_RISK_OVERRIDES and
    calls _lu_apply_risk_overrides to refresh all client records live.
    """

    def __init__(self, master, app, industries: list):
        super().__init__(master)
        self._app        = app
        self._industries = sorted(industries, key=str.lower)
        self._vars: dict       = {}   # industry -> tk.StringVar
        self._row_frames: list = []   # (frame, industry)

        self.title("⚙  Industry Risk Settings")
        self.resizable(True, True)
        self.minsize(540, 420)
        self.grab_set()
        self.focus_force()
        self.configure(bg=_CARD_WHITE)
        self.update_idletasks()
        pw = master.winfo_rootx()
        py = master.winfo_rooty()
        self.geometry(f"640x580+{pw + 60}+{py + 40}")
        self._build()

    # ── BUILD ──────────────────────────────────────────────────────────

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg=_NAVY_MID)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚙  Industry Risk Settings",
                 font=F(12, "bold"), fg=_WHITE, bg=_NAVY_MID,
                 padx=20, pady=12).pack(side="left")

        # Explanation note
        note = tk.Frame(self, bg=_NAVY_MIST,
                        highlightbackground=_BORDER_MID, highlightthickness=1)
        note.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(note,
                 text=("Toggle each industry between  🟠 HIGH  and  🟢 LOW  risk.\n"
                       "Click  ✔ Apply & Close  to re-score all clients immediately."),
                 font=F(8), fg=_TXT_SOFT, bg=_NAVY_MIST,
                 padx=12, pady=8, justify="left").pack(anchor="w")

        # Search bar
        search_row = tk.Frame(self, bg=_CARD_WHITE)
        search_row.pack(fill="x", padx=16, pady=(8, 2))
        tk.Label(search_row, text="🔍", font=F(10), bg=_CARD_WHITE,
                 fg=_TXT_SOFT).pack(side="left")
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._filter_rows())
        tk.Entry(search_row, textvariable=self._search_var,
                 font=F(9), relief="flat", bg=_OFF_WHITE,
                 fg=_TXT_NAVY, insertbackground=_TXT_NAVY,
                 highlightbackground=_BORDER_MID, highlightthickness=1
                 ).pack(side="left", fill="x", expand=True, padx=(6, 0), ipady=5)

        # Column header
        col_hdr = tk.Frame(self, bg=_NAVY_MID)
        col_hdr.pack(fill="x", padx=16, pady=(8, 0))
        tk.Label(col_hdr, text="  Industry / Sector", font=F(8, "bold"),
                 fg=_WHITE, bg=_NAVY_MID, pady=6,
                 anchor="w", width=42).pack(side="left")
        tk.Label(col_hdr, text="Risk Level  ", font=F(8, "bold"),
                 fg=_WHITE, bg=_NAVY_MID, pady=6).pack(side="right")

        # Scrollable list
        list_wrap = tk.Frame(self, bg=_CARD_WHITE)
        list_wrap.pack(fill="both", expand=True, padx=16, pady=(0, 4))
        sb = tk.Scrollbar(list_wrap, relief="flat", bg=_BORDER_LIGHT,
                          troughcolor=_OFF_WHITE, width=8, bd=0)
        sb.pack(side="right", fill="y")
        self._canvas = tk.Canvas(list_wrap, bg=_CARD_WHITE,
                                 highlightthickness=0, yscrollcommand=sb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        sb.config(command=self._canvas.yview)
        self._list_frame = tk.Frame(self._canvas, bg=_CARD_WHITE)
        win = self._canvas.create_window((0, 0), window=self._list_frame, anchor="nw")
        self._list_frame.bind(
            "<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")))
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfig(win, width=e.width))
        self._canvas.bind("<Enter>", lambda e: self._canvas.bind_all(
            "<MouseWheel>",
            lambda ev: self._canvas.yview_scroll(int(-1*(ev.delta/120)), "units")))
        self._canvas.bind("<Leave>",
                          lambda e: self._canvas.unbind_all("<MouseWheel>"))

        self._populate_rows()

        # Footer
        footer = tk.Frame(self, bg=_OFF_WHITE,
                          highlightbackground=_BORDER_MID, highlightthickness=1)
        footer.pack(fill="x", padx=16, pady=(4, 14))
        tk.Button(footer, text="Set ALL → HIGH",
                  font=F(8, "bold"), fg=_ACCENT_RED, bg="#FFE8E8",
                  relief="flat", bd=0, padx=10, pady=6, cursor="hand2",
                  command=self._set_all_high).pack(side="left", padx=(12, 4), pady=8)
        tk.Button(footer, text="Set ALL → LOW",
                  font=F(8, "bold"), fg=_ACCENT_SUCCESS, bg="#DCEDC8",
                  relief="flat", bd=0, padx=10, pady=6, cursor="hand2",
                  command=self._set_all_low).pack(side="left", padx=4, pady=8)
        tk.Button(footer, text="Cancel",
                  font=F(9), fg=_TXT_SOFT, bg=_OFF_WHITE,
                  relief="flat", bd=0, padx=10, pady=8, cursor="hand2",
                  command=self.destroy).pack(side="right", padx=(0, 4), pady=8)
        tk.Button(footer, text="  ✔  Apply & Close  ",
                  font=F(9, "bold"), fg=_WHITE, bg=_NAVY_MID,
                  activebackground=_NAVY_LIGHT, activeforeground=_WHITE,
                  relief="flat", bd=0, padx=14, pady=8, cursor="hand2",
                  command=self._apply).pack(side="right", padx=12, pady=8)

    # ── ROWS ───────────────────────────────────────────────────────────

    def _populate_rows(self):
        self._row_frames.clear()
        for i, industry in enumerate(self._industries):
            current = _INDUSTRY_RISK_OVERRIDES.get(industry, "LOW")
            var = tk.StringVar(value=current)
            self._vars[industry] = var

            row_bg = _WHITE if i % 2 == 0 else _OFF_WHITE
            row = tk.Frame(self._list_frame, bg=row_bg)
            row.pack(fill="x")
            tk.Frame(row, bg=_BORDER_LIGHT, height=1).pack(fill="x")

            inner = tk.Frame(row, bg=row_bg)
            inner.pack(fill="x", padx=8, pady=4)

            tk.Label(inner, text=industry, font=F(9), fg=_TXT_NAVY,
                     bg=row_bg, anchor="w", wraplength=340,
                     justify="left").pack(side="left", padx=(4, 0),
                                         fill="x", expand=True)

            btn_frame = tk.Frame(inner, bg=row_bg)
            btn_frame.pack(side="right", padx=6)

            # Use a closure to capture per-row references properly
            def _make_toggle(v, bhi, blo):
                def _pick(choice):
                    v.set(choice)
                    if choice == "HIGH":
                        bhi.config(relief="sunken", bg="#FFB3B3",
                                   fg=_ACCENT_RED,     font=F(8, "bold"))
                        blo.config(relief="flat",   bg="#F0FBE8",
                                   fg=_TXT_MUTED,   font=F(8))
                    else:
                        blo.config(relief="sunken", bg="#B7E8A0",
                                   fg=_ACCENT_SUCCESS, font=F(8, "bold"))
                        bhi.config(relief="flat",   bg="#FFE8E8",
                                   fg=_TXT_MUTED,   font=F(8))
                return _pick

            hi_btn = tk.Button(btn_frame, text="🟠 HIGH",
                               font=F(8), fg=_TXT_MUTED, bg="#FFE8E8",
                               relief="flat", bd=1, padx=8, pady=3,
                               cursor="hand2", activebackground="#FFB3B3")
            lo_btn = tk.Button(btn_frame, text="🟢 LOW",
                               font=F(8), fg=_TXT_MUTED, bg="#F0FBE8",
                               relief="flat", bd=1, padx=8, pady=3,
                               cursor="hand2", activebackground="#B7E8A0")

            toggle = _make_toggle(var, hi_btn, lo_btn)
            hi_btn.config(command=lambda t=toggle: t("HIGH"))
            lo_btn.config(command=lambda t=toggle: t("LOW"))
            hi_btn.pack(side="left", padx=(0, 4))
            lo_btn.pack(side="left")
            toggle(current)   # set initial visual state

            self._row_frames.append((row, industry))

    def _filter_rows(self):
        query = self._search_var.get().strip().lower()
        for row_frame, industry in self._row_frames:
            if query in industry.lower():
                row_frame.pack(fill="x")
            else:
                row_frame.pack_forget()

    # ── BULK ───────────────────────────────────────────────────────────

    def _set_all_high(self):
        _INDUSTRY_RISK_OVERRIDES.update(
            {ind: "HIGH" for ind in self._industries})
        self._rebuild_rows()

    def _set_all_low(self):
        _INDUSTRY_RISK_OVERRIDES.update(
            {ind: "LOW" for ind in self._industries})
        self._rebuild_rows()

    def _rebuild_rows(self):
        """Sync button visuals after a bulk change."""
        for w in self._list_frame.winfo_children():
            w.destroy()
        self._vars.clear()
        self._row_frames.clear()
        self._populate_rows()

    # ── APPLY ──────────────────────────────────────────────────────────

    def _apply(self):
        global _INDUSTRY_RISK_OVERRIDES
        _INDUSTRY_RISK_OVERRIDES.clear()
        for industry, var in self._vars.items():
            if var.get() == "HIGH":
                _INDUSTRY_RISK_OVERRIDES[industry] = "HIGH"
        # Also push the HIGH-risk industries into lu_core so every tab
        # and every re-render uses the same authoritative set.
        new_high = [ind for ind, lvl in _INDUSTRY_RISK_OVERRIDES.items() if lvl == "HIGH"]
        lu_core.set_high_risk_industries(new_high)
        self.destroy()
        _lu_apply_risk_overrides(self._app)


# ══════════════════════════════════════════════════════════════════════
#  OPEN RISK SETTINGS  (called by toolbar button)
# ══════════════════════════════════════════════════════════════════════

def _lu_open_risk_settings(self):
    all_data = getattr(self, "_lu_all_data", None)
    if not all_data or not all_data.get("general"):
        from tkinter import messagebox
        messagebox.showinfo(
            "No Data",
            "Load an Excel file first before changing risk settings.",
            parent=self)
        return
    industries = sorted({
        r.get("industry", "")
        for r in all_data.get("general", [])
        if r.get("industry")
    })
    if not industries:
        from tkinter import messagebox
        messagebox.showinfo("No Industries",
                            "No industry data found in the loaded file.",
                            parent=self)
        return
    _RiskSettingsDialog(self, self, industries)


# ══════════════════════════════════════════════════════════════════════
#  FILTER PILL
# ══════════════════════════════════════════════════════════════════════

def _lu_update_filter_pill(self):
    existing = getattr(self, "_lu_filter_pill", None)
    if existing:
        try:
            existing.destroy()
        except Exception:
            pass
        self._lu_filter_pill = None

    sectors = getattr(self, "_lu_analysis_filtered_sectors", None)
    risk_f = getattr(self, "_lu_analysis_risk_filter", None)
    prod = getattr(self, "_lu_analysis_product_substr", None)
    if not sectors and not risk_f and not prod:
        return

    pill_frame = tk.Frame(self._lu_client_bar, bg="#1E4080",
                          highlightbackground=_LIME_MID, highlightthickness=1)
    pill_frame.pack(side="left", padx=(6, 0), pady=10)

    bits = []
    if risk_f:
        bits.append(f"RISK {risk_f}")
    if prod:
        short = prod if len(prod) <= 36 else prod[:35] + "…"
        bits.append(f"PRODUCT: {short}")
    if sectors:
        ind = " · ".join(sectors[:5])
        if len(sectors) > 5:
            ind += "…"
        bits.append(f"INDUSTRY: {ind}")

    tk.Label(
        pill_frame,
        text=f"  {'  |  '.join(bits)}  ",
        font=F(8, "bold"), fg=_LIME_MID, bg="#1E4080",
        padx=6, pady=3,
    ).pack(side="left")
    clear_btn = tk.Label(pill_frame, text=" ✕ ", font=F(8, "bold"),
                         fg=_ACCENT_RED, bg="#1E4080", cursor="hand2", padx=2)
    clear_btn.pack(side="left")
    clear_btn.bind("<Button-1>", lambda e: _lu_clear_sector_filter(self))
    self._lu_filter_pill = pill_frame


def _lu_clear_sector_filter(self):
    self._lu_analysis_filtered_sectors = None
    self._lu_analysis_risk_filter = None
    self._lu_analysis_product_substr = None
    self._lu_search_var.set("")
    _lu_update_filter_pill(self)
    _lu_on_client_change(self, GENERAL_CLIENT)


# ══════════════════════════════════════════════════════════════════════
#  CLIENT SELECTOR + SEARCH
# ══════════════════════════════════════════════════════════════════════

def _lu_on_client_change(self, value: str):
    self._lu_analysis_active_client = value
    is_general = (value == GENERAL_CLIENT)
    active_sectors = getattr(self, "_lu_analysis_filtered_sectors", None)
    risk_f = getattr(self, "_lu_analysis_risk_filter", None)
    prod_sub = getattr(self, "_lu_analysis_product_substr", None)
    all_clients = self._lu_all_data.get("clients", {})

    if is_general:
        results = list(all_clients.values())
        if active_sectors:
            results = [r for r in results if (r.get("industry") in active_sectors)]
        if risk_f:
            rfl = risk_f.upper()
            results = [r for r in results if str(r.get("score_label") or "").upper() == rfl]
        if prod_sub:
            pl = prod_sub.lower()
            results = [r for r in results if pl in str(r.get("product_name") or "").lower()]
    else:
        results = [all_clients[value]] if value in all_clients else []

    self._lu_results = results

    if is_general:
        has_filter = bool(active_sectors or risk_f or prod_sub)
        if has_filter:
            parts = []
            if risk_f == "HIGH":
                parts.append("HIGH RISK")
            elif risk_f == "MEDIUM":
                parts.append("MEDIUM RISK")
            elif risk_f == "LOW":
                parts.append("LOW RISK")
            if active_sectors:
                parts.append("INDUSTRY")
            if prod_sub:
                parts.append("PRODUCT")
            self._lu_mode_badge.config(
                text=f"  FILTER · {' · '.join(parts)}  ",
                bg="#1E4080", fg=_LIME_MID)
        else:
            self._lu_mode_badge.config(text="  GENERAL VIEW  ", bg=_NAVY_MID, fg=_WHITE)
    else:
        rec = all_clients.get(value, {})
        label = rec.get("score_label", "N/A")
        accent = _CLIENT_HERO_ACCENT.get(label, _NAVY_PALE)
        self._lu_mode_badge.config(text=f"  PER-CLIENT  ·  {value[:35]}  ",
                                   bg=accent, fg=_WHITE)

    _lu_render_results(self, results)


def _lu_filter_by_search(self):
    raw = self._lu_search_var.get().strip()
    query = raw.lower()
    all_clients = self._lu_all_data.get("clients", {})
    clients = list(all_clients.keys())

    if not query:
        self._lu_analysis_filtered_sectors = None
        self._lu_analysis_risk_filter = None
        self._lu_analysis_product_substr = None
        _lu_update_filter_pill(self)
        _lu_on_client_change(self, GENERAL_CLIENT)
        return

    if query in _LU_SEARCH_HIGH_RISK:
        self._lu_analysis_filtered_sectors = None
        self._lu_analysis_product_substr = None
        self._lu_analysis_risk_filter = "HIGH"
        _lu_update_filter_pill(self)
        _lu_on_client_change(self, GENERAL_CLIENT)
        return

    if query in _LU_SEARCH_LOW_RISK:
        self._lu_analysis_filtered_sectors = None
        self._lu_analysis_product_substr = None
        self._lu_analysis_risk_filter = "LOW"
        _lu_update_filter_pill(self)
        _lu_on_client_change(self, GENERAL_CLIENT)
        return

    if query in _LU_SEARCH_MEDIUM_RISK:
        self._lu_analysis_filtered_sectors = None
        self._lu_analysis_product_substr = None
        self._lu_analysis_risk_filter = "MEDIUM"
        _lu_update_filter_pill(self)
        _lu_on_client_change(self, GENERAL_CLIENT)
        return

    # Normal text: do not keep the all-HIGH / all-LOW tier filter active.
    self._lu_analysis_risk_filter = None

    unique_industries = self._lu_all_data.get("unique_industries", [])
    unique_products = self._lu_all_data.get("unique_product_names", [])
    matched_industries = [ind for ind in unique_industries if query in ind.lower()]
    matched_products = [p for p in unique_products if query in p.lower()]

    client_hit_id = any(
        query in c.lower()
        or query in str(all_clients[c].get("client_id", "")).lower()
        or query in str(all_clients[c].get("pn", "")).lower()
        for c in clients
    )

    if matched_industries and not client_hit_id:
        self._lu_analysis_product_substr = None
        self._lu_analysis_filtered_sectors = matched_industries
        _lu_update_filter_pill(self)
        _lu_on_client_change(self, GENERAL_CLIENT)
        return

    if matched_products and not client_hit_id:
        self._lu_analysis_filtered_sectors = None
        self._lu_analysis_product_substr = raw
        _lu_update_filter_pill(self)
        _lu_on_client_change(self, GENERAL_CLIENT)
        return

    matched_clients = [
        c for c in clients
        if query in c.lower()
        or query in str(all_clients[c].get("client_id", "")).lower()
        or query in str(all_clients[c].get("pn", "")).lower()
        or query in str(all_clients[c].get("product_name") or "").lower()
    ]

    if matched_industries and matched_clients:
        if any(ind.lower() == query for ind in matched_industries):
            self._lu_analysis_product_substr = None
            self._lu_analysis_filtered_sectors = matched_industries
            _lu_update_filter_pill(self)
            _lu_on_client_change(self, GENERAL_CLIENT)
            return

    if matched_clients:
        self._lu_analysis_filtered_sectors = None
        self._lu_analysis_product_substr = None
        _lu_update_filter_pill(self)
        if len(matched_clients) == 1:
            _lu_on_client_change(self, matched_clients[0])
        else:
            self._lu_analysis_active_client = GENERAL_CLIENT
            self._lu_mode_badge.config(
                text=f"  {len(matched_clients)} CLIENTS MATCHED  ", bg=_NAVY_PALE, fg=_WHITE)
            results = [all_clients[c] for c in matched_clients if c in all_clients]
            self._lu_results = results
            _lu_render_results(self, results)
    elif matched_industries:
        self._lu_analysis_product_substr = None
        self._lu_analysis_filtered_sectors = matched_industries
        _lu_update_filter_pill(self)
        _lu_on_client_change(self, GENERAL_CLIENT)
    else:
        self._lu_analysis_filtered_sectors = None
        self._lu_analysis_product_substr = None
        _lu_update_filter_pill(self)
        self._lu_mode_badge.config(text="  NO MATCH  ", bg=_ACCENT_RED, fg=_WHITE)
        _lu_render_results(self, [])


def _lu_populate_client_dropdown(self):
    clients = list(self._lu_all_data.get("clients", {}).keys())
    self._lu_client_var.set(GENERAL_CLIENT)
    self._lu_analysis_active_client = GENERAL_CLIENT
    self._lu_analysis_filtered_sectors = None
    self._lu_analysis_risk_filter = None
    self._lu_analysis_product_substr = None
    _lu_update_filter_pill(self)
    self._lu_mode_badge.config(text="  GENERAL VIEW  ", bg=_NAVY_MID, fg=_WHITE)
    suffix = "client" if len(clients) == 1 else "clients"
    self._lu_client_count_lbl.config(text=f"{len(clients)} {suffix} loaded")


# ══════════════════════════════════════════════════════════════════════
#  EXPORT — same sector / client Excel as Loan Balance tab,
#  including HIGH + MEDIUM risk clients from Analysis data.
# ══════════════════════════════════════════════════════════════════════

def _lu_analysis_export_high_risk_sector_excel(self):
    """Sector summary + client breakdown for HIGH and MEDIUM risk rows (Analysis tab)."""
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        messagebox.showerror(
            "Missing Library",
            "openpyxl is not installed.\nRun:  pip install openpyxl",
            parent=self,
        )
        return
    all_data = getattr(self, "_lu_all_data", None) or {}
    general = all_data.get("general") or []
    export_rows = [
        r for r in general
        if str(r.get("score_label") or "").strip().upper() in ("HIGH", "MEDIUM", "MODERATE")
    ]
    if not export_rows:
        messagebox.showinfo(
            "No HIGH/MEDIUM risk clients",
            "There are no clients with Risk = HIGH or MEDIUM in the loaded data.",
            parent=self,
        )
        return
    from lu_loanbal_export_patch import (
        _build_loanbal_export_payload_from_records,
        _generate_loanbal_excel,
    )

    doc_title = "Risk Clients (High + Medium)"
    default_name = f"RiskClients_HighMedium_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    path = filedialog.asksaveasfilename(
        parent=self,
        title="Save Risk clients export",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=default_name,
    )
    if not path:
        return
    try:
        payload = _build_loanbal_export_payload_from_records(export_rows)
        _generate_loanbal_excel(
            payload,
            path,
            filepath=getattr(self, "_lu_filepath", "") or "",
            document_title=doc_title,
            client_sheet_title="Risk Clients (High + Medium) — Individual Breakdown",
            client_sheet_subtitle_suffix="Subset: clients with HIGH or MEDIUM risk.",
            export_scope_note=(
                "Analysis export — includes only clients with HIGH or MEDIUM risk. "
                "Detailed rules follow on this sheet."
            ),
            exported_by=getattr(self, "_current_username", None),
        )
        messagebox.showinfo("Export Complete", f"Excel saved to:\n{path}", parent=self)
    except Exception as ex:
        messagebox.showerror("Excel Export Error", str(ex), parent=self)


# ══════════════════════════════════════════════════════════════════════
#  BUILD ANALYSIS VIEW  (called once during panel construction)
# ══════════════════════════════════════════════════════════════════════

def _build_analysis_view(self, parent):
    """Build the scrollable Analysis sub-view inside parent."""
    self._lu_analysis_view = tk.Frame(parent, bg=_CARD_WHITE)
    self._lu_analysis_view.place(relx=0, rely=0, relwidth=1, relheight=1)

    # Risk Settings button — injected into the existing client toolbar.
    # _lu_client_bar is built by the outer app before this method is called.
    toolbar = getattr(self, "_lu_client_bar", None)
    if toolbar is not None and not getattr(self, "_product_risk_btn", None):
        self._product_risk_btn = tk.Button(
            toolbar,
            text="⚙  Product Risk",
            font=F(8, "bold"),
            fg=_LIME_MID,
            bg=_NAVY_LIGHT,
            activebackground=_NAVY_MID,
            activeforeground=_WHITE,
            relief="flat",
            bd=0,
            padx=10,
            pady=0,
            cursor="hand2",
            command=lambda: (
                self._open_product_risk_dialog()
                if hasattr(self, "_open_product_risk_dialog")
                else None
            ),
        )
        self._product_risk_btn.pack(side="right", padx=(0, 4), pady=10)
    if toolbar is not None and not getattr(self, "_expense_risk_btn", None):
        self._expense_risk_btn = tk.Button(
            toolbar,
            text="⚙  Expense Risk",
            font=F(8, "bold"),
            fg=_LIME_MID,
            bg=_NAVY_LIGHT,
            activebackground=_NAVY_MID,
            activeforeground=_WHITE,
            relief="flat",
            bd=0,
            padx=10,
            pady=0,
            cursor="hand2",
            command=lambda: (
                self._open_expense_risk_dialog()
                if hasattr(self, "_open_expense_risk_dialog")
                else None
            ),
        )
        self._expense_risk_btn.pack(side="right", padx=(0, 4), pady=10)
    if toolbar is not None and not getattr(self, "_risk_settings_btn", None):
        self._risk_settings_btn = tk.Button(
            toolbar,
            text="⚙  Industry Risk",
            font=F(8, "bold"),
            fg=_LIME_MID,
            bg=_NAVY_LIGHT,
            activebackground=_NAVY_MID,
            activeforeground=_WHITE,
            relief="flat",
            bd=0,
            padx=10,
            pady=0,
            cursor="hand2",
            command=lambda: (
                self._open_industry_risk_dialog()
                if hasattr(self, "_open_industry_risk_dialog")
                else _lu_open_risk_settings(self)
            ),
        )
        self._risk_settings_btn.pack(side="right", padx=(0, 4), pady=10)
    if toolbar is not None and not getattr(self, "_analysis_sector_export_btn", None):
        self._analysis_sector_export_btn = tk.Button(
            toolbar,
            text="💾  Risk Clients Excel",
            font=F(8, "bold"),
            fg=_LIME_MID,
            bg=_NAVY_LIGHT,
            activebackground=_NAVY_MID,
            activeforeground=_WHITE,
            relief="flat",
            bd=0,
            padx=10,
            pady=0,
            cursor="hand2",
            command=lambda: _lu_analysis_export_high_risk_sector_excel(self),
        )
        self._analysis_sector_export_btn.pack(side="right", padx=(0, 14), pady=10)

    self._lu_scroll_outer, self._lu_results_inner, self._lu_canvas = \
        _make_scrollable(self._lu_analysis_view, _CARD_WHITE)
    self._lu_results_frame = self._lu_results_inner
    _lu_show_placeholder(self)


# ══════════════════════════════════════════════════════════════════════
#  PLACEHOLDER + ERROR
# ══════════════════════════════════════════════════════════════════════

def _lu_show_placeholder(self):
    for w in self._lu_results_frame.winfo_children():
        w.destroy()
    ph = tk.Frame(self._lu_results_frame, bg=_CARD_WHITE)
    ph.pack(expand=True, fill="both", pady=60)
    tk.Label(ph, text="📊", font=("Segoe UI Emoji", 40),
             bg=_CARD_WHITE, fg=_TXT_MUTED).pack()
    tk.Label(ph, text="No analysis yet",
             font=F(14, "bold"), fg=_TXT_SOFT, bg=_CARD_WHITE).pack(pady=(8, 4))
    tk.Label(ph,
             text=("Load an Excel file to scan all clients.\n\n"
                   "Use Industry Risk / Expense Risk / Product Risk to mark HIGH or LOW.\n"
                   "Priority: Product > Expense > Industry.\n\n"
                   "Search: client, ID, PN, industry, or product text.\n"
                   "Type  high risk  or  low risk  to list all HIGH or LOW clients."),
             font=F(9), fg=_TXT_MUTED, bg=_CARD_WHITE, justify="center").pack()


def _lu_show_error(self, msg):
    for w in self._lu_results_frame.winfo_children():
        w.destroy()
    err = tk.Frame(self._lu_results_frame, bg=_CARD_WHITE)
    err.pack(expand=True, fill="both", pady=60)
    tk.Label(err, text="❌", font=("Segoe UI Emoji", 32), bg=_CARD_WHITE).pack()
    tk.Label(err, text="Analysis failed",
             font=F(13, "bold"), fg=_ACCENT_RED, bg=_CARD_WHITE).pack(pady=(8, 4))
    tk.Label(err, text=msg, font=F(9), fg=_TXT_SOFT,
             bg=_CARD_WHITE, wraplength=500, justify="center").pack()


# ══════════════════════════════════════════════════════════════════════
#  RESULTS RENDERER (dispatcher)
# ══════════════════════════════════════════════════════════════════════

def _lu_render_results(self, results: list):
    for w in self._lu_results_frame.winfo_children():
        w.destroy()
    for w in self._lu_analysis_view.winfo_children():
        if w is not self._lu_scroll_outer:
            w.destroy()

    if not results:
        self._lu_scroll_outer.pack(fill="both", expand=True)
        ph = tk.Frame(self._lu_results_frame, bg=_CARD_WHITE)
        ph.pack(expand=True, fill="both", pady=60)
        tk.Label(ph, text="⚠️  No clients found.",
                 font=F(11), fg=_ACCENT_GOLD, bg=_CARD_WHITE).pack(pady=20)
        tk.Label(ph, text="Load an Excel file with the required columns.",
                 font=F(9), fg=_TXT_SOFT, bg=_CARD_WHITE, justify="center").pack()
        return

    if getattr(self, "_lu_analysis_active_client", GENERAL_CLIENT) == GENERAL_CLIENT:
        self._lu_scroll_outer.pack_forget()
        direct = tk.Frame(self._lu_analysis_view, bg=_CARD_WHITE)
        direct.pack(fill="both", expand=True)
        _lu_render_general_view(self, results, direct)
    else:
        self._lu_scroll_outer.pack(fill="both", expand=True)
        _lu_render_client_view(self, results)
        self._lu_canvas.yview_moveto(0)


# ══════════════════════════════════════════════════════════════════════
#  GENERAL VIEW  (treeview)
# ══════════════════════════════════════════════════════════════════════

def _lu_render_general_view(self, results: list, parent: tk.Frame):
    total_lb      = sum(r.get("loan_balance") or 0 for r in results)
    total_net     = sum(r.get("net_income")   or 0 for r in results)
    industry_counts = {}
    for r in results:
        industry_counts[r["industry"]] = industry_counts.get(r["industry"], 0) + 1
    high_risk_clients = sum(
        1 for r in results
        if str(r.get("score_label") or "").strip().upper() == "HIGH"
    )
    medium_risk_clients = sum(
        1 for r in results
        if str(r.get("score_label") or "").strip().upper() in ("MEDIUM", "MODERATE")
    )

    active_industries = getattr(self, "_lu_analysis_filtered_sectors", None)
    stats_bg     = "#0E2040" if active_industries else _NAVY_MIST
    stats_border = _LIME_MID if active_industries else _BORDER_MID

    stats_bar = tk.Frame(parent, bg=stats_bg,
                         highlightbackground=stats_border, highlightthickness=1)
    stats_bar.pack(fill="x", padx=20, pady=(16, 0))

    if active_industries:
        fi = tk.Frame(stats_bar, bg=stats_bg)
        fi.pack(side="left", padx=(12, 0), pady=10)
        icon = "🏭"
        tk.Label(fi, text=f"{icon}  Filtered: {' · '.join(active_industries)}",
                 font=F(8, "bold"), fg=_LIME_MID, bg=stats_bg).pack(anchor="w")
        tk.Label(fi, text="Showing industry subset only",
                 font=F(7), fg=_TXT_MUTED, bg=stats_bg).pack(anchor="w")

    for lbl, val in [
        ("👥 Clients",              str(len(results))),
        ("💰 Total Loan Bal",       f"₱{total_lb:,.2f}"),
        ("📈 Total Net Income",     f"₱{total_net:,.2f}"),
        ("🏭 Industries",          str(len(industry_counts))),
        ("🟠 High Risk Clients",    str(high_risk_clients)),
        ("🟡 Medium Risk Clients",  str(medium_risk_clients)),
    ]:
        c = tk.Frame(stats_bar, bg=stats_bg)
        c.pack(side="left", padx=20, pady=10)
        tk.Label(c, text=lbl,  font=F(7),        fg=_TXT_SOFT, bg=stats_bg).pack(anchor="w")
        tk.Label(c, text=val,  font=F(12, "bold"),
                 fg=_WHITE if active_industries else _TXT_NAVY,
                 bg=stats_bg).pack(anchor="w")

    tk.Label(parent, text="Click any row to view full client details",
             font=F(7), fg=_TXT_MUTED, bg=_CARD_WHITE).pack(anchor="e", padx=20, pady=(4, 2))

    style = ttk.Style()
    style.configure("LU.Treeview",
                    background=_WHITE, foreground=_TXT_NAVY,
                    rowheight=26, fieldbackground=_WHITE,
                    bordercolor=_BORDER_MID, font=("Segoe UI", 9))
    style.configure("LU.Treeview.Heading",
                    background=_NAVY_DEEP, foreground=_WHITE,
                    font=("Segoe UI", 9, "bold"), relief="flat", borderwidth=0)
    style.map("LU.Treeview.Heading", background=[("active", _NAVY_LIGHT)])
    style.map("LU.Treeview",
              background=[("selected", _NAVY_GHOST)],
              foreground=[("selected", _TXT_NAVY)])
    style.layout("LU.Treeview", [("LU.Treeview.treearea", {"sticky": "nswe"})])

    tree_frame = tk.Frame(parent, bg=_CARD_WHITE)
    tree_frame.pack(fill="both", expand=True, padx=20, pady=(0, 16))
    vsb = ttk.Scrollbar(tree_frame, orient="vertical")
    vsb.pack(side="right", fill="y")
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
    hsb.pack(side="bottom", fill="x")

    COLS = tuple(c[0] for c in LU_CLIENT_TREE_SPEC)
    COL_LABELS = {c[0]: c[1] for c in LU_CLIENT_TREE_SPEC}
    COL_WIDTHS = {c[0]: c[3] for c in LU_CLIENT_TREE_SPEC}
    COL_ANCHOR = {c[0]: c[4] for c in LU_CLIENT_TREE_SPEC}

    tree = ttk.Treeview(tree_frame, columns=COLS, show="headings",
                        style="LU.Treeview",
                        yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.pack(side="left", fill="both", expand=True)
    vsb.config(command=tree.yview)
    hsb.config(command=tree.xview)

    _stretch_cols = {
        "client", "industry", "source_income", "biz_exp_detail", "hhld_exp_detail",
        "product_name", "personal_assets", "business_assets", "business_inventory",
    }
    for col in COLS:
        tree.heading(col, text=COL_LABELS[col],
                     command=lambda c=col: _tv_sort(tree, c, False))
        tree.column(col, width=COL_WIDTHS[col], minwidth=36,
                    anchor=COL_ANCHOR[col], stretch=(col in _stretch_cols))

    tree.tag_configure("HIGH",     background="#FFF5F5", foreground=_ACCENT_RED)
    tree.tag_configure("MEDIUM",   background="#FFFBF0", foreground=_ACCENT_GOLD)
    tree.tag_configure("LOW",      background="#F0FBE8", foreground=_ACCENT_SUCCESS)
    tree.tag_configure("NA",       background=_WHITE,    foreground=_TXT_MUTED)
    tree.tag_configure("alt",      background=_OFF_WHITE)

    _iid_to_client = {}
    for idx, rec in enumerate(results):
        rl = rec.get("score_label", "N/A")
        values = lu_client_row_tuple(rec)
        tag = rl if rl in ("HIGH", "MEDIUM", "LOW") else "NA"
        if tag == "NA" and idx % 2 == 1:
            tag = "alt"
        iid = tree.insert("", "end", values=values, tags=(tag,))
        _iid_to_client[iid] = rec.get("client", "")

    tree.bind("<MouseWheel>",
              lambda ev: tree.yview_scroll(int(-1*(ev.delta/120)), "units") or "break")

    def _on_row_click(event):
        iid = tree.identify_row(event.y)
        if iid:
            client_name = _iid_to_client.get(iid, "")
            if client_name:
                self._lu_client_var.set(client_name)
                _lu_on_client_change(self, client_name)

    tree.bind("<ButtonRelease-1>", _on_row_click)
    tree.bind("<Return>", _on_row_click)
    self._lu_general_tree = tree


def _tv_sort(tree, col: str, reverse: bool):
    data = [(tree.set(iid, col), iid) for iid in tree.get_children("")]
    def _key(val):
        v = val[0].replace("₱", "").replace(",", "").strip()
        try:
            return (0, float(v))
        except ValueError:
            return (1, v.lower())
    data.sort(key=_key, reverse=reverse)
    for idx, (_, iid) in enumerate(data):
        tree.move(iid, "", idx)
    tree.heading(col, command=lambda c=col: _tv_sort(tree, c, not reverse))


# ══════════════════════════════════════════════════════════════════════
#  PER-CLIENT VIEW (simple industry-based risk)
# ══════════════════════════════════════════════════════════════════════

def _lu_render_client_view(self, results: list):
    if not results:
        return
    rec         = results[0]
    client_name = rec["client"]
    label       = rec.get("score_label", "N/A")
    score       = rec.get("score", 0.0)
    hero_bg     = _CLIENT_HERO_BG.get(label, "#0A1628")
    hero_accent = _CLIENT_HERO_ACCENT.get(label, _NAVY_PALE)

    pad = tk.Frame(self._lu_results_frame, bg=_CARD_WHITE)
    pad.pack(fill="both", expand=True)

    # Hero section
    hero = tk.Frame(pad, bg=hero_bg)
    hero.pack(fill="x")
    hi = tk.Frame(hero, bg=hero_bg)
    hi.pack(fill="x", padx=28, pady=20)
    left = tk.Frame(hi, bg=hero_bg)
    left.pack(side="left", fill="y")
    tk.Label(left, text="PER-CLIENT ANALYSIS",
             font=F(7, "bold"), fg=hero_accent, bg=hero_bg).pack(anchor="w")
    tk.Label(left, text=f"👤  {client_name}",
             font=F(18, "bold"), fg=_WHITE, bg=hero_bg).pack(anchor="w", pady=(4, 2))
    tk.Label(left, text=f"Industry: {rec.get('industry', '—')}",
             font=F(9), fg=hero_accent, bg=hero_bg).pack(anchor="w")

    right = tk.Frame(hi, bg=hero_bg)
    right.pack(side="right", fill="y")
    score_icons = {"HIGH": "🟠", "MEDIUM": "🟡", "MODERATE": "🟡", "LOW": "🟢", "N/A": "⚪"}
    tk.Label(right, text=score_icons.get(label, "⚪"),
             font=("Segoe UI Emoji", 32), bg=hero_bg).pack()
    tk.Label(right, text=label, font=F(16, "bold"), fg=hero_accent, bg=hero_bg).pack()
    tk.Label(right, text=f"Risk Score  {score:.2f}",
             font=F(9), fg=_WHITE, bg=hero_bg).pack()

    # Financial summary
    fin_bar = tk.Frame(pad, bg="#EEF3FB")
    fin_bar.pack(fill="x")
    for lbl, val in [
        ("Client ID",          rec.get("client_id", "—")),
        ("PN",                 rec.get("pn", "—")),
        ("Total Source",       f"₱{rec['total_source']:,.2f}" if rec.get("total_source") else "—"),
        ("Net Income",         f"₱{rec['net_income']:,.2f}"   if rec.get("net_income")   else "—"),
        ("Amort History",      f"₱{rec['amort_history']:,.2f}" if rec.get("amort_history") else "—"),
        ("Current Amort",      f"₱{rec['current_amort']:,.2f}" if rec.get("current_amort") else "—"),
        ("Loan Balance",       f"₱{rec['loan_balance']:,.2f}" if rec.get("loan_balance") else "—"),
    ]:
        c = tk.Frame(fin_bar, bg="#EEF3FB")
        c.pack(side="left", padx=12, pady=10)
        tk.Label(c, text=lbl, font=F(7),        fg=_TXT_SOFT,  bg="#EEF3FB").pack(anchor="w")
        tk.Label(c, text=val, font=F(9, "bold"), fg=_TXT_NAVY,  bg="#EEF3FB").pack(anchor="w")
    tk.Frame(pad, bg=_BORDER_MID, height=1).pack(fill="x")

    # Risk info
    info_frame = tk.Frame(pad, bg=_CARD_WHITE)
    info_frame.pack(fill="x", padx=28, pady=16)
    tk.Label(info_frame, text="Risk Classification",
             font=F(11, "bold"), fg=_TXT_NAVY, bg=_CARD_WHITE).pack(anchor="w", pady=(0, 6))
    risk_badge = tk.Frame(info_frame, bg=_RISK_BADGE_BG.get(label, _OFF_WHITE),
                          highlightbackground=_RISK_COLOR.get(label, _TXT_MUTED),
                          highlightthickness=1)
    risk_badge.pack(anchor="w")
    tk.Label(risk_badge, text=f"  {label} RISK  ", font=F(9, "bold"),
             fg=_RISK_COLOR.get(label, _TXT_NAVY), bg=risk_badge.cget("bg"),
             padx=10, pady=4).pack()

    if rec.get("source_income"):
        sf = tk.Frame(pad, bg="#EEF3FB",
                      highlightbackground=_BORDER_MID, highlightthickness=1)
        sf.pack(fill="x", padx=28, pady=(14, 0))
        tk.Label(sf, text="💰  Source of Income",
                 font=F(9, "bold"), fg=_NAVY_MID, bg="#EEF3FB",
                 padx=12, pady=6).pack(anchor="w")
        tk.Label(sf, text=rec["source_income"],
                 font=F(8), fg=_TXT_NAVY, bg="#EEF3FB",
                 padx=12, pady=8, anchor="w", justify="left",
                 wraplength=800).pack(fill="x")

    # Simple note about risk basis
    note = tk.Frame(pad, bg=_CARD_WHITE)
    note.pack(fill="x", padx=28, pady=(16, 8))
    tk.Label(note, text="ℹ️  Risk uses product overrides, expense overrides, automatic "
                        "fuel/LPG signals, then industry. Use the ⚙ buttons to adjust overrides.",
             font=F(8), fg=_TXT_MUTED, bg=_CARD_WHITE, justify="left").pack(anchor="w")

    # Full record list (all fields, readable key/value layout)
    tk.Label(
        pad,
        text="Full Client Record",
        font=F(11, "bold"),
        fg=_TXT_NAVY,
        bg=_CARD_WHITE,
    ).pack(anchor="w", padx=28, pady=(18, 6))

    list_wrap = tk.Frame(
        pad,
        bg=_CARD_WHITE,
        highlightbackground=_BORDER_LIGHT,
        highlightthickness=1,
    )
    list_wrap.pack(fill="both", expand=True, padx=28, pady=(0, 24))

    row_idx = 0
    for _cid, heading, field, _w, _a, kind in LU_CLIENT_TREE_SPEC:
        if field not in rec:
            continue
        raw_val = rec.get(field)
        if raw_val is None or str(raw_val).strip() == "":
            continue
        # Use the same formatter as table mode for consistent value display.
        value = lu_format_lu_cell(rec, field, kind, text_limit=500)

        bg = _WHITE if row_idx % 2 == 0 else _OFF_WHITE
        row = tk.Frame(list_wrap, bg=bg)
        row.pack(fill="x")

        tk.Label(
            row,
            text=heading,
            font=F(8, "bold"),
            fg=_TXT_SOFT,
            bg=bg,
            width=26,
            anchor="nw",
            justify="left",
            padx=10,
            pady=7,
        ).pack(side="left", fill="y")

        tk.Label(
            row,
            text=value,
            font=F(9),
            fg=_TXT_NAVY,
            bg=bg,
            anchor="w",
            justify="left",
            wraplength=760,
            padx=8,
            pady=7,
        ).pack(side="left", fill="x", expand=True)

        tk.Frame(list_wrap, bg=_BORDER_LIGHT, height=1).pack(fill="x")
        row_idx += 1


# ══════════════════════════════════════════════════════════════════════
#  ATTACH
# ══════════════════════════════════════════════════════════════════════

def attach(cls):
    """Attach Analysis-tab methods to the app class."""
    cls._build_analysis_view         = _build_analysis_view
    cls._lu_show_placeholder          = _lu_show_placeholder
    cls._lu_show_error                = _lu_show_error
    cls._lu_render_results            = _lu_render_results
    cls._lu_render_general_view       = _lu_render_general_view
    cls._lu_render_client_view        = _lu_render_client_view
    cls._lu_on_client_change          = _lu_on_client_change
    cls._lu_filter_by_search          = _lu_filter_by_search
    cls._lu_populate_client_dropdown  = _lu_populate_client_dropdown
    cls._lu_update_filter_pill        = _lu_update_filter_pill
    cls._lu_clear_sector_filter       = _lu_clear_sector_filter
    cls._lu_open_risk_settings        = _lu_open_risk_settings
    cls._lu_apply_risk_overrides      = _lu_apply_risk_overrides
    cls._lu_analysis_export_high_risk_sector_excel = _lu_analysis_export_high_risk_sector_excel
    cls._tv_sort                      = staticmethod(_tv_sort)
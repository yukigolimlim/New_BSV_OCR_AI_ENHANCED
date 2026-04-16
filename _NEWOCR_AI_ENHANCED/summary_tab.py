"""
summary_tab.py — DocExtract Pro
=================================
"Look-Up Summary" tab: persistent database-backed view of ALL applicants
ever processed by the Look-Up tab, across sessions.

DEDUPLICATION POLICY (client_id):
  • A non-empty client_id is the primary unique key.
  • On insert/merge: if a row with the same client_id already exists,
    only NULL / empty fields in the existing row are patched with the
    incoming values — existing data is NEVER overwritten.
  • On startup: if the DB already contains duplicate client_ids (legacy
    data), they are collapsed into one row (most-complete record kept).

DATA POPULATION SOURCES:
  • From Look-Up tab (automatic on extraction):
      applicant_name, residence_address, office_address,
      income_items, income_total, business_items, business_total,
      household_items, household_total, net_income,
      petrol_risk, transport_risk, results_json, page_map,
      amort_history_total, source_file, status, session_id
  • From "Other Data" import (manual):
      client_id, pn, industry_name, loan_balance, amortized_cost
  • From "Amort." import (manual):
      amort_current_total
  • From "P.Loan" import (manual):
      principal_loan, maturity, interest_rate, plus expanded fields
"""

import re
import io
import csv
import json
import sqlite3
import threading
import tkinter as tk
import tkinter.ttk as ttk
import customtkinter as ctk
from pathlib import Path
from datetime import datetime
from tkinter import filedialog, messagebox

from app_constants import (
    NAVY_DEEP, NAVY_LIGHT, NAVY_MID, NAVY_PALE, NAVY_MIST, NAVY_GHOST,
    WHITE, OFF_WHITE, CARD_WHITE,
    LIME_BRIGHT, LIME_DARK, LIME_MID, LIME_PALE, LIME_MIST,
    TXT_NAVY, TXT_SOFT, TXT_MUTED, TXT_ON_LIME,
    ACCENT_RED, ACCENT_GOLD, ACCENT_SUCCESS,
    BORDER_LIGHT, BORDER_MID,
    SIDEBAR_BG, SIDEBAR_ITEM, SIDEBAR_HVR,
    F, FF, FMONO,
)

DB_DIR  = Path(__file__).parent / "lookup_summary_results"
DB_PATH = DB_DIR / "applicants.db"

PAD          = 20
PAGE_SIZE    = 50
HDR_BG       = "#93C47D"
HDR_FG       = "#FFFFFF"
ROW_BG_EVEN  = "#F3F9F0"
ROW_BG_ODD   = WHITE
NET_GREEN    = "#1F6B28"
TOT_BG       = "#D9EAD3"
SEC_BG       = "#E8F0FA"
SEC_FG       = NAVY_MID

STATUS_COLORS = {
    "done":    ("#F0FDF4", "#166534", "✓ Done"),
    "error":   ("#FEF2F2", "#991B1B", "✗ Error"),
    "running": ("#FFFBEB", "#92400E", "⟳ Running"),
    "waiting": ("#F3F4F6", "#6B7280", "… Waiting"),
}

# ── Virtual columns: derived from results_json at render time, not real DB cols.
# ── These must never be used in ORDER BY or WHERE clauses.
_VIRTUAL_COLS = {
    "spouse_info",
    "personal_assets",
    "business_assets",
    "business_inventory",
}

# ── TABLE_COLS definition ─────────────────────────────────────────────
# (db_col, label, width_px, is_monetary, is_text_block)
# Columns populated by Look-Up are marked with ★ in comments.
TABLE_COLS = [
    # ── Populated by Other Data import ───────────────────────────────
    ("client_id",           "Client ID",                           130, False, False),
    ("pn",                  "PN",                                  100, False, False),
    # ── ★ Populated by Look-Up ────────────────────────────────────────
    ("applicant_name",      "Applicant",                           200, False, False),
    ("residence_address",   "Residence Address",                   220, False, True),
    ("office_address",      "Office Address",                      180, False, True),
    # ── Populated by Other Data import ───────────────────────────────
    ("industry_name",       "Industry Name",                       160, False, False),
    # ── ★ Virtual — derived from results_json (Look-Up) ──────────────
    ("spouse_info",         "Spouse Info",                         220, False, True),
    ("personal_assets",     "Personal Assets",                     220, False, True),
    ("business_assets",     "Business Assets",                     220, False, True),
    ("business_inventory",  "Business Inventory",                  200, False, True),
    # ── ★ Populated by Look-Up ────────────────────────────────────────
    ("income_items",        "Source of Income",                    200, False, True),
    ("income_total",        "Total Income",                        130, True,  False),
    ("business_items",      "Business Expenses",                   200, False, True),
    ("business_total",      "Total Business",                      130, True,  False),
    ("household_items",     "Household Expenses",                  200, False, True),
    ("household_total",     "Total Household",                     130, True,  False),
    ("net_income",          "Total Net Income",                    130, True,  False),
    # ── ★ Populated by Look-Up (real DB column) ───────────────────────
    ("amort_history_total", "Total Amort. History",                150, True,  False),
    # ── Populated by Amort. import ────────────────────────────────────
    ("amort_current_total", "Total Current Amort.",                150, True,  False),
    # ── Populated by Other Data import ───────────────────────────────
    ("loan_balance",        "Loan Balance",                        150, True,  False),
    ("amortized_cost",      "Total Amortized Cost",                160, True,  False),
    # ── Populated by P.Loan import ────────────────────────────────────
    ("principal_loan",      "Principal Loan",                      150, True,  False),
    ("maturity",            "Maturity",                            140, False, False),
    ("interest_rate",       "Interest Rate",                       120, False, False),
    # ── Populated by P.Loan import (expanded fields) ──────────────────
    ("branch",              "Branch",                              130, False, False),
    ("loan_class_name",     "Loan Class",                          140, False, False),
    ("product_name",        "Product Name",                        160, False, False),
    ("loan_date",           "Loan Date",                           120, False, False),
    ("term_unit",           "Term Unit",                           90,  False, False),
    ("term",                "Term",                                80,  False, False),
    ("security",            "Security",                            160, False, True ),
    ("release_tag",         "Release Tag",                         120, False, False),
    ("loan_status",         "Loan Status",                         120, False, False),
    ("ao_name",             "AO Name",                             160, False, False),
]

TREE_COLS = [c[0] for c in TABLE_COLS]

LOOKUP_ROWS = [
    ("cibi_place_of_work",      "CI/BI Report",      "Office Address"),
    ("cibi_temp_residence",     "CI/BI Report",      "Residence Address"),
    ("cibi_spouse",             "CI/BI Report",      "Spouse / Employment"),
    ("cibi_spouse_office",      "CI/BI Report",      "Spouse Office Address"),
    ("cibi_personal_assets",    "CI/BI Report",      "Personal Assets"),
    ("cibi_business_assets",    "CI/BI Report",      "Business Assets"),
    ("cibi_business_inventory", "CI/BI Report",      "Business Inventory"),
    ("cibi_petrol_products",    "CI/BI Report",      "Petrol / Plastics / PVC Risk"),
    ("cibi_transport_services", "CI/BI Report",      "Transport Services Risk"),
    ("credit_history_amort",    "CI/BI Report",      "Credit History Amort."),
    ("income_remittance",       "Cashflow Analysis", "Source of Income"),
    ("cfa_business_expenses",   "Cashflow Analysis", "Business Expenses"),
    ("cfa_household_expenses",  "Cashflow Analysis", "Household / Personal Expenses"),
    ("ws_food_grocery",         "Worksheet",         "Food / Grocery"),
    ("ws_fuel_transport",       "Worksheet",         "Fuel and Transportation"),
    ("ws_electricity",          "Worksheet",         "Electricity Expense"),
    ("ws_fertilizer",           "Worksheet",         "Fertilizer"),
    ("ws_forwarding",           "Worksheet",         "Forwarding / Trucking / Hauling"),
    ("ws_fuel_diesel",          "Worksheet",         "Fuel / Gas / Diesel"),
    ("ws_equipment",            "Worksheet",         "Cost of Rent of Equipment"),
]
NON_MONETARY = {"cibi_place_of_work", "cibi_temp_residence",
                "cibi_spouse", "cibi_spouse_office",
                "cibi_personal_assets", "cibi_business_assets",
                "cibi_business_inventory"}


# ═══════════════════════════════════════════════════════════════════════
#  HELPERS — derive display fields from results_json
# ═══════════════════════════════════════════════════════════════════════

def _extract_spouse_info(results: dict) -> str:
    parts = []

    raw_spouse = results.get("cibi_spouse", {})
    spouse_items = raw_spouse.get("items", []) if isinstance(raw_spouse, dict) else []

    raw_office = results.get("cibi_spouse_office", {})
    office_items = raw_office.get("items", []) if isinstance(raw_office, dict) else []

    def _clean(text: str) -> str:
        return re.sub(r"\s*\[N/A\]\s*$", "", text).strip()

    for item in spouse_items:
        parts.append(_clean(item))
    for item in office_items:
        cleaned = _clean(item)
        if cleaned:
            parts.append(f"Office: {cleaned}")

    return "  ·  ".join(parts) if parts else ""


def _extract_asset_items(results: dict, key: str) -> str:
    raw = results.get(key, {})
    if not isinstance(raw, dict):
        return ""
    items = raw.get("items", [])
    return "  ·  ".join(items) if items else ""


def _parse_amort_history_total(results_json_str: str) -> float | None:
    """
    Derive amort_history_total from results_json string.
    Returns float or None.
    """
    try:
        blob = json.loads(results_json_str or "{}")
        data = blob.get("credit_history_amort", {})
        raw  = data.get("total", "") if isinstance(data, dict) else ""
        if raw:
            cleaned = re.sub(r"[^\d.]", "", str(raw).replace(",", ""))
            return float(cleaned) if cleaned else None
    except Exception:
        pass
    return None


# ═══════════════════════════════════════════════════════════════════════
#  EXPORT CHECKLIST DIALOG
# ═══════════════════════════════════════════════════════════════════════

_CHECKLIST_PREVIEW_COLS = [
    ("Applicant",                                180, False),
    ("Client ID",                                 90, False),
    ("PN",                                        80, False),
    ("Industry Name",                            120, False),
    ("Spouse Info",                              160, False),
    ("Personal Assets",                          160, False),
    ("Business Assets",                          160, False),
    ("Business Inventory",                       140, False),
    ("Total Source Of Income",                   120, True),
    ("Total Business Expenses",                  120, True),
    ("Total Household / Personal Expenses",      130, True),
    ("Total Net Income",                         120, True),
    ("Total Amortization History",               130, True),
    ("Total Current Amortization",               130, True),
    ("Loan Balance",                             110, True),
    ("Total Amortized Cost",                     120, True),
    ("Principal Loan",                           120, True),
    ("Maturity",                                 120, False),
    ("Interest Rate",                            110, False),
    # ── P.Loan expanded ───────────────────────────────────────────────
    ("Branch",                                   110, False),
    ("Loan Class",                               120, False),
    ("Product Name",                             140, False),
    ("Loan Date",                                100, False),
    ("Term Unit",                                80,  False),
    ("Term",                                     70,  False),
    ("Loan Amount",                              120, True ),
    ("Loan Status",                              100, False),
    ("AO Name",                                  140, False),
]

_ALL_EXPORT_COLS = [
    ("Client ID",                                 90, False),
    ("PN",                                        80, False),
    ("Applicant",                                180, False),
    ("Residence Address",                        160, False),
    ("Office Address",                           140, False),
    ("Industry Name",                            120, False),
    ("Spouse Info",                              200, False),
    ("Personal Assets",                          200, False),
    ("Business Assets",                          200, False),
    ("Business Inventory",                       180, False),
    ("Source of Income",                         160, False),
    ("Total Source Of Income",                   120, True),
    ("Business Expenses",                        160, False),
    ("Total Business Expenses",                  120, True),
    ("Household / Personal Expenses",            160, False),
    ("Total Household / Personal Expenses",      130, True),
    ("Total Net Income",                         120, True),
    ("Total Amortization History",               130, True),
    ("Total Current Amortization",               130, True),
    ("Loan Balance",                             110, True),
    ("Total Amortized Cost",                     120, True),
    ("Principal Loan",                           120, True),
    ("Maturity",                                 120, False),
    ("Interest Rate",                            110, False),
    # ── P.Loan expanded ───────────────────────────────────────────────
    ("Branch",                                   120, False),
    ("Loan Class",                               130, False),
    ("Product Name",                             150, False),
    ("Loan Date",                                110, False),
    ("Term Unit",                                90,  False),
    ("Term",                                     80,  False),
    ("Security",                                 160, False),
    ("Release Tag",                              110, False),
    ("Loan Amount",                              130, True ),
    ("Loan Status",                              110, False),
    ("AO Name",                                  150, False),
]


def _checklist_fmt_currency(val) -> str:
    if val in (None, ""):
        return "—"
    try:
        return f"P{float(val):,.2f}"
    except Exception:
        return str(val) or "—"


def _apply_checklist_tree_style():
    style = ttk.Style()
    style.configure(
        "Checklist.Treeview",
        background=WHITE, foreground=TXT_NAVY, fieldbackground=WHITE,
        rowheight=30, font=("Segoe UI", 9), borderwidth=0, relief="flat",
    )
    style.configure(
        "Checklist.Treeview.Heading",
        background=HDR_BG, foreground=HDR_FG,
        font=("Segoe UI", 9, "bold"), relief="flat", borderwidth=0, padding=(6, 6),
    )
    style.map("Checklist.Treeview.Heading",
        background=[("active", "#7AB567")], relief=[("active", "flat")])
    style.map("Checklist.Treeview",
        background=[("selected", "#C8E6C9")], foreground=[("selected", NAVY_DEEP)])


def _show_export_checklist(parent: tk.Widget, flat_rows: list) -> tuple | None:
    if not flat_rows:
        return ([], [c[0] for c in _ALL_EXPORT_COLS])

    _apply_checklist_tree_style()

    win = tk.Toplevel(parent)
    win.title("Select Rows & Columns to Export")
    win.configure(bg=CARD_WHITE)
    win.resizable(True, True)
    win.grab_set()

    p_x = parent.winfo_rootx(); p_y = parent.winfo_rooty()
    p_w = parent.winfo_width(); p_h = parent.winfo_height()
    w_w, w_h = 1140, 720
    cx = p_x + (p_w - w_w) // 2; cy = p_y + (p_h - w_h) // 2
    win.geometry(f"{w_w}x{w_h}+{cx}+{cy}")
    win.minsize(760, 500)

    result = [None]
    checked:     dict[str, bool] = {str(i): True for i in range(len(flat_rows))}
    col_checked: dict[str, bool] = {c[0]: True for c in _ALL_EXPORT_COLS}

    hdr = tk.Frame(win, bg=NAVY_DEEP)
    hdr.pack(fill="x")
    title_f = tk.Frame(hdr, bg=NAVY_DEEP)
    title_f.pack(side="left", padx=16, pady=10)
    tk.Label(title_f, text="Export Preview — Select Rows & Columns to Include",
             font=("Segoe UI", 13, "bold"), fg=WHITE, bg=NAVY_DEEP).pack(anchor="w")
    tk.Label(title_f,
             text="Tick the rows and columns you want. Only checked items will appear in the Excel file.",
             font=("Segoe UI", 8), fg="#8DA8C8", bg=NAVY_DEEP).pack(anchor="w", pady=(1, 0))

    count_var = tk.StringVar()
    tk.Label(hdr, textvariable=count_var, font=("Segoe UI", 11, "bold"),
             fg="#B9F5A0", bg=NAVY_DEEP, padx=16).pack(side="right", pady=10)

    def _refresh_count():
        nr = sum(1 for v in checked.values() if v)
        nc = sum(1 for v in col_checked.values() if v)
        count_var.set(f"{nr} / {len(flat_rows)} rows  ·  {nc} / {len(_ALL_EXPORT_COLS)} cols")

    _refresh_count()

    col_outer = tk.Frame(win, bg="#E8F0FB",
                         highlightbackground=BORDER_MID, highlightthickness=1)
    col_outer.pack(fill="x", padx=16, pady=(8, 0))
    col_hdr = tk.Frame(col_outer, bg="#E8F0FB")
    col_hdr.pack(fill="x", padx=8, pady=(6, 2))
    tk.Label(col_hdr, text="📋  Columns to export:",
             font=("Segoe UI", 8, "bold"), fg=NAVY_MID, bg="#E8F0FB").pack(side="left")

    def _col_set_all(val: bool):
        for k, v in _col_vars.items():
            v.set(val); col_checked[k] = val
        _refresh_count()

    tk.Button(col_hdr, text="All", font=("Segoe UI", 7, "bold"),
              fg=TXT_ON_LIME, bg=LIME_MID, activebackground=LIME_BRIGHT,
              activeforeground=TXT_ON_LIME, relief="flat", bd=0,
              padx=8, pady=2, cursor="hand2",
              command=lambda: _col_set_all(True)).pack(side="right", padx=(4, 0))
    tk.Button(col_hdr, text="None", font=("Segoe UI", 7, "bold"),
              fg=TXT_NAVY, bg="#D8E4F4", activebackground="#C0D0EC",
              relief="flat", bd=0, padx=8, pady=2, cursor="hand2",
              command=lambda: _col_set_all(False)).pack(side="right", padx=(0, 4))

    col_grid = tk.Frame(col_outer, bg="#E8F0FB")
    col_grid.pack(fill="x", padx=8, pady=(0, 6))
    _col_vars: dict[str, tk.BooleanVar] = {}
    COLS_PER_ROW = 6

    for idx, (col_key, col_w, is_mon) in enumerate(_ALL_EXPORT_COLS):
        var = tk.BooleanVar(value=True)
        _col_vars[col_key] = var

        def _on_toggle(k=col_key, v=var):
            col_checked[k] = v.get(); _refresh_count()

        row_idx = idx // COLS_PER_ROW; col_idx = idx % COLS_PER_ROW
        pill = tk.Frame(col_grid, bg="#D8E8FB",
                        highlightbackground="#B8CFF0", highlightthickness=1)
        pill.grid(row=row_idx, column=col_idx, padx=3, pady=3, sticky="w")
        tk.Checkbutton(pill, text=col_key, variable=var,
                       font=("Segoe UI", 8), fg=TXT_NAVY, bg="#D8E8FB",
                       activebackground="#C0D4F0", selectcolor=WHITE,
                       relief="flat", bd=0, padx=6, pady=3,
                       command=_on_toggle).pack()

    toolbar = tk.Frame(win, bg="#F0F4FA",
                       highlightbackground=BORDER_MID, highlightthickness=1)
    toolbar.pack(fill="x", padx=16, pady=(8, 0))

    def _mk_toolbar_btn(text, bg, fg, hov, cmd):
        return tk.Button(toolbar, text=text, font=("Segoe UI", 8, "bold"),
                         fg=fg, bg=bg, activebackground=hov, activeforeground=fg,
                         relief="flat", bd=0, padx=10, pady=4,
                         cursor="hand2", command=cmd)

    _mk_toolbar_btn("☑  Select All Rows", LIME_MID, TXT_ON_LIME, LIME_BRIGHT,
                    lambda: _set_all(True)).pack(side="left", padx=(8, 4), pady=6)
    _mk_toolbar_btn("☐  Deselect All Rows", "#E0E8F0", TXT_NAVY, "#C8D8EC",
                    lambda: _set_all(False)).pack(side="left", padx=(0, 12), pady=6)

    tk.Frame(toolbar, bg=BORDER_MID, width=1).pack(side="left", fill="y", pady=4)
    tk.Label(toolbar, text="Quick filter:", font=("Segoe UI", 8),
             fg=TXT_MUTED, bg="#F0F4FA").pack(side="left", padx=(12, 4), pady=6)
    filter_var = tk.StringVar()
    filter_var.trace_add("write", lambda *_: _apply_filter())
    tk.Entry(toolbar, textvariable=filter_var, font=("Segoe UI", 9),
             fg=TXT_NAVY, bg=WHITE, relief="solid", bd=1,
             insertbackground=NAVY_MID, width=24).pack(side="left", pady=6, ipady=3)

    CHECK_COL = "#check"
    TCOLS     = [CHECK_COL] + [c[0] for c in _CHECKLIST_PREVIEW_COLS]

    tbl_outer = tk.Frame(win, bg=BORDER_LIGHT)
    tbl_outer.pack(fill="both", expand=True, padx=16, pady=(6, 0))
    tbl_wrap = tk.Frame(tbl_outer, bg=CARD_WHITE)
    tbl_wrap.pack(fill="both", expand=True, padx=1, pady=1)
    tbl_wrap.rowconfigure(0, weight=1); tbl_wrap.columnconfigure(0, weight=1)

    vscroll = tk.Scrollbar(tbl_wrap, orient="vertical", relief="flat",
                           troughcolor=OFF_WHITE, bg=BORDER_LIGHT, width=8, bd=0)
    vscroll.grid(row=0, column=1, sticky="ns")
    hscroll = tk.Scrollbar(tbl_wrap, orient="horizontal", relief="flat",
                           troughcolor=OFF_WHITE, bg=BORDER_LIGHT, bd=0)
    hscroll.grid(row=1, column=0, columnspan=2, sticky="ew")

    tree = ttk.Treeview(tbl_wrap, columns=TCOLS, show="headings",
                        style="Checklist.Treeview",
                        yscrollcommand=vscroll.set, xscrollcommand=hscroll.set,
                        selectmode="browse")
    tree.grid(row=0, column=0, sticky="nsew")
    vscroll.config(command=tree.yview); hscroll.config(command=tree.xview)
    tree.bind("<Enter>", lambda e: tree.bind_all(
        "<MouseWheel>",
        lambda ev: tree.yview_scroll(int(-1 * (ev.delta / 120)), "units")))
    tree.bind("<Leave>", lambda e: tree.unbind_all("<MouseWheel>"))

    _all_hdr = [True]

    def _header_toggle():
        new_val = not _all_hdr[0]; _all_hdr[0] = new_val; _set_all(new_val)

    tree.heading(CHECK_COL, text="☑ All", command=_header_toggle)
    tree.column(CHECK_COL, width=62, minwidth=50, anchor="center", stretch=False)
    for col_name, col_w, is_mon in _CHECKLIST_PREVIEW_COLS:
        tree.heading(col_name, text=col_name)
        tree.column(col_name, width=col_w, minwidth=50,
                    anchor="e" if is_mon else "w", stretch=False)

    tree.tag_configure("even_on",  background=ROW_BG_EVEN, foreground=TXT_NAVY)
    tree.tag_configure("odd_on",   background=ROW_BG_ODD,  foreground=TXT_NAVY)
    tree.tag_configure("even_off", background="#F0F0F0",    foreground="#BBBBBB")
    tree.tag_configure("odd_off",  background="#EBEBEB",    foreground="#BBBBBB")

    _visible_iids: list[str] = []

    def _row_values(idx: int, row_dict: dict) -> list:
        vals = ["☑" if checked[str(idx)] else "☐"]
        for col_name, _, is_mon in _CHECKLIST_PREVIEW_COLS:
            raw = row_dict.get(col_name, "")
            if is_mon:
                vals.append(_checklist_fmt_currency(raw))
            else:
                vals.append(str(raw) if raw is not None else "")
        return vals

    def _row_tag(idx: int) -> str:
        even = (idx % 2 == 0)
        return ("even_on" if even else "odd_on") if checked[str(idx)] \
               else ("even_off" if even else "odd_off")

    def _apply_filter():
        term = filter_var.get().strip().lower()
        tree.delete(*tree.get_children()); _visible_iids.clear()
        for i, row_dict in enumerate(flat_rows):
            iid = str(i)
            if term:
                haystack = " ".join(
                    str(row_dict.get(c, "") or "").lower()
                    for c, _, __ in _CHECKLIST_PREVIEW_COLS)
                if term not in haystack:
                    continue
            _visible_iids.append(iid)
            tree.insert("", "end", iid=iid,
                        values=_row_values(i, row_dict), tags=(_row_tag(i),))
        _refresh_count()

    _apply_filter()

    def _update_row_display(iid: str):
        idx = int(iid)
        tree.item(iid, values=_row_values(idx, flat_rows[idx]),
                  tags=(_row_tag(idx),))

    def _toggle_row(iid: str):
        checked[iid] = not checked[iid]; _update_row_display(iid); _refresh_count()
        n = sum(1 for v in checked.values() if v)
        if n == len(flat_rows):
            _all_hdr[0] = True; tree.heading(CHECK_COL, text="☑ All")
        elif n == 0:
            _all_hdr[0] = False; tree.heading(CHECK_COL, text="☐ All")
        else:
            _all_hdr[0] = False; tree.heading(CHECK_COL, text="— All")

    def _set_all(val: bool):
        _all_hdr[0] = val
        for iid in checked:
            checked[iid] = val
        for iid in _visible_iids:
            _update_row_display(iid)
        tree.heading(CHECK_COL, text="☑ All" if val else "☐ All")
        _refresh_count()

    tree.bind("<ButtonRelease-1>", lambda e: _toggle_row(tree.identify_row(e.y))
              if tree.identify_row(e.y) else None)
    tree.bind("<space>", lambda e: _toggle_row(tree.focus()) if tree.focus() else None)

    btn_bar = tk.Frame(win, bg=CARD_WHITE,
                       highlightbackground=BORDER_MID, highlightthickness=1)
    btn_bar.pack(fill="x", padx=16, pady=(4, 12))

    def _mk_btn(parent, text, bg, fg, hov, cmd):
        return tk.Button(parent, text=text, font=("Segoe UI", 9, "bold"),
                         fg=fg, bg=bg, activebackground=hov, activeforeground=fg,
                         relief="flat", bd=0, padx=18, pady=7,
                         cursor="hand2", command=cmd)

    def _on_export():
        selected_rows = [flat_rows[int(i)]
                         for i in sorted(checked, key=int) if checked[i]]
        selected_cols = [c[0] for c in _ALL_EXPORT_COLS if col_checked.get(c[0], True)]
        result[0] = (selected_rows, selected_cols); win.destroy()

    def _on_cancel():
        result[0] = None; win.destroy()

    _mk_btn(btn_bar, "✕  Cancel", "#F5F5F5", TXT_SOFT, "#E0E0E0",
            _on_cancel).pack(side="right", padx=(4, 8), pady=6)
    _mk_btn(btn_bar, "📊  Export Selected to Excel", LIME_MID, TXT_ON_LIME, LIME_BRIGHT,
            _on_export).pack(side="right", padx=(0, 4), pady=6)
    tk.Label(btn_bar, textvariable=count_var, font=("Segoe UI", 9),
             fg=TXT_MUTED, bg=CARD_WHITE).pack(side="left", padx=10, pady=6)

    win.protocol("WM_DELETE_WINDOW", _on_cancel)
    parent.wait_window(win)
    return result[0]


# ═══════════════════════════════════════════════════════════════════════
#  DATABASE LAYER
# ═══════════════════════════════════════════════════════════════════════

def _db_connect() -> sqlite3.Connection:
    DB_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=10, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


# ── All columns that can be patched (never overwrite existing non-null data)
_PATCHABLE_COLS = [
    # ── Set by Look-Up ───────────────────────────────────────────────
    "applicant_name", "residence_address", "office_address",
    "income_items", "income_total",
    "business_items", "business_total",
    "household_items", "household_total",
    "net_income", "petrol_risk", "transport_risk",
    "results_json", "page_map",
    "amort_history_total",          # ← real DB column, set by Look-Up
    "source_file", "status",
    # ── Set by Other Data import ─────────────────────────────────────
    "client_id", "pn", "industry_name", "loan_balance", "amortized_cost",
    # ── Set by Amort. import ─────────────────────────────────────────
    "amort_current_total",
    # ── Set by P.Loan import ─────────────────────────────────────────
    "principal_loan", "maturity", "interest_rate",
    "branch", "loan_class_name", "product_name",
    "loan_date", "term_unit", "term", "security", "release_tag",
    "loan_amount", "loan_status", "ao_name",
]


def _patch_existing(conn: sqlite3.Connection, existing_id: int,
                    incoming: dict) -> None:
    """Fill in only columns that are currently NULL/empty. Never overwrites."""
    existing = dict(conn.execute(
        "SELECT * FROM applicants WHERE id=?", (existing_id,)
    ).fetchone())

    parts, vals = [], []
    for col in _PATCHABLE_COLS:
        if col not in incoming:
            continue
        existing_val = existing.get(col)
        is_empty = (existing_val is None or
                    (isinstance(existing_val, str) and existing_val.strip() == ""))
        if is_empty and incoming[col] not in (None, ""):
            parts.append(f"{col}=?")
            vals.append(incoming[col])

    if parts:
        conn.execute(
            f"UPDATE applicants SET {', '.join(parts)} WHERE id=?",
            vals + [existing_id]
        )


def _db_init():
    with _db_connect() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS applicants (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id          TEXT    NOT NULL,
            processed_at        TEXT    NOT NULL,
            source_file         TEXT,
            status              TEXT    DEFAULT 'done',
            applicant_name      TEXT,
            residence_address   TEXT,
            office_address      TEXT,
            income_items        TEXT,
            income_total        REAL,
            business_items      TEXT,
            business_total      REAL,
            household_items     TEXT,
            household_total     REAL,
            net_income          REAL,
            petrol_risk         INTEGER DEFAULT 0,
            transport_risk      INTEGER DEFAULT 0,
            results_json        TEXT,
            page_map            TEXT,
            amort_history_total REAL,
            amort_current_total REAL,
            client_id           TEXT,
            pn                  TEXT,
            industry_name       TEXT,
            loan_balance        REAL,
            amortized_cost      REAL,
            principal_loan      REAL,
            maturity            TEXT,
            interest_rate       TEXT,
            branch              TEXT,
            loan_class_name     TEXT,
            product_name        TEXT,
            loan_date           TEXT,
            term_unit           TEXT,
            term                TEXT,
            security            TEXT,
            release_tag         TEXT,
            loan_amount         REAL,
            loan_status         TEXT,
            ao_name             TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_session   ON applicants(session_id);
        CREATE INDEX IF NOT EXISTS idx_name      ON applicants(applicant_name COLLATE NOCASE);
        CREATE INDEX IF NOT EXISTS idx_status    ON applicants(status);
        CREATE INDEX IF NOT EXISTS idx_processed ON applicants(processed_at);
        CREATE INDEX IF NOT EXISTS idx_clientid  ON applicants(client_id);
        """)

        # ── schema migrations for legacy DBs ──────────────────────────
        cols = [r[1] for r in conn.execute(
            "PRAGMA table_info(applicants)").fetchall()]
        migrations = [
            ("amort_history_total", "REAL"),
            ("amort_current_total", "REAL"),
            ("client_id",           "TEXT"),
            ("pn",                  "TEXT"),
            ("industry_name",       "TEXT"),
            ("loan_balance",        "REAL"),
            ("amortized_cost",      "REAL"),
            ("principal_loan",      "REAL"),
            ("maturity",            "TEXT"),
            ("interest_rate",       "TEXT"),
            ("branch",              "TEXT"),
            ("loan_class_name",     "TEXT"),
            ("product_name",        "TEXT"),
            ("industry_name_ploan", "TEXT"),
            ("loan_date",           "TEXT"),
            ("term_unit",           "TEXT"),
            ("term",                "TEXT"),
            ("security",            "TEXT"),
            ("release_tag",         "TEXT"),
            ("loan_amount",         "REAL"),
            ("loan_balance_ploan",  "REAL"),
            ("amort_ploan",         "REAL"),
            ("loan_status",         "TEXT"),
            ("ao_name",             "TEXT"),
        ]
        for col_name, col_type in migrations:
            if col_name not in cols:
                conn.execute(
                    f"ALTER TABLE applicants ADD COLUMN {col_name} {col_type}")

        # ── backfill amort_history_total for legacy rows ───────────────
        rows_needing_backfill = conn.execute(
            "SELECT id, results_json FROM applicants "
            "WHERE amort_history_total IS NULL AND results_json IS NOT NULL"
        ).fetchall()
        for r in rows_needing_backfill:
            val = _parse_amort_history_total(r["results_json"])
            if val is not None:
                conn.execute(
                    "UPDATE applicants SET amort_history_total=? WHERE id=?",
                    (val, r["id"]))

    _db_deduplicate_client_ids()


def _db_deduplicate_client_ids() -> int:
    """
    Collapse rows that share the same non-empty client_id.
    Keeps the most-complete record; patches missing fields from duplicates;
    deletes the extras. Returns number of duplicate rows removed.
    """
    removed = 0
    with _db_connect() as conn:
        dupes = conn.execute("""
            SELECT client_id, COUNT(*) as cnt
            FROM applicants
            WHERE client_id IS NOT NULL AND TRIM(client_id) != ''
            GROUP BY client_id HAVING cnt > 1
        """).fetchall()

        for dup in dupes:
            cid  = dup["client_id"]
            rows = [dict(r) for r in conn.execute(
                "SELECT * FROM applicants WHERE client_id=? ORDER BY id ASC",
                (cid,)).fetchall()]

            def _score(r):
                return sum(1 for v in r.values()
                           if v is not None and str(v).strip() != "")

            rows.sort(key=_score, reverse=True)
            keeper_id = rows[0]["id"]
            keeper    = rows[0]

            for dup_row in rows[1:]:
                for col in _PATCHABLE_COLS:
                    if col not in keeper:
                        continue
                    keeper_val  = keeper.get(col)
                    is_empty    = (keeper_val is None or
                                   (isinstance(keeper_val, str) and
                                    keeper_val.strip() == ""))
                    incoming    = dup_row.get(col)
                    has_value   = (incoming is not None and
                                   str(incoming).strip() != "")
                    if is_empty and has_value:
                        conn.execute(
                            f"UPDATE applicants SET {col}=? WHERE id=?",
                            (incoming, keeper_id))
                        keeper[col] = incoming
                conn.execute("DELETE FROM applicants WHERE id=?",
                             (dup_row["id"],))
                removed += 1
    return removed


def _db_upsert(session_id: str, row_data: dict) -> int:
    """
    Insert or patch a row.

    Match priority:
      1. Non-empty client_id  → unique business key
      2. session_id + source_file → technical fallback

    Existing data is NEVER overwritten — only empty fields are filled.
    """
    with _db_connect() as conn:
        existing_id = None
        client_id   = (row_data.get("client_id") or "").strip()

        if client_id:
            row = conn.execute(
                "SELECT id FROM applicants WHERE TRIM(client_id)=?",
                (client_id,)).fetchone()
            if row:
                existing_id = row["id"]

        if existing_id is None:
            row = conn.execute(
                "SELECT id FROM applicants WHERE session_id=? AND source_file=?",
                (session_id, row_data.get("source_file", ""))).fetchone()
            if row:
                existing_id = row["id"]

        if existing_id is not None:
            _patch_existing(conn, existing_id, row_data)
            return existing_id

        # ── Full INSERT covering all columns set by Look-Up ────────────
        cur = conn.execute("""
            INSERT INTO applicants (
                session_id, processed_at, source_file, status,
                applicant_name, residence_address, office_address,
                income_items, income_total,
                business_items, business_total,
                household_items, household_total,
                net_income, petrol_risk, transport_risk,
                results_json, page_map,
                amort_history_total,
                amort_current_total,
                client_id, pn, industry_name,
                loan_balance, amortized_cost,
                principal_loan, maturity, interest_rate,
                branch, loan_class_name, product_name,
                loan_date, term_unit, term, security, release_tag,
                loan_status, ao_name
            ) VALUES (
                :session_id, :processed_at, :source_file, :status,
                :applicant_name, :residence_address, :office_address,
                :income_items, :income_total,
                :business_items, :business_total,
                :household_items, :household_total,
                :net_income, :petrol_risk, :transport_risk,
                :results_json, :page_map,
                :amort_history_total,
                :amort_current_total,
                :client_id, :pn, :industry_name,
                :loan_balance, :amortized_cost,
                :principal_loan, :maturity, :interest_rate,
                :branch, :loan_class_name, :product_name,
                :loan_date, :term_unit, :term, :security, :release_tag,
                :loan_status, :ao_name
            )
        """, row_data)
        return cur.lastrowid


# ── Safe sortable columns (real DB columns only, no virtuals) ──────────
_SORTABLE_COLS = (
    {c[0] for c in TABLE_COLS}
    - _VIRTUAL_COLS
    | {"processed_at", "session_id", "source_file", "id"}
)


def _db_query(search: str = "", session_id: str = "",
              sort_col: str = "processed_at", sort_asc: bool = False,
              offset: int = 0, limit: int = PAGE_SIZE,
              adv_filters: dict = None) -> tuple:
    order_col = sort_col if sort_col in _SORTABLE_COLS else "processed_at"
    direction = "ASC" if sort_asc else "DESC"
    where_parts, params = [], []

    terms = [t.strip() for t in search.split(",") if t.strip()] if search else []
    for term in terms:
        like = f"%{term}%"
        where_parts.append(
            "(applicant_name LIKE ? OR residence_address LIKE ? "
            "OR office_address LIKE ? OR income_items LIKE ? "
            "OR business_items LIKE ? OR household_items LIKE ? "
            "OR source_file LIKE ? OR client_id LIKE ? "
            "OR pn LIKE ? OR industry_name LIKE ?)")
        params.extend([like] * 10)

    if session_id:
        where_parts.append("session_id = ?")
        params.append(session_id)

    # ── Advanced per-column filters ───────────────────────────────────
    # client_id uses exact match (TRIM + case-insensitive equality).
    # All other columns use partial LIKE matching.
    _EXACT_MATCH_COLS = {"client_id"}

    if adv_filters:
        for col, values in adv_filters.items():
            if col not in _SORTABLE_COLS or not values:
                continue
            if col in _EXACT_MATCH_COLS:
                # OR across each supplied value, but each is an exact match
                placeholders = " OR ".join(
                    [f"TRIM(UPPER({col})) = TRIM(UPPER(?))" for _ in values])
                where_parts.append(f"({placeholders})")
                params.extend(values)
            else:
                placeholders = " OR ".join([f"{col} LIKE ?" for _ in values])
                where_parts.append(f"({placeholders})")
                params.extend([f"%{v}%" for v in values])

    where = ("WHERE " + " AND ".join(where_parts)) if where_parts else ""
    with _db_connect() as conn:
        total = conn.execute(
            f"SELECT COUNT(*) FROM applicants {where}", params).fetchone()[0]
        rows  = conn.execute(
            f"SELECT * FROM applicants {where} "
            f"ORDER BY {order_col} {direction} LIMIT ? OFFSET ?",
            params + [limit, offset]).fetchall()
    return rows, total


def _db_totals(session_id: str = "", search: str = "",
               adv_filters: dict = None) -> dict:
    where_parts, params = [], []

    terms = [t.strip() for t in search.split(",") if t.strip()] if search else []
    for term in terms:
        like = f"%{term}%"
        where_parts.append(
            "(applicant_name LIKE ? OR residence_address LIKE ? "
            "OR office_address LIKE ? OR income_items LIKE ? "
            "OR business_items LIKE ? OR household_items LIKE ? "
            "OR source_file LIKE ? OR client_id LIKE ? "
            "OR pn LIKE ? OR industry_name LIKE ?)")
        params.extend([like] * 10)

    if session_id:
        where_parts.append("session_id = ?")
        params.append(session_id)

    # ── Same exact-match rule for client_id ───────────────────────────
    _EXACT_MATCH_COLS = {"client_id"}

    if adv_filters:
        for col, values in adv_filters.items():
            if col not in _SORTABLE_COLS or not values:
                continue
            if col in _EXACT_MATCH_COLS:
                placeholders = " OR ".join(
                    [f"TRIM(UPPER({col})) = TRIM(UPPER(?))" for _ in values])
                where_parts.append(f"({placeholders})")
                params.extend(values)
            else:
                placeholders = " OR ".join([f"{col} LIKE ?" for _ in values])
                where_parts.append(f"({placeholders})")
                params.extend([f"%{v}%" for v in values])

    where = ("WHERE " + " AND ".join(where_parts)) if where_parts else ""
    with _db_connect() as conn:
        row = conn.execute(f"""
            SELECT
                COUNT(*)                                            AS total,
                SUM(CASE WHEN status='done'  THEN 1 ELSE 0 END)   AS done,
                SUM(CASE WHEN status='error' THEN 1 ELSE 0 END)   AS errors,
                SUM(income_total)                                  AS income,
                SUM(net_income)                                    AS net,
                SUM(amort_current_total)                           AS amort_current
            FROM applicants {where}
        """, params).fetchone()
    return dict(row) if row else {}


def _db_delete_row(row_id: int):
    with _db_connect() as conn:
        conn.execute("DELETE FROM applicants WHERE id=?", (row_id,))


def _db_clear_all():
    with _db_connect() as conn:
        conn.execute("DELETE FROM applicants")


def _db_update_amort_current(row_id: int, value: float) -> bool:
    with _db_connect() as conn:
        conn.execute("UPDATE applicants SET amort_current_total=? WHERE id=?",
                     (value, row_id))
    return True


def _db_update_amort_all(matches: list, value: float) -> int:
    count = 0
    for row_id, _ in matches:
        _db_update_amort_current(row_id, value)
        count += 1
    return count


def _db_update_other_data_all(matches: list, client_id: str, pn_joined: str,
                               industry_name: str, loan_balance,
                               amortized_cost) -> int:
    count = 0
    with _db_connect() as conn:
        for row_id, _ in matches:
            parts, vals = [], []
            if client_id:
                parts.append("client_id=?");      vals.append(client_id)
            if pn_joined:
                parts.append("pn=?");             vals.append(pn_joined)
            if industry_name:
                parts.append("industry_name=?");  vals.append(industry_name)
            if loan_balance is not None:
                parts.append("loan_balance=?");   vals.append(loan_balance)
            if amortized_cost is not None:
                parts.append("amortized_cost=?"); vals.append(amortized_cost)
            if parts:
                conn.execute(
                    f"UPDATE applicants SET {', '.join(parts)} WHERE id=?",
                    vals + [row_id])
                count += 1
    return count


# ───────────────────────────────────────────────────────────────────────
#  CHANGE 2 (added constants + _db_update_cell)
# ───────────────────────────────────────────────────────────────────────

# Columns the user may edit directly in the table.
# Virtual columns (derived from results_json) and read-only audit
# fields (session_id, processed_at, source_file, status, page_map,
# results_json, petrol_risk, transport_risk) are intentionally excluded.
_EDITABLE_COLS = {
    "client_id", "pn", "applicant_name",
    # virtual columns — written back into results_json
    "spouse_info", "personal_assets", "business_assets", "business_inventory",
    # item-text columns — real DB columns
    "income_items", "business_items", "household_items",
    "residence_address", "office_address", "industry_name",
    "income_total", "business_total", "household_total", "net_income",
    "amort_history_total", "amort_current_total",
    "loan_balance", "amortized_cost",
    "principal_loan", "maturity", "interest_rate",
    "branch", "loan_class_name", "product_name",
    "loan_date", "term_unit", "term", "security", "release_tag",
    "loan_amount", "loan_status", "ao_name",
}

# Subset of _EDITABLE_COLS that must be stored as REAL in SQLite.
_MONETARY_COLS = {
    "income_total", "business_total", "household_total", "net_income",
    "amort_history_total", "amort_current_total",
    "loan_balance", "amortized_cost", "principal_loan",
    "loan_amount",
}


def _db_update_cell(row_id: int, col_name: str, raw_value: str) -> str:
    """
    Write a single edited cell back to the DB.

    Returns the normalised display string (formatted if monetary),
    or raises ValueError if the value cannot be parsed for monetary cols.
    """
    if col_name not in _EDITABLE_COLS:
        raise ValueError(f"Column '{col_name}' is not editable.")

    if col_name in _MONETARY_COLS:
        cleaned = re.sub(r"[^\d.]", "", raw_value.replace(",", "").strip())
        if cleaned == "":
            db_val      = None
            display_val = "—"
        else:
            try:
                db_val      = float(cleaned)
                display_val = f"P{db_val:,.2f}"
            except ValueError:
                raise ValueError(
                    f"'{raw_value}' is not a valid number for '{col_name}'.")
    else:
        db_val      = raw_value.strip() or None
        display_val = raw_value.strip() if raw_value.strip() else "—"

    with _db_connect() as conn:
        conn.execute(
            f"UPDATE applicants SET {col_name}=? WHERE id=?",
            (db_val, row_id))

    return display_val


# Map from virtual column name → (results_json top-level key, sub-key)
# For spouse_info two keys are involved; handled specially in the function.
_VIRTUAL_TO_JSON: dict[str, tuple] = {
    "spouse_info":        ("cibi_spouse",             "items"),   # special
    "personal_assets":    ("cibi_personal_assets",    "items"),
    "business_assets":    ("cibi_business_assets",    "items"),
    "business_inventory": ("cibi_business_inventory", "items"),
}


def _db_update_virtual_cell(row_id: int, col_name: str, raw_value: str) -> str:
    """
    Write an edited virtual-column value back into results_json.

    The display format uses '  ·  ' as an item separator (set by _render_tree).
    We split on that, strip, and store back as a JSON items list.

    For spouse_info the plain items go into cibi_spouse["items"]; the
    'Office: …' prefixed items go into cibi_spouse_office["items"].

    Returns the normalised display string (items joined with '  ·  ').
    """
    if col_name not in _VIRTUAL_TO_JSON:
        raise ValueError(f"'{col_name}' is not a known virtual column.")

    raw_stripped = raw_value.strip()
    if raw_stripped in ("", "—"):
        items = []
    else:
        items = [p.strip() for p in raw_stripped.split("  ·  ") if p.strip()]

    with _db_connect() as conn:
        row = conn.execute(
            "SELECT results_json FROM applicants WHERE id=?", (row_id,)
        ).fetchone()
        if row is None:
            raise ValueError(f"Row id={row_id} not found.")

        try:
            blob = json.loads(row["results_json"] or "{}")
        except Exception:
            blob = {}

        if col_name == "spouse_info":
            plain_items  = [it for it in items if not it.startswith("Office: ")]
            office_items = [it[len("Office: "):] for it in items
                            if it.startswith("Office: ")]

            if "cibi_spouse" not in blob or not isinstance(blob["cibi_spouse"], dict):
                blob["cibi_spouse"] = {}
            blob["cibi_spouse"]["items"] = plain_items

            if "cibi_spouse_office" not in blob or \
               not isinstance(blob["cibi_spouse_office"], dict):
                blob["cibi_spouse_office"] = {}
            blob["cibi_spouse_office"]["items"] = office_items
        else:
            json_key, sub_key = _VIRTUAL_TO_JSON[col_name]
            if json_key not in blob or not isinstance(blob[json_key], dict):
                blob[json_key] = {}
            blob[json_key][sub_key] = items

        conn.execute(
            "UPDATE applicants SET results_json=? WHERE id=?",
            (json.dumps(blob, ensure_ascii=False), row_id)
        )

    return "  ·  ".join(items) if items else "—"


# ═══════════════════════════════════════════════════════════════════════
#  NAME-MATCHING FUNCTIONS  (similarity-based)
# ═══════════════════════════════════════════════════════════════════════

SUFFIXES = {"JR", "SR", "II", "III", "IV", "V", "ESQ", "PHD", "MD", "CPA"}

# PATCH: lowered similarity threshold from 0.72 to 0.68
SIMILARITY_THRESHOLD = 0.68

try:
    from rapidfuzz import fuzz as _rfuzz
    def _similarity(a: str, b: str) -> float:
        return _rfuzz.token_sort_ratio(a, b) / 100.0
except ImportError:
    from difflib import SequenceMatcher
    def _similarity(a: str, b: str) -> float:
        a_sorted = " ".join(sorted(a.split()))
        b_sorted = " ".join(sorted(b.split()))
        return SequenceMatcher(None, a_sorted, b_sorted).ratio()


def _normalise_for_sim(name: str) -> str:
    """
    Normalise a name for similarity comparison.
    - Flips 'LAST, FIRST' format to 'FIRST LAST'
    - Uppercases, removes punctuation
    - Strips suffixes (JR, SR, etc.)
    - Strips single-letter initials (e.g. 'I.' -> removed)
    """
    if "," in name:
        parts = name.split(",", 1)
        name  = parts[1].strip() + " " + parts[0].strip()
    tokens = re.split(r"[\s.]+", name.strip().upper())
    tokens = [re.sub(r"[^A-Z]", "", t) for t in tokens]
    tokens = [t for t in tokens if t and t not in SUFFIXES]
    tokens = [t for t in tokens if len(t) > 1]  # strip bare initials
    return " ".join(tokens)


def _extract_first_last(norm_name: str) -> tuple:
    """
    From a normalised name string, extract:
      - first token  (first name)
      - last_key     (last name, compound-aware)
      - middle_tokens (everything in between)

    Handles compound Filipino/Spanish surnames that begin with a
    particle: DELA, DE, DEL, LOS, LAS, SAN, SANTA, VAN, VON, etc.
    When the second-to-last token is a known particle the last TWO
    tokens are joined as the surname key, e.g.:
        ['JUAN', 'DELA', 'CRUZ']  ->  first='JUAN', last='DELA CRUZ'
        ['MARIA', 'DE', 'LOS', 'SANTOS'] ->  first='MARIA', last='DE LOS SANTOS'  (*)

    (*) 3-token particles: the loop walks backwards collecting
        consecutive particle tokens so DE LOS SANTOS is fully captured.

    Returns: (first, last_key, middle_tokens)
    """
    PARTICLES = {
        "DE", "DEL", "DELA", "LOS", "LAS", "SAN", "SANTA",
        "VAN", "VON", "DER", "DEN", "TEN", "TER", "LE", "LA",
    }

    tokens = norm_name.split()
    if not tokens:
        return ("", "", [])
    if len(tokens) == 1:
        return (tokens[0], tokens[0], [])

    first = tokens[0]

    # Walk backwards from the end collecting consecutive particle tokens
    # so that multi-word surnames like DE LOS SANTOS are fully captured.
    surname_start = len(tokens) - 1          # index of last token
    while surname_start > 1 and tokens[surname_start - 1] in PARTICLES:
        surname_start -= 1

    last_key = " ".join(tokens[surname_start:])  # e.g. "DELA CRUZ"
    middle   = tokens[1:surname_start]           # tokens between first and surname

    return (first, last_key, middle)


def _first_last_match(norm_a: str, norm_b: str) -> bool:
    """
    Returns True if two normalised names refer to the same person,
    defined as: same first name AND same compound-aware last name,
    regardless of middle name / initial presence.

    Examples that return True:
      JUAN DELA CRUZ          <-> JUAN ISIDRO DELA CRUZ
      JUAN DELA CRUZ          <-> JUAN I DELA CRUZ          (initial stripped by normalise)
      DELA CRUZ JUAN ISIDRO   <-> JUAN DELA CRUZ            (post comma-flip)
      MARIA DE LOS SANTOS     <-> MARIA CLARA DE LOS SANTOS

    Examples that return False:
      JUAN DELA CRUZ          <-> PEDRO DELA CRUZ           (different first)
      JUAN DELA CRUZ          <-> JUAN DELA SANTOS          (different last)
    """
    first_a, last_a, _ = _extract_first_last(norm_a)
    first_b, last_b, _ = _extract_first_last(norm_b)

    if not first_a or not last_a or not first_b or not last_b:
        return False

    return first_a == first_b and last_a == last_b


def _db_candidates_all():
    with _db_connect() as conn:
        return conn.execute(
            "SELECT id, applicant_name, processed_at FROM applicants "
            "WHERE applicant_name IS NOT NULL ORDER BY processed_at DESC"
        ).fetchall()


# ── UPDATED FUNCTION ──────────────────────────────────────────────────
# The change: first+last match now takes priority over fuzzy scoring.
def _resolve_name_similarity(client_name: str,
                              threshold: float = SIMILARITY_THRESHOLD) -> tuple:
    candidates = _db_candidates_all()
    needle     = _normalise_for_sim(client_name)
    scored     = []

    for c in candidates:
        hay   = _normalise_for_sim(c["applicant_name"])
        score = _similarity(needle, hay)

        # ── Primary: first + last name match regardless of middle ──────
        # Handles: "JUAN DELA CRUZ"        <-> "JUAN ISIDRO DELA CRUZ"
        #          "JUAN IBIS DELA CRUZ"   <-> "JUAN DELA CRUZ"
        #          "DELA CRUZ, JUAN"       <-> "JUAN DELA CRUZ"    (post-flip)
        #          "DELA CRUZ, JUAN ISIDRO" <-> "JUAN DELA CRUZ"
        #          "DELA CRUZ, JUAN I."    <-> "JUAN DELA CRUZ"
        if _first_last_match(needle, hay):
            floored = max(score, threshold)
            scored.append((floored, c["id"], c["applicant_name"]))
            continue  # already matched — skip fuzzy check to avoid double-add

        # ── Secondary: fuzzy score meets threshold ────────────────────
        if score >= threshold:
            scored.append((score, c["id"], c["applicant_name"]))

    if not scored:
        return [], ""

    top  = max(s for s, _, __ in scored)
    hits = [(rid, rname) for score, rid, rname in scored if score >= top - 0.02]

    label = "exact" if top == 1.0 else ("high" if top >= 0.90 else "similar")
    return hits, label


# ═══════════════════════════════════════════════════════════════════════
#  PUBLIC WRITER  (called from lookup_tab.py)
# ═══════════════════════════════════════════════════════════════════════

def db_save_applicant(session_id: str, results: dict):
    """
    Persist one applicant's Look-Up results to SQLite.

    Only data that comes directly from the Look-Up extraction is written
    here. client_id, pn, industry_name, loan_balance, amortized_cost,
    principal_loan, maturity, interest_rate are left as NULL/empty
    intentionally — they are populated later via the import buttons.
    """
    _db_init()
    gate = results.get("_gate_data", {})

    def _items(key):
        return "\n".join(results.get(key, {}).get("items", []))

    def _total(key):
        t = results.get(key, {}).get("total", "")
        if not t:
            return None
        try:
            return float(re.sub(r"[^\d.]", "", t.replace(",", "")))
        except Exception:
            return None

    # ── Net income: prefer explicit Gemini value, fall back to calculated ──
    raw_net = str(results.get("_cfa_net_income", "")).strip()
    net_val = None
    if raw_net:
        try:
            net_val = float(re.sub(r"[^\d.]", "", raw_net.replace(",", "")))
        except Exception:
            pass
    if net_val is None:
        inc = _total("income_remittance") or 0
        biz = _total("cfa_business_expenses") or 0
        hh  = _total("cfa_household_expenses") or 0
        if inc or biz or hh:
            net_val = inc - biz - hh

    petrol    = bool(results.get("cibi_petrol_products",    {}).get("items"))
    transport = bool(results.get("cibi_transport_services", {}).get("items"))

    # ── Slim results_json (exclude internal _ keys) ────────────────────
    slim = {k: v for k, v in results.items()
            if not k.startswith("_") and isinstance(v, dict)}

    results_json_str = json.dumps(slim, ensure_ascii=False)

    # ── amort_history_total derived from the already-parsed results ────
    amort_hist = _parse_amort_history_total(results_json_str)

    row_data = {
        "session_id":        session_id,
        "processed_at":      datetime.now().isoformat(timespec="seconds"),
        "source_file":       results.get("_source_file", ""),
        "status":            "done",
        # ── Look-Up extracted fields ───────────────────────────────────
        "applicant_name":    results.get("_applicant_name", ""),
        "residence_address": gate.get("residence_address", ""),
        "office_address":    gate.get("office_address", ""),
        "income_items":      _items("income_remittance"),
        "income_total":      _total("income_remittance"),
        "business_items":    _items("cfa_business_expenses"),
        "business_total":    _total("cfa_business_expenses"),
        "household_items":   _items("cfa_household_expenses"),
        "household_total":   _total("cfa_household_expenses"),
        "net_income":        net_val,
        "petrol_risk":       1 if petrol    else 0,
        "transport_risk":    1 if transport else 0,
        "results_json":      results_json_str,
        "page_map":          results.get("_page_map", ""),
        "amort_history_total": amort_hist,
        # ── Intentionally NULL — set later by import buttons ──────────
        "amort_current_total": None,
        "client_id":           "",
        "pn":                  "",
        "industry_name":       "",
        "loan_balance":        None,
        "amortized_cost":      None,
        "principal_loan":      None,
        "maturity":            "",
        "interest_rate":       "",
        # Expanded P.Loan fields (initially empty)
        "branch":              "",
        "loan_class_name":     "",
        "product_name":        "",
        "loan_date":           "",
        "term_unit":           "",
        "term":                "",
        "security":            "",
        "release_tag":         "",
        "loan_status":         "",
        "ao_name":             "",
    }
    _db_upsert(session_id, row_data)


# ═══════════════════════════════════════════════════════════════════════
#  TREEVIEW STYLE
# ═══════════════════════════════════════════════════════════════════════

def _apply_tree_style():
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Summary.Treeview",
        background=WHITE, foreground=TXT_NAVY, fieldbackground=WHITE,
        rowheight=36, font=("Segoe UI", 9), borderwidth=0, relief="flat")
    style.configure("Summary.Treeview.Heading",
        background=HDR_BG, foreground=HDR_FG,
        font=("Segoe UI", 9, "bold"), relief="flat", borderwidth=0, padding=(8, 8))
    style.map("Summary.Treeview.Heading",
        background=[("active", "#7AB567")], relief=[("active", "flat")])
    style.map("Summary.Treeview",
        background=[("selected", "#C8E6C9")], foreground=[("selected", NAVY_DEEP)])


# ═══════════════════════════════════════════════════════════════════════
#  PANEL BUILDER
# ═══════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════
#  ADVANCED FILTER DIALOG
# ═══════════════════════════════════════════════════════════════════════

# Columns available in the advanced filter (real DB cols only, user-facing)
_ADV_FILTER_COLS = [
    ("applicant_name",   "Applicant Name"),
    ("client_id",        "Client ID"),
    ("pn",               "PN"),
    ("industry_name",    "Industry Name"),
    ("residence_address","Residence Address"),
    ("office_address",   "Office Address"),
    ("loan_status",      "Loan Status"),
    ("ao_name",          "AO Name"),
    ("branch",           "Branch"),
    ("loan_class_name",  "Loan Class"),
    ("product_name",     "Product Name"),
    ("maturity",         "Maturity"),
    ("interest_rate",    "Interest Rate"),
    ("term_unit",        "Term Unit"),
    ("release_tag",      "Release Tag"),
]


def _open_advanced_filter(self):
    win = tk.Toplevel(self)
    win.title("Advanced Column Filter")
    win.configure(bg=CARD_WHITE)
    win.resizable(True, True)
    win.grab_set()

    p_x = self.winfo_rootx(); p_y = self.winfo_rooty()
    p_w = self.winfo_width(); p_h = self.winfo_height()
    w_w, w_h = 680, 560
    win.geometry(f"{w_w}x{w_h}+{p_x + (p_w-w_w)//2}+{p_y + (p_h-w_h)//2}")
    win.minsize(500, 400)

    # ── Header ────────────────────────────────────────────────────────
    hdr = tk.Frame(win, bg=NAVY_DEEP)
    hdr.pack(fill="x")
    tk.Label(hdr, text="⧉  Advanced Column Filter",
             font=("Segoe UI", 12, "bold"), fg=WHITE, bg=NAVY_DEEP,
             padx=16, pady=10).pack(side="left")
    tk.Label(hdr,
             text="Each row = one column filter. Values in a column are OR-matched.\nDifferent columns are AND-matched.",
             font=("Segoe UI", 8), fg="#8DA8C8", bg=NAVY_DEEP,
             padx=16, justify="left").pack(side="left", pady=8)

    # ── Scrollable body ───────────────────────────────────────────────
    body_outer = tk.Frame(win, bg=CARD_WHITE)
    body_outer.pack(fill="both", expand=True, padx=16, pady=(10, 0))

    canvas  = tk.Canvas(body_outer, bg=CARD_WHITE, highlightthickness=0)
    vscroll = tk.Scrollbar(body_outer, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vscroll.set)
    vscroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)

    rows_frame = tk.Frame(canvas, bg=CARD_WHITE)
    cwin = canvas.create_window((0, 0), window=rows_frame, anchor="nw")

    def _on_cfg(e):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(cwin, width=canvas.winfo_width())
    rows_frame.bind("<Configure>", _on_cfg)
    canvas.bind("<Configure>", _on_cfg)
    canvas.bind_all("<MouseWheel>",
        lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

    # ── Column header labels ───────────────────────────────────────────
    col_hdr = tk.Frame(rows_frame, bg="#E8EEF8")
    col_hdr.pack(fill="x", pady=(0, 4))
    tk.Label(col_hdr, text="Column", font=("Segoe UI", 8, "bold"),
             fg=NAVY_MID, bg="#E8EEF8", width=22, anchor="w",
             padx=8, pady=4).grid(row=0, column=0, sticky="w")
    tk.Label(col_hdr, text="Filter Values  (comma-separated, OR logic)",
             font=("Segoe UI", 8, "bold"),
             fg=NAVY_MID, bg="#E8EEF8", padx=8, pady=4).grid(row=0, column=1, sticky="w")
    tk.Label(col_hdr, text="", bg="#E8EEF8", width=4).grid(row=0, column=2)

    # ── Filter rows (one per column) ──────────────────────────────────
    filter_vars: dict[str, tk.StringVar] = {}

    current = getattr(self, "_sum_adv_filters", {})

    for i, (db_col, label) in enumerate(_ADV_FILTER_COLS):
        row_bg = ROW_BG_EVEN if i % 2 == 0 else ROW_BG_ODD
        row_f  = tk.Frame(rows_frame, bg=row_bg,
                          highlightbackground=BORDER_LIGHT, highlightthickness=1)
        row_f.pack(fill="x", pady=1)
        row_f.columnconfigure(1, weight=1)

        tk.Label(row_f, text=label, font=("Segoe UI", 8, "bold"),
                 fg=NAVY_DEEP, bg=row_bg, width=22, anchor="w",
                 padx=8, pady=6).grid(row=0, column=0, sticky="w")

        # Pre-populate from existing active filters
        existing = ", ".join(current.get(db_col, []))
        var = tk.StringVar(value=existing)
        filter_vars[db_col] = var

        entry = tk.Entry(row_f, textvariable=var, font=("Segoe UI", 9),
                         fg=TXT_NAVY, bg=WHITE if not existing else "#FFFDE7",
                         relief="solid", bd=1,
                         insertbackground=NAVY_MID)
        entry.grid(row=0, column=1, sticky="ew", padx=(0, 4), pady=5)

        # Yellow bg when typed, white when empty
        def _on_var_change(v=var, e=entry):
            e.config(bg="#FFFDE7" if v.get().strip() else WHITE)
        var.trace_add("write", lambda *_, v=var, e=entry: _on_var_change(v, e))

        # Clear button per row
        def _clear_row(v=var):
            v.set("")
        tk.Button(row_f, text="✕", font=("Segoe UI", 7),
                  fg=TXT_MUTED, bg=row_bg, activebackground=ACCENT_RED,
                  activeforeground=WHITE, relief="flat", bd=0,
                  padx=6, pady=4, cursor="hand2",
                  command=_clear_row).grid(row=0, column=2, padx=(0, 4))

    # ── Hint label ────────────────────────────────────────────────────
    hint = tk.Frame(win, bg="#F0F4FA",
                    highlightbackground=BORDER_MID, highlightthickness=1)
    hint.pack(fill="x", padx=16, pady=(6, 0))
    tk.Label(hint,
             text='💡  Example: Industry Name → "Agriculture, Wholesale"  '
                  'shows rows where industry contains Agriculture OR Wholesale.',
             font=("Segoe UI", 8), fg=TXT_MUTED, bg="#F0F4FA",
             pady=5, padx=10, anchor="w").pack(fill="x")

    # ── Button bar ────────────────────────────────────────────────────
    btn_bar = tk.Frame(win, bg=CARD_WHITE,
                       highlightbackground=BORDER_MID, highlightthickness=1)
    btn_bar.pack(fill="x", padx=16, pady=(6, 12))

    active_count_var = tk.StringVar()

    def _count_active():
        n = sum(1 for v in filter_vars.values() if v.get().strip())
        active_count_var.set(
            f"{n} column filter(s) active" if n else "No filters active")

    for v in filter_vars.values():
        v.trace_add("write", lambda *_: _count_active())
    _count_active()

    tk.Label(btn_bar, textvariable=active_count_var,
             font=("Segoe UI", 8), fg=LIME_DARK, bg=CARD_WHITE,
             padx=10).pack(side="left", pady=8)

    def _clear_all_filters():
        for v in filter_vars.values():
            v.set("")

    def _apply():
        new_filters = {}
        for db_col, var in filter_vars.items():
            raw = var.get().strip()
            if raw:
                # Split on comma, strip each, drop empties
                values = [t.strip() for t in raw.split(",") if t.strip()]
                if values:
                    new_filters[db_col] = values
        self._sum_adv_filters = new_filters
        self._sum_page        = 0

        # Update the indicator label on the main panel
        n = len(new_filters)
        if n:
            cols_used = ", ".join(
                label for db_col, label in _ADV_FILTER_COLS
                if db_col in new_filters)
            self._sum_adv_filter_lbl.config(
                text=f"⧉ {n} filter(s): {cols_used}")
            self._sum_adv_filter_btn.configure(
                fg_color=LIME_MID, text_color=TXT_ON_LIME)
        else:
            self._sum_adv_filter_lbl.config(text="")
            self._sum_adv_filter_btn.configure(
                fg_color="#2A3A6C", text_color=WHITE)

        _load_and_render(self)
        canvas.unbind_all("<MouseWheel>")
        win.destroy()

    def _on_cancel():
        canvas.unbind_all("<MouseWheel>")
        win.destroy()

    tk.Button(btn_bar, text="Clear All", font=("Segoe UI", 8),
              fg=TXT_SOFT, bg="#F0F0F0", activebackground="#E0E0E0",
              relief="flat", bd=0, padx=12, pady=6, cursor="hand2",
              command=_clear_all_filters).pack(side="right", padx=(4, 8), pady=6)
    tk.Button(btn_bar, text="✕  Cancel", font=("Segoe UI", 8),
              fg=TXT_SOFT, bg="#F0F0F0", activebackground="#E0E0E0",
              relief="flat", bd=0, padx=12, pady=6, cursor="hand2",
              command=_on_cancel).pack(side="right", padx=(0, 4), pady=6)
    ctk.CTkButton(btn_bar, text="✔  Apply Filters", command=_apply,
                  width=120, height=30, corner_radius=6,
                  fg_color=LIME_MID, hover_color=LIME_BRIGHT,
                  text_color=TXT_ON_LIME,
                  font=FF(8, "bold")).pack(side="right", padx=(0, 4), pady=6)

    win.protocol("WM_DELETE_WINDOW", _on_cancel)

def _build_lookup_summary_panel(self, parent):
    _db_init()
    _apply_tree_style()

    outer = tk.Frame(parent, bg=CARD_WHITE)
    self._lookup_summary_frame = outer
    main = tk.Frame(outer, bg=CARD_WHITE)
    main.pack(fill="both", expand=True)

    # ── Header band ───────────────────────────────────────────────────
    header_band = tk.Frame(main, bg=NAVY_DEEP)
    header_band.pack(fill="x", padx=PAD, pady=(16, 0))
    title_block = tk.Frame(header_band, bg=NAVY_DEEP)
    title_block.pack(side="left", padx=16, pady=12)
    tk.Label(title_block, text="Look-Up Summary",
             font=F(15, "bold"), fg=WHITE, bg=NAVY_DEEP).pack(anchor="w")
    tk.Label(title_block,
             text="Persistent across sessions  ·  SQLite backed  ·  live view",
             font=F(8), fg="#8DA8C8", bg=NAVY_DEEP).pack(anchor="w", pady=(2, 0))

    btn_block = tk.Frame(header_band, bg=NAVY_DEEP)
    btn_block.pack(side="right", padx=12, pady=10)

    self._sum_export_csv_btn = ctk.CTkButton(
        btn_block, text="⬇  CSV", command=lambda: _export_csv(self),
        width=68, height=30, corner_radius=6, fg_color="transparent",
        hover_color="#1E3A5F", text_color="#8DA8C8", font=FF(8, "bold"),
        border_width=1, border_color="#2E4E72")
    self._sum_export_csv_btn.pack(side="left", padx=(0, 4))

    self._sum_export_xl_btn = ctk.CTkButton(
        btn_block, text="📊  Excel", command=lambda: _export_excel(self),
        width=74, height=30, corner_radius=6, fg_color=LIME_MID,
        hover_color=LIME_BRIGHT, text_color=TXT_ON_LIME,
        font=FF(8, "bold"), border_width=0)
    self._sum_export_xl_btn.pack(side="left", padx=(0, 4))

    self._sum_refresh_btn = ctk.CTkButton(
        btn_block, text="↺  Refresh", command=lambda: _refresh_summary(self),
        width=74, height=30, corner_radius=6, fg_color="#1A3A5C",
        hover_color="#1E4A72", text_color=WHITE,
        font=FF(8, "bold"), border_width=0)
    self._sum_refresh_btn.pack(side="left", padx=(0, 4))

    self._sum_clear_all_btn = ctk.CTkButton(
        btn_block, text="🗑  Clear All", command=lambda: _clear_all(self),
        width=80, height=30, corner_radius=6, fg_color="#3D1010",
        hover_color="#5C1A1A", text_color="#FF8A80",
        font=FF(8, "bold"), border_width=0)
    self._sum_clear_all_btn.pack(side="left")

    self._sum_import_amort_btn = ctk.CTkButton(
        btn_block, text="⬆  Amort.", command=lambda: _import_amort_file(self),
        width=74, height=30, corner_radius=6, fg_color="#1A3A5C",
        hover_color="#1E4A72", text_color=WHITE,
        font=FF(8, "bold"), border_width=0)
    self._sum_import_amort_btn.pack(side="left", padx=(4, 0))

    self._sum_import_other_btn = ctk.CTkButton(
        btn_block, text="⬆  Other Data", command=lambda: _import_other_data_file(self),
        width=88, height=30, corner_radius=6, fg_color="#2A1A5C",
        hover_color="#3A2472", text_color="#C8B8FF",
        font=FF(8, "bold"), border_width=0)
    self._sum_import_other_btn.pack(side="left", padx=(4, 0))

    self._sum_import_ploan_btn = ctk.CTkButton(
        btn_block, text="⬆  P.Loan", command=lambda: _import_ploan_file(self),
        width=74, height=30, corner_radius=6, fg_color="#1A4A3C",
        hover_color="#256050", text_color="#A0FFD8",
        font=FF(8, "bold"), border_width=0)
    self._sum_import_ploan_btn.pack(side="left", padx=(4, 0))

    self._sum_merge_db_btn = ctk.CTkButton(
        btn_block, text="⛁  Merge DB", command=lambda: _merge_db_files(self),
        width=80, height=30, corner_radius=6, fg_color="#2D4A1E",
        hover_color="#3D6128", text_color="#B9F5A0",
        font=FF(8, "bold"), border_width=0)
    self._sum_merge_db_btn.pack(side="left", padx=(4, 0))

    self._sum_merge_xl_btn = ctk.CTkButton(
        btn_block, text="⛁  Merge Excel", command=lambda: _merge_excel_files(self),
        width=90, height=30, corner_radius=6, fg_color="#1E3D4A",
        hover_color="#255262", text_color="#A0E4F5",
        font=FF(8, "bold"), border_width=0)
    self._sum_merge_xl_btn.pack(side="left", padx=(4, 0))

    self._sum_dedup_btn = ctk.CTkButton(
        btn_block, text="🔗  Dedup", command=lambda: _run_dedup(self),
        width=72, height=30, corner_radius=6, fg_color="#4A2D1E",
        hover_color="#623D28", text_color="#FFD0A0",
        font=FF(8, "bold"), border_width=0)
    self._sum_dedup_btn.pack(side="left", padx=(4, 0))

    self._sum_validate_btn = ctk.CTkButton(
        btn_block, text="✔  Validate", command=lambda: _validate_clients(self),
        width=80, height=30, corner_radius=6, fg_color="#1A4A2A",
        hover_color="#256035", text_color="#A0FFB8",
        font=FF(8, "bold"), border_width=0)
    self._sum_validate_btn.pack(side="left", padx=(4, 0))

    # ── Controls row ──────────────────────────────────────────────────
    controls_row = tk.Frame(main, bg="#F0F4FA",
                            highlightbackground=BORDER_MID, highlightthickness=1)
    controls_row.pack(fill="x", padx=PAD)

    stats_group = tk.Frame(controls_row, bg="#F0F4FA")
    stats_group.pack(side="left", padx=10, pady=8)
    self._sum_stat_labels = {}
    for key, label, color, pill_bg in [
        ("total",         "Total",          NAVY_DEEP,      "#E8EEF8"),
        ("done",          "Done",           ACCENT_SUCCESS, "#F0FDF4"),
        ("errors",        "Errors",         ACCENT_RED,     "#FFF0F0"),
        ("income",        "Income",         NAVY_MID,       "#EEF3FA"),
        ("net",           "Net",            NET_GREEN,      "#F0FDF4"),
        ("amort_current", "Curr. Amort.",   ACCENT_GOLD,    "#FFFBF0"),
    ]:
        pill = tk.Frame(stats_group, bg=pill_bg,
                        highlightbackground=BORDER_MID, highlightthickness=1)
        pill.pack(side="left", padx=(0, 6))
        tk.Label(pill, text=label, font=F(7, "bold"), fg=TXT_MUTED,
                 bg=pill_bg, padx=10, pady=3).pack()
        lbl = tk.Label(pill, text="—", font=F(11, "bold"),
                       fg=color, bg=pill_bg, padx=10, pady=4)
        lbl.pack()
        self._sum_stat_labels[key] = lbl

    tk.Frame(controls_row, bg=BORDER_MID, width=1).pack(
        side="left", fill="y", pady=6, padx=4)

    search_wrap = tk.Frame(controls_row, bg=WHITE,
                           highlightbackground=BORDER_MID, highlightthickness=1)
    search_wrap.pack(side="left", fill="x", expand=True, padx=8, pady=8)
    tk.Label(search_wrap, text="🔍", font=("Segoe UI Emoji", 9),
             bg=WHITE, fg=NAVY_PALE).pack(side="left", padx=(8, 2))
    self._sum_search_var = tk.StringVar()
    self._sum_search_var.trace_add("write", lambda *a: _on_search_change(self))
    _HINT = "Search all fields… separate terms with commas for AND filtering"
    _se = tk.Entry(search_wrap, textvariable=self._sum_search_var,
                   font=F(9), fg=TXT_MUTED, bg=WHITE, relief="flat", bd=0,
                   insertbackground=NAVY_MID)
    _se.pack(side="left", fill="x", expand=True, pady=6)
    _se.insert(0, _HINT)

    def _se_focus_in(e):
        if self._sum_search_var.get() == _HINT:
            _se.delete(0, "end"); _se.config(fg=TXT_NAVY)

    def _se_focus_out(e):
        if not self._sum_search_var.get().strip():
            _se.config(fg=TXT_MUTED); _se.insert(0, _HINT)

    _se.bind("<FocusIn>",  _se_focus_in)
    _se.bind("<FocusOut>", _se_focus_out)

    self._sum_adv_filter_btn = ctk.CTkButton(
        controls_row, text="⧉  Filter",
        command=lambda: _open_advanced_filter(self),
        width=74, height=30, corner_radius=6,
        fg_color="#2A3A6C", hover_color="#3A4A8C",
        text_color=WHITE, font=FF(8, "bold"), border_width=0)
    self._sum_adv_filter_btn.pack(side="left", padx=(4, 0), pady=8)

    self._sum_adv_filter_lbl = tk.Label(
        controls_row, text="", font=F(7), fg=LIME_MID, bg="#F0F4FA")
    self._sum_adv_filter_lbl.pack(side="left", padx=(2, 0), pady=8)

    right_ctrl = tk.Frame(controls_row, bg="#F0F4FA")
    right_ctrl.pack(side="right", padx=10, pady=8)
    self._sum_count_lbl = tk.Label(right_ctrl, text="",
                                    font=F(8), fg=TXT_SOFT, bg="#F0F4FA")
    self._sum_count_lbl.pack(side="top", anchor="e")
    pg_sub = tk.Frame(right_ctrl, bg="#F0F4FA")
    pg_sub.pack(side="top", pady=(4, 0))
    self._sum_prev_btn = ctk.CTkButton(
        pg_sub, text="◀", command=lambda: _page_prev(self),
        width=30, height=24, corner_radius=5, fg_color=CARD_WHITE,
        hover_color=NAVY_MIST, text_color=NAVY_MID, font=FF(8, "bold"),
        border_width=1, border_color=BORDER_MID, state="disabled")
    self._sum_prev_btn.pack(side="left", padx=(0, 4))
    self._sum_page_lbl = tk.Label(pg_sub, text="Page 1",
                                  font=F(8), fg=TXT_SOFT, bg="#F0F4FA")
    self._sum_page_lbl.pack(side="left", padx=4)
    self._sum_next_btn = ctk.CTkButton(
        pg_sub, text="▶", command=lambda: _page_next(self),
        width=30, height=24, corner_radius=5, fg_color=CARD_WHITE,
        hover_color=NAVY_MIST, text_color=NAVY_MID, font=FF(8, "bold"),
        border_width=1, border_color=BORDER_MID, state="disabled")
    self._sum_next_btn.pack(side="left", padx=(4, 0))

    # ── Treeview ──────────────────────────────────────────────────────
    tbl_outer = tk.Frame(main, bg=BORDER_LIGHT)
    tbl_outer.pack(fill="both", expand=True, padx=PAD, pady=(8, PAD))
    tbl_wrap = tk.Frame(tbl_outer, bg=CARD_WHITE)
    tbl_wrap.pack(fill="both", expand=True, padx=1, pady=1)
    tbl_wrap.rowconfigure(0, weight=1); tbl_wrap.columnconfigure(0, weight=1)

    vscroll = tk.Scrollbar(tbl_wrap, orient="vertical", relief="flat",
                           troughcolor=OFF_WHITE, bg=BORDER_LIGHT, width=8, bd=0)
    vscroll.grid(row=0, column=1, sticky="ns")
    hscroll = tk.Scrollbar(tbl_wrap, orient="horizontal", relief="flat",
                           troughcolor=OFF_WHITE, bg=BORDER_LIGHT, bd=0)
    hscroll.grid(row=1, column=0, columnspan=2, sticky="ew")

    self._sum_tree = ttk.Treeview(
        tbl_wrap, columns=TREE_COLS, show="headings",
        style="Summary.Treeview",
        yscrollcommand=vscroll.set, xscrollcommand=hscroll.set,
        selectmode="browse")
    self._sum_tree.grid(row=0, column=0, sticky="nsew")
    vscroll.config(command=self._sum_tree.yview)
    hscroll.config(command=self._sum_tree.xview)

    for db_col, label, width_px, is_mon, is_txt in TABLE_COLS:
        anchor = "e" if is_mon else "w"
        self._sum_tree.heading(db_col, text=label,
                               command=lambda c=db_col: _sort_by(self, c))
        self._sum_tree.column(db_col, width=width_px, minwidth=60,
                              anchor=anchor, stretch=False)

    self._sum_tree.tag_configure("even", background=ROW_BG_EVEN)
    self._sum_tree.tag_configure("odd",  background=ROW_BG_ODD)

    # ── CHANGE 4 (replace the single <Double-1> bind with two lines) ──
    self._sum_tree.bind("<Double-1>", lambda e: _on_tree_double_click(self, e))
    self._sum_tree.bind("<Return>",   lambda e: _on_tree_return_key(self, e))

    self._sum_tree.bind("<Button-3>", lambda e: _on_tree_right_click(self, e))
    self._sum_tree.bind("<Enter>",
        lambda e: self._sum_tree.bind_all("<MouseWheel>",
            lambda ev: self._sum_tree.yview_scroll(int(-1*(ev.delta/120)), "units")))
    self._sum_tree.bind("<Leave>", lambda e: self._sum_tree.unbind_all("<MouseWheel>"))

    self._sum_sort_col       = "processed_at"
    self._sum_sort_asc       = False
    self._sum_page           = 0
    self._sum_total_rows     = 0
    self._sum_session_filter = ""
    self._sum_search_after   = None
    self._sum_row_data       = {}
    self._sum_adv_filters    = {}
    _refresh_summary(self)


# ═══════════════════════════════════════════════════════════════════════
#  DATA LOADING + RENDERING
# ═══════════════════════════════════════════════════════════════════════

def _on_search_change(self):
    if "separate terms with commas" in self._sum_search_var.get():
        return
    if self._sum_search_after:
        self.after_cancel(self._sum_search_after)
    self._sum_search_after = self.after(300, lambda: _load_and_render(self))


def _refresh_summary(self):
    _load_and_render(self)


def _load_and_render(self):
    raw    = self._sum_search_var.get().strip()
    search = "" if "separate terms with commas" in raw else raw
    offset = self._sum_page * PAGE_SIZE
    rows, total = _db_query(
        search=search, session_id=self._sum_session_filter,
        sort_col=self._sum_sort_col, sort_asc=self._sum_sort_asc,
        offset=offset, limit=PAGE_SIZE,
        adv_filters=getattr(self, "_sum_adv_filters", {}))
    self._sum_total_rows = total
    _update_stats(self)
    _update_pagination(self, total)
    _render_tree(self, rows)
    shown_start = offset + 1 if total > 0 else 0
    shown_end   = min(offset + PAGE_SIZE, total)
    self._sum_count_lbl.config(
        text=f"{shown_start}–{shown_end} of {total} applicant(s)")


def _update_stats(self):
    raw    = self._sum_search_var.get().strip()
    search = "" if "separate terms with commas" in raw else raw
    tots   = _db_totals(session_id=self._sum_session_filter, search=search,
                        adv_filters=getattr(self, "_sum_adv_filters", {}))

    def _fmt(val):
        if val is None:
            return "—"
        try:
            return f"P{float(val):,.0f}"
        except Exception:
            return "—"

    self._sum_stat_labels["total"].config(text=str(tots.get("total", 0)))
    self._sum_stat_labels["done"].config(text=str(tots.get("done", 0)))
    self._sum_stat_labels["errors"].config(text=str(tots.get("errors", 0)))
    self._sum_stat_labels["income"].config(text=_fmt(tots.get("income")))
    self._sum_stat_labels["net"].config(text=_fmt(tots.get("net")))
    self._sum_stat_labels["amort_current"].config(text=_fmt(tots.get("amort_current")))


def _update_pagination(self, total: int):
    total_pages = max(1, -(-total // PAGE_SIZE))
    cur_page    = self._sum_page + 1
    self._sum_page_lbl.config(text=f"Page {cur_page} / {total_pages}")
    self._sum_prev_btn.configure(
        state="normal" if self._sum_page > 0 else "disabled")
    self._sum_next_btn.configure(
        state="normal" if cur_page < total_pages else "disabled")


def _page_prev(self):
    if self._sum_page > 0:
        self._sum_page -= 1; _load_and_render(self)


def _page_next(self):
    total_pages = max(1, -(-self._sum_total_rows // PAGE_SIZE))
    if self._sum_page + 1 < total_pages:
        self._sum_page += 1; _load_and_render(self)


def _fmt_money(val) -> str:
    if val in (None, ""):
        return "—"
    try:
        return f"P{float(val):,.2f}"
    except Exception:
        return str(val) or "—"


def _render_tree(self, rows):
    self._sum_tree.delete(*self._sum_tree.get_children())
    self._sum_row_data = {}

    for i, row in enumerate(rows):
        row    = dict(row)
        row_id = row.get("id")

        # ── Parse results_json once for all virtual fields ─────────────
        try:
            results_blob = json.loads(row.get("results_json", "") or "{}")
        except Exception:
            results_blob = {}

        # ── Virtual columns — derived at render time ───────────────────
        row["spouse_info"]        = _extract_spouse_info(results_blob)
        row["personal_assets"]    = _extract_asset_items(results_blob, "cibi_personal_assets")
        row["business_assets"]    = _extract_asset_items(results_blob, "cibi_business_assets")
        row["business_inventory"] = _extract_asset_items(results_blob, "cibi_business_inventory")

        # ── amort_history_total: use real DB value; backfill if missing ─
        if row.get("amort_history_total") is None and results_blob:
            row["amort_history_total"] = _parse_amort_history_total(
                row.get("results_json", ""))

        tag    = "even" if i % 2 == 0 else "odd"
        values = []
        for db_col, label, width_px, is_monetary, is_text_block in TABLE_COLS:
            raw = row.get(db_col, "") or ""
            if is_monetary:
                values.append(_fmt_money(raw))
            elif is_text_block:
                values.append(str(raw).replace("\n", "  ·  "))
            else:
                values.append(str(raw))
        self._sum_tree.insert("", "end", iid=str(row_id),
                              values=values, tags=(tag,))
        self._sum_row_data[str(row_id)] = row


def _sort_by(self, col_key: str):
    # ── Virtual columns cannot be sorted by DB; fall back to name ─────
    if col_key in _VIRTUAL_COLS:
        col_key = "applicant_name"

    if self._sum_sort_col == col_key:
        self._sum_sort_asc = not self._sum_sort_asc
    else:
        self._sum_sort_col = col_key
        self._sum_sort_asc = True

    self._sum_page = 0
    for db_col, label, _, __, ___ in TABLE_COLS:
        active = (self._sum_sort_col == db_col)
        ind = (" ▲" if active and self._sum_sort_asc else
               " ▼" if active else "")
        self._sum_tree.heading(db_col, text=label + ind)
    _load_and_render(self)


# ═══════════════════════════════════════════════════════════════════════
#  ROW INTERACTION (including CHANGE 3 + CHANGE 5)
# ═══════════════════════════════════════════════════════════════════════

# ── CHANGE 3 (add _start_cell_edit) ───────────────────────────────────
def _start_cell_edit(self, iid: str, col_id: str):
    """
    Pop a lightweight Entry widget directly over the clicked Treeview
    cell so the user can type a new value.

    Behaviour:
      <Return> or <KP_Enter>  — commit and close
      <Tab>                   — commit and move to next editable column
      <Escape>                — cancel, restore original value
      FocusOut                — commit and close
    """
    if col_id not in _EDITABLE_COLS:
        return

    tree   = self._sum_tree
    row_id = int(iid)

    # ── Get the cell's pixel bounding box ─────────────────────────────
    bbox = tree.bbox(iid, col_id)
    if not bbox:
        # Row is off-screen — scroll it into view first, then retry once
        tree.see(iid)
        tree.update_idletasks()
        bbox = tree.bbox(iid, col_id)
        if not bbox:
            return
    x, y, width, height = bbox

    # ── Current raw value ─────────────────────────────────────────────
    col_index   = TREE_COLS.index(col_id)
    cur_display = tree.item(iid, "values")[col_index]

    if col_id in _MONETARY_COLS:
        # Strip "P" prefix, commas, and "—" placeholder
        cur_edit = re.sub(r"[P,]", "", cur_display).strip()
        if cur_edit == "—":
            cur_edit = ""
    elif col_id in _VIRTUAL_TO_JSON or col_id in {
            "income_items", "business_items", "household_items"}:
        # Tree renders newlines as '  ·  '; restore for editing
        cur_edit = cur_display.replace("  ·  ", "\n") \
            if cur_display not in ("—", "") else ""
    else:
        cur_edit = "" if cur_display in ("—", "") else cur_display

    # ── Build the overlay Entry ────────────────────────────────────────
    var   = tk.StringVar(value=cur_edit)
    entry = tk.Entry(
        tree,
        textvariable=var,
        font=("Segoe UI", 9),
        fg=TXT_NAVY,
        bg="#FFFDE7",                   # soft yellow — "edit mode" signal
        insertbackground=NAVY_MID,
        relief="solid",
        bd=1,
        highlightthickness=1,
        highlightbackground=LIME_MID,   # green border, matches app theme
        highlightcolor=LIME_BRIGHT,
    )
    entry.place(x=x, y=y, width=max(width, 120), height=height)
    entry.focus_set()
    entry.select_range(0, "end")

    # Pause mousewheel on the treeview while the Entry is open so the
    # row cannot scroll out from under the widget.
    tree.unbind_all("<MouseWheel>")

    committed = [False]   # guard against double-commit on FocusOut + Return

    def _commit(event=None):
        if committed[0]:
            return
        committed[0] = True
        new_raw = var.get()
        try:
            if col_id in _VIRTUAL_TO_JSON:
                display_val = _db_update_virtual_cell(row_id, col_id, new_raw)
            elif col_id in {"income_items", "business_items", "household_items"}:
                display_val = _db_update_cell(row_id, col_id, new_raw)
            else:
                display_val = _db_update_cell(row_id, col_id, new_raw)
        except ValueError as exc:
            committed[0] = False   # allow re-edit after validation error
            entry.config(bg="#FFE0E0", highlightbackground=ACCENT_RED)
            messagebox.showerror("Invalid Value", str(exc))
            entry.focus_set()
            return

        # Patch the in-memory row cache so the detail window stays current
        if iid in self._sum_row_data:
            if col_id in _MONETARY_COLS:
                cleaned = re.sub(r"[^\d.]", "",
                                 new_raw.replace(",", "")).strip()
                self._sum_row_data[iid][col_id] = \
                    float(cleaned) if cleaned else None
            elif col_id in _VIRTUAL_TO_JSON:
                # Virtual cols live in results_json; update the cache directly
                # so the detail window reflects the edit without a full reload.
                self._sum_row_data[iid][col_id] = new_raw.strip() or None
                # Also keep results_json in sync so _extract_* helpers work
                # if the detail window re-reads from the row dict.
                try:
                    blob = json.loads(
                        self._sum_row_data[iid].get("results_json") or "{}")
                    items = [p.strip()
                             for p in new_raw.strip().split("  ·  ")
                             if p.strip()]
                    if col_id == "spouse_info":
                        plain  = [it for it in items
                                  if not it.startswith("Office: ")]
                        office = [it[len("Office: "):] for it in items
                                  if it.startswith("Office: ")]
                        blob.setdefault("cibi_spouse", {})["items"]        = plain
                        blob.setdefault("cibi_spouse_office", {})["items"] = office
                    else:
                        json_key, sub_key = _VIRTUAL_TO_JSON[col_id]
                        blob.setdefault(json_key, {})[sub_key] = items
                    self._sum_row_data[iid]["results_json"] = \
                        json.dumps(blob, ensure_ascii=False)
                except Exception:
                    pass
            else:
                self._sum_row_data[iid][col_id] = new_raw.strip() or None

        # Update just this one cell in the treeview — no full reload
        vals = list(tree.item(iid, "values"))
        vals[col_index] = display_val
        tree.item(iid, values=vals)

        entry.destroy()
        _update_stats(self)   # keep the aggregate stat pills in sync

    def _cancel(event=None):
        committed[0] = True
        entry.destroy()

    def _tab_next(event=None):
        """Commit and jump to the next editable column on the same row."""
        _commit()
        if not committed[0]:
            return   # commit was blocked by a validation error
        # Find the next editable column to the right
        start = col_index + 1
        for next_idx in range(start, len(TREE_COLS)):
            next_col = TREE_COLS[next_idx]
            if next_col in _EDITABLE_COLS:
                self.after(50, lambda c=next_col: _start_cell_edit(self, iid, c))
                break
        return "break"   # suppress default Tab focus behaviour

    entry.bind("<Return>",   _commit)
    entry.bind("<KP_Enter>", _commit)
    entry.bind("<Escape>",   _cancel)
    entry.bind("<Tab>",      _tab_next)
    entry.bind("<FocusOut>", _commit)

    # Restore mousewheel when the entry is destroyed
    entry.bind("<Destroy>", lambda e: tree.bind_all(
        "<MouseWheel>",
        lambda ev: tree.yview_scroll(int(-1 * (ev.delta / 120)), "units")))


# ── CHANGE 5 (replace _on_tree_double_click and add _on_tree_return_key) ──
def _on_tree_double_click(self, event):
    """
    Double-click routing:
      • Editable column   → open inline cell editor
      • Non-editable col  → open the detail window (original behaviour)
    """
    tree   = self._sum_tree
    iid    = tree.identify_row(event.y)
    col_id = tree.identify_column(event.x)   # e.g. "#3"

    if not iid:
        return

    # Convert Treeview "#N" index to db_col name
    try:
        col_index = int(col_id.lstrip("#")) - 1
        db_col    = TREE_COLS[col_index]
    except (ValueError, IndexError):
        db_col = ""

    if db_col in _EDITABLE_COLS:
        _start_cell_edit(self, iid, db_col)
    else:
        # Non-editable or virtual column — show detail window
        row = self._sum_row_data.get(iid)
        if row:
            _open_detail_window(self, row)


def _on_tree_return_key(self, event):
    """
    Pressing <Return> on a focused row opens the inline editor on the
    first editable column (applicant_name), letting power users navigate
    and edit without the mouse.
    """
    iid = self._sum_tree.focus()
    if iid:
        _start_cell_edit(self, iid, "applicant_name")


def _on_tree_right_click(self, event):
    iid = self._sum_tree.identify_row(event.y)
    if not iid:
        return
    self._sum_tree.selection_set(iid)
    self._sum_tree.focus(iid)
    menu = tk.Menu(self._sum_tree, tearoff=0, bg=CARD_WHITE, fg=TXT_NAVY,
                   activebackground=NAVY_MIST, font=("Segoe UI", 9))
    menu.add_command(label="👁  View Details",
                     command=lambda: _open_detail_window(
                         self, self._sum_row_data.get(iid, {})))
    menu.add_separator()
    menu.add_command(label="✕  Delete",
                     command=lambda: _delete_row(self, int(iid)))
    try:
        menu.tk_popup(event.x_root, event.y_root)
    finally:
        menu.grab_release()


def _open_detail_window(self, row: dict):
    win = tk.Toplevel(self)
    win.title(f"Detail — {row.get('applicant_name', '')}")
    win.geometry("780x640")
    win.configure(bg=CARD_WHITE)
    win.grab_set()
    canvas  = tk.Canvas(win, bg=CARD_WHITE, highlightthickness=0)
    vscroll = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vscroll.set)
    vscroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    body = tk.Frame(canvas, bg=CARD_WHITE)
    cwin = canvas.create_window((0, 0), window=body, anchor="nw")

    def _on_cfg(e):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(cwin, width=canvas.winfo_width())

    body.bind("<Configure>", _on_cfg)
    canvas.bind("<Configure>", _on_cfg)
    canvas.bind_all("<MouseWheel>",
                    lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
    _build_detail_panel(self, row, body)
    ctk.CTkButton(body, text="Close", command=win.destroy,
                  width=100, height=32, corner_radius=7,
                  fg_color=NAVY_LIGHT, hover_color=NAVY_PALE,
                  text_color=WHITE, font=FF(9, "bold")).pack(pady=(12, 16))


def _build_detail_panel(self, row: dict, parent: tk.Frame):
    info_strip = tk.Frame(parent, bg=NAVY_MIST)
    info_strip.pack(fill="x")

    for label, value in [
        ("Client ID",            row.get("client_id",         "—") or "—"),
        ("PN",                   row.get("pn",                "—") or "—"),
        ("Applicant",            row.get("applicant_name",    "—") or "—"),
        ("Industry",             row.get("industry_name",     "—") or "—"),
        ("Residence Address",    row.get("residence_address", "—") or "—"),
        ("Office Address",       row.get("office_address",    "—") or "—"),
        ("Loan Balance",         _fmt_money(row.get("loan_balance"))   or "—"),
        ("Total Amortized Cost", _fmt_money(row.get("amortized_cost")) or "—"),
        ("Principal Loan",       _fmt_money(row.get("principal_loan")) or "—"),
        ("Maturity",             row.get("maturity",          "—") or "—"),
        ("Interest Rate",        row.get("interest_rate",     "—") or "—"),
        ("Branch",               row.get("branch",            "—") or "—"),
        ("Loan Class",           row.get("loan_class_name",   "—") or "—"),
        ("Product Name",         row.get("product_name",      "—") or "—"),
        ("Loan Date",            row.get("loan_date",         "—") or "—"),
        ("Term Unit",            row.get("term_unit",         "—") or "—"),
        ("Term",                 row.get("term",              "—") or "—"),
        ("Security",             row.get("security",          "—") or "—"),
        ("Release Tag",          row.get("release_tag",       "—") or "—"),
        ("Loan Amount",          _fmt_money(row.get("loan_amount"))         or "—"),
        ("Loan Status",          row.get("loan_status",       "—") or "—"),
        ("AO Name",              row.get("ao_name",           "—") or "—"),
        ("Amort. History",       _fmt_money(row.get("amort_history_total")) or "—"),
        ("Curr. Amort.",         _fmt_money(row.get("amort_current_total")) or "—"),
        ("Source File",          row.get("source_file",       "—") or "—"),
        ("Processed At",         (row.get("processed_at", "") or "")[:16].replace("T", "  ")),
        ("Session",              (row.get("session_id",   "") or "")[:19].replace("T", "  ")),
    ]:
        col = tk.Frame(info_strip, bg=NAVY_MIST)
        col.pack(side="left", padx=12, pady=8, anchor="w")
        tk.Label(col, text=label, font=F(7, "bold"),
                 fg=TXT_MUTED, bg=NAVY_MIST).pack(anchor="w")
        tk.Label(col, text=value or "—", font=F(9, "bold"), fg=NAVY_DEEP,
                 bg=NAVY_MIST, wraplength=180, justify="left").pack(anchor="w")

    flags = tk.Frame(info_strip, bg=NAVY_MIST)
    flags.pack(side="right", padx=12, pady=8)
    if row.get("petrol_risk"):
        tk.Label(flags, text="⚠ Petrol Risk", font=F(8, "bold"),
                 fg=ACCENT_RED, bg=NAVY_MIST).pack(anchor="e")
    if row.get("transport_risk"):
        tk.Label(flags, text="⚠ Transport Risk", font=F(8, "bold"),
                 fg=ACCENT_RED, bg=NAVY_MIST).pack(anchor="e")

    tk.Frame(parent, bg=BORDER_MID, height=1).pack(fill="x")

    try:
        results = json.loads(row.get("results_json", "") or "{}")
    except Exception:
        results = {}

    last_section = None
    PAD_X        = 16

    for idx, (key, section, field_label) in enumerate(LOOKUP_ROWS):
        field_data = results.get(key, {})
        items  = field_data.get("items", []) if isinstance(field_data, dict) else []
        total  = field_data.get("total", "") if isinstance(field_data, dict) else ""
        non_m  = key in NON_MONETARY

        if section != last_section:
            last_section = section
            sec_bar = tk.Frame(parent, bg=SEC_BG)
            sec_bar.pack(fill="x", padx=PAD_X, pady=(8, 0))
            tk.Label(sec_bar, text=f"  {section.upper()}",
                     font=F(8, "bold"), fg=SEC_FG, bg=SEC_BG, pady=4).pack(side="left")

        row_bg = ROW_BG_EVEN if idx % 2 == 0 else ROW_BG_ODD
        row_f  = tk.Frame(parent, bg=row_bg,
                          highlightbackground="#E5EAF3", highlightthickness=1)
        row_f.pack(fill="x", padx=PAD_X)
        tk.Label(row_f, text=field_label, font=F(8, "bold"), fg=NAVY_DEEP,
                 bg=row_bg, padx=8, pady=6, anchor="w", width=26).pack(side="left")
        amt_txt = total if (total and not non_m) else "—"
        tk.Label(row_f, text=amt_txt,
                 font=F(9, "bold") if amt_txt != "—" else F(8),
                 fg=NAVY_MID if amt_txt != "—" else TXT_MUTED,
                 bg=row_bg, padx=8, width=14, anchor="e").pack(side="left")
        det_txt = ("\n".join(f"• {it}" for it in items) if items else "No data found")
        tk.Label(row_f, text=det_txt, font=F(8),
                 fg=TXT_NAVY if items else TXT_MUTED,
                 bg=row_bg, padx=8, anchor="w",
                 wraplength=440, justify="left").pack(side="left", fill="x", expand=True)

    page_map = row.get("page_map", "") or ""
    if page_map:
        pm = tk.Frame(parent, bg=CARD_WHITE)
        pm.pack(fill="x", padx=PAD_X, pady=(8, 12))
        tk.Label(pm, text="Page Map:", font=F(7, "bold"),
                 fg=TXT_MUTED, bg=CARD_WHITE).pack(anchor="w")
        tk.Label(pm, text=page_map, font=FMONO(7), fg=TXT_SOFT,
                 bg=CARD_WHITE, justify="left", anchor="w").pack(anchor="w", padx=8)
    else:
        tk.Frame(parent, bg=CARD_WHITE, height=10).pack()


# ═══════════════════════════════════════════════════════════════════════
#  DELETE / CLEAR / DEDUP
# ═══════════════════════════════════════════════════════════════════════

def _delete_row(self, row_id: int):
    if not messagebox.askyesno("Delete Record",
            "Remove this applicant from the database?\n\nThis cannot be undone."):
        return
    _db_delete_row(row_id)
    _refresh_summary(self)


def _clear_all(self):
    totals = _db_totals()
    total  = totals.get("total", 0) or 0
    if total == 0:
        messagebox.showinfo("Clear All", "There are no records to delete.")
        return
    if not messagebox.askyesno(
            "Clear All Records",
            f"This will permanently delete ALL {total} applicant record(s) "
            f"from the database.\n\nThis cannot be undone.\n\nContinue?",
            icon="warning"):
        return
    _db_clear_all()
    _refresh_summary(self)


def _run_dedup(self):
    if not messagebox.askyesno(
            "Deduplicate by Client ID",
            "This will scan the database and merge any rows that share the "
            "same Client ID.\n\n"
            "• The most complete record is kept.\n"
            "• Missing fields are filled in from duplicates.\n"
            "• Duplicate rows are deleted.\n\nContinue?"):
        return

    _flash_btn(self, self._sum_dedup_btn, "⟳  Working…", 60_000)

    def _worker():
        try:
            removed = _db_deduplicate_client_ids()
            self.after(0, lambda: _refresh_summary(self))
            self.after(0, lambda: (
                _flash_btn(self, self._sum_dedup_btn, "✓  Done!", 2500),
                messagebox.showinfo(
                    "Deduplication Complete",
                    (f"Deduplication finished.\n\n"
                     f"✓  Duplicate rows removed : {removed}\n\n"
                     f"All unique Client IDs now have a single consolidated record.")
                    if removed > 0 else
                    "No duplicate Client IDs found — the database is already clean."
                )
            ))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: (
                _flash_btn(self, self._sum_dedup_btn, "✗  Error", 3000),
                messagebox.showerror("Dedup Error", err)
            ))

    threading.Thread(target=_worker, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  AMORTIZATION IMPORT  (CHANGE 1)
# ═══════════════════════════════════════════════════════════════════════

def _import_amort_file(self):
    path = filedialog.askopenfilename(
        title="Import Amortization Values",
        filetypes=[("Excel & CSV files", "*.xlsx *.csv"),
                   ("Excel files", "*.xlsx"),
                   ("CSV files", "*.csv"),
                   ("All files", "*.*")])
    if not path:
        return

    _flash_btn(self, self._sum_import_amort_btn, "⟳  Reading…", 60_000)

    def _worker():
        try:
            if path.lower().endswith(".csv"):
                import csv as _csv
                with open(path, newline="", encoding="utf-8-sig") as f:
                    reader  = _csv.DictReader(f)
                    records = [dict(row) for row in reader]
                all_cols = list(records[0].keys()) if records else []
            else:
                import openpyxl
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                ws = wb.active
                header_row = next(ws.iter_rows(min_row=1, max_row=1), None)
                if header_row is None:
                    raise ValueError("The Excel file appears to be empty.")
                all_cols = [
                    str(cell.value).strip() if cell.value is not None else ""
                    for cell in header_row
                ]
                records = []
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if all(v is None for v in row):
                        continue
                    records.append({
                        all_cols[i]: (str(v).strip() if v is not None else "")
                        for i, v in enumerate(row) if i < len(all_cols)
                    })
                wb.close()

            if not records:
                raise ValueError("No data rows found in the file.")

            def _find_col(cols, *keywords):
                for kw in keywords:
                    kw_norm = re.sub(r"[\s_]", "", kw.lower())
                    for c in cols:
                        c_norm = re.sub(r"[\s_]", "", c.lower())
                        if kw_norm in c_norm:
                            return c
                return None

            # ── Detect required columns ────────────────────────────────
            col_client   = _find_col(all_cols, "applicant", "client", "name")
            col_amort    = _find_col(all_cols,
                                     "monthlypaymentamount", "monthly payment amount",
                                     "monthlypayment", "paymentamount",
                                     "currentamort", "totalcurrentamort", "amort")
            # ── client_id column (optional but preferred) ──────────────
            col_clientid = _find_col(all_cols, "clientid", "client id",
                                     "client_id", "cid")

            missing = []
            if not col_client: missing.append("Applicant / Client")
            if not col_amort:  missing.append("MonthlyPaymentAmount / Amortization")
            if missing:
                raise ValueError(
                    f"Could not detect column(s): {', '.join(missing)}\n\n"
                    f"File has: {', '.join(all_cols)}")

            # ── Aggregate amort values per client ──────────────────────
            aggregated: dict[str, dict] = {}
            bad_rows:   list[tuple]     = []

            for file_row in records:
                client_name = str(file_row.get(col_client)   or "").strip()
                raw_val     = str(file_row.get(col_amort)    or "").strip()
                raw_cid     = str(file_row.get(col_clientid) or "").strip() \
                              if col_clientid else ""

                if not client_name:
                    continue

                try:
                    cleaned   = re.sub(r"[^\d.]", "", raw_val.replace(",", ""))
                    amort_val = float(cleaned) if cleaned else None
                except Exception:
                    amort_val = None

                name_key = client_name.upper()

                if amort_val is None:
                    bad_rows.append((name_key, f"bad value: '{raw_val}'"))
                    continue

                if name_key not in aggregated:
                    aggregated[name_key] = {
                        "name":      client_name,
                        "client_id": raw_cid.upper(),
                        "total":     0.0,
                    }
                if not aggregated[name_key]["client_id"] and raw_cid:
                    aggregated[name_key]["client_id"] = raw_cid.upper()

                aggregated[name_key]["total"] += amort_val

            # ── Build client_id → db row id lookup ─────────────────────
            with _db_connect() as _conn:
                cid_to_dbid: dict[str, int] = {
                    str(r[0]).strip().upper(): r[1]
                    for r in _conn.execute(
                        "SELECT client_id, id FROM applicants "
                        "WHERE client_id IS NOT NULL AND TRIM(client_id) != ''"
                    ).fetchall()
                }

            # ── Match & update ─────────────────────────────────────────
            updated_by_id   = []
            updated_by_name = []
            updated_relaxed = []
            skipped_names   = list(bad_rows)

            for name_key, entry in aggregated.items():
                amort_val   = entry["total"]
                client_name = entry["name"]
                file_cid    = entry["client_id"]

                # PRIMARY: match by client_id
                if file_cid and file_cid in cid_to_dbid:
                    db_id = cid_to_dbid[file_cid]
                    _db_update_amort_current(db_id, amort_val)
                    updated_by_id.append((name_key, db_id))
                    continue

                # FALLBACK: name similarity
                hits, sim_label = _resolve_name_similarity(client_name)
                if hits:
                    _db_update_amort_all(hits, amort_val)
                    if sim_label in ("exact", "high"):
                        updated_by_name.append((name_key, hits[0][1]))
                    else:
                        updated_relaxed.append((name_key, hits[0][1]))
                else:
                    reason = ("client_id not in DB, no name match"
                              if file_cid else "no client_id, no name match")
                    skipped_names.append((name_key, reason))

            self.after(0, lambda: _refresh_summary(self))

            cid_detected = f"'{col_clientid}'" if col_clientid else "not detected"
            msg  = "Amort. import complete.\n\n"
            msg += f"Client ID column      : {cid_detected}\n"
            msg += f"✓  Matched by ID      : {len(updated_by_id)} record(s)\n"
            msg += f"✓  Matched by name    : {len(updated_by_name)} record(s)\n"
            msg += f"–  Skipped            : {len(skipped_names)} row(s)\n"

            if not col_clientid:
                msg += "\nℹ  No Client ID column found in file — used name matching only.\n"

            if updated_relaxed:
                msg += f"\n⚠  {len(updated_relaxed)} matched via relaxed similarity — please verify:\n"
                for file_n, db_id in updated_relaxed[:10]:
                    msg += f"  • File: {file_n}  →  DB id: {db_id}\n"
                if len(updated_relaxed) > 10:
                    msg += f"  … and {len(updated_relaxed) - 10} more\n"

            if skipped_names:
                msg += "\nSkipped rows:\n"
                for name, reason in skipped_names[:10]:
                    msg += f"  • {name}  ({reason})\n"
                if len(skipped_names) > 10:
                    msg += f"  … and {len(skipped_names) - 10} more"

            self.after(0, lambda: (
                _flash_btn(self, self._sum_import_amort_btn, "✓  Done!", 2500),
                messagebox.showinfo("Amort. Import Result", msg)
            ))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: (
                _flash_btn(self, self._sum_import_amort_btn, "✗  Error", 3000),
                messagebox.showerror("Import Error", err)
            ))

    threading.Thread(target=_worker, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  OTHER DATA IMPORT  (PATCHED)
# ═══════════════════════════════════════════════════════════════════════

def _import_other_data_file(self):
    path = filedialog.askopenfilename(
        title="Import Other Data (Client ID / PN / Industry / Loan Balance)",
        filetypes=[("Excel & CSV files", "*.xlsx *.csv"),
                   ("Excel files", "*.xlsx"),
                   ("CSV files", "*.csv"),
                   ("All files", "*.*")])
    if not path:
        return

    _flash_btn(self, self._sum_import_other_btn, "⟳  Reading…", 60_000)

    def _worker():
        try:
            if path.lower().endswith(".csv"):
                import csv as _csv
                with open(path, newline="", encoding="utf-8-sig") as f:
                    reader  = _csv.DictReader(f)
                    records = [dict(row) for row in reader]
                all_cols = list(records[0].keys()) if records else []
            else:
                import openpyxl
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                ws = wb.active
                header_row = next(ws.iter_rows(min_row=1, max_row=1), None)
                if header_row is None:
                    raise ValueError("The Excel file appears to be empty.")
                all_cols = [
                    str(cell.value).strip() if cell.value is not None else ""
                    for cell in header_row
                ]
                records = []
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if all(v is None for v in row):
                        continue
                    records.append({
                        all_cols[i]: (str(v).strip() if v is not None else "")
                        for i, v in enumerate(row) if i < len(all_cols)
                    })
                wb.close()

            if not records:
                raise ValueError("No data rows found in the file.")

            def _find_col(cols, *keywords):
                for kw in keywords:
                    kw_norm = re.sub(r"[\s_\-]", "", kw.lower())
                    for c in cols:
                        c_norm = re.sub(r"[\s_\-]", "", c.lower())
                        if kw_norm == c_norm or kw_norm in c_norm:
                            return c
                return None

            col_name      = _find_col(all_cols, "clientname", "client name",
                                      "applicant", "applicantname", "name")
            col_clientid  = _find_col(all_cols, "clientid", "client id",
                                      "client_id", "cid")
            col_pn        = _find_col(all_cols, "pnid", "pn id", "pn_id",
                                      "pn", "promissorynote")
            col_industry  = _find_col(all_cols, "industryname", "industry name",
                                      "industry_name", "industry")
            col_loanbal   = _find_col(all_cols, "loanbalance", "loan balance",
                                      "loan_balance", "loanbal", "balance")
            col_amortcost = _find_col(all_cols, "amortizedcost", "amortized cost",
                                      "amortized_cost", "amortcost",
                                      "amortised cost", "amortisedcost")

            if not col_name:
                raise ValueError(
                    f"Could not detect a client name column.\n\n"
                    f"File has: {', '.join(all_cols)}")

            if not any([col_clientid, col_pn, col_industry,
                        col_loanbal, col_amortcost]):
                raise ValueError(
                    f"No data columns found (clientid / pnid / industryname / "
                    f"loanbalance / amortizedcost).\n\n"
                    f"File has: {', '.join(all_cols)}")

            # ── B-1: pn_collect keyed on normalised name ───────────────
            # Using _normalise_for_sim ensures "DELA CRUZ, JUAN" and
            # "JUAN DELA CRUZ" accumulate into the same PN bucket.
            pn_collect: dict[str, list] = {}
            if col_pn:
                for file_row in records:
                    client_name = str(file_row.get(col_name) or "").strip()
                    if not client_name:
                        continue
                    pn_val   = str(file_row.get(col_pn) or "").strip()
                    norm_key = _normalise_for_sim(client_name)   # ← normalised
                    if norm_key not in pn_collect:
                        pn_collect[norm_key] = []
                    if pn_val and pn_val not in pn_collect[norm_key]:
                        pn_collect[norm_key].append(pn_val)

            def _agg_numeric(col_key, records, col_name_field, bad_list, label):
                agg = {}
                for file_row in records:
                    client_name = str(
                        file_row.get(col_name_field) or "").strip()
                    if not client_name:
                        continue
                    raw_val  = str(file_row.get(col_key) or "").strip()
                    # ← also normalise numeric-agg keys
                    norm_key = _normalise_for_sim(client_name)
                    try:
                        cleaned = re.sub(r"[^\d.]", "",
                                         raw_val.replace(",", ""))
                        val     = float(cleaned) if cleaned else None
                    except Exception:
                        val = None
                    if val is None:
                        if raw_val:
                            bad_list.append(
                                (norm_key, f"bad {label}: '{raw_val}'"))
                        if norm_key not in agg:
                            agg[norm_key] = None
                    else:
                        agg[norm_key] = (agg.get(norm_key) or 0.0) + val
                return agg

            loan_bal_bad  = []
            amortcost_bad = []
            loan_bal_agg  = (
                _agg_numeric(col_loanbal,   records, col_name,
                             loan_bal_bad,  "loanbalance")
                if col_loanbal   else {})
            amortcost_agg = (
                _agg_numeric(col_amortcost, records, col_name,
                             amortcost_bad, "amortizedcost")
                if col_amortcost else {})

            # ── B-2: dedup keyed on normalised name ────────────────────
            # "DELA CRUZ, JUAN" normalises to "JUAN DELA CRUZ" — same key
            # as a row that already spells it in natural order.
            seen_norm_keys: set[str] = set()
            deduped: list[dict]      = []
            dup_count = 0

            for file_row in records:
                client_name = str(file_row.get(col_name) or "").strip()
                if not client_name:
                    continue
                norm_key = _normalise_for_sim(client_name)
                if norm_key in seen_norm_keys:
                    dup_count += 1
                    continue
                seen_norm_keys.add(norm_key)
                pn_joined = "\n".join(
                    pn_collect.get(norm_key, [])) if col_pn else ""
                deduped.append({
                    "name":           client_name,   # original display name
                    "norm_key":       norm_key,       # normalised form
                    "client_id":      str(
                        file_row.get(col_clientid) or "").strip()
                        if col_clientid else "",
                    "pn_joined":      pn_joined,
                    "industry":       str(
                        file_row.get(col_industry) or "").strip()
                        if col_industry else "",
                    "loan_balance":   loan_bal_agg.get(norm_key)
                                      if col_loanbal   else None,
                    "amortized_cost": amortcost_agg.get(norm_key)
                                      if col_amortcost else None,
                })

            # ── B-3: pre-match by client_id (mirrors Amort import) ─────
            # Build a lookup of client_id → db row id from the current DB.
            updated_by_cid:  list = []
            name_match_queue: list[dict] = []

            if col_clientid:
                with _db_connect() as _conn:
                    cid_to_dbid: dict[str, int] = {
                        str(r[0]).strip().upper(): r[1]
                        for r in _conn.execute(
                            "SELECT client_id, id FROM applicants "
                            "WHERE client_id IS NOT NULL "
                            "AND TRIM(client_id) != ''"
                        ).fetchall()
                    }

                for entry in deduped:
                    file_cid = entry["client_id"].upper()
                    if file_cid and file_cid in cid_to_dbid:
                        db_id = cid_to_dbid[file_cid]
                        rows_written = _db_update_other_data_all(
                            [(db_id, None)],
                            entry["client_id"],
                            entry["pn_joined"],
                            entry["industry"],
                            entry["loan_balance"],
                            entry["amortized_cost"],
                        )
                        if rows_written:
                            updated_by_cid.append(
                                (entry["norm_key"], db_id))
                    else:
                        # No client_id match — queue for name matching
                        name_match_queue.append(entry)
            else:
                # No client_id column in file — everything goes to name match
                name_match_queue = deduped

            # ── Name-similarity matching for remaining entries ─────────
            updated_strict   = []
            updated_relaxed  = []
            skipped_no_match = []

            for entry in name_match_queue:
                hits, sim_label = _resolve_name_similarity(entry["name"])
                if not hits:
                    skipped_no_match.append(
                        entry["norm_key"]); continue
                rows_written = _db_update_other_data_all(
                    hits,
                    entry["client_id"],
                    entry["pn_joined"],
                    entry["industry"],
                    entry["loan_balance"],
                    entry["amortized_cost"],
                )
                if rows_written:
                    if sim_label == "exact":
                        updated_strict.append(
                            (entry["norm_key"], hits[0][1]))
                    else:
                        updated_relaxed.append(
                            (entry["norm_key"], hits[0][1]))

            _db_deduplicate_client_ids()
            self.after(0, lambda: _refresh_summary(self))

            # ── B-4: updated summary message ──────────────────────────
            total_updated = (len(updated_by_cid)
                             + len(updated_strict)
                             + len(updated_relaxed))
            cols_imported = ", ".join(filter(None, [
                "Client ID"            if col_clientid  else "",
                "PN (all)"             if col_pn        else "",
                "Industry Name"        if col_industry  else "",
                "Loan Balance"         if col_loanbal   else "",
                "Total Amortized Cost" if col_amortcost else "",
            ]))
            msg  = "Other Data import complete.\n\n"
            msg += f"Columns imported       : {cols_imported}\n"
            msg += f"✓  Matched by ID       : {len(updated_by_cid)} record(s)\n"
            msg += f"✓  Matched by name     : {len(updated_strict) + len(updated_relaxed)} record(s)\n"
            msg += f"↩  Duplicates skipped  : {dup_count}\n"
            msg += f"–  No DB match         : {len(skipped_no_match)} name(s)\n"
            if loan_bal_bad:
                msg += (f"⚠  Unparseable loan balance values: "
                        f"{len(loan_bal_bad)}\n")
            if amortcost_bad:
                msg += (f"⚠  Unparseable amortized cost values: "
                        f"{len(amortcost_bad)}\n")
            if not col_clientid:
                msg += ("\nℹ  No Client ID column found — "
                        "used name matching only.\n")
            if updated_relaxed:
                msg += (f"\n⚠  {len(updated_relaxed)} matched via "
                        f"similarity — please verify:\n")
                for file_n, db_id in updated_relaxed[:10]:
                    msg += f"  • File: {file_n}  →  DB id: {db_id}\n"
                if len(updated_relaxed) > 10:
                    msg += f"  … and {len(updated_relaxed) - 10} more\n"
            if skipped_no_match:
                msg += "\nNames with no DB match:\n"
                for name in skipped_no_match[:10]:
                    msg += f"  • {name}\n"
                if len(skipped_no_match) > 10:
                    msg += f"  … and {len(skipped_no_match) - 10} more"

            self.after(0, lambda: (
                _flash_btn(self, self._sum_import_other_btn, "✓  Done!", 2500),
                messagebox.showinfo("Other Data Import Result", msg)
            ))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: (
                _flash_btn(self, self._sum_import_other_btn, "✗  Error", 3000),
                messagebox.showerror("Other Data Import Error", err)
            ))

    threading.Thread(target=_worker, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  PRINCIPAL LOAN IMPORT
# ═══════════════════════════════════════════════════════════════════════

def _import_ploan_file(self):
    path = filedialog.askopenfilename(
        title="Import Principal Loan Data",
        filetypes=[("Excel & CSV files", "*.xlsx *.csv"),
                   ("Excel files", "*.xlsx"),
                   ("CSV files", "*.csv"),
                   ("All files", "*.*")])
    if not path:
        return

    _flash_btn(self, self._sum_import_ploan_btn, "⟳  Reading…", 60_000)

    def _worker():
        try:
            # ── 1. Read file ─────────────────────────────────────────────
            if path.lower().endswith(".csv"):
                import csv as _csv
                with open(path, newline="", encoding="utf-8-sig") as f:
                    reader  = _csv.DictReader(f)
                    records = [dict(row) for row in reader]
                all_cols = list(records[0].keys()) if records else []
            else:
                import openpyxl
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                ws = wb.active
                header_row = next(ws.iter_rows(min_row=1, max_row=1), None)
                if header_row is None:
                    raise ValueError("The Excel file appears to be empty.")
                all_cols = [
                    str(cell.value).strip() if cell.value is not None else ""
                    for cell in header_row
                ]
                records = []
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if all(v is None for v in row):
                        continue
                    records.append({
                        all_cols[i]: (str(v).strip() if v is not None else "")
                        for i, v in enumerate(row) if i < len(all_cols)
                    })
                wb.close()

            if not records:
                raise ValueError("No data rows found in the file.")

            # ── 2. Column detection helper (exact match first, partial second) ──
            def _find_col(cols, *keywords):
                # Pass 1: exact match only
                for kw in keywords:
                    kw_norm = re.sub(r"[\s_\-]", "", kw.lower())
                    for c in cols:
                        c_norm = re.sub(r"[\s_\-]", "", c.lower())
                        if kw_norm == c_norm:
                            return c
                # Pass 2: partial match fallback
                for kw in keywords:
                    kw_norm = re.sub(r"[\s_\-]", "", kw.lower())
                    for c in cols:
                        c_norm = re.sub(r"[\s_\-]", "", c.lower())
                        if kw_norm in c_norm:
                            return c
                return None

            # ── 3. Detect columns ────────────────────────────────────────
            col_clientname = _find_col(all_cols,
                                       "clientname", "client name", "client_name",
                                       "applicant", "applicantname", "name")
            if not col_clientname:
                raise ValueError(
                    f"Could not detect a client name column.\n\n"
                    f"File has: {', '.join(all_cols)}")

            # ── Column map: (db_col, is_monetary, keyword hints...) ──────
            # is_monetary=True  → values are summed across multiple loans
            # is_monetary=False → unique values are collected and joined with ", "
            # NOTE: "term" uses only the exact column name to avoid matching "termunit"
            COL_MAP_DEFS = [
            ("client_id",           False, "clientid", "client id", "client_id", "cid"),
            ("pn",                  False, "pnid", "pn id", "pn_id", "pn", "promissorynote"),
            ("branch",              False, "branch", "branchname", "branch name"),
            ("loan_class_name",     False, "loanclassname", "loan class name", "loan class", "loanclass"),
            ("product_name",        False, "productname", "product name", "product_name", "product"),
            ("industry_name",       False, "industryname", "industry name", "industry_name", "industry"),
            ("loan_date",           False, "loandate", "loan date", "loan_date", "dateofrelease", "releasedate"),
            ("maturity",            False, "maturity", "maturitydate", "maturity date", "duedate", "due date"),
            ("interest_rate",       False, "interest", "interestrate", "interest rate", "interest_rate", "rate", "intrate"),
            ("term_unit",           False, "termunit", "term unit", "term_unit", "paymentfrequency", "frequency"),
            ("term",                False, "term"),
            ("security",            False, "security", "collateral", "securitydescription"),
            ("release_tag",         False, "releasetag", "release tag", "release_tag", "tag"),
            ("loan_amount",         True,  "loanamount", "loan amount", "principalloan", "principal loan", "amount"),
            ("loan_balance",        True,  "loanbalance", "loan balance"),
            ("amort_current_total", True,  "ammortization", "amortization"),
            ("loan_status",         False, "loanstatus", "loan status", "status", "accountstatus"),
            ("ao_name",             False, "aoname", "ao name", "ao_name", "accountofficer", "account officer", "ao"),
        ]

            # Detect which file column maps to each db column
            # detected: db_col -> (file_col, is_monetary)
            detected: dict[str, tuple] = {}
            for entry in COL_MAP_DEFS:
                db_col      = entry[0]
                is_monetary = entry[1]
                keywords    = entry[2:]
                file_col    = _find_col(all_cols, *keywords)
                detected[db_col] = (file_col, is_monetary)

            # ── 4. Aggregate all rows per client ─────────────────────────
            aggregated: dict[str, dict] = {}
            bad_rows:   list[tuple]     = []

            for file_row in records:
                client_name = str(file_row.get(col_clientname) or "").strip()
                if not client_name:
                    continue

                norm_key = _normalise_for_sim(client_name)

                if norm_key not in aggregated:
                    bucket: dict = {"_display_name": client_name}
                    for db_col, (file_col, is_monetary) in detected.items():
                        bucket[db_col] = 0.0 if is_monetary else []
                    aggregated[norm_key] = bucket

                bucket = aggregated[norm_key]

                for db_col, (file_col, is_monetary) in detected.items():
                    if file_col is None:
                        continue
                    raw = str(file_row.get(file_col) or "").strip()

                    if is_monetary:
                        cleaned = re.sub(r"[^\d.]", "", raw.replace(",", ""))
                        try:
                            val = float(cleaned) if cleaned else None
                        except ValueError:
                            val = None
                            if raw:
                                bad_rows.append(
                                    (norm_key, f"bad {db_col}: '{raw}'"))
                        if val is not None:
                            bucket[db_col] = (bucket[db_col] or 0.0) + val
                    else:
                        if raw and raw not in bucket[db_col]:
                            bucket[db_col].append(raw)

            # ── 5. Flatten aggregated buckets into final write-ready dicts ─
            write_ready: dict[str, dict] = {}

            for norm_key, bucket in aggregated.items():
                record: dict = {"_display_name": bucket["_display_name"]}
                for db_col, (file_col, is_monetary) in detected.items():
                    if file_col is None:
                        record[db_col] = None
                        continue
                    raw_val = bucket[db_col]
                    if is_monetary:
                        record[db_col] = raw_val if raw_val != 0.0 else None
                    else:
                        joined = ", ".join(raw_val) if raw_val else None
                        record[db_col] = joined
                write_ready[norm_key] = record

            # ── 6. Build client_id → db row id lookup ────────────────────
            with _db_connect() as _conn:
                cid_to_dbid: dict[str, int] = {
                    str(r[0]).strip().upper(): r[1]
                    for r in _conn.execute(
                        "SELECT client_id, id FROM applicants "
                        "WHERE client_id IS NOT NULL AND TRIM(client_id) != ''"
                    ).fetchall()
                }

            # ── 7. Match & write to DB ────────────────────────────────────
            updated_by_id   = []
            updated_by_name = []
            updated_relaxed = []
            not_found       = []

            def _write_entry(conn, db_row_id: int, record: dict):
                parts, vals = [], []
                for db_col in detected:
                    val = record.get(db_col)
                    if val is None:
                        continue
                    parts.append(f"{db_col}=?")
                    vals.append(val)
                if record.get("loan_amount") is not None:
                    parts.append("principal_loan=?")
                    vals.append(record["loan_amount"])
                if parts:
                    conn.execute(
                        f"UPDATE applicants SET {', '.join(parts)} WHERE id=?",
                        vals + [db_row_id])

            with _db_connect() as conn:
                for norm_key, record in write_ready.items():
                    display_name = record["_display_name"]

                    file_cid = str(record.get("client_id") or "").strip().upper()
                    if file_cid and file_cid in cid_to_dbid:
                        db_id = cid_to_dbid[file_cid]
                        _write_entry(conn, db_id, record)
                        updated_by_id.append((norm_key, db_id))
                        continue

                    hits, sim_label = _resolve_name_similarity(display_name)
                    if not hits:
                        not_found.append(norm_key)
                        continue

                    for hit_id, _ in hits:
                        _write_entry(conn, hit_id, record)

                    if sim_label in ("exact", "high"):
                        updated_by_name.append((norm_key, hits[0][1]))
                    else:
                        updated_relaxed.append((norm_key, hits[0][1]))

            self.after(0, lambda: _refresh_summary(self))

            # ── 8. Build result message ───────────────────────────────────
            detected_display = ", ".join(
                f"{db_col}='{fc}'" for db_col, (fc, _) in detected.items() if fc
            )
            total_updated = len(updated_by_id) + len(updated_by_name) + len(updated_relaxed)
            msg  = "Principal Loan import complete.\n\n"
            msg += f"Columns detected      : {detected_display}\n\n"
            msg += f"✓  Matched by ID      : {len(updated_by_id):,} record(s)\n"
            msg += f"✓  Matched by name    : {len(updated_by_name):,} record(s)\n"
            msg += f"Total updated         : {total_updated:,} record(s)\n"
            msg += f"–  Not found          : {len(not_found):,} name(s)\n"
            if bad_rows:
                msg += f"⚠  Bad numeric values : {len(bad_rows):,} row(s)\n"
            if not detected.get("client_id") or not detected["client_id"][0]:
                msg += "\nℹ  No Client ID column found — used name matching only.\n"
            if updated_relaxed:
                msg += f"\n⚠  {len(updated_relaxed)} matched via relaxed similarity — please verify:\n"
                for file_n, db_id in updated_relaxed[:10]:
                    msg += f"  • {file_n}  →  DB id: {db_id}\n"
                if len(updated_relaxed) > 10:
                    msg += f"  … and {len(updated_relaxed) - 10} more\n"
            if not_found:
                msg += "\nNames with no DB match:\n"
                for k in not_found[:15]:
                    msg += f"  • {k}\n"
                if len(not_found) > 15:
                    msg += f"  … and {len(not_found) - 15} more\n"
            if bad_rows:
                msg += "\nRows with unparseable numeric values:\n"
                for k, reason in bad_rows[:10]:
                    msg += f"  • {k}  ({reason})\n"
                if len(bad_rows) > 10:
                    msg += f"  … and {len(bad_rows) - 10} more\n"

            self.after(0, lambda: (
                _flash_btn(self, self._sum_import_ploan_btn, "✓  Done!", 2500),
                messagebox.showinfo("P.Loan Import Result", msg)
            ))

        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: (
                _flash_btn(self, self._sum_import_ploan_btn, "✗  Error", 3000),
                messagebox.showerror("P.Loan Import Error", err)
            ))

    threading.Thread(target=_worker, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  MERGE DB
# ═══════════════════════════════════════════════════════════════════════

def _merge_db_files(self):
    paths = filedialog.askopenfilenames(
        title="Select DB files to merge into current database",
        filetypes=[("SQLite DB files", "*.db"), ("All files", "*.*")])
    if not paths:
        return

    src_paths = [p for p in paths if Path(p).resolve() != DB_PATH.resolve()]
    if not src_paths:
        messagebox.showwarning("Merge DB",
            "All selected files are the current database — nothing to merge.")
        return

    _flash_btn(self, self._sum_merge_db_btn, "⟳  Merging…", 60_000)

    # All real DB columns (excludes virtual display cols)
    _COLS = [
        "session_id", "processed_at", "source_file", "status",
        "applicant_name", "residence_address", "office_address",
        "income_items", "income_total",
        "business_items", "business_total",
        "household_items", "household_total",
        "net_income", "petrol_risk", "transport_risk",
        "results_json", "page_map",
        "amort_history_total",
        "amort_current_total",
        "client_id", "pn", "industry_name",
        "loan_balance", "amortized_cost",
        "principal_loan", "maturity", "interest_rate",
        "branch", "loan_class_name", "product_name",
        "loan_date", "term_unit", "term", "security", "release_tag",
        "loan_amount",
        "loan_status", "ao_name",
    ]
    _INSERT = (
        f"INSERT INTO applicants ({', '.join(_COLS)}) "
        f"VALUES ({', '.join(['?' for _ in _COLS])})"
    )

    def _ensure_col(conn, col, col_type):
        existing = [r[1] for r in conn.execute(
            "PRAGMA table_info(applicants)").fetchall()]
        if col not in existing:
            conn.execute(f"ALTER TABLE applicants ADD COLUMN {col} {col_type}")

    _MIGRATIONS = [
        ("amort_history_total", "REAL"),
        ("amort_current_total", "REAL"),
        ("client_id",           "TEXT"),
        ("pn",                  "TEXT"),
        ("industry_name",       "TEXT"),
        ("loan_balance",        "REAL"),
        ("amortized_cost",      "REAL"),
        ("principal_loan",      "REAL"),
        ("maturity",            "TEXT"),
        ("interest_rate",       "TEXT"),
        ("branch",              "TEXT"),
        ("loan_class_name",     "TEXT"),
        ("product_name",        "TEXT"),
        ("loan_date",           "TEXT"),
        ("term_unit",           "TEXT"),
        ("term",                "TEXT"),
        ("security",            "TEXT"),
        ("release_tag",         "TEXT"),
        ("loan_amount",         "REAL"),
        ("loan_status",         "TEXT"),
        ("ao_name",             "TEXT"),
    ]

    def _worker():
        total_inserted = 0; total_skipped = 0; total_patched = 0
        file_results   = []

        try:
            with _db_connect() as out_conn:
                for col, ctype in _MIGRATIONS:
                    _ensure_col(out_conn, col, ctype)

                for src in src_paths:
                    try:
                        s_conn = sqlite3.connect(str(src), timeout=10)
                        s_conn.row_factory = sqlite3.Row
                        for col, ctype in _MIGRATIONS:
                            _ensure_col(s_conn, col, ctype)
                        src_rows = s_conn.execute(
                            "SELECT * FROM applicants").fetchall()
                        s_conn.close()
                    except Exception as e:
                        file_results.append((Path(src).name, 0, 0, 0, str(e)))
                        continue

                    existing_by_clientid = {
                        str(r[0]).strip().upper(): r[1]
                        for r in out_conn.execute(
                            "SELECT client_id, id FROM applicants "
                            "WHERE client_id IS NOT NULL AND TRIM(client_id) != ''"
                        ).fetchall()
                    }
                    existing_primary = {
                        (r[0], r[1])
                        for r in out_conn.execute(
                            "SELECT session_id, source_file FROM applicants"
                        ).fetchall()
                    }
                    existing_fallback = {
                        (str(r[0]).strip().upper(), str(r[1]).strip().upper())
                        for r in out_conn.execute(
                            "SELECT applicant_name, source_file FROM applicants"
                        ).fetchall()
                    }

                    ins = skp = pat = 0
                    for row in src_rows:
                        rd           = dict(row)
                        incoming_cid = str(rd.get("client_id") or "").strip().upper()
                        pk           = (rd.get("session_id", ""),
                                        rd.get("source_file", ""))
                        fk           = (
                            str(rd.get("applicant_name") or "").strip().upper(),
                            str(rd.get("source_file")    or "").strip().upper(),
                        )

                        if incoming_cid and incoming_cid in existing_by_clientid:
                            _patch_existing(out_conn,
                                            existing_by_clientid[incoming_cid], rd)
                            pat += 1; continue
                        if pk in existing_primary or fk in existing_fallback:
                            skp += 1; continue

                        out_conn.execute(_INSERT, [rd.get(c) for c in _COLS])
                        if incoming_cid:
                            existing_by_clientid[incoming_cid] = out_conn.execute(
                                "SELECT last_insert_rowid()").fetchone()[0]
                        existing_primary.add(pk)
                        existing_fallback.add(fk)
                        ins += 1

                    out_conn.commit()
                    file_results.append((Path(src).name, ins, pat, skp, None))
                    total_inserted += ins; total_patched += pat; total_skipped += skp

            _db_deduplicate_client_ids()
            self.after(0, lambda: _refresh_summary(self))

            msg  = "Merge complete.\n\n"
            msg += f"✓  Inserted : {total_inserted:,} new record(s)\n"
            msg += f"🔧  Patched  : {total_patched:,} existing record(s)\n"
            msg += f"–  Skipped  : {total_skipped:,} (exact duplicates)\n\n"
            msg += "Per file:\n"
            for fname, ins, pat, skp, err in file_results:
                if err:
                    msg += f"  ✗  {fname}  →  Error: {err}\n"
                else:
                    msg += (f"  ✓  {fname}  →  {ins:,} inserted, "
                            f"{pat:,} patched, {skp:,} skipped\n")

            self.after(0, lambda: (
                _flash_btn(self, self._sum_merge_db_btn, "✓  Done!", 2500),
                messagebox.showinfo("Merge DB Result", msg)
            ))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: (
                _flash_btn(self, self._sum_merge_db_btn, "✗  Error", 3000),
                messagebox.showerror("Merge DB Error", err)
            ))

    threading.Thread(target=_worker, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  MERGE EXCEL
# ═══════════════════════════════════════════════════════════════════════

def _merge_excel_files(self):
    paths = filedialog.askopenfilenames(
        title="Select Excel summaries to merge",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if not paths or len(paths) < 1:
        return

    out_path = filedialog.asksaveasfilename(
        title="Save merged Excel as…",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=f"Merged_LookUp_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    if not out_path:
        return

    _flash_btn(self, self._sum_merge_xl_btn, "⟳  Merging…", 60_000)

    def _worker():
        _HEADERS = [
            "Client ID", "PN",
            "Applicant", "Residence Address", "Office Address", "Industry Name",
            "Spouse Info", "Personal Assets", "Business Assets", "Business Inventory",
            "Source of Income", "Total Source Of Income",
            "Business Expenses", "Total Business Expenses",
            "Household / Personal Expenses", "Total Household / Personal Expenses",
            "Total Net Income",
            "Total Amortization History", "Total Current Amortization",
            "Loan Balance", "Total Amortized Cost", "Principal Loan",
            "Maturity", "Interest Rate",
            # P.Loan expanded
            "Branch", "Loan Class", "Product Name",
            "Loan Date", "Term Unit", "Term", "Security", "Release Tag",
            "Loan Amount",
            "Loan Status", "AO Name",
        ]
        _MONETARY = {
            "Total Source Of Income", "Total Business Expenses",
            "Total Household / Personal Expenses",
            "Total Amortization History", "Total Current Amortization",
            "Loan Balance", "Total Amortized Cost", "Principal Loan",
            "Loan Amount", "Loan Balance (PLoan)", "Amort. (PLoan)",
        }
        _NET_COL    = "Total Net Income"
        _SUM_COLS   = _MONETARY | {_NET_COL}
        _COL_WIDTHS = {
            "Client ID": 14, "PN": 12,
            "Applicant": 22, "Residence Address": 30, "Office Address": 26,
            "Industry Name": 20,
            "Spouse Info": 28, "Personal Assets": 28,
            "Business Assets": 28, "Business Inventory": 24,
            "Source of Income": 32, "Total Source Of Income": 22,
            "Business Expenses": 32, "Total Business Expenses": 22,
            "Household / Personal Expenses": 36,
            "Total Household / Personal Expenses": 24,
            "Total Net Income": 20,
            "Total Amortization History": 26, "Total Current Amortization": 26,
            "Loan Balance": 22, "Total Amortized Cost": 24,
            "Principal Loan": 22, "Maturity": 20, "Interest Rate": 18,
            "Branch": 18, "Loan Class": 20, "Product Name": 22,
            "Industry (PLoan)": 22, "Loan Date": 16, "Term Unit": 14,
            "Term": 10, "Security": 24, "Release Tag": 16,
            "Loan Amount": 20, "Loan Balance (PLoan)": 22,
            "Amort. (PLoan)": 18, "Loan Status": 16, "AO Name": 22,
        }

        def _to_float(val):
            if val is None or val == "":
                return None
            try:
                return float(str(val).replace(",", "")
                             .replace("(", "-").replace(")", ""))
            except Exception:
                return None

        def _read_one(path):
            import openpyxl as _xl
            wb = _xl.load_workbook(str(path), read_only=True, data_only=True)
            ws = wb.active
            rows_iter  = ws.iter_rows(values_only=True)
            header_row = next(rows_iter, None)
            if header_row is None:
                wb.close(); return [], []
            file_hdrs = [str(h).strip() if h is not None else "" for h in header_row]
            col_map   = {h: i for i, h in enumerate(file_hdrs) if h in _HEADERS}
            records   = []
            for row in rows_iter:
                first = str(row[0]).strip().upper() if row[0] is not None else ""
                if first in ("TOTAL", "AVERAGE") or all(v is None for v in row):
                    continue
                rec = {h: (row[col_map[h]] if h in col_map else None)
                       for h in _HEADERS}
                records.append(rec)
            wb.close()
            return records, [h for h in _HEADERS if h not in col_map]

        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            all_rows       = []
            seen_clientids: set[str] = set()
            seen_names:     set[str] = set()
            dup_rows       = []
            file_results   = []

            for path in paths:
                rows, missing = _read_one(path)
                ins = 0
                for row in rows:
                    cid  = str(row.get("Client ID") or "").strip().upper()
                    name = str(row.get("Applicant")  or "").strip().upper()
                    if cid and cid in seen_clientids:
                        dup_rows.append((name, Path(path).name, "dup client_id"))
                        continue
                    if not cid and name in seen_names:
                        dup_rows.append((name, Path(path).name, "dup name"))
                        continue
                    if cid: seen_clientids.add(cid)
                    if name: seen_names.add(name)
                    all_rows.append(row); ins += 1
                file_results.append((Path(path).name, ins, missing))

            if not all_rows:
                raise ValueError("No data rows found across selected files.")

            hdr_fill  = PatternFill("solid", fgColor="93C47D")
            tot_fill  = PatternFill("solid", fgColor="D9EAD3")
            avg_fill  = PatternFill("solid", fgColor="B8D0E5")
            even_fill = PatternFill("solid", fgColor="FFFFFF")
            odd_fill  = PatternFill("solid", fgColor="F3F9F0")
            hdr_font  = Font(name="Roboto", bold=True, color="FFFFFF", size=10)
            body_font = Font(name="Roboto", size=9)
            bold_font = Font(name="Roboto", bold=True, size=9)
            net_font  = Font(name="Roboto", bold=True, size=9, color="1F6B28")
            tot_font  = Font(name="Roboto", bold=True, size=10)
            tot_font_j= Font(name="Roboto", bold=True, size=10, color="1F6B28")
            avg_font  = Font(name="Roboto", bold=True, size=10, color="1A3A5C")
            thin      = Side(style="thin",   color="CCCCCC")
            med       = Side(style="medium", color="555555")
            cell_bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
            tot_bdr   = Border(left=med,  right=med,  top=med,  bottom=med)
            wrap_al   = Alignment(horizontal="left",  vertical="top", wrap_text=True)
            right_al  = Alignment(horizontal="right", vertical="top")
            right_c   = Alignment(horizontal="right", vertical="center")
            left_c    = Alignment(horizontal="left",  vertical="center")
            CURRENCY  = '#,##0.00;(#,##0.00);"-"'

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Look-Up Summary (Merged)"

            for ci, h in enumerate(_HEADERS, 1):
                cell = ws.cell(row=1, column=ci, value=h)
                cell.font = hdr_font; cell.fill = hdr_fill
                cell.alignment = wrap_al; cell.border = cell_bdr
                ws.column_dimensions[get_column_letter(ci)].width = \
                    _COL_WIDTHS.get(h, 20)
            ws.row_dimensions[1].height = 28

            for ri, row_dict in enumerate(all_rows, 2):
                bg_fill    = even_fill if (ri - 2) % 2 == 0 else odd_fill
                text_lines = []
                for ci, h in enumerate(_HEADERS, 1):
                    val  = row_dict.get(h)
                    fval = _to_float(val) if h in _SUM_COLS else None
                    cell = ws.cell(row=ri, column=ci,
                                   value=fval if fval is not None
                                   else (val if val is not None else None))
                    cell.fill = bg_fill; cell.border = cell_bdr
                    if h == _NET_COL:
                        cell.number_format = CURRENCY
                        cell.font = net_font; cell.alignment = right_al
                    elif h in _MONETARY:
                        cell.number_format = CURRENCY
                        cell.font = bold_font; cell.alignment = right_al
                    elif h == "Applicant":
                        cell.font = bold_font; cell.alignment = wrap_al
                    else:
                        cell.font = body_font; cell.alignment = wrap_al
                        if val:
                            text_lines.append(len(str(val).split("\n")))
                ws.row_dimensions[ri].height = max(
                    18, min((max(text_lines) if text_lines else 1) * 15, 150))

            first_data = 2; last_data = len(all_rows) + 1
            tot_row    = last_data + 1; avg_row = last_data + 2

            for ci, h in enumerate(_HEADERS, 1):
                col_l = get_column_letter(ci)
                cell  = ws.cell(row=tot_row, column=ci)
                cell.fill = tot_fill; cell.border = tot_bdr
                if h == "Applicant":
                    cell.value = "TOTAL"; cell.font = tot_font; cell.alignment = left_c
                elif h == _NET_COL:
                    cell.value = f"=SUM({col_l}{first_data}:{col_l}{last_data})"
                    cell.number_format = CURRENCY
                    cell.font = tot_font_j; cell.alignment = right_c
                elif h in _SUM_COLS:
                    cell.value = f"=SUM({col_l}{first_data}:{col_l}{last_data})"
                    cell.number_format = CURRENCY
                    cell.font = tot_font; cell.alignment = right_c
                else:
                    cell.font = tot_font; cell.alignment = left_c

                cell_avg = ws.cell(row=avg_row, column=ci)
                cell_avg.fill = avg_fill; cell_avg.border = tot_bdr
                if h == "Applicant":
                    cell_avg.value = "AVERAGE"
                    cell_avg.font = avg_font; cell_avg.alignment = left_c
                elif h == _NET_COL:
                    cell_avg.value = f"=AVERAGE({col_l}{first_data}:{col_l}{last_data})"
                    cell_avg.number_format = CURRENCY
                    cell_avg.font = avg_font; cell_avg.alignment = right_c
                elif h in _SUM_COLS:
                    cell_avg.value = f"=AVERAGE({col_l}{first_data}:{col_l}{last_data})"
                    cell_avg.number_format = CURRENCY
                    cell_avg.font = avg_font; cell_avg.alignment = right_c
                else:
                    cell_avg.font = avg_font; cell_avg.alignment = left_c

            ws.row_dimensions[tot_row].height = 22
            ws.row_dimensions[avg_row].height = 22
            ws.freeze_panes = "A2"
            wb.save(out_path)

            msg  = f"Merge complete.\n\n"
            msg += f"✓  Total rows : {len(all_rows):,}\n"
            msg += f"–  Duplicates : {len(dup_rows):,} (skipped)\n\n"
            msg += "Per file:\n"
            for fname, ins, missing in file_results:
                msg += f"  ✓  {fname}  →  {ins:,} rows added\n"
                if missing:
                    msg += f"     ⚠ Missing cols (blanked): {', '.join(missing)}\n"
            if dup_rows:
                msg += "\nDuplicate applicants skipped (first file wins):\n"
                for name, src, reason in dup_rows[:15]:
                    msg += f"  • {name}  (from {src})\n"
                if len(dup_rows) > 15:
                    msg += f"  … and {len(dup_rows) - 15} more\n"

            self.after(0, lambda: (
                _flash_btn(self, self._sum_merge_xl_btn, "✓  Saved!", 2500),
                messagebox.showinfo("Merge Excel Result", msg)
            ))
        except ImportError:
            self.after(0, lambda: (
                _flash_btn(self, self._sum_merge_xl_btn, "openpyxl not installed", 3000),
                messagebox.showerror("Merge Excel Error",
                    "openpyxl is required.\nRun: pip install openpyxl")
            ))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: (
                _flash_btn(self, self._sum_merge_xl_btn, "✗  Error", 3000),
                messagebox.showerror("Merge Excel Error", err)
            ))

    threading.Thread(target=_worker, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  EXPORT HELPERS
# ═══════════════════════════════════════════════════════════════════════

def _get_all_filtered_rows(self) -> list:
    raw    = self._sum_search_var.get().strip()
    search = "" if "separate terms with commas" in raw else raw
    rows, _ = _db_query(
        search=search, session_id=self._sum_session_filter,
        sort_col=self._sum_sort_col, sort_asc=self._sum_sort_asc,
        offset=0, limit=100_000,
        adv_filters=getattr(self, "_sum_adv_filters", {}))
    return [dict(r) for r in rows]


def _row_to_export_dict(row: dict) -> dict:
    def _fmt(val):
        try:
            return float(val) if val not in (None, "") else None
        except Exception:
            return None

    # Derive virtual fields from results_json
    try:
        results = json.loads(row.get("results_json", "") or "{}")
    except Exception:
        results = {}

    spouse_info        = _extract_spouse_info(results)
    personal_assets    = _extract_asset_items(results, "cibi_personal_assets")
    business_assets    = _extract_asset_items(results, "cibi_business_assets")
    business_inventory = _extract_asset_items(results, "cibi_business_inventory")

    # amort_history_total: prefer real DB column, fall back to JSON parse
    amort_hist = row.get("amort_history_total")
    if amort_hist is None:
        amort_hist = _parse_amort_history_total(row.get("results_json", ""))

    return {
        "Client ID":                           row.get("client_id",          "") or "",
        "PN":                                  row.get("pn",                 "") or "",
        "Applicant":                           row.get("applicant_name",     "") or "",
        "Residence Address":                   row.get("residence_address",  "") or "",
        "Office Address":                      row.get("office_address",     "") or "",
        "Industry Name":                       row.get("industry_name",      "") or "",
        "Spouse Info":                         spouse_info,
        "Personal Assets":                     personal_assets,
        "Business Assets":                     business_assets,
        "Business Inventory":                  business_inventory,
        "Source of Income":                    row.get("income_items",       "") or "",
        "Total Source Of Income":              _fmt(row.get("income_total")),
        "Business Expenses":                   row.get("business_items",     "") or "",
        "Total Business Expenses":             _fmt(row.get("business_total")),
        "Household / Personal Expenses":       row.get("household_items",    "") or "",
        "Total Household / Personal Expenses": _fmt(row.get("household_total")),
        "Total Net Income":                    _fmt(row.get("net_income")),
        "Total Amortization History":          _fmt(amort_hist),
        "Total Current Amortization":          _fmt(row.get("amort_current_total")),
        "Loan Balance":                        _fmt(row.get("loan_balance")),
        "Total Amortized Cost":                _fmt(row.get("amortized_cost")),
        "Principal Loan":                      _fmt(row.get("principal_loan")),
        "Maturity":                            row.get("maturity",           "") or "",
        "Interest Rate":                       row.get("interest_rate",      "") or "",
        # ── P.Loan expanded fields ─────────────────────────────────────
        "Branch":                              row.get("branch",             "") or "",
        "Loan Class":                          row.get("loan_class_name",    "") or "",
        "Product Name":                        row.get("product_name",       "") or "",
        "Loan Date":                           row.get("loan_date",          "") or "",
        "Term Unit":                           row.get("term_unit",          "") or "",
        "Term":                                row.get("term",               "") or "",
        "Security":                            row.get("security",           "") or "",
        "Release Tag":                         row.get("release_tag",        "") or "",
        "Loan Amount":                         _fmt(row.get("loan_amount")),
        "Loan Status":                         row.get("loan_status",        "") or "",
        "AO Name":                             row.get("ao_name",            "") or "",
    }


def _export_csv(self):
    rows = _get_all_filtered_rows(self)
    if not rows:
        return
    path = filedialog.asksaveasfilename(
        title="Export to CSV", defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        initialfile=f"LookUp_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    if not path:
        return
    flat    = [_row_to_export_dict(r) for r in rows]
    headers = list(flat[0].keys()) if flat else []
    try:
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            writer.writerows(flat)
        _flash_btn(self, self._sum_export_csv_btn, "✓  Saved!", 2000)
    except Exception as e:
        _flash_btn(self, self._sum_export_csv_btn, f"Error: {e}", 3000)


def _export_excel(self):
    rows = _get_all_filtered_rows(self)
    if not rows:
        messagebox.showinfo("Export", "No records to export.")
        return

    flat = [_row_to_export_dict(r) for r in rows]
    if not flat:
        return

    checklist_result = _show_export_checklist(self, flat)
    if checklist_result is None:
        return
    flat, selected_col_keys = checklist_result
    if not flat:
        messagebox.showinfo("Export", "No rows were selected — nothing exported.")
        return
    if not selected_col_keys:
        messagebox.showinfo("Export", "No columns were selected — nothing exported.")
        return

    path = filedialog.asksaveasfilename(
        title="Export to Excel", defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=f"LookUp_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    if not path:
        return

    def _worker():
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            all_possible_headers = [c[0] for c in _ALL_EXPORT_COLS]
            selected_set         = set(selected_col_keys)
            headers              = [h for h in all_possible_headers
                                    if h in selected_set]
            if not headers:
                return

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Look-Up Summary"

            hdr_fill  = PatternFill("solid", fgColor="93C47D")
            tot_fill  = PatternFill("solid", fgColor="D9EAD3")
            avg_fill  = PatternFill("solid", fgColor="B8D0E5")
            even_fill = PatternFill("solid", fgColor="FFFFFF")
            odd_fill  = PatternFill("solid", fgColor="F3F9F0")
            hdr_font  = Font(name="Roboto", bold=True, color="FFFFFF", size=10)
            body_font = Font(name="Roboto", size=9)
            bold_font = Font(name="Roboto", bold=True, size=9)
            net_font  = Font(name="Roboto", bold=True, size=9, color="1F6B28")
            tot_font  = Font(name="Roboto", bold=True, size=10)
            tot_font_j= Font(name="Roboto", bold=True, size=10, color="1F6B28")
            avg_font  = Font(name="Roboto", bold=True, size=10, color="1A3A5C")
            thin      = Side(style="thin",   color="CCCCCC")
            med       = Side(style="medium", color="555555")
            cell_bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
            tot_bdr   = Border(left=med,  right=med,  top=med,  bottom=med)
            wrap_al   = Alignment(horizontal="left",  vertical="top", wrap_text=True)
            right_al  = Alignment(horizontal="right", vertical="top")
            right_c   = Alignment(horizontal="right", vertical="center")
            left_c    = Alignment(horizontal="left",  vertical="center")
            CURRENCY  = '#,##0.00;(#,##0.00);"-"'

            TOTAL_COLS = {
                "Total Source Of Income", "Total Business Expenses",
                "Total Household / Personal Expenses",
                "Total Amortization History", "Total Current Amortization",
                "Loan Balance", "Total Amortized Cost", "Principal Loan",
                "Loan Amount",
            }
            NET_COL  = "Total Net Income"
            SUM_COLS = TOTAL_COLS | {NET_COL}

            col_widths = {
                "Client ID": 14, "PN": 12,
                "Applicant": 22, "Residence Address": 30, "Office Address": 26,
                "Industry Name": 20,
                "Spouse Info": 28, "Personal Assets": 28,
                "Business Assets": 28, "Business Inventory": 24,
                "Source of Income": 32, "Total Source Of Income": 22,
                "Business Expenses": 32, "Total Business Expenses": 22,
                "Household / Personal Expenses": 36,
                "Total Household / Personal Expenses": 24,
                "Total Net Income": 20,
                "Total Amortization History": 26,
                "Total Current Amortization": 26,
                "Loan Balance": 22, "Total Amortized Cost": 24,
                "Principal Loan": 22, "Maturity": 20, "Interest Rate": 18,
                "Branch": 18, "Loan Class": 20, "Product Name": 22,
                "Loan Date": 16, "Term Unit": 14,
                "Term": 10, "Security": 24, "Release Tag": 16,
                "Loan Amount": 20,
                "Loan Status": 16, "AO Name": 22,
            }

            for ci, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=ci, value=h)
                cell.font = hdr_font; cell.fill = hdr_fill
                cell.alignment = wrap_al; cell.border = cell_bdr
                ws.column_dimensions[get_column_letter(ci)].width = \
                    col_widths.get(h, 20)
            ws.row_dimensions[1].height = 28

            for ri, row_dict in enumerate(flat, 2):
                bg_fill    = even_fill if (ri - 2) % 2 == 0 else odd_fill
                text_lines = []
                for ci, h in enumerate(headers, 1):
                    val  = row_dict.get(h)
                    cell = ws.cell(row=ri, column=ci,
                                   value=val if val is not None else None)
                    cell.fill = bg_fill; cell.border = cell_bdr
                    if h == NET_COL:
                        cell.number_format = CURRENCY
                        cell.font = net_font; cell.alignment = right_al
                    elif h in TOTAL_COLS:
                        cell.number_format = CURRENCY
                        cell.font = bold_font; cell.alignment = right_al
                    elif h in ("Applicant", "Client ID", "PN"):
                        cell.font = bold_font; cell.alignment = wrap_al
                    else:
                        cell.font = body_font; cell.alignment = wrap_al
                        if val:
                            text_lines.append(len(str(val).split("\n")))
                max_lines = max(text_lines) if text_lines else 1
                ws.row_dimensions[ri].height = max(18, min(max_lines * 15, 150))

            first_data = 2; last_data = len(flat) + 1
            total_row  = last_data + 1; avg_row = last_data + 2

            for ci, h in enumerate(headers, 1):
                col_l = get_column_letter(ci)
                cell  = ws.cell(row=total_row, column=ci)
                cell.fill = tot_fill; cell.border = tot_bdr
                if h == "Applicant":
                    cell.value = "TOTAL"; cell.font = tot_font; cell.alignment = left_c
                elif h == NET_COL:
                    cell.value = f"=SUM({col_l}{first_data}:{col_l}{last_data})"
                    cell.number_format = CURRENCY
                    cell.font = tot_font_j; cell.alignment = right_c
                elif h in SUM_COLS:
                    cell.value = f"=SUM({col_l}{first_data}:{col_l}{last_data})"
                    cell.number_format = CURRENCY
                    cell.font = tot_font; cell.alignment = right_c
                else:
                    cell.font = tot_font; cell.alignment = left_c

                cell_avg = ws.cell(row=avg_row, column=ci)
                cell_avg.fill = avg_fill; cell_avg.border = tot_bdr
                if h == "Applicant":
                    cell_avg.value = "AVERAGE"
                    cell_avg.font = avg_font; cell_avg.alignment = left_c
                elif h == NET_COL:
                    cell_avg.value = f"=AVERAGE({col_l}{first_data}:{col_l}{last_data})"
                    cell_avg.number_format = CURRENCY
                    cell_avg.font = avg_font; cell_avg.alignment = right_c
                elif h in SUM_COLS:
                    cell_avg.value = f"=AVERAGE({col_l}{first_data}:{col_l}{last_data})"
                    cell_avg.number_format = CURRENCY
                    cell_avg.font = avg_font; cell_avg.alignment = right_c
                else:
                    cell_avg.font = avg_font; cell_avg.alignment = left_c

            ws.row_dimensions[total_row].height = 22
            ws.row_dimensions[avg_row].height   = 22
            ws.freeze_panes = "A2"
            wb.save(path)

            self.after(0, lambda: (
                _flash_btn(self, self._sum_export_xl_btn, "✓  Saved!", 2000),
                messagebox.showinfo("Export Excel",
                    f"Saved {len(flat):,} record(s) · {len(headers)} column(s)\n"
                    f"with TOTAL and AVERAGE rows to:\n{path}")
            ))
        except ImportError:
            self.after(0, lambda: (
                _flash_btn(self, self._sum_export_xl_btn, "openpyxl missing", 3000),
                messagebox.showerror("Export Excel Error",
                    "openpyxl is not installed.\nRun: pip install openpyxl")
            ))
        except Exception as e:
            self.after(0, lambda err=str(e): (
                _flash_btn(self, self._sum_export_xl_btn, "✗  Error", 3000),
                messagebox.showerror("Export Excel Error", err)
            ))

    threading.Thread(target=_worker, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  VALIDATE CLIENT ID ↔ APPLICANT NAME
# ═══════════════════════════════════════════════════════════════════════

def _validate_clients(self):
    path = filedialog.askopenfilename(
        title="Select Reference File (Client ID + Client Name)",
        filetypes=[("Excel & CSV files", "*.xlsx *.csv"),
                   ("Excel files", "*.xlsx"),
                   ("CSV files", "*.csv"),
                   ("All files", "*.*")])
    if not path:
        return

    out_path = filedialog.asksaveasfilename(
        title="Save Validation Report as…",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=f"Validation_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    if not out_path:
        return

    _flash_btn(self, self._sum_validate_btn, "⟳  Validating…", 60_000)

    def _worker():
        try:
            def _find_col(cols, *keywords):
                for kw in keywords:
                    kw_norm = re.sub(r"[\s_\-]", "", kw.lower())
                    for c in cols:
                        c_norm = re.sub(r"[\s_\-]", "", c.lower())
                        if kw_norm == c_norm or kw_norm in c_norm:
                            return c
                return None

            if path.lower().endswith(".csv"):
                import csv as _csv
                with open(path, newline="", encoding="utf-8-sig") as f:
                    reader      = _csv.DictReader(f)
                    ref_records = [dict(row) for row in reader]
                all_cols = list(ref_records[0].keys()) if ref_records else []
            else:
                import openpyxl as _oxl
                wb = _oxl.load_workbook(path, read_only=True, data_only=True)
                ws = wb.active
                header_row = next(ws.iter_rows(min_row=1, max_row=1), None)
                if header_row is None:
                    raise ValueError("Reference file appears to be empty.")
                all_cols = [
                    str(cell.value).strip() if cell.value is not None else ""
                    for cell in header_row
                ]
                ref_records = []
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if all(v is None for v in row):
                        continue
                    ref_records.append({
                        all_cols[i]: (str(v).strip() if v is not None else "")
                        for i, v in enumerate(row) if i < len(all_cols)
                    })
                wb.close()

            if not ref_records:
                raise ValueError("No data rows found in the reference file.")

            col_cid  = _find_col(all_cols, "clientid", "client id", "client_id", "cid")
            col_name = _find_col(all_cols, "clientname", "client name", "applicant",
                                 "applicantname", "name")
            if not col_cid:
                raise ValueError(
                    f"Could not detect a Client ID column.\n\nFile has: {', '.join(all_cols)}")
            if not col_name:
                raise ValueError(
                    f"Could not detect a Client Name column.\n\nFile has: {', '.join(all_cols)}")

            ref_map:       dict[str, str]  = {}
            ref_order:     list[tuple]     = []
            seen_ref_cids: set[str]        = set()

            for rec in ref_records:
                cid  = str(rec.get(col_cid,  "") or "").strip()
                name = str(rec.get(col_name, "") or "").strip()
                cid_up = cid.upper()
                if cid and name and cid_up not in seen_ref_cids:
                    ref_map[cid_up] = name
                    ref_order.append((cid, name))
                    seen_ref_cids.add(cid_up)

            with _db_connect() as conn:
                db_rows = conn.execute(
                    "SELECT id, client_id, applicant_name, pn, industry_name, "
                    "residence_address, office_address, "
                    "income_total, business_total, household_total, net_income, "
                    "amort_history_total, amort_current_total, "
                    "loan_balance, amortized_cost, "
                    "principal_loan, maturity, interest_rate, "
                    "branch, loan_class_name, product_name, loan_date, security, loan_status, ao_name, "
                    "loan_amount "
                    "FROM applicants"
                ).fetchall()
            db_rows = [dict(r) for r in db_rows]

            matched:   list[tuple] = []
            unmatched: list[tuple] = []
            no_cid:    list[str]   = []
            db_cid_set: set[str]   = set()

            for r in db_rows:
                db_cid  = str(r.get("client_id",      "") or "").strip()
                db_name = str(r.get("applicant_name", "") or "").strip()
                if not db_cid:
                    no_cid.append(db_name or "(no name)"); continue
                db_cid_up = db_cid.upper()
                db_cid_set.add(db_cid_up)
                if db_cid_up not in ref_map:
                    unmatched.append((db_cid, db_name, "Client ID not in reference"))
                    continue
                ref_name = ref_map[db_cid_up]
                score    = _similarity(
                    _normalise_for_sim(db_name.upper()),
                    _normalise_for_sim(ref_name.upper()))
                if score >= SIMILARITY_THRESHOLD:
                    matched.append((db_cid, db_name, ref_name))
                else:
                    unmatched.append((db_cid, db_name,
                                      f"Name mismatch — reference: {ref_name}"))

            ref_not_in_db = [(cid, name) for cid, name in ref_order
                             if cid.upper() not in db_cid_set]

            CHECKED_COLS = [
                ("client_id",           "Client ID"),
                ("pn",                  "PN"),
                ("applicant_name",      "Applicant Name"),
                ("residence_address",   "Residence Address"),
                ("office_address",      "Office Address"),
                ("industry_name",       "Industry Name"),
                ("income_total",        "Total Income"),
                ("business_total",      "Total Business Expenses"),
                ("household_total",     "Total Household Expenses"),
                ("net_income",          "Total Net Income"),
                ("amort_history_total", "Total Amortization History"),
                ("amort_current_total", "Total Current Amortization"),
                ("loan_balance",        "Loan Balance"),
                ("amortized_cost",      "Total Amortized Cost"),
                ("principal_loan",      "Principal Loan"),
                ("maturity",            "Maturity"),
                ("interest_rate",       "Interest Rate"),
                ("branch",              "Branch"),
                ("loan_class_name",     "Loan Class"),
                ("product_name",        "Product Name"),
                ("loan_date",           "Loan Date"),
                ("security",            "Security"),
                ("loan_status",         "Loan Status"),
                ("ao_name",             "AO Name"),
                ("loan_amount",         "Loan Amount"),
            ]

            missing_info_rows: list[dict] = []
            for r in db_rows:
                missing_cols = []
                for col_key, col_label in CHECKED_COLS:
                    val = r.get(col_key)
                    if val is None or (isinstance(val, str) and val.strip() == ""):
                        missing_cols.append(col_label)
                if missing_cols:
                    missing_info_rows.append({
                        "client_id":      r.get("client_id",      "") or "",
                        "applicant_name": r.get("applicant_name", "") or "(no name)",
                        "missing":        missing_cols,
                    })

            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            wb_out = openpyxl.Workbook()
            thin     = Side(style="thin",   color="CCCCCC")
            med      = Side(style="medium", color="555555")
            cell_bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
            hdr_bdr  = Border(left=med,  right=med,  top=med,  bottom=med)
            wrap_al  = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
            ctr_al   = Alignment(horizontal="center", vertical="center")
            left_c   = Alignment(horizontal="left",   vertical="center")
            hdr_font = Font(name="Segoe UI", bold=True, color="FFFFFF", size=10)
            body_f   = Font(name="Segoe UI", size=9)
            bold_f   = Font(name="Segoe UI", bold=True, size=9)

            fill_hdr_green  = PatternFill("solid", fgColor="2D6A4F")
            fill_hdr_red    = PatternFill("solid", fgColor="9B2226")
            fill_hdr_gray   = PatternFill("solid", fgColor="4A4E69")
            fill_match      = PatternFill("solid", fgColor="D8F3DC")
            fill_unmatch    = PatternFill("solid", fgColor="FFE8E8")
            fill_nocid      = PatternFill("solid", fgColor="FFF3CD")
            fill_notindb    = PatternFill("solid", fgColor="E8F0FF")
            fill_section    = PatternFill("solid", fgColor="F0F4FA")
            fill_missing    = PatternFill("solid", fgColor="FFF8E8")

            def _hc(ws, row, col, value, fill):
                c = ws.cell(row=row, column=col, value=value)
                c.font = hdr_font; c.fill = fill
                c.alignment = ctr_al; c.border = hdr_bdr
                return c

            def _bc(ws, row, col, value, fill, fnt=None, al=None):
                c = ws.cell(row=row, column=col, value=value)
                c.font = fnt or body_f; c.fill = fill
                c.alignment = al or wrap_al; c.border = cell_bdr
                return c

            def _section_label(ws, row, text, fill_bg, txt_color, ncols):
                c = ws.cell(row=row, column=1, value=text)
                c.font = Font(name="Segoe UI", bold=True, size=9, color=txt_color)
                c.fill = fill_bg; c.alignment = wrap_al; c.border = cell_bdr
                if ncols > 1:
                    ws.merge_cells(start_row=row, start_column=1,
                                   end_row=row,   end_column=ncols)
                ws.row_dimensions[row].height = 18

            # ── Sheet 1: Summary ───────────────────────────────────────
            ws1 = wb_out.active
            ws1.title = "Summary"
            ws1.column_dimensions["A"].width = 38
            ws1.column_dimensions["B"].width = 18
            ws1.cell(row=1, column=1,
                     value="Validation Report — Summary").font = \
                Font(name="Segoe UI", bold=True, size=13, color="1A3A5C")
            ws1.cell(row=1, column=1).alignment = ctr_al
            ws1.merge_cells("A1:B1")
            ws1.row_dimensions[1].height = 28
            ws1.cell(row=2, column=1,
                     value=f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}").font = \
                Font(name="Segoe UI", size=8, color="888888")
            ws1.merge_cells("A2:B2")
            ws1.row_dimensions[2].height = 16

            total_db    = len(db_rows)
            n_matched   = len(matched)
            n_unmatched = len(unmatched)
            n_no_cid    = len(no_cid)
            n_notindb   = len(ref_not_in_db)
            n_miss_rows = len(missing_info_rows)
            rate        = f"{n_matched / total_db * 100:.1f}%" if total_db else "N/A"

            _hc(ws1, 3, 1, "Category", fill_hdr_gray)
            _hc(ws1, 3, 2, "Count",    fill_hdr_gray)
            ws1.row_dimensions[3].height = 22

            for ri, (label, count, fill) in enumerate([
                ("Total Records in DB",                  total_db,    fill_section),
                ("✓  Matched (ID + Name correct)",        n_matched,   fill_match),
                ("✗  Unmatched (ID or Name mismatch)",    n_unmatched, fill_unmatch),
                ("⚠  No Client ID in DB (unvalidatable)", n_no_cid,    fill_nocid),
                ("—  In Reference but not in DB",         n_notindb,   fill_notindb),
                ("⚠  Records with missing column data",   n_miss_rows, fill_missing),
                ("Total Unique Entries in Reference",     len(ref_map), fill_section),
            ], 4):
                _bc(ws1, ri, 1, label, fill, bold_f, left_c)
                _bc(ws1, ri, 2, count, fill, bold_f, ctr_al)
                ws1.row_dimensions[ri].height = 20

            rate_row = 4 + 7
            c1 = ws1.cell(row=rate_row, column=1, value="Match Rate")
            c1.font = Font(name="Segoe UI", bold=True, size=10, color="1F6B28")
            c1.alignment = left_c
            c2 = ws1.cell(row=rate_row, column=2, value=rate)
            c2.font = Font(name="Segoe UI", bold=True, size=10, color="1F6B28")
            c2.alignment = ctr_al
            ws1.row_dimensions[rate_row].height = 22
            ws1.freeze_panes = "A4"

            # ── Sheet 2: Unmatched & Not in DB ─────────────────────────
            ws2 = wb_out.create_sheet("Unmatched & Not In DB")
            ws2.column_dimensions["A"].width = 18
            ws2.column_dimensions["B"].width = 30
            ws2.column_dimensions["C"].width = 38
            ws2.column_dimensions["D"].width = 16
            _hc(ws2, 1, 1, "Client ID",     fill_hdr_red)
            _hc(ws2, 1, 2, "Applicant",     fill_hdr_red)
            _hc(ws2, 1, 3, "Reason / Note", fill_hdr_red)
            _hc(ws2, 1, 4, "Issue Type",    fill_hdr_red)
            ws2.row_dimensions[1].height = 24

            ri = 2
            if unmatched or no_cid:
                _section_label(ws2, ri,
                    "SECTION A — DB Records: Unmatched or Missing Client ID",
                    PatternFill("solid", fgColor="FFE8E8"), "9B2226", 4)
                ri += 1
                for db_cid, db_name, reason in unmatched:
                    issue = "Name Mismatch" if "mismatch" in reason.lower() else "ID Not in Ref"
                    _bc(ws2, ri, 1, db_cid,  fill_unmatch, bold_f)
                    _bc(ws2, ri, 2, db_name, fill_unmatch)
                    _bc(ws2, ri, 3, reason,  fill_unmatch)
                    _bc(ws2, ri, 4, issue,   fill_unmatch, bold_f, ctr_al)
                    ws2.row_dimensions[ri].height = 18; ri += 1
                for name in no_cid:
                    _bc(ws2, ri, 1, "(no ID)",               fill_nocid, bold_f)
                    _bc(ws2, ri, 2, name,                    fill_nocid)
                    _bc(ws2, ri, 3, "Missing Client ID in DB", fill_nocid)
                    _bc(ws2, ri, 4, "No ID",                 fill_nocid, bold_f, ctr_al)
                    ws2.row_dimensions[ri].height = 18; ri += 1

            ri += 1
            _section_label(ws2, ri,
                "SECTION B — In Reference File but NOT found in DB",
                PatternFill("solid", fgColor="E8F0FF"), "1A3A5C", 4)
            ri += 1
            if ref_not_in_db:
                for ref_cid, ref_name in ref_not_in_db:
                    _bc(ws2, ri, 1, ref_cid,           fill_notindb, bold_f)
                    _bc(ws2, ri, 2, ref_name,           fill_notindb)
                    _bc(ws2, ri, 3, "Not found in DB",  fill_notindb)
                    _bc(ws2, ri, 4, "Not in DB",        fill_notindb, bold_f, ctr_al)
                    ws2.row_dimensions[ri].height = 18; ri += 1
            else:
                c = ws2.cell(row=ri, column=1,
                             value="✓  All reference entries exist in the DB.")
                c.font = Font(name="Segoe UI", bold=True, size=9, color="1F6B28")
                c.fill = fill_notindb; c.alignment = wrap_al
                ws2.merge_cells(start_row=ri, start_column=1,
                                end_row=ri, end_column=4)
                ws2.row_dimensions[ri].height = 18
            ws2.freeze_panes = "A2"

            # ── Sheet 3: Missing Info ───────────────────────────────────
            ws3 = wb_out.create_sheet("Missing Info")
            ws3.column_dimensions["A"].width = 18
            ws3.column_dimensions["B"].width = 30
            ws3.column_dimensions["C"].width = 16
            ws3.column_dimensions["D"].width = 55
            _hc(ws3, 1, 1, "Client ID",      fill_hdr_red)
            _hc(ws3, 1, 2, "Applicant",       fill_hdr_red)
            _hc(ws3, 1, 3, "Missing Count",   fill_hdr_red)
            _hc(ws3, 1, 4, "Missing Columns", fill_hdr_red)
            ws3.row_dimensions[1].height = 24

            if missing_info_rows:
                missing_info_rows.sort(key=lambda x: len(x["missing"]), reverse=True)
                for ri3, entry in enumerate(missing_info_rows, 2):
                    n_miss = len(entry["missing"])
                    fill_row = (PatternFill("solid", fgColor="FFD0D0") if n_miss >= 8
                                else PatternFill("solid", fgColor="FFE8CC") if n_miss >= 4
                                else fill_missing)
                    _bc(ws3, ri3, 1, entry["client_id"],         fill_row, bold_f)
                    _bc(ws3, ri3, 2, entry["applicant_name"],     fill_row)
                    _bc(ws3, ri3, 3, n_miss,                      fill_row, bold_f, ctr_al)
                    _bc(ws3, ri3, 4, ", ".join(entry["missing"]), fill_row)
                    ws3.row_dimensions[ri3].height = 20
            else:
                c = ws3.cell(row=2, column=1,
                             value="✓  All records have complete information.")
                c.font = Font(name="Segoe UI", bold=True, size=10, color="1F6B28")
                c.alignment = wrap_al
                ws3.merge_cells("A2:D2")
                ws3.row_dimensions[2].height = 22
            ws3.freeze_panes = "A2"

            wb_out.save(out_path)

            msg  = "Validation complete.\n\n"
            msg += f"Total DB records      : {total_db:,}\n"
            msg += f"✓  Matched            : {n_matched:,}\n"
            msg += f"✗  Unmatched          : {n_unmatched:,}\n"
            msg += f"⚠  No Client ID in DB : {n_no_cid:,}\n"
            msg += f"—  In Ref, not in DB  : {n_notindb:,}\n"
            msg += f"⚠  Missing col info   : {n_miss_rows:,} record(s)\n"
            msg += f"Match Rate            : {rate}\n\n"
            msg += f"Report saved to:\n{out_path}"

            self.after(0, lambda: (
                _flash_btn(self, self._sum_validate_btn, "✓  Done!", 2500),
                messagebox.showinfo("Validation Result", msg)
            ))
        except ImportError:
            self.after(0, lambda: (
                _flash_btn(self, self._sum_validate_btn, "openpyxl missing", 3000),
                messagebox.showerror("Validate Error",
                    "openpyxl is not installed.\nRun: pip install openpyxl")
            ))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: (
                _flash_btn(self, self._sum_validate_btn, "✗  Error", 3000),
                messagebox.showerror("Validate Error", err)
            ))

    threading.Thread(target=_worker, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════════════════

def _flash_btn(self, btn, msg: str, ms: int):
    try:
        orig = btn.cget("text")
        btn.configure(text=msg)
        self.after(ms, lambda: btn.configure(text=orig)
                   if btn.winfo_exists() else None)
    except Exception:
        pass


def lookup_summary_notify(self):
    if getattr(self, "_current_tab", "") == "lookup_summary":
        _refresh_summary(self)
    elif hasattr(self, "_sum_stat_labels"):
        _update_stats(self)


# ═══════════════════════════════════════════════════════════════════════
#  ATTACH
# ═══════════════════════════════════════════════════════════════════════

def attach(cls):
    cls._build_lookup_summary_panel = _build_lookup_summary_panel
    cls.lookup_summary_notify        = lookup_summary_notify
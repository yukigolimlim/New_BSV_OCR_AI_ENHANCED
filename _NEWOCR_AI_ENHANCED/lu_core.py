"""
lu_core.py — LU Analysis: domain knowledge, Excel parsing, risk engine
=======================================================================
Pure-Python module.  No tkinter / matplotlib / reportlab imports here.
Everything that touches a workbook or produces a result dict lives here.

Updated to parse the Look-Up Summary sheet format:
  Client ID | PN | Applicant | Residence Address | Office Address |
  Industry Name | Source of Income | Total Source Of Income |
  Business Expenses | Total Business Expenses |
  Household / Personal Expenses | Total Household / Personal Expenses |
  Total Net Income | Total Amortization History |
  Total Current Amortization | Loan Balance
"""

import re
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False


# ══════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════════════════

GENERAL_CLIENT = "📊  General (All Clients)"

_MAX_HEADER_SCAN_ROWS = 10
_MAX_DATA_ROWS        = 2000

_DEFAULT_RISK = ("LOW", "No specific sector-expense sensitivity rule defined; treat as low risk.")
_RISK_ORDER   = {"HIGH": 0, "MODERATE": 1, "LOW": 2}

_SCORE_BANDS = [
    (2.5, "CRITICAL", "#B71C1C", "#FFEBEE"),
    (1.8, "HIGH",     "#E53E3E", "#FFF5F5"),
    (1.2, "MODERATE", "#D4A017", "#FFFBF0"),
    (0.0, "LOW",      "#2E7D32", "#F0FBE8"),
]

# ── Canonical sector names used throughout ──────────────────────────
SECTOR_WHOLESALE   = "Wholesale / Retail Trade"
SECTOR_AGRICULTURE = "Agriculture (Fishing & Forestry)"
SECTOR_TRANSPORT   = "Transport"
SECTOR_REMITTANCE  = "Remittance"
SECTOR_CONSUMER    = "Consumer Loan"
SECTOR_OTHER       = "Other / Household"


# ══════════════════════════════════════════════════════════════════════
#  INDUSTRY → SECTOR MAPPING
# ══════════════════════════════════════════════════════════════════════

# Keywords matched against the "Industry Name" column (case-insensitive substring)
INDUSTRY_SECTOR_MAP: list[tuple[str, str]] = [
    # Wholesale / Retail
    ("wholesale",           SECTOR_WHOLESALE),
    ("retail",              SECTOR_WHOLESALE),
    ("trading",             SECTOR_WHOLESALE),
    ("trade",               SECTOR_WHOLESALE),
    ("sari-sari",           SECTOR_WHOLESALE),
    ("sari sari",           SECTOR_WHOLESALE),
    ("store",               SECTOR_WHOLESALE),
    ("bigasan",             SECTOR_WHOLESALE),
    ("rice",                SECTOR_WHOLESALE),
    ("grocery",             SECTOR_WHOLESALE),
    ("supermarket",         SECTOR_WHOLESALE),
    ("repair of motor",     SECTOR_WHOLESALE),
    # Agriculture / Fishing / Forestry
    ("fishing",             SECTOR_AGRICULTURE),
    ("fishery",             SECTOR_AGRICULTURE),
    ("aquaculture",         SECTOR_AGRICULTURE),
    ("aqua",                SECTOR_AGRICULTURE),
    ("fish",                SECTOR_AGRICULTURE),
    ("forestry",            SECTOR_AGRICULTURE),
    ("farming",             SECTOR_AGRICULTURE),
    ("agriculture",         SECTOR_AGRICULTURE),
    ("agri",                SECTOR_AGRICULTURE),
    ("feeds",               SECTOR_AGRICULTURE),
    ("poultry",             SECTOR_AGRICULTURE),
    ("livestock",           SECTOR_AGRICULTURE),
    # Transport
    ("transport",           SECTOR_TRANSPORT),
    ("trucking",            SECTOR_TRANSPORT),
    ("hauling",             SECTOR_TRANSPORT),
    ("tricycle",            SECTOR_TRANSPORT),
    ("jeepney",             SECTOR_TRANSPORT),
    ("gasoline station",    SECTOR_TRANSPORT),
    ("gas station",         SECTOR_TRANSPORT),
    # Remittance
    ("remittance",          SECTOR_REMITTANCE),
    ("money transfer",      SECTOR_REMITTANCE),
    ("forex",               SECTOR_REMITTANCE),
    # Consumer Loan
    ("consumer loan",       SECTOR_CONSUMER),
    ("personal loan",       SECTOR_CONSUMER),
    ("salary loan",         SECTOR_CONSUMER),
    ("lending",             SECTOR_CONSUMER),
    ("microfinance",        SECTOR_CONSUMER),
    ("cooperative",         SECTOR_CONSUMER),
    ("coop",                SECTOR_CONSUMER),
]

# Also scan Source of Income text for sector hints
SOURCE_SECTOR_MAP: list[tuple[str, str]] = [
    ("remittance",          SECTOR_REMITTANCE),
    ("ofw",                 SECTOR_REMITTANCE),
    ("grocery",             SECTOR_WHOLESALE),
    ("store",               SECTOR_WHOLESALE),
    ("sari-sari",           SECTOR_WHOLESALE),
    ("rice",                SECTOR_WHOLESALE),
    ("fish",                SECTOR_AGRICULTURE),
    ("fishing",             SECTOR_AGRICULTURE),
    ("transport",           SECTOR_TRANSPORT),
    ("trucking",            SECTOR_TRANSPORT),
    ("hauling",             SECTOR_TRANSPORT),
    ("lending",             SECTOR_CONSUMER),
    ("salary loan",         SECTOR_CONSUMER),
]


# ══════════════════════════════════════════════════════════════════════
#  EXPENSE RISK TABLES  (per sector)
# ══════════════════════════════════════════════════════════════════════

SECTOR_KEYWORDS = {}          # kept for backward compat shim
SECTOR_KEYWORDS_EXTENDED = {} # kept for backward compat shim

EXPENSE_PATTERNS: list[tuple[re.Pattern, str]] = [
    (re.compile(r'\b(fuel|diesel|gasoline|gas|petrol)\b',          re.I), "Fuel / Diesel"),
    (re.compile(r'\b(oil|lubricant|lube)\b',                        re.I), "Oil & Lubricants"),
    (re.compile(r'\b(electricity|power|electric bill|elec|light)\b',re.I), "Utilities (Power)"),
    (re.compile(r'\b(water|water bill)\b',                          re.I), "Utilities (Water)"),
    (re.compile(r'\b(salary|salaries|wages|payroll|labor)\b',       re.I), "Salaries / Wages"),
    (re.compile(r'\b(rent|lease|rental)\b',                         re.I), "Rent / Lease"),
    (re.compile(r'\b(repair|maintenance|parts|spare)\b',            re.I), "Repairs & Maintenance"),
    (re.compile(r'\b(insurance|premium)\b',                         re.I), "Insurance Premium"),
    (re.compile(r'\b(feed|feeds|bait|fishmeal)\b',                  re.I), "Feed / Bait"),
    (re.compile(r'\b(ice|cold storage|refriger)\b',                 re.I), "Ice / Cold Storage"),
    (re.compile(r'\b(toll|tollway)\b',                              re.I), "Toll Fees"),
    (re.compile(r'\b(tax|taxes|bir|vat)\b',                         re.I), "Taxes"),
    (re.compile(r'\b(communication|telco|phone|internet|load)\b',   re.I), "Communication"),
    (re.compile(r'\b(supplies|office supplies|packaging)\b',        re.I), "Supplies"),
    (re.compile(r'\b(depreciation|amortization)\b',                 re.I), "Depreciation"),
    (re.compile(r'\b(freight|shipping|cargo|delivery|logistics)\b', re.I), "Freight / Logistics"),
    (re.compile(r'\b(interest|bank charge|financing)\b',            re.I), "Interest / Bank Charges"),
    (re.compile(r'\b(miscellaneous|misc|other expense)\b',          re.I), "Miscellaneous"),
    (re.compile(r'\b(household|personal expense|food|groceries)\b', re.I), "Household Expenses"),
]

SECTOR_EXPENSE_RISK: dict[str, dict[str, tuple[str, str]]] = {
    SECTOR_AGRICULTURE: {
        "Fuel / Diesel":         ("HIGH",     "Fishing vessels are heavily fuel-dependent; oil price spikes directly compress margins."),
        "Oil & Lubricants":      ("HIGH",     "Engine maintenance requires constant oil supply; shortages raise downtime risk."),
        "Feed / Bait":           ("HIGH",     "Bait and fishmeal prices track global commodity prices and are highly volatile."),
        "Ice / Cold Storage":    ("MODERATE", "Cold-chain demand increases with catch volume; power cost fluctuations apply."),
        "Utilities (Power)":     ("MODERATE", "Cold storage and processing equipment are electricity-intensive."),
        "Repairs & Maintenance": ("MODERATE", "Marine/farm equipment requires frequent maintenance; parts may face import-cost pressure."),
        "Salaries / Wages":      ("LOW",      "Crew/farm wages are relatively stable but may rise with commodity price pressure."),
        "Insurance Premium":     ("MODERATE", "Marine/agri insurance rises with environmental risk."),
        "Freight / Logistics":   ("MODERATE", "Land transport of catch/produce to markets is diesel-sensitive."),
        "Household Expenses":    ("LOW",      "Personal household costs; monitor for pressure on disposable income."),
        "Miscellaneous":         ("LOW",      "Indirect exposure; monitor for cost creep."),
    },
    SECTOR_TRANSPORT: {
        "Fuel / Diesel":         ("HIGH",     "Diesel is the primary operating cost; any price increase directly hits profitability."),
        "Oil & Lubricants":      ("HIGH",     "Fleet maintenance depends on oil products; supply disruptions affect schedules."),
        "Freight / Logistics":   ("HIGH",     "Subcontracted haulage costs move directly with diesel prices."),
        "Toll Fees":             ("MODERATE", "Toll adjustments compound with fuel costs on high-frequency routes."),
        "Repairs & Maintenance": ("MODERATE", "Aging fleets under high mileage face accelerated wear."),
        "Salaries / Wages":      ("MODERATE", "Driver wages may rise if operators compete for scarce licensed drivers."),
        "Insurance Premium":     ("MODERATE", "Commercial vehicle insurance tracks fuel-cost risk environment."),
        "Depreciation":          ("LOW",      "Straight-line depreciation is fixed; less sensitive unless asset replacement is imminent."),
        "Utilities (Power)":     ("LOW",      "Minimal direct exposure unless operating electric vehicles."),
        "Taxes":                 ("LOW",      "Stable unless excise tax regime changes."),
        "Household Expenses":    ("LOW",      "Personal household costs; monitor for pressure on disposable income."),
        "Miscellaneous":         ("LOW",      "Indirect; monitor."),
    },
    SECTOR_WHOLESALE: {
        "Freight / Logistics":     ("HIGH",     "Wholesale margins are thin; rising logistics and fuel costs reduce competitiveness."),
        "Fuel / Diesel":           ("HIGH",     "Delivery vehicles and generator backup are heavily fuel-dependent; price spikes directly raise operating costs."),
        "Oil & Lubricants":        ("HIGH",     "Vehicle and equipment maintenance requires constant oil supply; cost increases compound with fuel exposure."),
        "Utilities (Power)":       ("MODERATE", "Warehousing and refrigerated storage are power-intensive."),
        "Ice / Cold Storage":      ("MODERATE", "Perishable goods wholesalers face cold-chain cost pressure."),
        "Interest / Bank Charges": ("MODERATE", "Higher interest rate environment raises working capital financing cost."),
        "Salaries / Wages":        ("LOW",      "Wage bill is relatively stable; volume-driven hiring adjusts slowly."),
        "Rent / Lease":            ("LOW",      "Fixed lease terms provide short-term shelter from market rate increases."),
        "Supplies":                ("LOW",      "Packaging and office supplies show modest inflation sensitivity."),
        "Insurance Premium":       ("LOW",      "Cargo insurance is relatively stable absent major weather events."),
        "Taxes":                   ("LOW",      "Stable VAT and local business tax environment assumed."),
        "Household Expenses":      ("LOW",      "Personal household costs; monitor for pressure on disposable income."),
        "Miscellaneous":           ("LOW",      "Indirect; monitor."),
    },
    SECTOR_REMITTANCE: {
        "Salaries / Wages":        ("MODERATE", "Staff compensation is a primary operating cost; minimum wage adjustments apply."),
        "Rent / Lease":            ("MODERATE", "Branch network rental costs are sensitive to commercial real-estate trends."),
        "Communication":           ("LOW",      "Telco and internet costs are relatively stable."),
        "Utilities (Power)":       ("LOW",      "Branch office power consumption is modest."),
        "Insurance Premium":       ("LOW",      "Fidelity and liability premiums are generally stable."),
        "Interest / Bank Charges": ("LOW",      "Float management costs; mildly sensitive to rate environment."),
        "Taxes":                   ("LOW",      "Regulatory fees and DST are stable."),
        "Household Expenses":      ("LOW",      "Personal household costs; monitor for pressure on disposable income."),
        "Miscellaneous":           ("LOW",      "Indirect; monitor."),
    },
    SECTOR_CONSUMER: {
        "Interest / Bank Charges": ("HIGH",     "Rising benchmark rates directly compress net interest margins on fixed-rate consumer loans."),
        "Salaries / Wages":        ("MODERATE", "Collections and loan officer headcount drives cost-to-income ratio."),
        "Communication":           ("LOW",      "Collections outreach costs are modest."),
        "Utilities (Power)":       ("LOW",      "Office operations have limited power exposure."),
        "Rent / Lease":            ("LOW",      "Branch footprint is typically fixed-term leased."),
        "Insurance Premium":       ("LOW",      "Credit life insurance costs are passed to borrowers."),
        "Taxes":                   ("LOW",      "DST and documentary charges are volume-linked but stable per unit."),
        "Household Expenses":      ("LOW",      "Personal household costs; monitor for pressure on disposable income."),
        "Miscellaneous":           ("LOW",      "Indirect; monitor."),
    },
    SECTOR_OTHER: {
        "Household Expenses":      ("LOW",      "Standard household/personal expenses; no specific sector risk."),
        "Salaries / Wages":        ("LOW",      "Employment income is the primary source; relatively stable."),
        "Utilities (Power)":       ("LOW",      "Standard household utility costs."),
        "Miscellaneous":           ("LOW",      "General living expenses; monitor for over-commitment."),
    },
}


# ══════════════════════════════════════════════════════════════════════
#  CELL / NUMERIC HELPERS
# ══════════════════════════════════════════════════════════════════════

def _cell_str(cell) -> str:
    if cell is None:
        return ""
    val = getattr(cell, "value", None) if hasattr(cell, "value") else cell
    if val is None:
        return ""
    return str(val).strip()


def _parse_numeric(raw) -> float | None:
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        v = float(raw)
        return v if v == v else None
    txt = str(raw).strip()
    if not txt or txt.startswith("="):
        return None
    negative = False
    if txt.startswith("(") and txt.endswith(")"):
        txt = txt[1:-1].strip()
        negative = True
    txt = (txt.replace("₱", "").replace("$", "").replace("€", "")
              .replace("£", "").replace(",", "").strip().rstrip("%").strip())
    try:
        v = float(txt)
        return -v if negative else v
    except (ValueError, TypeError):
        return None


def _fmt_value(vals: list) -> str:
    nums = [v for v in vals if isinstance(v, float)]
    if nums:
        total = sum(nums)
        avg   = total / len(nums)
        if total >= 1_000_000:
            return f"₱{total:,.2f}  (avg ₱{avg:,.2f} over {len(nums)} entries)"
        return f"₱{total:,.2f}  (avg ₱{avg:,.2f})"
    return "; ".join(str(v) for v in vals[:5]) + ("…" if len(vals) > 5 else "")


def _numeric_total(vals: list) -> float:
    return sum(v for v in vals if isinstance(v, float))


# ══════════════════════════════════════════════════════════════════════
#  RISK SCORE
# ══════════════════════════════════════════════════════════════════════

def _compute_risk_score(expenses: list[dict]) -> tuple[float, str, str, str]:
    if not expenses:
        return (0.0, "N/A", "#9AAACE", "#F5F7FA")
    weight = {"HIGH": 3, "MODERATE": 2, "LOW": 1}
    total  = sum(weight.get(e["risk"], 1) for e in expenses)
    score  = total / len(expenses)
    for threshold, label, fg, bg in _SCORE_BANDS:
        if score >= threshold:
            return (score, label, fg, bg)
    return (score, "LOW", "#2E7D32", "#F0FBE8")


# ══════════════════════════════════════════════════════════════════════
#  SECTOR DETECTION  (from Industry Name + Source of Income text)
# ══════════════════════════════════════════════════════════════════════

def _detect_sector(industry_name: str, source_of_income: str) -> str:
    """Return the canonical sector for a client row."""
    text_industry = (industry_name or "").lower()
    text_source   = (source_of_income or "").lower()

    # Industry name takes priority
    for keyword, sector in INDUSTRY_SECTOR_MAP:
        if keyword in text_industry:
            return sector

    # Fall back to source of income hints
    for keyword, sector in SOURCE_SECTOR_MAP:
        if keyword in text_source:
            return sector

    # Household catch-all
    if "household" in text_industry or "personal" in text_industry:
        return SECTOR_OTHER

    return SECTOR_OTHER


# ══════════════════════════════════════════════════════════════════════
#  EXPENSE ITEM PARSER  (from multiline cell text like "Rent [8,000]")
# ══════════════════════════════════════════════════════════════════════

_ITEM_AMOUNT_PAT = re.compile(
    r'([A-Za-z][^[\n\r]*?)\s*\[\s*([0-9,]+(?:\.[0-9]+)?)\s*\]', re.S)


def _parse_expense_items(text: str) -> list[tuple[str, float]]:
    """Parse 'Label  [amount]\\nLabel2  [amount2]' → [(label, amount), ...]"""
    if not text:
        return []
    items = []
    for m in _ITEM_AMOUNT_PAT.finditer(str(text)):
        label  = m.group(1).strip().rstrip("(").strip()
        amount = _parse_numeric(m.group(2))
        if label and amount is not None and amount > 0:
            items.append((label, amount))
    return items


def _classify_expense_item(label: str) -> str:
    """Map a free-text expense label to a canonical expense category."""
    low = label.lower()
    for pat, name in EXPENSE_PATTERNS:
        if pat.search(low):
            return name
    return "Miscellaneous"


# ══════════════════════════════════════════════════════════════════════
#  COLUMN INDEX FINDER  (for the Look-Up Summary header row)
# ══════════════════════════════════════════════════════════════════════

_COL_PATTERNS = {
    "client_id":         re.compile(r'client\s*id',                          re.I),
    "pn":                re.compile(r'^pn$',                                  re.I),
    "applicant":         re.compile(r'applicant',                             re.I),
    "residence":         re.compile(r'residence',                             re.I),
    "office":            re.compile(r'office\s*address',                      re.I),
    "industry":          re.compile(r'industry\s*name',                       re.I),
    "source_income":     re.compile(r'source\s+of\s+income',                  re.I),
    "total_source":      re.compile(r'total\s+source\s+of\s+income',          re.I),
    "biz_exp_detail":    re.compile(r'^business\s+expenses?$',                re.I),
    "total_biz_exp":     re.compile(r'total\s+business\s+expenses?',          re.I),
    "hhld_exp_detail":   re.compile(r'household\s*/?\s*personal\s+expenses?', re.I),
    "total_hhld_exp":    re.compile(r'total\s+household',                     re.I),
    "total_net_income":  re.compile(r'total\s+net\s+income',                  re.I),
    "amort_history":     re.compile(r'total\s+amortization\s+history',        re.I),
    "current_amort":     re.compile(r'total\s+current\s+amortization',        re.I),
    "principal_loan":    re.compile(r'principal\s*loan',                      re.I),
    "loan_balance":      re.compile(r'loan\s+balance',                        re.I),
}


def _find_columns(header_row: tuple) -> dict[str, int]:
    """Return {field_name: 0-based col index} for recognised columns."""
    cols: dict[str, int] = {}
    for i, cell_val in enumerate(header_row):
        if cell_val is None:
            continue
        s = str(cell_val).strip()
        for field, pat in _COL_PATTERNS.items():
            if field not in cols and pat.search(s):
                cols[field] = i
    return cols


# ══════════════════════════════════════════════════════════════════════
#  ROW → CLIENT RECORD
# ══════════════════════════════════════════════════════════════════════

def _row_to_client(row: tuple, cols: dict[str, int]) -> dict | None:
    """Convert a data row into a client analysis dict.  Returns None to skip."""

    def get(field):
        idx = cols.get(field)
        return row[idx] if idx is not None and idx < len(row) else None

    applicant = str(get("applicant") or "").strip()
    if not applicant or applicant.upper() in ("TOTAL", "SUBTOTAL", "GRAND TOTAL", ""):
        return None

    client_id = str(get("client_id") or "").strip()
    pn        = str(get("pn")        or "").strip()

    industry  = str(get("industry")      or "").strip()
    src_text  = str(get("source_income") or "").strip()
    biz_text  = str(get("biz_exp_detail")  or "").strip()
    hhld_text = str(get("hhld_exp_detail") or "").strip()

    total_source  = _parse_numeric(get("total_source"))
    total_biz     = _parse_numeric(get("total_biz_exp"))
    total_hhld    = _parse_numeric(get("total_hhld_exp"))
    total_net     = _parse_numeric(get("total_net_income"))
    amort_hist     = _parse_numeric(get("amort_history"))
    current_amort  = _parse_numeric(get("current_amort"))
    loan_balance   = _parse_numeric(get("loan_balance"))
    principal_loan = _parse_numeric(get("principal_loan"))

    residence = str(get("residence") or "").strip()
    office    = str(get("office")    or "").strip()

    # ── Sector detection ──────────────────────────────────────────────
    sector = _detect_sector(industry, src_text)

    # ── Build expense breakdown from detail cells ─────────────────────
    risk_table = SECTOR_EXPENSE_RISK.get(sector, {})

    # Parse individual items from biz and household detail cells
    biz_items  = _parse_expense_items(biz_text)
    hhld_items = _parse_expense_items(hhld_text)

    # Bucket items into canonical expense categories
    category_totals: dict[str, float] = {}
    for label, amt in biz_items + hhld_items:
        cat = _classify_expense_item(label)
        category_totals[cat] = category_totals.get(cat, 0.0) + amt

    # If no itemised data, fall back to column totals
    if not category_totals:
        if total_biz:
            category_totals["Business Expenses (Total)"] = total_biz
        if total_hhld:
            category_totals["Household Expenses"] = total_hhld

    # Build expense list
    expenses_out: list[dict] = []
    seen_cats: set[str] = set()

    for cat, total in sorted(category_totals.items(), key=lambda x: -x[1]):
        risk, reason = risk_table.get(cat, _DEFAULT_RISK)
        expenses_out.append({
            "name":       cat,
            "risk":       risk,
            "reason":     reason,
            "value_str":  f"₱{total:,.2f}",
            "has_values": True,
            "total":      total,
        })
        seen_cats.add(cat)

    # Advisory items (HIGH/MODERATE) not found in data
    for exp_name, (adv_risk, adv_reason) in risk_table.items():
        if adv_risk in ("HIGH", "MODERATE") and exp_name not in seen_cats:
            expenses_out.append({
                "name":       exp_name,
                "risk":       adv_risk,
                "reason":     adv_reason,
                "value_str":  "(not found — advisory)",
                "has_values": False,
                "total":      0.0,
            })

    expenses_out.sort(key=lambda x: (_RISK_ORDER.get(x["risk"], 9), -x["total"]))
    score, score_label, score_fg, score_bg = _compute_risk_score(expenses_out)

    # ── Sector-level override: these sectors are always HIGH risk ─────────────
    _HIGH_RISK_SECTORS = {
        SECTOR_AGRICULTURE,   # Agriculture (Fishing & Forestry)
        SECTOR_WHOLESALE,     # Wholesale / Retail Trade
        SECTOR_CONSUMER,      # Consumer Loan
        SECTOR_REMITTANCE,    # Remittance
        SECTOR_TRANSPORT,     # Transport
    }
    if sector in _HIGH_RISK_SECTORS:
        score       = max(score, 1.8)   # ensure score sits in the HIGH band
        score_label = "HIGH"
        score_fg    = "#E53E3E"
        score_bg    = "#FFF5F5"

    return {
        "client_id":      client_id,
        "pn":             pn,
        "client":         applicant,
        "residence":      residence,
        "office":         office,
        "industry":       industry,
        "sector":         sector,
        "source_income":  src_text,
        "biz_expenses":   biz_text,
        "hhld_expenses":  hhld_text,
        "total_source":   total_source,
        "total_biz":      total_biz,
        "total_hhld":     total_hhld,
        "net_income":     total_net,
        "amort_history":  amort_hist,
        "current_amort":  current_amort,
        "principal_loan": principal_loan,
        "loan_balance":   loan_balance,
        "expenses":       expenses_out,
        "score":          score,
        "score_label":    score_label,
        "score_fg":       score_fg,
        "score_bg":       score_bg,
        "sheet":          "",   # filled by caller
    }


# ══════════════════════════════════════════════════════════════════════
#  BACKWARD COMPAT STUBS (used by lu_ui report / export)
# ══════════════════════════════════════════════════════════════════════

def _is_summary_sheet(ws) -> bool:
    return True   # all sheets in new format are summary-style

def _analyse_sheet(ws, sheet_name: str) -> list[dict]:
    return []   # not used; run_lu_analysis handles everything

def _read_income_from_summary(wb) -> dict[str, dict]:
    return {}   # income is now embedded in client records


# ══════════════════════════════════════════════════════════════════════
#  PUBLIC ENTRY POINT
# ══════════════════════════════════════════════════════════════════════

def run_lu_analysis(filepath: str) -> dict:
    """
    Load an Excel workbook (Look-Up Summary format) and return:
    {
        "general":    [client_dict, ...],       # all clients
        "clients":    {applicant_name: client_dict},
        "income_map": {applicant_name: {"gross": float|None, "net": float|None}},
        "sector_map": {sector_name: [client_dict, ...]},
        "totals": {
            "loan_balance":   float,
            "total_source":   float,
            "total_net":      float,
            "current_amort":  float,
        }
    }
    """
    if not _HAS_OPENPYXL:
        raise RuntimeError("openpyxl is not installed.\nRun:  pip install openpyxl")

    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    try:
        general:    list[dict]           = []
        clients:    dict[str, dict]      = {}
        income_map: dict[str, dict]      = {}
        sector_map: dict[str, list]      = {}

        for sheet_name in wb.sheetnames:
            ws   = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            # Find header row (scan first 10 rows)
            cols      = {}
            data_start = 1
            for scan_idx, row in enumerate(rows[:_MAX_HEADER_SCAN_ROWS]):
                cols = _find_columns(row)
                if len(cols) >= 5:
                    data_start = scan_idx + 1
                    break

            if not cols:
                continue

            for row in rows[data_start:data_start + _MAX_DATA_ROWS]:
                rec = _row_to_client(row, cols)
                if rec is None:
                    continue
                rec["sheet"] = sheet_name

                name = rec["client"]
                clients[name] = rec
                general.append(rec)

                income_map[name] = {
                    "gross": rec["total_source"],
                    "net":   rec["net_income"],
                }

                sec = rec["sector"]
                sector_map.setdefault(sec, []).append(rec)

    finally:
        wb.close()

    # Grand totals
    totals = {
        "loan_balance":  sum(r["loan_balance"]   or 0 for r in general),
        "total_source":  sum(r["total_source"]   or 0 for r in general),
        "total_net":     sum(r["net_income"]      or 0 for r in general),
        "current_amort": sum(r["current_amort"]  or 0 for r in general),
    }

    return {
        "general":    general,
        "clients":    clients,
        "income_map": income_map,
        "sector_map": sector_map,
        "totals":     totals,
    }
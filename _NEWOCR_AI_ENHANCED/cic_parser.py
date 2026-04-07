"""
cic_parser.py — CIC Credit Report Parser for DocExtract Pro
=============================================================
Detects, parses, and structures Philippine Credit Information Corporation
(CIC) credit reports into a normalized dict for use by the analysis engine.

Exports
-------
  is_cic_report(text)          → bool
  parse_cic_report(text)       → dict
  format_cic_for_analysis(text)→ str   (human-readable enriched summary)
  CIC_ANALYSIS_PROMPT_BLOCK    → str   (system-prompt injection for Groq)
"""

from __future__ import annotations
import re
from typing import Any


# ══════════════════════════════════════════════════════════════════════════════
#  DETECTION
# ══════════════════════════════════════════════════════════════════════════════

# Phrases that strongly identify a CIC credit report
_CIC_FINGERPRINTS = [
    r"credit\s+information\s+corporation",
    r"cic\s+subject\s+code",
    r"cic\s+contract\s+code",
    r"provider\s+code\s+encrypted",
    r"reorganized\s+credit\s+indicator",
    r"overdue\s+payments\s+number",
    r"installments\s+detail",
    r"creditinfo\.gov\.ph",
    r"subject\s+matched",          # header badge on CIC reports
]


def is_cic_report(text: str) -> bool:
    """Return True if the extracted text looks like a CIC credit report."""
    t = text.lower()
    hits = sum(1 for p in _CIC_FINGERPRINTS if re.search(p, t, re.I))
    return hits >= 2


# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _find(pattern: str, text: str, group: int = 1,
          flags: int = re.I) -> str:
    m = re.search(pattern, text, flags)
    return m.group(group).strip() if m else ""


def _find_all(pattern: str, text: str, flags: int = re.I) -> list[str]:
    return [m.strip() for m in re.findall(pattern, text, flags)]


def _amount(raw: str) -> float:
    """Convert a string like '3,000,000' or '96,666' to float."""
    try:
        return float(raw.replace(",", "").replace("₱", "").strip())
    except (ValueError, AttributeError):
        return 0.0


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN PARSER
# ══════════════════════════════════════════════════════════════════════════════

def parse_cic_report(text: str) -> dict[str, Any]:
    """
    Parse a CIC credit report (as extracted text) into a structured dict.

    Returns a dict with these top-level keys:
      subject, employment, contracts, credit_cards,
      legal_flags, risk_summary, raw_text
    """
    data: dict[str, Any] = {
        "subject":      _parse_subject(text),
        "employment":   _parse_employment(text),
        "contracts":    _parse_contracts(text),
        "credit_cards": _parse_credit_cards(text),
        "legal_flags":  _parse_legal_flags(text),
        "raw_text":     text,
    }
    data["risk_summary"] = _build_risk_summary(data)
    return data


# ── SUBJECT ───────────────────────────────────────────────────────────────────

def _parse_subject(text: str) -> dict:
    s: dict[str, Any] = {}

    # CIC Subject Code
    s["cic_subject_code"] = _find(
        r"cic\s+subject\s+code[\s:]+([A-Z0-9]+)", text)

    # Name — try explicit labels first, then fall back to positional
    first  = _find(r"first\s+name[\s:]+([A-Z][A-Z\s]+?)(?:\n|last\s+name)", text)
    middle = _find(r"middle\s+name[\s:]+([A-Z][A-Z\s]+?)(?:\n|suffix)", text)
    last   = _find(r"last\s+name[\s:]+([A-Z][A-Z\s]+?)(?:\n|previous)", text)
    suffix = _find(r"suffix[\s:]+([IVX]{1,4}|JR|SR)(?:\s|\n)", text)
    s["full_name"] = " ".join(p for p in [first, middle, last, suffix] if p) or "Unknown"

    s["date_of_birth"] = _find(
        r"date\s+of\s+birth[\s:]+(\d{2}/\d{2}/\d{4})", text)
    s["gender"]        = _find(r"gender[\s:]+(Male|Female)", text)
    s["civil_status"]  = _find(r"civil\s+status[\s:]+(Single|Married|Widowed|Separated)", text)
    s["nationality"]   = _find(r"nationality[\s:]+([A-Z]+)", text)
    s["resident"]      = _find(r"resident[\s:]+(Yes|No)", text)
    s["dependents"]    = _find(r"number\s+of\s+dependents[\s:]+(\d+)", text)
    s["cars_owned"]    = _find(r"cars\s+owned[\s:]+(\d+)", text)

    # Spouse
    sp_first = _find(r"spouse[\s\S]{0,30}first\s+name[\s:]+([A-Z][A-Z\s]+?)(?:\n|last\s+name)", text)
    sp_last  = _find(r"spouse[\s\S]{0,60}last\s+name[\s:]+([A-Z][A-Z\s]+?)(?:\n)", text)
    s["spouse"] = f"{sp_first} {sp_last}".strip() if sp_first or sp_last else ""

    # IDs
    s["tin"] = _find(r"tin[\s:]+(\d{6,12})", text)
    s["sss"] = _find(r"sss\s*card[\s:]+([A-Z0-9]+)", text)

    # Address (first address block)
    s["address"] = _find(
        r"individual\s*[-–]\s*main\s+address[^\n]*\n\s*([^\n]{10,})", text)

    return s


# ── EMPLOYMENT ────────────────────────────────────────────────────────────────

def _parse_employment(text: str) -> dict:
    e: dict[str, Any] = {}

    e["employer"]          = _find(r"company\s+trade\s+name[\s:]+([^\n]+)", text)
    e["industry"]          = _find(r"psic[\s:]+([^\n]{5,})", text)
    e["occupation"]        = _find(r"occupation[\s:]+([^\n]+)", text)
    e["occupation_status"] = _find(r"occupation\s+status[\s:]+([^\n]+)", text)

    gross_raw = _find(r"gross\s+income[\s:]+([0-9,]+)", text)
    e["gross_income_raw"]  = gross_raw
    e["gross_income"]      = _amount(gross_raw)

    freq = _find(r"annual/monthly[\s:]+(Annual|Monthly)", text)
    e["income_frequency"]  = freq

    # Normalise to monthly
    if e["gross_income"] > 0:
        if freq.lower() == "annual":
            e["monthly_income"] = e["gross_income"] / 12
        else:
            e["monthly_income"] = e["gross_income"]
    else:
        e["monthly_income"] = 0.0

    e["hired_from"] = _find(r"hired\s+from[\s:]+([0-9/]+)", text)
    e["hired_to"]   = _find(r"hired\s+to[\s:]+([0-9/]+)", text)

    # Compute years of service
    if e["hired_from"]:
        try:
            from datetime import datetime
            hf = datetime.strptime(e["hired_from"], "%d/%m/%Y")
            ht_str = e.get("hired_to", "")
            if ht_str and ht_str not in ("-", ""):
                ht = datetime.strptime(ht_str, "%d/%m/%Y")
            else:
                # Use report request date if available, else today
                rd = _find(r"request\s+date\s+(\d{2}/\d{2}/\d{4})", text)
                ht = datetime.strptime(rd, "%d/%m/%Y") if rd else datetime.today()
            e["years_of_service"] = round((ht - hf).days / 365.25, 1)
        except (ValueError, TypeError):
            e["years_of_service"] = None
    else:
        e["years_of_service"] = None

    return e


# ── CONTRACTS (INSTALLMENTS) ──────────────────────────────────────────────────

def _parse_contracts(text: str) -> list[dict]:
    contracts = []

    # ── Active / Closed contracts table ──────────────────────────────────────
    # Pattern: CIC contract code  type  financed  outstanding  overdue
    active_pattern = re.compile(
        r"([A-Z]\d{8})\s+"                          # CIC contract code
        r"([\w/\s]+?)\s+"                           # contract type
        r"([\d,]+)\s+"                              # financed amount
        r"([\d,]+)\s+"                              # outstanding balance
        r"([\d,]+)\s+"                              # overdue amount
        r"(\d{2}/\d{2}/\d{4})\s+"                  # contract start
        r"(\d{2}/\d{2}/\d{4})",                    # contract end
        re.I,
    )
    for m in active_pattern.finditer(text):
        contracts.append({
            "cic_contract_code": m.group(1),
            "contract_type":     m.group(2).strip(),
            "financed_amount":   _amount(m.group(3)),
            "outstanding":       _amount(m.group(4)),
            "overdue_amount":    _amount(m.group(5)),
            "start_date":        m.group(6),
            "end_date":          m.group(7),
            "phase":             "active",
        })

    # ── Requested / Refused contracts ────────────────────────────────────────
    req_pattern = re.compile(
        r"([A-Z]\d{8})\s+"                          # code
        r"(Vehicle Loan|Personal Loan|Salary Loan"
        r"|Housing Loan|Agricultural Loan"
        r"|Micro(?:finance)? Loan|[A-Za-z ]+Loan)\s+"
        r"(Requested|Refused|Renounced)\s+"
        r"(Borrower|Co-Borrower|Guarantor)\s+"
        r"(\d{2}/\d{2}/\d{4})",                    # request date
        re.I,
    )
    for m in req_pattern.finditer(text):
        contracts.append({
            "cic_contract_code": m.group(1),
            "contract_type":     m.group(2).strip(),
            "financed_amount":   0.0,
            "outstanding":       0.0,
            "overdue_amount":    0.0,
            "start_date":        m.group(5),
            "end_date":          "",
            "phase":             m.group(3).lower(),
            "role":              m.group(4),
        })

    # ── Delinquency history from detail blocks ────────────────────────────────
    # Look for "Past Due" status lines
    past_due_blocks = re.findall(
        r"(\d{4}/\d{1,2})[\s\S]{0,60}?"
        r"(\d+-\d+\s+days\s+delay[^\n]*|91[-–]180\s+days[^\n]*|"
        r"more\s+than\s+3\s+cycles[^\n]*)[\s\S]{0,20}?"
        r"(Past\s+Due)",
        text, re.I,
    )
    for blk in past_due_blocks:
        # Attach to last contract if available
        if contracts:
            c = contracts[-1]
            if "delinquency_history" not in c:
                c["delinquency_history"] = []
            c["delinquency_history"].append({
                "period":   blk[0],
                "delay":    blk[1].strip(),
                "status":   "Past Due",
            })

    return contracts


# ── CREDIT CARDS ──────────────────────────────────────────────────────────────

def _parse_credit_cards(text: str) -> list[dict]:
    cards = []

    # From the Credit Cards summary table
    card_pattern = re.compile(
        r"(P\d{8}|R\d{8})\s+"                      # CIC contract code
        r"\d+\s+"                                   # card reference no.
        r"Credit\s+Card\s+"
        r"([\d,]+)\s+"                              # credit limit
        r"([\d,]+)\s+"                              # outstanding balance
        r"([\d,]+)\s+"                              # overdue payments amount
        r"([\d,]+)",                                # outstanding balance unbilled
        re.I,
    )
    for m in card_pattern.finditer(text):
        cards.append({
            "cic_contract_code":     m.group(1),
            "credit_limit":          _amount(m.group(2)),
            "outstanding_balance":   _amount(m.group(3)),
            "overdue_amount":        _amount(m.group(4)),
            "unbilled_balance":      _amount(m.group(5)),
        })

    # If the table pattern didn't match, try detail-block extraction
    if not cards:
        detail_blocks = re.split(r"Detail\s+of\s+Credit\s+Card\s+\d+", text, flags=re.I)
        for block in detail_blocks[1:]:
            c: dict[str, Any] = {}
            c["credit_limit"]        = _amount(_find(r"credit\s+limit[\s:]+([\d,]+)", block))
            c["outstanding_balance"] = _amount(_find(r"outstanding\s+balance[\s:]+([\d,]+)", block))
            c["overdue_amount"]      = _amount(_find(r"overdue\s+payments\s+amount[\s:]+([\d,]+)", block))
            c["unbilled_balance"]    = _amount(_find(r"outstanding\s+balance\s*[-–]\s*unbilled[\s:]+([\d,]+)", block))
            c["min_payment_status"]  = _find(r"min\s+payment\s+indicator[\s:]+([^\n]+)", block)
            c["overdue_days"]        = _find(r"overdue\s+days[\s:]+([^\n]+)", block)
            c["contract_type"]       = _find(r"transaction\s+type[\s:]+([^\n]+)", block)
            if c["credit_limit"] > 0:
                cards.append(c)

    return cards


# ── LEGAL FLAGS ───────────────────────────────────────────────────────────────

def _parse_legal_flags(text: str) -> list[dict]:
    flags = []

    # "A legal action has been taken" blocks
    legal_pattern = re.compile(
        r"(\d{2}/\d{2}/\d{4})\s+"                  # event date
        r"(A\s+legal\s+action\s+has\s+been\s+taken|"
        r"legal\s+action[^\n]*)\s+"
        r"(LEGAL\s+ACTION|[^\n]+)",
        re.I,
    )
    for m in legal_pattern.finditer(text):
        flags.append({
            "event_date":   m.group(1),
            "event_type":   m.group(2).strip(),
            "event_detail": m.group(3).strip(),
        })

    # Central Bureau of Financial Risks mention
    if re.search(r"central\s+bureau\s+of\s+financial\s+risks", text, re.I):
        flags.append({
            "event_date":   "",
            "event_type":   "Bureau Listing",
            "event_detail": "Subject is linked to Central Bureau of Financial Risks",
        })

    return flags


# ── RISK SUMMARY ──────────────────────────────────────────────────────────────

def _build_risk_summary(data: dict) -> dict:
    """Derive high-level risk metrics from parsed CIC data."""
    rs: dict[str, Any] = {}

    emp = data["employment"]
    contracts = data["contracts"]
    cards = data["credit_cards"]
    legal = data["legal_flags"]

    # Monthly income
    rs["monthly_income"] = emp.get("monthly_income", 0.0)

    # Total outstanding debt (loans + cards)
    loan_outstanding = sum(c.get("outstanding", 0.0) for c in contracts
                           if c.get("phase") == "active")
    card_outstanding = sum(c.get("outstanding_balance", 0.0) for c in cards)
    rs["total_outstanding_debt"] = loan_outstanding + card_outstanding

    # Total monthly obligations (rough estimate: overdue + min card payments)
    loan_overdue   = sum(c.get("overdue_amount", 0.0) for c in contracts)
    card_overdue   = sum(c.get("overdue_amount", 0.0) for c in cards)
    rs["total_overdue"] = loan_overdue + card_overdue

    # Credit utilisation (cards)
    total_limit = sum(c.get("credit_limit", 0.0) for c in cards)
    total_card_bal = sum(c.get("outstanding_balance", 0.0) for c in cards)
    rs["credit_utilisation_pct"] = (
        round(total_card_bal / total_limit * 100, 1) if total_limit > 0 else None
    )

    # Delinquency level
    has_past_due_91_plus = any(
        "91" in str(entry.get("delay", "")) or "more than 3 cycle" in str(entry.get("delay", "")).lower()
        for c in contracts
        for entry in c.get("delinquency_history", [])
    )
    has_any_overdue = rs["total_overdue"] > 0
    rs["has_severe_delinquency"] = has_past_due_91_plus
    rs["has_any_overdue"]        = has_any_overdue

    # Legal flag
    rs["has_legal_action"]       = len([f for f in legal
                                        if "legal action" in f.get("event_type", "").lower()]) > 0
    rs["has_bureau_listing"]     = len([f for f in legal
                                        if "bureau" in f.get("event_detail", "").lower()]) > 0

    # Number of active loans / cards
    rs["active_loan_count"]   = sum(1 for c in contracts if c.get("phase") == "active")
    rs["active_card_count"]   = len(cards)
    rs["requested_loan_count"]= sum(1 for c in contracts if c.get("phase") == "requested")

    # Employment tenure
    rs["years_of_service"] = emp.get("years_of_service")

    # Overall risk tier
    if rs["has_legal_action"] or rs["has_severe_delinquency"] or rs["has_bureau_listing"]:
        rs["risk_tier"] = "HIGH"
    elif rs["has_any_overdue"] or (rs["credit_utilisation_pct"] or 0) > 70:
        rs["risk_tier"] = "MODERATE"
    else:
        rs["risk_tier"] = "LOW"

    return rs


# ══════════════════════════════════════════════════════════════════════════════
#  FORMATTED SUMMARY (for display in Extracted Text tab)
# ══════════════════════════════════════════════════════════════════════════════

def format_cic_for_analysis(text: str) -> str:
    """
    Return a human-readable enriched summary of a CIC report that the
    analysis engine (Groq) can consume directly.
    """
    if not is_cic_report(text):
        return text  # pass through unchanged if not a CIC report

    d = parse_cic_report(text)
    s  = d["subject"]
    emp= d["employment"]
    rs = d["risk_summary"]
    contracts = d["contracts"]
    cards     = d["credit_cards"]
    legal     = d["legal_flags"]

    lines: list[str] = []
    add = lines.append

    add("═" * 62)
    add("  CIC CREDIT REPORT — STRUCTURED EXTRACT")
    add("  (Parsed by DocExtract Pro / BSV AI-OCR)")
    add("═" * 62)

    # ── Subject ───────────────────────────────────────────────────────────
    add("\n1. BORROWER PROFILE")
    add("─" * 44)
    add(f"  Full Name        : {s.get('full_name', 'N/A')}")
    add(f"  CIC Subject Code : {s.get('cic_subject_code', 'N/A')}")
    add(f"  Date of Birth    : {s.get('date_of_birth', 'N/A')}")
    add(f"  Gender           : {s.get('gender', 'N/A')}")
    add(f"  Civil Status     : {s.get('civil_status', 'N/A')}")
    add(f"  Nationality      : {s.get('nationality', 'N/A')}")
    add(f"  Resident         : {s.get('resident', 'N/A')}")
    add(f"  Dependents       : {s.get('dependents', 'N/A')}")
    add(f"  Spouse           : {s.get('spouse', 'N/A')}")
    add(f"  TIN              : {s.get('tin', 'N/A')}")
    add(f"  SSS Card         : {s.get('sss', 'N/A')}")
    add(f"  Address          : {s.get('address', 'N/A')}")

    # ── Employment ────────────────────────────────────────────────────────
    add("\n2. EMPLOYMENT & INCOME")
    add("─" * 44)
    add(f"  Employer         : {emp.get('employer', 'N/A')}")
    add(f"  Industry         : {emp.get('industry', 'N/A')}")
    add(f"  Occupation       : {emp.get('occupation', 'N/A')}")
    add(f"  Status           : {emp.get('occupation_status', 'N/A')}")
    add(f"  Gross Income     : {emp.get('income_frequency','Annual')} ₱{emp.get('gross_income',0):,.0f}")
    add(f"  Monthly Income   : ₱{emp.get('monthly_income',0):,.2f}")
    add(f"  Hired From       : {emp.get('hired_from', 'N/A')}")
    yos = emp.get("years_of_service")
    yos_str = f"{yos} yrs" if yos else "N/A"
    add(f"  Years of Service : {yos_str}")

    # ── Loans / Installments ──────────────────────────────────────────────
    add("\n3. INSTALLMENT CONTRACTS")
    add("─" * 44)
    if contracts:
        for i, c in enumerate(contracts, 1):
            add(f"  [{i}] {c.get('contract_type','').upper()}  |  Phase: {c.get('phase','').upper()}")
            add(f"      Code            : {c.get('cic_contract_code','N/A')}")
            if c.get('financed_amount'):
                add(f"      Financed Amount : ₱{c['financed_amount']:,.0f}")
            if c.get('outstanding'):
                add(f"      Outstanding Bal : ₱{c['outstanding']:,.0f}")
            if c.get('overdue_amount'):
                add(f"      Overdue Amount  : ₱{c['overdue_amount']:,.0f}  ⚠")
            if c.get('start_date'):
                add(f"      Start / End     : {c.get('start_date')} → {c.get('end_date','–')}")
            for dh in c.get("delinquency_history", []):
                add(f"      ⚠  Delinquent   : {dh['period']}  |  {dh['delay']}  |  {dh['status']}")
    else:
        add("  No installment contracts found.")

    # ── Credit Cards ──────────────────────────────────────────────────────
    add("\n4. CREDIT CARDS")
    add("─" * 44)
    if cards:
        for i, c in enumerate(cards, 1):
            add(f"  [{i}] Code: {c.get('cic_contract_code','N/A')}")
            add(f"      Credit Limit      : ₱{c.get('credit_limit',0):,.0f}")
            add(f"      Outstanding Bal   : ₱{c.get('outstanding_balance',0):,.0f}")
            add(f"      Overdue Amount    : ₱{c.get('overdue_amount',0):,.0f}")
            add(f"      Unbilled Balance  : ₱{c.get('unbilled_balance',0):,.0f}")
            if c.get("min_payment_status"):
                add(f"      Min Payment       : {c['min_payment_status']}")
            if c.get("overdue_days"):
                add(f"      Overdue Days      : {c['overdue_days']}")
    else:
        add("  No credit card records found.")

    # ── Legal / Adverse ───────────────────────────────────────────────────
    add("\n5. LEGAL & ADVERSE INFORMATION")
    add("─" * 44)
    if legal:
        for f in legal:
            add(f"  ❌ {f.get('event_date','')}  |  {f.get('event_type','')}  |  {f.get('event_detail','')}")
    else:
        add("  ✅  No legal actions or adverse information found.")

    # ── Risk Summary ──────────────────────────────────────────────────────
    add("\n6. RISK ASSESSMENT SUMMARY")
    add("─" * 44)
    tier_sym = {"HIGH": "❌", "MODERATE": "⚠", "LOW": "✅"}.get(rs.get("risk_tier",""), "•")
    add(f"  Overall Risk Tier       : {tier_sym}  {rs.get('risk_tier','UNKNOWN')}")
    add(f"  Monthly Income          : ₱{rs.get('monthly_income',0):,.2f}")
    add(f"  Total Outstanding Debt  : ₱{rs.get('total_outstanding_debt',0):,.2f}")
    add(f"  Total Overdue Amount    : ₱{rs.get('total_overdue',0):,.2f}")
    util = rs.get("credit_utilisation_pct")
    util_str = f"{util}%" if util is not None else "N/A"
    add(f"  Credit Card Utilisation : {util_str}")
    add(f"  Active Loan Count       : {rs.get('active_loan_count',0)}")
    add(f"  Active Card Count       : {rs.get('active_card_count',0)}")
    add(f"  Severe Delinquency      : {'YES ⚠' if rs.get('has_severe_delinquency') else 'No'}")
    add(f"  Any Overdue             : {'YES ⚠' if rs.get('has_any_overdue') else 'No'}")
    add(f"  Legal Action on Record  : {'YES ❌' if rs.get('has_legal_action') else 'No'}")
    add(f"  Bureau Listing          : {'YES ❌' if rs.get('has_bureau_listing') else 'No'}")
    tenure = rs.get("years_of_service")
    tenure_str = f"{tenure} yrs" if tenure else "N/A"
    add(f"  Employment Tenure       : {tenure_str}")

    add("\n" + "═" * 62)
    add("  END OF CIC STRUCTURED EXTRACT")
    add("═" * 62)

    # Append the full raw text below so the AI can also reference it
    add("\n\n── RAW EXTRACTED TEXT (for reference) ──\n")
    add(text)

    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
#  ANALYSIS PROMPT INJECTION BLOCK
# ══════════════════════════════════════════════════════════════════════════════

CIC_ANALYSIS_PROMPT_BLOCK = """
════════════════════════════════════════════════════
  CIC CREDIT REPORT — ANALYSIS INSTRUCTIONS
════════════════════════════════════════════════════

The document provided is a Philippine CIC (Credit Information Corporation)
Credit Report. You must analyse it as a credit officer would when evaluating
a loan application. Follow these specific instructions:

DOCUMENT SECTIONS TO ANALYSE:
  1. Borrower Profile — verify completeness, cross-check IDs (TIN, SSS)
  2. Employment & Income — assess stability, tenure, income level
  3. Installment Contracts — review all active, closed, requested, and refused loans
  4. Credit Cards — check utilisation rate, overdue amounts, min-payment behaviour
  5. Legal / Adverse Info — flag any legal actions, bureau listings, or derogatory info
  6. Delinquency History — assess payment patterns (days overdue, cycle lates)

KEY METRICS TO COMPUTE:
  - Debt Service Ratio (DSR) = Total monthly obligations ÷ Monthly net income
    (If DSR > 0.35 → manageable; > 0.50 → high risk; > 0.70 → decline territory)
  - Credit Card Utilisation = Outstanding balance ÷ Total credit limit
    (< 30% healthy; 30–70% moderate; > 70% concerning)
  - Delinquency Bucket:
      Current = no overdue
      Stage 1 = 1–30 days overdue
      Stage 2 = 31–90 days overdue
      Stage 3 = 91–180 days (PAST DUE)
      Stage 4 = 180+ days (LOSS / NPA territory)

RED FLAGS (auto-escalate to DECLINE or CONDITIONAL):
  ❌ Legal action on record
  ❌ Central Bureau of Financial Risks listing
  ❌ 91–180 days past due ("more than 3 cycles late")
  ❌ Multiple consecutive overdue months
  ❌ Payment below minimum on credit cards
  ❌ Refused/declined loan in recent history
  ⚠  DSR > 50%
  ⚠  Credit utilisation > 70%
  ⚠  Requested loan recently not yet funded

POSITIVE FACTORS:
  ✅ Long employment tenure (10+ years = strong)
  ✅ High income relative to obligations
  ✅ Real estate collateral
  ✅ Co-borrower or guarantor present
  ✅ No legal adverse info

OUTPUT FORMAT — use the standard BSV credit analysis report format:
  1. Document Type & Verification
  2. Borrower Profile
  3. Income & Employment Analysis
  4. Existing Credit Obligations
  5. Payment Behaviour & Delinquency Analysis
  6. Credit Card Analysis
  7. Legal & Adverse Information
  8. Risk Flags
  9. Credit Score Estimate (300–900 scale)
  10. Recommended BSV Loan Products & Eligibility
  11. Final Verdict: APPROVE / CONDITIONALLY APPROVE / DECLINE
  12. Conditions / Recommendations

Always reference specific figures from the CIC report to support your conclusions.
Never fabricate data not present in the document.
"""
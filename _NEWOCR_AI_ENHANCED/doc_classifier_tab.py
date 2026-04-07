"""
doc_classifier_tab.py — DocExtract Pro
========================================
Replaces the plain "Extracted" tab with a smart document classifier panel.

Features
--------
  • Auto-detects document type from 14+ BSV-specific categories
  • Shows a colored tag/label card at the top
  • Extracts & highlights key financial figures
  • Shows structured fields specific to the document type
  • Auto-routes to the correct CIBI slot
  • Raw text still accessible via toggle

Document Categories
-------------------
  FINANCIAL     : Bank Statement, Receipt
  TRANSPORT     : TODA (ORCR)
  PROPERTY      : Rent / Lease
  REMITTANCE    : GCash, Maya, Padala, Online Transfer
  SPOUSE_INCOME : Spouse remittance (GCash, Maya, Padala)
  FARMING       : Copra, Palay, Pinya, other crops
  LIVESTOCK     : Chicken, Piggery, Cows
  PAYSLIP       : Payslip / Payroll
  ITR           : Income Tax Return
  SALN          : SALN / Net Worth
  CIC           : CIC Credit Report
  BANK_CI       : Bank CI Certification
  GENERAL       : Any other document

CIBI Slot Routing
-----------------
  PAYSLIP  → cibi_slots["PAYSLIP"]
  ITR      → cibi_slots["ITR"]
  SALN     → cibi_slots["SALN"]
  CIC      → cibi_slots["CIC"]
  BANK_CI  → cibi_slots["BANK_CI"]
  others   → no auto-route (shown as info only)

FIXES APPLIED
-------------
- Duplicate _clf_show_placeholder removed; the full pill-grid version is
  kept and _clf_clear_panel now calls _clf_show_placeholder correctly.
- self.after() calls that passed a dict as a positional arg replaced with
  lambda wrappers so config() is called properly.
- Dead code: unused `text_inline` variable in extract_fields() removed.
- Imports (customtkinter, threading, os, filedialog, PIL) moved to
  module-level to avoid repeated in-function import overhead and to make
  the dependency tree visible. Guard blocks handle ImportError gracefully.
- FIX: _toggle_raw in _clf_section_bank_risk_result now swaps the canvas
  out and shows _clf_raw_frame so "View Full Report" actually works.
"""
from __future__ import annotations

import re
import threading
import os
import tkinter as tk
from pathlib import Path
from typing import Optional
import json as _json

# Optional dependencies — imported at module level with graceful fallbacks
try:
    import customtkinter as ctk
    _HAS_CTK = True
except ImportError:
    _HAS_CTK = False

try:
    from tkinter import filedialog as _filedialog
    _HAS_FILEDIALOG = True
except ImportError:
    _HAS_FILEDIALOG = False

try:
    from bank_statement_risk import (
        assess_bank_statement,
        bank_statement_risk_to_text,
        BankStatementRiskResult,
        VERDICT_GOOD, VERDICT_BAD, VERDICT_UNCERTAIN,
    )
    _HAS_BANK_RISK = True
except ImportError:
    _HAS_BANK_RISK = False

# ── Colour palette (mirrors app.py constants) ─────────────────────────────────
NAVY_DEEP    = "#0A1628"
NAVY         = "#112240"
NAVY_MID     = "#1B3A6B"
NAVY_LIGHT   = "#2D5FA6"
NAVY_PALE    = "#5B8FD4"
NAVY_GHOST   = "#C5D8F5"
NAVY_MIST    = "#EAF1FB"
WHITE        = "#FFFFFF"
OFF_WHITE    = "#F7F9FC"
CARD_WHITE   = "#FFFFFF"
LIME_BRIGHT  = "#00E5A0"
LIME         = "#00C48A"
LIME_MID     = "#00A876"
LIME_DARK    = "#007A56"
LIME_PALE    = "#B3F5E2"
LIME_MIST    = "#E6FBF5"
BORDER_LIGHT = "#DDE6F4"
BORDER_MID   = "#B8CCEA"
TXT_NAVY     = "#0A1628"
TXT_NAVY_MID = "#1B3A6B"
TXT_SOFT     = "#5B7BAD"
TXT_MUTED    = "#96AFCC"
ACCENT_GOLD    = "#F59E0B"
ACCENT_SUCCESS = "#00C48A"
ACCENT_RED     = "#EF4444"

# ── Document type definitions ─────────────────────────────────────────────────

DOC_TYPES = {
    # key: (label, icon, color_bg, color_fg, cibi_slot, keywords)
    "BANK_STATEMENT": (
        "Bank Statement", "🏦", "#EFF6FF", "#1D4ED8",
        None,
        ["account number", "account no", "balance", "transaction", "deposit",
         "withdrawal", "debit", "credit", "statement of account", "soa",
         "available balance", "current balance", "savings account",
         "checking account", "bank statement"],
    ),
    "RECEIPT": (
        "Receipt / Official Receipt", "🧾", "#FFF7ED", "#C2410C",
        None,
        ["official receipt", "o.r. no", "or no", "receipt no", "received from",
         "amount received", "payment for", "official receipt number",
         "tin:", "vat", "non-vat", "acknowledgment receipt"],
    ),
    "TODA_ORCR": (
        "TODA — ORCR", "🚗", "#F0FDF4", "#15803D",
        None,
        ["orcr", "official receipt", "certificate of registration",
         "cr no", "plate number", "plate no", "mv file no",
         "toda", "tricycle", "motor vehicle", "land transportation",
         "lto", "engine number", "chassis number"],
    ),
    "RENT": (
        "Rent / Lease Agreement", "🏠", "#FDF4FF", "#7E22CE",
        None,
        ["rent", "lease", "lessor", "lessee", "monthly rental",
         "rental fee", "contract of lease", "rental agreement",
         "monthly rent", "rental income", "space rental"],
    ),
    "GCASH": (
        "GCash Remittance", "📱", "#F0FDF4", "#166534",
        None,
        ["gcash", "g-cash", "gcash ref", "reference number", "mobile money",
         "send money", "gcash transaction", "gcash receipt",
         "transaction id", "globe", "gcash wallet"],
    ),
    "MAYA": (
        "Maya / PayMaya Remittance", "💳", "#EEF2FF", "#4338CA",
        None,
        ["maya", "paymaya", "pay maya", "maya wallet", "maya transaction",
         "maya receipt", "maya ref", "maya reference"],
    ),
    "PADALA": (
        "Padala / Remittance Center", "📦", "#FFFBEB", "#B45309",
        None,
        ["padala", "remittance", "money transfer", "western union",
         "palawan express", "lbc", "cebuana", "mlhuillier", "m lhuillier",
         "jrs express", "sender", "receiver", "control number",
         "padala center", "remittance center"],
    ),
    "ONLINE_TRANSFER": (
        "Online Bank Transfer", "💻", "#F0F9FF", "#0369A1",
        None,
        ["instapay", "pesonet", "online transfer", "fund transfer",
         "bank transfer", "bdo", "bpi", "metrobank", "landbank",
         "unionbank", "security bank", "transaction reference",
         "transferred to", "transferred from"],
    ),
    "SPOUSE_REMITTANCE": (
        "Spouse Remittance", "👫", "#FFF1F2", "#BE123C",
        None,
        ["spouse", "husband", "wife", "asawa", "ofw", "remittance",
         "send money", "overseas", "abroad", "monthly remittance",
         "family allotment"],
    ),
    "FARMING": (
        "Farming / Crop Income", "🌾", "#F7FEE7", "#3F6212",
        None,
        ["copra", "palay", "pinya", "farmgate", "farm produce",
         "rice", "pineapple", "corn", "mais",
         "sugarcane", "tubo", "banana", "saging", "cassava", "kamoteng kahoy",
         "mango", "mangga", "vegetable", "gulay", "harvest", "ani",
         "crop", "farm produce receipt", "farmgate price",
         "kilo", "cavan", "sack", "cavans", "farm", "agricultural",
         "price per kg", "per cavan", "per kilo"],
    ),
    "LIVESTOCK": (
        "Livestock Income", "🐄", "#FFF9F0", "#92400E",
        None,
        ["chicken", "manok", "piggery", "baboy", "pig", "hog",
         "cow", "baka", "cattle", "goat", "kambing", "carabao",
         "kalabaw", "duck", "pato", "livestock", "animal",
         "poultry", "head count", "live weight", "dressed"],
    ),
    "PAYSLIP": (
        "Payslip / Payroll", "💵", "#F0FDF4", "#166534",
        "PAYSLIP",
        ["payslip", "pay slip", "payroll", "basic pay", "gross pay",
         "net pay", "deductions", "sss", "philhealth", "pagibig",
         "withholding tax", "employee id", "period covered",
         "salary", "allowance", "overtime"],
    ),
    "ITR": (
        "Income Tax Return (ITR)", "📊", "#EFF6FF", "#1E40AF",
        "ITR",
        ["income tax return", "itr", "bir form", "bir 2316",
         "annual income", "taxable income", "tax due",
         "bureau of internal revenue", "tin no", "tin number",
         "employer tin", "compensation income"],
    ),
    "SALN": (
        "SALN / Net Worth Statement", "📄", "#F8FAFC", "#334155",
        "SALN",
        ["saln", "statement of assets", "net worth", "liabilities",
         "real property", "personal property", "total assets",
         "total liabilities", "net worth", "assets and liabilities"],
    ),
    "CIC": (
        "CIC Credit Report", "📋", "#EFF6FF", "#1E40AF",
        "CIC",
        ["cic", "credit information corporation", "credit report",
         "credit score", "credit history", "loan account",
         "credit card", "outstanding balance", "past due",
         "credit standing"],
    ),
    "BANK_CI": (
        "Bank CI Certification", "🏦", "#F0FDF4", "#166534",
        "BANK_CI",
        ["bank certification", "bank ci", "credit investigation",
         "no contrary data", "ncd", "no adverse", "good standing",
         "certify", "certified", "informant", "bank manager",
         "branch manager", "to whom it may concern"],
    ),
}

# ── Field definitions per document type ──────────────────────────────────────
FIELD_DEFS: dict[str, list[tuple]] = {
    "BANK_STATEMENT": [
        ("Account Name",    r"(?:account\s+name\s*[:\-]\s*([^\n]{3,60})|^([A-Z][A-Z\s\.]+?)(?=\s+BRGY|\s+PUROK|\s+LOT|\s+BLOCK|\s+\d{3,}|\s+FOR\s|\s+ACCOUNT)|^([A-Z][A-Z\s\.]{4,40})$)", "👤"),
        ("Account Number",  r"(?:account\s+(?:no|number))[\s\n]*([\d\-]{6,})", "🔢"),
        ("Bank Name",       r"(?i)\b(BDO|BPI|Metrobank|Landbank|UnionBank|Security\s*Bank|PNB|RCBC|Chinabank|EastWest|PSBank|AlliedBank|DBP|UCPB)\b", "🏦"),
        ("Period",          r"(?:for\s+([A-Za-z]+\s+\d+\s*[-\u2013]\s*[A-Za-z]*\s*\d+,?\s*\d{4})|(?:period|statement\s+period)\s*[:\-]\s*([^\n]+))", "📅"),
        ("Closing Balance",  r"(?:(?:account\s+summary[^\n]*?)balance\s*[P₱]\s*([\d,]+\.\d{2})|(?:closing|ending|available)\s+balance[^\d\n]*([\d,]+\.\d{2}))", "💰"),
        ("Total Deposits",   r"(?:deposits?|total\s+(?:credit|deposit)s?)\s*[P₱]\s*([\d,]+(?:\.\d+)?)", "⬆"),
        ("Total Withdrawals",r"(?:withdrawals?|total\s+(?:debit|withdrawal)s?)\s*[P₱]\s*([\d,]+(?:\.\d+)?)", "⬇"),
    ],
    "RECEIPT": [
        ("OR Number",     r"(?:o\.?r\.?\s*(?:no|number)|receipt\s+no)\s*[:\-]?\s*([\w\-]+)",  "🔢"),
        ("Date",          r"(?:date)\s*[:\-]\s*([^\n]+)",                                      "📅"),
        ("Received From", r"(?:received\s+from|payor)\s*[:\-]\s*([^\n]+)",                     "👤"),
        ("Amount",        r"(?:amount|total)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)",             "💰"),
        ("Purpose",       r"(?:payment\s+for|purpose|particulars)\s*[:\-]\s*([^\n]+)",         "📝"),
        ("Issued By",     r"(?:issued\s+by|cashier|received\s+by)\s*[:\-]\s*([^\n]+)",         "✍"),
    ],
    "TODA_ORCR": [
        ("Plate Number",   r"(?:plate\s+(?:no|number))\s*[:\-]\s*([\w\s\-]+)",                "🚗"),
        ("MV File No",     r"(?:mv\s+file\s+no|file\s+number)\s*[:\-]\s*([\w\-]+)",           "🔢"),
        ("OR Number",      r"(?:o\.?r\.?\s*(?:no|number))\s*[:\-]?\s*([\w\-]+)",              "📋"),
        ("CR Number",      r"(?:c\.?r\.?\s*(?:no|number))\s*[:\-]?\s*([\w\-]+)",              "📋"),
        ("Owner",          r"(?:owner|registered\s+owner)\s*[:\-]\s*([^\n]+)",                 "👤"),
        ("Vehicle Type",   r"(?:body\s+type|vehicle\s+type|make|motor\s+vehicle)\s*[:\-]\s*([^\n]+)", "🚌"),
        ("Expiry Date",    r"(?:expiry|expiration|valid\s+until|validity)\s*[:\-]\s*([^\n]+)", "📅"),
    ],
    "RENT": [
        ("Lessor",          r"(?:lessor|landlord|owner)\s*[:\-]\s*([^\n]+)",                   "👤"),
        ("Lessee",          r"(?:lessee|tenant|renter)\s*[:\-]\s*([^\n]+)",                    "👤"),
        ("Monthly Rent",    r"(?:monthly\s+rent(?:al)?|rental\s+(?:fee|amount|rate))\s*[:\-]?\s*[P₱]?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Property Address",r"(?:property|address|location|premises)\s*[:\-]\s*([^\n]+)",      "🏠"),
        ("Lease Period",    r"(?:lease\s+period|term|duration)\s*[:\-]\s*([^\n]+)",            "📅"),
        ("Contract Date",   r"(?:date|dated|contract\s+date)\s*[:\-]\s*([^\n]+)",              "📅"),
    ],
    "GCASH": [
        ("Reference No",    r"(?:ref(?:erence)?\s*(?:no|number)?|transaction\s+id)\s*[:\-]?\s*([\w\d]+)", "🔢"),
        ("Amount",          r"(?:amount|total|php)\s*[:\-]?\s*[P₱]?\s*([\d,]+(?:\.\d+)?)",       "💰"),
        ("Sender",          r"(?:from|sender|sent\s+by)\s*[:\-]\s*([^\n]+)",                   "👤"),
        ("Receiver",        r"(?:to|receiver|recipient)\s*[:\-]\s*([^\n]+)",                   "👤"),
        ("Date & Time",     r"(?:date|datetime|transaction\s+date)\s*[:\-]\s*([^\n]+)",        "📅"),
        ("Mobile Number",   r"(?:mobile|number|contact)\s*[:\-]?\s*(09\d{9}|\+639\d{9})",     "📱"),
    ],
    "MAYA": [
        ("Reference No",    r"(?:ref(?:erence)?\s*(?:no|number)?|transaction\s+id)\s*[:\-]?\s*([\w\d]+)", "🔢"),
        ("Amount",          r"(?:amount|total|php)\s*[:\-]?\s*[P₱]?\s*([\d,]+(?:\.\d+)?)",       "💰"),
        ("Sender",          r"(?:from|sender)\s*[:\-]\s*([^\n]+)",                             "👤"),
        ("Receiver",        r"(?:to|receiver|recipient)\s*[:\-]\s*([^\n]+)",                   "👤"),
        ("Date & Time",     r"(?:date|datetime)\s*[:\-]\s*([^\n]+)",                           "📅"),
    ],
    "PADALA": [
        ("Control Number",  r"(?:control\s+(?:no|number)|tracking|claim\s+(?:no|number))\s*[:\-]?\s*([\w\d][\w\d\-]+)", "🔢"),
        ("Amount",          r"(?:amount|total|php)\s*[:\-]?\s*[P₱]?\s*([\d,]+(?:\.\d+)?)",       "💰"),
        ("Sender",          r"(?:sender|from|sent\s+by)\s*[:\-]\s*([^\n]+)",                   "👤"),
        ("Receiver",        r"(?:receiver|beneficiary|recipient|to)\s*[:\-]\s*([^\n]+)",       "👤"),
        ("Center",          r"(?:outlet|branch|center|agent)\s*[:\-]\s*([^\n]+)",              "🏪"),
        ("Date",            r"(?:date|transaction\s+date)\s*[:\-]\s*([^\n]+)",                 "📅"),
    ],
    "ONLINE_TRANSFER": [
        ("Reference No",    r"(?:ref(?:erence)?\s*(?:no|number)?|transaction\s+(?:id|ref))\s*[:\-]?\s*([\w\d]+)", "🔢"),
        ("Amount",          r"(?:amount|transfer\s+amount|php)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("From Account",    r"(?:from|debit(?:ed)?\s+(?:account|acct))\s*[:\-]\s*([^\n]+)",   "🏦"),
        ("To Account",      r"(?:to|credit(?:ed)?\s+(?:account|acct))\s*[:\-]\s*([^\n]+)",    "🏦"),
        ("Date & Time",     r"(?:date|datetime|value\s+date)\s*[:\-]\s*([^\n]+)",              "📅"),
        ("Bank",            r"(?:bank|institution)\s*[:\-]\s*([^\n]+)",                        "🏦"),
    ],
    "SPOUSE_REMITTANCE": [
        ("Spouse Name",     r"(?:name|sender|from)\s*[:\-]\s*([^\n]+)",                        "👤"),
        ("Amount",          r"(?:amount|php|remittance)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)",  "💰"),
        ("Reference No",    r"(?:ref(?:erence)?\s*(?:no|number)?|control\s+no)\s*[:\-]?\s*([\w\d]+)", "🔢"),
        ("Date",            r"(?:date|transaction\s+date)\s*[:\-]\s*([^\n]+)",                 "📅"),
        ("Channel",         r"(?:via|through|channel|sent\s+via)\s*[:\-]\s*([^\n]+)",         "📡"),
        ("Frequency",       r"(?:frequency|monthly|weekly|every)\s*[:\-]\s*([^\n]+)",         "🔄"),
    ],
    "FARMING": [
        ("Crop Type",       r"(?:crop|produce|commodity|item)\s*[:\-]\s*([^\n]+)",             "🌾"),
        ("Quantity",        r"(?:quantity|qty|weight|volume|cavans?|sacks?|kilos?|kg)\s*[:\-]?\s*([\d,]+(?:\.\d+)?(?:\s*(?:kg|kgs|cavan|sack|pc))?)", "📦"),
        ("Price per Unit",  r"(?:price\s+per|unit\s+price|farmgate\s+price|per\s+(?:kg|kilo|cavan|sack))\s*[:\-]?\s*[P₱]?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Total Amount",    r"(?:total|gross|amount)\s*[:\-]?\s*[P₱]?\s*([\d,]+(?:\.\d+)?)",     "💵"),
        ("Buyer/Trader",    r"(?:buyer|trader|purchased\s+by|sold\s+to)\s*[:\-]\s*([^\n]+)",  "👤"),
        ("Date of Sale",    r"(?:date|sale\s+date|sold\s+on)\s*[:\-]\s*([^\n]+)",             "📅"),
        ("Seller",          r"(?:seller|grower|farmer|sold\s+by)\s*[:\-]\s*([^\n]+)",         "👨‍🌾"),
    ],
    "LIVESTOCK": [
        ("Animal Type",     r"(?:animal|type|species|kind)\s*[:\-]\s*([^\n]+)",                "🐄"),
        ("Head Count",      r"(?:head(?:s)?|count|quantity|number\s+of)\s*[:\-]?\s*(\d+(?:\s*(?:head|pcs|units))?)", "🔢"),
        ("Live Weight",     r"(?:live\s+weight|liveweight|weight)\s*[:\-]?\s*([\d,]+(?:\.\d+)?(?:\s*kg)?)", "⚖"),
        ("Price per kg",    r"(?:price\s+per\s+(?:kg|kilo)|per\s+(?:kilo|kg|head))\s*[:\-]?\s*[P₱]?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Total Amount",    r"(?:total|gross|amount|proceeds)\s*[:\-]?\s*[P₱]?\s*([\d,]+(?:\.\d+)?)", "💵"),
        ("Buyer",           r"(?:buyer|purchased\s+by|sold\s+to)\s*[:\-]\s*([^\n]+)",         "👤"),
        ("Date",            r"(?:date|sale\s+date)\s*[:\-]\s*([^\n]+)",                        "📅"),
    ],
    "PAYSLIP": [
        ("Employee Name",   r"(?:employee\s+name|name)\s*[:\-]\s*([^\n]+)",                    "👤"),
        ("Employee ID",     r"(?:employee\s+(?:id|no|number))\s*[:\-]\s*([\w\d\-]+)",         "🔢"),
        ("Period Covered",  r"(?:period\s+covered?|pay\s+period|payroll\s+period)\s*[:\-]\s*([^\n]+)", "📅"),
        ("Basic Pay",       r"(?:basic\s+(?:pay|salary))\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Gross Pay",       r"(?:gross\s+(?:pay|income|salary))\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Total Deductions",r"(?:total\s+deductions?)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)",    "⬇"),
        ("Net Pay",         r"(?:net\s+(?:pay|income|salary))\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💵"),
        ("Employer",        r"(?:employer|company|establishment)\s*[:\-]\s*([^\n]+)",          "🏢"),
    ],
    "SPOUSE_PAYSLIP": [
        ("Employee Name",   r"(?:employee\s+name|name)\s*[:\-]\s*([^\n]+)",                    "👤"),
        ("Employee ID",     r"(?:employee\s+(?:id|no|number))\s*[:\-]\s*([\w\d\-]+)",         "🔢"),
        ("Period Covered",  r"(?:period\s+covered?|pay\s+period|payroll\s+period)\s*[:\-]\s*([^\n]+)", "📅"),
        ("Basic Pay",       r"(?:basic\s+(?:pay|salary))\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Gross Pay",       r"(?:gross\s+(?:pay|income|salary))\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Total Deductions",r"(?:total\s+deductions?)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)",    "⬇"),
        ("Net Pay",         r"(?:net\s+(?:pay|income|salary))\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💵"),
        ("Employer",        r"(?:employer|company|establishment)\s*[:\-]\s*([^\n]+)",          "🏢"),
    ],
    "ITR": [
        ("Taxpayer Name",   r"(?:taxpayer\s+name|name\s+of\s+taxpayer)\s*[:\-]\s*([^\n]+)",   "👤"),
        ("TIN",             r"(?:tin(?:\s+(?:no|number))?)\s*[:\-]?\s*([\d\-]+)",             "🔢"),
        ("Tax Year",        r"(?:taxable\s+year|year|for\s+the\s+year)\s*[:\-]\s*(\d{4})",    "📅"),
        ("Gross Income",    r"(?:gross\s+(?:compensation|income))\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Taxable Income",  r"(?:taxable\s+(?:compensation|income))\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Tax Due",         r"(?:tax\s+due|income\s+tax\s+due)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "📊"),
        ("Employer",        r"(?:employer|withholding\s+agent)\s*[:\-]\s*([^\n]+)",            "🏢"),
    ],
    "SALN": [
        ("Declarant Name",  r"(?:name|declarant)\s*[:\-]\s*([^\n]+)",                          "👤"),
        ("Position",        r"(?:position|designation)\s*[:\-]\s*([^\n]+)",                    "💼"),
        ("Total Assets",    r"(?:total\s+assets?)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)",        "💰"),
        ("Total Liabilities",r"(?:total\s+liabilities)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)",  "📉"),
        ("Net Worth",       r"(?:net\s+worth)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)",            "💵"),
        ("Date Filed",      r"(?:date\s+(?:filed|signed))\s*[:\-]\s*([^\n]+)",                 "📅"),
    ],
    "CIC": [
        ("Borrower Name",   r"(?:full\s+name|name|subject)\s*[:\-]\s*([^\n]+)",                "👤"),
        ("Date of Birth",   r"(?:date\s+of\s+birth|dob|birthday)\s*[:\-]\s*([^\n]+)",         "📅"),
        ("Credit Score",    r"(?:credit\s+score|score)\s*[:\-]?\s*(\d{3})",                    "⭐"),
        ("Active Loans",    r"(?:active\s+loans?|number\s+of\s+(?:loans?|accounts?))\s*[:\-]?\s*(\d+)", "📋"),
        ("Total Balance",   r"(?:total\s+(?:outstanding|balance))\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "💰"),
        ("Overdue Amount",  r"(?:(?:total\s+)?overdue|past\s+due)\s*[:\-]?\s*₱?\s*([\d,]+(?:\.\d+)?)", "⚠"),
        ("Report Date",     r"(?:report\s+date|as\s+of|date)\s*[:\-]\s*([^\n]+)",              "📅"),
    ],
    "BANK_CI": [
        ("Subject Name",    r"(?:subject|name|borrower)\s*[:\-]\s*([^\n]+)",                   "👤"),
        ("Bank Name",       r"(?:bank|institution)\s*[:\-]\s*([^\n]+)",                        "🏦"),
        ("Verdict",         r"(?:ncd|no\s+contrary\s+data|no\s+adverse|no\s+derogatory|dishonored|past\s+due|delinquent)", "⚖"),
        ("Remarks",         r"(?:remarks?)\s*[:\-]\s*([^\n]+)",                                "📝"),
        ("Date",            r"(?:date|dated)\s*[:\-]\s*([^\n]+)",                              "📅"),
        ("Certified By",    r"(?:certified\s+by|authorized\s+by|branch\s+manager)\s*[:\-]\s*([^\n]+)", "✍"),
    ],
}

GENERAL_FIELDS = [
    ("Name / Borrower",  r"(?:name|borrower|applicant)\s*[:\-]\s*([^\n]+)",                    "👤"),
    ("Date",             r"(?:date)\s*[:\-]\s*([^\n]+)",                                       "📅"),
    ("Amount",           r"₱\s*([\d,]+(?:\.\d+)?)",                                           "💰"),
    ("Reference No",     r"(?:ref(?:erence)?\s*(?:no|number)?|control\s+no)\s*[:\-]?\s*([\w\d]+)", "🔢"),
    ("Address",          r"(?:address)\s*[:\-]\s*([^\n]+)",                                    "📍"),
]


# ── VLM schemas & metadata ────────────────────────────────────────────────────

_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp", ".gif"}

_VLM_SCHEMAS: dict[str, dict] = {
    "BANK_STATEMENT": {
        "account_name": None, "account_number": None, "bank_name": None,
        "period": None, "closing_balance": None,
        "total_deposits": None, "total_withdrawals": None,
    },
    "RECEIPT": {
        "or_number": None, "date": None, "received_from": None,
        "amount": None, "purpose": None, "issued_by": None,
    },
    "TODA_ORCR": {
        "plate_number": None, "mv_file_no": None, "or_number": None,
        "cr_number": None, "owner": None, "vehicle_type": None, "expiry_date": None,
    },
    "RENT": {
        "lessor": None, "lessee": None, "monthly_rent": None,
        "property_address": None, "lease_period": None, "contract_date": None,
    },
    "GCASH": {
        "reference_no": None, "amount": None, "sender": None,
        "receiver": None, "date_time": None, "mobile_number": None,
    },
    "MAYA": {
        "reference_no": None, "amount": None, "sender": None,
        "receiver": None, "date_time": None,
    },
    "PADALA": {
        "control_number": None, "amount": None, "sender": None,
        "receiver": None, "center": None, "date": None,
    },
    "ONLINE_TRANSFER": {
        "reference_no": None, "amount": None, "from_account": None,
        "to_account": None, "date_time": None, "bank": None,
    },
    "SPOUSE_REMITTANCE": {
        "spouse_name": None, "amount": None, "reference_no": None,
        "date": None, "channel": None, "frequency": None,
    },
    "FARMING": {
        "crop_type": None, "quantity": None, "price_per_unit": None,
        "total_amount": None, "buyer_trader": None, "date_of_sale": None, "seller": None,
    },
    "LIVESTOCK": {
        "animal_type": None, "head_count": None, "live_weight": None,
        "price_per_kg": None, "total_amount": None, "buyer": None, "date": None,
    },
    "PAYSLIP": {
        "employee_name": None, "employee_id": None, "period_covered": None,
        "basic_pay": None, "gross_pay": None, "total_deductions": None,
        "net_pay": None, "employer": None,
    },
    "SPOUSE_PAYSLIP": {
        "employee_name": None, "employee_id": None, "period_covered": None,
        "basic_pay": None, "gross_pay": None, "total_deductions": None,
        "net_pay": None, "employer": None,
    },
    "ITR": {
        "taxpayer_name": None, "tin": None, "tax_year": None,
        "gross_income": None, "taxable_income": None, "tax_due": None, "employer": None,
    },
    "SALN": {
        "declarant_name": None, "position": None, "total_assets": None,
        "total_liabilities": None, "net_worth": None, "date_filed": None,
    },
    "CIC": {
        "borrower_name": None, "date_of_birth": None, "credit_score": None,
        "active_loans": None, "total_balance": None, "overdue_amount": None, "report_date": None,
    },
    "BANK_CI": {
        "subject_name": None, "bank_name": None, "verdict": None,
        "remarks": None, "date": None, "certified_by": None,
    },
}

_VLM_FIELD_META: dict[str, dict[str, tuple]] = {
    "BANK_STATEMENT": {
        "account_name": ("Account Name", "👤"), "account_number": ("Account Number", "🔢"),
        "bank_name": ("Bank Name", "🏦"), "period": ("Period", "📅"),
        "closing_balance": ("Closing Balance", "💰"),
        "total_deposits": ("Total Deposits", "⬆"), "total_withdrawals": ("Total Withdrawals", "⬇"),
    },
    "RECEIPT": {
        "or_number": ("OR Number", "🔢"), "date": ("Date", "📅"),
        "received_from": ("Received From", "👤"), "amount": ("Amount", "💰"),
        "purpose": ("Purpose", "📝"), "issued_by": ("Issued By", "✍"),
    },
    "TODA_ORCR": {
        "plate_number": ("Plate Number", "🚗"), "mv_file_no": ("MV File No", "🔢"),
        "or_number": ("OR Number", "📋"), "cr_number": ("CR Number", "📋"),
        "owner": ("Owner", "👤"), "vehicle_type": ("Vehicle Type", "🚌"),
        "expiry_date": ("Expiry Date", "📅"),
    },
    "RENT": {
        "lessor": ("Lessor", "👤"), "lessee": ("Lessee", "👤"),
        "monthly_rent": ("Monthly Rent", "💰"), "property_address": ("Property Address", "🏠"),
        "lease_period": ("Lease Period", "📅"), "contract_date": ("Contract Date", "📅"),
    },
    "GCASH": {
        "reference_no": ("Reference No", "🔢"), "amount": ("Amount", "💰"),
        "sender": ("Sender", "👤"), "receiver": ("Receiver", "👤"),
        "date_time": ("Date & Time", "📅"), "mobile_number": ("Mobile Number", "📱"),
    },
    "MAYA": {
        "reference_no": ("Reference No", "🔢"), "amount": ("Amount", "💰"),
        "sender": ("Sender", "👤"), "receiver": ("Receiver", "👤"),
        "date_time": ("Date & Time", "📅"),
    },
    "PADALA": {
        "control_number": ("Control Number", "🔢"), "amount": ("Amount", "💰"),
        "sender": ("Sender", "👤"), "receiver": ("Receiver", "👤"),
        "center": ("Center", "🏪"), "date": ("Date", "📅"),
    },
    "ONLINE_TRANSFER": {
        "reference_no": ("Reference No", "🔢"), "amount": ("Amount", "💰"),
        "from_account": ("From Account", "🏦"), "to_account": ("To Account", "🏦"),
        "date_time": ("Date & Time", "📅"), "bank": ("Bank", "🏦"),
    },
    "SPOUSE_REMITTANCE": {
        "spouse_name": ("Spouse Name", "👤"), "amount": ("Amount", "💰"),
        "reference_no": ("Reference No", "🔢"), "date": ("Date", "📅"),
        "channel": ("Channel", "📡"), "frequency": ("Frequency", "🔄"),
    },
    "FARMING": {
        "crop_type": ("Crop Type", "🌾"), "quantity": ("Quantity", "📦"),
        "price_per_unit": ("Price per Unit", "💰"), "total_amount": ("Total Amount", "💵"),
        "buyer_trader": ("Buyer/Trader", "👤"), "date_of_sale": ("Date of Sale", "📅"),
        "seller": ("Seller", "👨‍🌾"),
    },
    "LIVESTOCK": {
        "animal_type": ("Animal Type", "🐄"), "head_count": ("Head Count", "🔢"),
        "live_weight": ("Live Weight", "⚖"), "price_per_kg": ("Price per kg", "💰"),
        "total_amount": ("Total Amount", "💵"), "buyer": ("Buyer", "👤"),
        "date": ("Date", "📅"),
    },
    "PAYSLIP": {
        "employee_name": ("Employee Name", "👤"), "employee_id": ("Employee ID", "🔢"),
        "period_covered": ("Period Covered", "📅"), "basic_pay": ("Basic Pay", "💰"),
        "gross_pay": ("Gross Pay", "💰"), "total_deductions": ("Total Deductions", "⬇"),
        "net_pay": ("Net Pay", "💵"), "employer": ("Employer", "🏢"),
    },
    "SPOUSE_PAYSLIP": {
        "employee_name": ("Employee Name", "👤"), "employee_id": ("Employee ID", "🔢"),
        "period_covered": ("Period Covered", "📅"), "basic_pay": ("Basic Pay", "💰"),
        "gross_pay": ("Gross Pay", "💰"), "total_deductions": ("Total Deductions", "⬇"),
        "net_pay": ("Net Pay", "💵"), "employer": ("Employer", "🏢"),
    },
    "ITR": {
        "taxpayer_name": ("Taxpayer Name", "👤"), "tin": ("TIN", "🔢"),
        "tax_year": ("Tax Year", "📅"), "gross_income": ("Gross Income", "💰"),
        "taxable_income": ("Taxable Income", "💰"), "tax_due": ("Tax Due", "📊"),
        "employer": ("Employer", "🏢"),
    },
    "SALN": {
        "declarant_name": ("Declarant Name", "👤"), "position": ("Position", "💼"),
        "total_assets": ("Total Assets", "💰"), "total_liabilities": ("Total Liabilities", "📉"),
        "net_worth": ("Net Worth", "💵"), "date_filed": ("Date Filed", "📅"),
    },
    "CIC": {
        "borrower_name": ("Borrower Name", "👤"), "date_of_birth": ("Date of Birth", "📅"),
        "credit_score": ("Credit Score", "⭐"), "active_loans": ("Active Loans", "📋"),
        "total_balance": ("Total Balance", "💰"), "overdue_amount": ("Overdue Amount", "⚠"),
        "report_date": ("Report Date", "📅"),
    },
    "BANK_CI": {
        "subject_name": ("Subject Name", "👤"), "bank_name": ("Bank Name", "🏦"),
        "verdict": ("Verdict", "⚖"), "remarks": ("Remarks", "📝"),
        "date": ("Date", "📅"), "certified_by": ("Certified By", "✍"),
    },
}

_VLM_MODELS = ["gemini-2.5-flash", "gemini-2.0-flash"]

# Maximum number of financial figures to display
_MAX_FIGURES = 12


def _is_image_file(file_path: str) -> bool:
    if not file_path:
        return False
    return Path(file_path).suffix.lower() in _IMAGE_EXTS


def vlm_extract_fields(
    doc_type: str,
    file_path: str,
    api_key: str | None = None,
) -> list[tuple[str, str, str]] | None:
    if not _is_image_file(file_path):
        return None

    schema = _VLM_SCHEMAS.get(doc_type)
    meta   = _VLM_FIELD_META.get(doc_type)
    if not schema or not meta:
        return None

    key = api_key or os.environ.get("GEMINI_API_KEY", "")
    if not key or key == "YOUR_GEMINI_API_KEY_HERE":
        return None

    try:
        from google import genai as _genai
        from google.genai import types as _gtypes
        import PIL.Image as _PILImage

        client  = _genai.Client(api_key=key)
        pil_img = _PILImage.open(file_path).convert("RGB")

        schema_str = _json.dumps(schema, indent=2)
        prompt = (
            f"You are a highly accurate document data extraction assistant for "
            f"Banco San Vicente (BSV), a rural bank in the Philippines.\n\n"
            f"Extract ALL visible fields from this {doc_type.replace('_', ' ').title()} "
            f"document image.\n\n"
            f"RULES:\n"
            f"1. Return ONLY valid JSON — no markdown, no explanation, no extra text.\n"
            f"2. Use null for any field not found or not visible.\n"
            f"3. For monetary amounts: include the number only (e.g. \"28,948.36\"), "
            f"   no currency symbols.\n"
            f"4. For dates: use the format as it appears in the document.\n"
            f"5. Do NOT guess or fabricate values.\n\n"
            f"Return this exact JSON structure:\n{schema_str}"
        )

        resp = None
        for model in _VLM_MODELS:
            try:
                resp = client.models.generate_content(
                    model    = model,
                    contents = [prompt, pil_img],
                    config   = _gtypes.GenerateContentConfig(
                        max_output_tokens=1024,
                        temperature=0.0,
                    ),
                )
                break
            except Exception as e:
                if any(kw in str(e).lower() for kw in
                       ("429", "quota", "resource_exhausted")):
                    continue
                raise

        if resp is None:
            return None

        raw = ""
        try:
            raw = resp.text or ""
        except Exception:
            try:
                raw = "".join(
                    p.text for p in resp.candidates[0].content.parts
                    if hasattr(p, "text") and p.text
                )
            except Exception:
                pass

        if not raw.strip():
            return None

        cleaned = re.sub(r"^```(?:json)?\s*", "", raw.strip(), flags=re.I)
        cleaned = re.sub(r"\s*```$", "", cleaned.strip())

        try:
            data = _json.loads(cleaned)
        except _json.JSONDecodeError:
            m = re.search(r"\{[\s\S]+\}", cleaned)
            if m:
                try:
                    data = _json.loads(m.group(0))
                except Exception:
                    return None
            else:
                return None

        results = []
        for key_name, (label, icon) in meta.items():
            val = data.get(key_name)
            if val is None or str(val).strip() in ("", "null", "None"):
                results.append((icon, label, "[not found]"))
            else:
                results.append((icon, label, str(val).strip()[:80]))

        return results

    except ImportError:
        return None
    except Exception as e:
        print(f"[vlm_extract_fields] {e}")
        return None


def classify_document(text: str, filename: str = "") -> str:
    if not text or not text.strip():
        return "GENERAL"

    t_lower    = text.lower()
    name_lower = filename.lower()

    filename_hints = {
        "cic":      "CIC",
        "bank_ci":  "BANK_CI",
        "bankci":   "BANK_CI",
        "payslip":  "PAYSLIP",
        "itr":      "ITR",
        "saln":     "SALN",
        "gcash":    "GCASH",
        "maya":     "MAYA",
        "paymaya":  "MAYA",
        "orcr":     "TODA_ORCR",
        "toda":     "TODA_ORCR",
        "padala":   "PADALA",
        "remit":    "PADALA",
        "rent":     "RENT",
        "lease":    "RENT",
        "copra":    "FARMING",
        "palay":    "FARMING",
        "pinya":    "FARMING",
        "harvest":  "FARMING",
        "farmgate": "FARMING",
        "livestock":"LIVESTOCK",
        "piggery":  "LIVESTOCK",
        "chicken":  "LIVESTOCK",
        "baboy":    "LIVESTOCK",
        "baka":     "LIVESTOCK",
        "statement":"BANK_STATEMENT",
        "receipt":  "RECEIPT",
    }
    for hint, doc_type in filename_hints.items():
        if hint in name_lower:
            return doc_type

    scores: dict[str, int] = {}
    for key, (label, icon, bg, fg, slot, keywords) in DOC_TYPES.items():
        score  = sum(2 if kw in t_lower else 0 for kw in keywords[:5])
        score += sum(1 if kw in t_lower else 0 for kw in keywords[5:])
        scores[key] = score

    best_key   = max(scores, key=lambda k: scores[k])
    best_score = scores[best_key]

    if best_key == "RECEIPT":
        if scores.get("FARMING", 0) >= 3:
            best_key = "FARMING"
        elif scores.get("LIVESTOCK", 0) >= 3:
            best_key = "LIVESTOCK"

    return best_key if best_score >= 3 else "GENERAL"


def extract_fields(doc_type: str, text: str) -> list[tuple[str, str, str]]:
    """
    Extract structured fields from text using regex.
    FIX: removed unused `text_inline` variable (dead code).
    """
    fields  = FIELD_DEFS.get(doc_type, GENERAL_FIELDS)
    results = []

    for (label, pattern, icon) in fields:
        if pattern is None:
            results.append((icon, label, "[see raw text]"))
            continue

        m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if m:
            val = next((g.strip() for g in m.groups() if g), m.group(0).strip())
            val = val.split("\n")[0].strip()[:80]
            results.append((icon, label, val))
            continue

        # Fallback: label on one line, value on the next
        label_pat   = re.sub(r"\\s\+", r"\\s+", re.escape(label))
        next_line_m = re.search(
            label_pat + r"[\t ]*\n\s*([^\n]{1,80})",
            text, re.IGNORECASE
        )
        if next_line_m:
            results.append((icon, label, next_line_m.group(1).strip()[:80]))
        else:
            results.append((icon, label, "[not found]"))

    return results


def extract_financial_figures(text: str) -> list[tuple[str, str]]:
    figures = []
    seen    = set()

    for m in re.finditer(
        r"[₱P]\s*([\d,]+(?:\.\d{2})?)"
        r"|PHP\s*([\d,]+(?:\.\d{2})?)"
        r"|(?:amount|total|balance|deposits?|withdrawals?|pay|salary|income|proceeds)"
        r"\s*[:\-]?\s*[P₱]?\s*([\d,]{4,}(?:\.\d{2})?)",
        text, re.IGNORECASE
    ):
        raw = (m.group(1) or m.group(2) or m.group(3) or "").replace(",", "")
        if not raw:
            continue
        try:
            val = float(raw)
        except ValueError:
            continue
        if val < 100 or val > 999_999_999:
            continue
        key = f"{val:.2f}"
        if key in seen:
            continue
        seen.add(key)
        start   = max(0, m.start() - 30)
        ctx_raw = text[start:m.start()].replace("\n", " ").strip()
        ctx     = (ctx_raw[-30:] if ctx_raw else "").strip()
        figures.append((f"₱{val:,.2f}", ctx))

    figures.sort(
        key=lambda x: float(x[0].replace("₱", "").replace(",", "")),
        reverse=True,
    )
    return figures[:_MAX_FIGURES]


# ── Tkinter mixin ─────────────────────────────────────────────────────────────

class DocClassifierTabMixin:
    """
    Mixin for DocExtractorApp.
    Call _build_classifier_panel(parent_card) in _build_right().
    After extraction, call show_classified_result(text, file_path).
    """

    def _F(self, size, weight="normal"):
        return ("Segoe UI", size, weight)

    def _FMONO(self, size, weight="normal"):
        return ("Consolas", size, weight)

    # ── Build ──────────────────────────────────────────────────────────────
    def _build_classifier_panel(self, parent):
        self._txt_frame = tk.Frame(parent, bg=CARD_WHITE)

        self._clf_tag_bar = tk.Frame(self._txt_frame, bg=CARD_WHITE)
        self._clf_tag_bar.pack(fill="x", padx=0, pady=0)

        clf_body_outer = tk.Frame(self._txt_frame, bg=CARD_WHITE)
        clf_body_outer.pack(fill="both", expand=True)

        clf_sb = tk.Scrollbar(clf_body_outer, relief="flat",
                              troughcolor=OFF_WHITE, bg=BORDER_LIGHT, width=8, bd=0)
        clf_sb.pack(side="right", fill="y")

        self._clf_canvas = tk.Canvas(
            clf_body_outer, bg=OFF_WHITE,
            highlightthickness=0,
            yscrollcommand=clf_sb.set,
        )
        self._clf_canvas.pack(side="left", fill="both", expand=True)
        clf_sb.config(command=self._clf_canvas.yview)

        self._clf_inner = tk.Frame(self._clf_canvas, bg=OFF_WHITE)
        self._clf_inner_win = self._clf_canvas.create_window(
            (0, 0), window=self._clf_inner, anchor="nw"
        )

        self._clf_inner.bind(
            "<Configure>",
            lambda e: self._clf_canvas.configure(
                scrollregion=self._clf_canvas.bbox("all")
            ),
        )
        self._clf_canvas.bind(
            "<Configure>",
            lambda e: self._clf_canvas.itemconfig(
                self._clf_inner_win, width=e.width
            ),
        )
        self._clf_canvas.bind(
            "<MouseWheel>",
            lambda e: self._clf_canvas.yview_scroll(
                int(-1 * (e.delta / 120)), "units"
            ),
        )
        self._clf_inner.bind(
            "<MouseWheel>",
            lambda e: self._clf_canvas.yview_scroll(
                int(-1 * (e.delta / 120)), "units"
            ),
        )

        # Raw text pane (hidden by default)
        self._clf_raw_frame = tk.Frame(self._txt_frame, bg=CARD_WHITE)
        raw_sb = tk.Scrollbar(self._clf_raw_frame, relief="flat",
                              troughcolor=OFF_WHITE, bg=BORDER_LIGHT, width=8, bd=0)
        raw_sb.pack(side="right", fill="y")
        self._textbox = tk.Text(
            self._clf_raw_frame,
            wrap="word",
            font=self._FMONO(10),
            fg=TXT_NAVY, bg=CARD_WHITE,
            relief="flat", bd=0,
            padx=24, pady=16,
            spacing1=3, spacing2=2, spacing3=3,
            insertbackground=LIME_DARK,
            yscrollcommand=raw_sb.set,
            state="disabled",
            cursor="arrow",
            selectbackground=NAVY_GHOST,
            selectforeground=TXT_NAVY,
        )
        self._textbox.pack(side="left", fill="both", expand=True)
        raw_sb.config(command=self._textbox.yview)
        self._textbox.tag_configure(
            "search_match",   background=LIME_PALE,   foreground=TXT_NAVY)
        self._textbox.tag_configure(
            "search_current", background=LIME_BRIGHT, foreground=NAVY_DEEP)

        self._clf_show_raw = False

        self._clf_show_placeholder()

    # ── Placeholder (single definition — pill grid version) ───────────────
    def _clf_show_placeholder(self):
        """
        Shows the blank 'waiting for extraction' state with a pill grid of
        supported document types.
        FIX: was defined twice; the simple version is removed and this one
             (the full pill-grid version) is the single authoritative definition.
        """
        self._clf_clear_tag_bar()
        for w in self._clf_inner.winfo_children():
            w.destroy()

        outer = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        outer.pack(expand=True, fill="both", pady=60)

        tk.Label(outer, text="📄",
                 font=("Segoe UI Emoji", 44),
                 fg=BORDER_MID, bg=OFF_WHITE).pack()
        tk.Label(outer, text="Smart Document Extractor",
                 font=self._F(14, "bold"),
                 fg=TXT_MUTED, bg=OFF_WHITE).pack(pady=(10, 4))
        tk.Label(
            outer,
            text="Browse a file and click  ⚡ Extract Text\n"
                 "The system will auto-detect the document type\n"
                 "and extract the relevant financial fields.",
            font=self._F(9),
            fg=TXT_MUTED, bg=OFF_WHITE,
            justify="center",
        ).pack()

        tk.Frame(outer, bg=BORDER_LIGHT, height=1).pack(
            fill="x", padx=40, pady=(20, 12))

        tk.Label(
            outer, text="SUPPORTED DOCUMENT TYPES",
            font=self._F(7, "bold"),
            fg=TXT_MUTED, bg=OFF_WHITE,
        ).pack()

        pills = [
            ("🏦 Bank Statement", "#EFF6FF", "#1D4ED8"),
            ("🧾 Receipt",        "#FFF7ED", "#C2410C"),
            ("🚗 TODA/ORCR",      "#F0FDF4", "#15803D"),
            ("🏠 Rent",           "#FDF4FF", "#7E22CE"),
            ("📱 GCash/Maya",     "#F0FDF4", "#166534"),
            ("📦 Padala",         "#FFFBEB", "#B45309"),
            ("💻 Online Transfer","#F0F9FF", "#0369A1"),
            ("👫 Spouse Income",  "#FFF1F2", "#BE123C"),
            ("🌾 Farming",        "#F7FEE7", "#3F6212"),
            ("🐄 Livestock",      "#FFF9F0", "#92400E"),
            ("💵 Payslip",        "#F0FDF4", "#166534"),
            ("👫 Spouse Payslip", "#FFF1F2", "#BE123C"),
            ("📊 ITR",            "#EFF6FF", "#1E40AF"),
            ("📄 SALN",           "#F8FAFC", "#334155"),
            ("📋 CIC",            "#EFF6FF", "#1E40AF"),
            ("🏦 Bank CI",        "#F0FDF4", "#166534"),
        ]

        row_frame = None
        for i, (label, bg, fg) in enumerate(pills):
            if i % 5 == 0:
                row_frame = tk.Frame(outer, bg=OFF_WHITE)
                row_frame.pack(pady=2)
            tk.Label(
                row_frame, text=label,
                font=self._F(7, "bold"),
                fg=fg, bg=bg,
                padx=8, pady=3,
            ).pack(side="left", padx=3)

    # ── Main display ───────────────────────────────────────────────────────
    def show_classified_result(
        self,
        text:      str,
        file_path: str = "",
        file_list: list | None = None,
    ):
        self._textbox.config(state="normal")
        self._textbox.delete("1.0", "end")
        self._textbox.insert("end", text)
        self._textbox.config(state="disabled")

        fname    = Path(file_path).name if file_path else ""
        doc_type = classify_document(text, fname)
        doc_info = DOC_TYPES.get(doc_type)

        if doc_info:
            label, icon, bg_color, fg_color, cibi_slot, _ = doc_info
        else:
            label, icon, bg_color, fg_color, cibi_slot = (
                "General Document", "📄", NAVY_MIST, NAVY_MID, None)

        self._clf_build_tag_bar(
            doc_type, label, icon, bg_color, fg_color,
            cibi_slot, text, file_path,
        )

        for w in self._clf_inner.winfo_children():
            w.destroy()
        self._clf_show_raw = False

        if cibi_slot and file_path:
            self._clf_auto_route(cibi_slot, file_path, text)

        figures = extract_financial_figures(text)
        if figures:
            self._clf_section_figures(figures)

        if _is_image_file(file_path):
            self._clf_section_fields_loading(label, icon, fg_color, bg_color)
            self._clf_canvas.yview_moveto(0)

            def _vlm_worker():
                api_key    = os.environ.get("GEMINI_API_KEY", "")
                vlm_result = vlm_extract_fields(doc_type, file_path, api_key)
                if vlm_result:
                    self.after(
                        0,
                        lambda: self._clf_replace_fields_section(
                            label, icon, fg_color, bg_color, vlm_result,
                            text, "🤖 VLM", doc_type, file_path,
                        ),
                    )
                else:
                    regex_result = extract_fields(doc_type, text)
                    self.after(
                        0,
                        lambda: self._clf_replace_fields_section(
                            label, icon, fg_color, bg_color, regex_result,
                            text, "🔍 Regex", doc_type, file_path,
                        ),
                    )

            threading.Thread(target=_vlm_worker, daemon=True).start()
        else:
            fields = extract_fields(doc_type, text)
            self._clf_section_fields(label, icon, fg_color, bg_color, fields)
            self._clf_section_raw_toggle(text)
            self._clf_canvas.yview_moveto(0)

    # ── Tag bar ────────────────────────────────────────────────────────────
    def _clf_clear_tag_bar(self):
        for w in self._clf_tag_bar.winfo_children():
            w.destroy()

    def _clf_build_tag_bar(
        self, doc_type, label, icon,
        bg_color, fg_color, cibi_slot,
        text, file_path,
    ):
        self._clf_clear_tag_bar()

        bar = tk.Frame(self._clf_tag_bar, bg=bg_color)
        bar.pack(fill="x")

        tk.Canvas(bar, height=3, bg=fg_color, highlightthickness=0).pack(fill="x")

        inner = tk.Frame(bar, bg=bg_color)
        inner.pack(fill="x", padx=16, pady=8)

        left = tk.Frame(inner, bg=bg_color)
        left.pack(side="left", fill="y")

        tk.Label(left, text=icon,
                 font=("Segoe UI Emoji", 22),
                 fg=fg_color, bg=bg_color).pack(side="left", padx=(0, 10))

        label_col = tk.Frame(left, bg=bg_color)
        label_col.pack(side="left")
        tk.Label(label_col, text=label,
                 font=self._F(11, "bold"),
                 fg=fg_color, bg=bg_color, anchor="w").pack(anchor="w")
        tk.Label(label_col, text="Document type auto-detected",
                 font=self._F(7),
                 fg=fg_color, bg=bg_color, anchor="w").pack(anchor="w")

        right = tk.Frame(inner, bg=bg_color)
        right.pack(side="right", fill="y")

        if cibi_slot:
            tk.Label(
                right,
                text=f"✅  Routed → {cibi_slot}",
                font=self._F(8, "bold"),
                fg=WHITE, bg=LIME_DARK,
                padx=10, pady=4,
            ).pack(side="right", padx=(8, 0))

        if _HAS_CTK:
            clear_btn = ctk.CTkButton(
                right,
                text="✕  Clear",
                command=self._clf_clear_panel,
                width=80, height=26, corner_radius=6,
                fg_color=NAVY_MIST,
                hover_color="#E8D5D5",
                text_color=NAVY_MID,
                font=ctk.CTkFont("Segoe UI", 8, weight="bold"),
                border_width=1,
                border_color=BORDER_MID,
            )
        else:
            clear_btn = tk.Button(
                right, text="✕  Clear",
                command=self._clf_clear_panel,
                font=self._F(8), fg=NAVY_MID, bg=NAVY_MIST,
                relief="flat", bd=1,
            )
        clear_btn.pack(side="right")

    # ── Section: Financial figures ─────────────────────────────────────────
    def _clf_section_figures(self, figures: list):
        sec = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        sec.pack(fill="x", padx=16, pady=(12, 0))

        self._clf_section_header(sec, "💰  KEY FINANCIAL FIGURES", LIME_DARK)

        grid = tk.Frame(sec, bg=OFF_WHITE)
        grid.pack(fill="x", pady=(6, 0))

        cols = 4
        for i, (amount, ctx) in enumerate(figures):
            col = i % cols
            row = i // cols

            card_outer = tk.Frame(grid, bg=BORDER_LIGHT, padx=1, pady=1)
            card_outer.grid(row=row, column=col, padx=4, pady=4, sticky="ew")
            grid.columnconfigure(col, weight=1)

            card = tk.Frame(card_outer, bg=LIME_MIST)
            card.pack(fill="both", expand=True)

            tk.Label(card, text=amount,
                     font=self._F(11, "bold"),
                     fg=LIME_DARK, bg=LIME_MIST,
                     anchor="w", padx=10).pack(anchor="w", pady=(8, 2))
            if ctx:
                tk.Label(card, text=ctx,
                         font=self._F(7),
                         fg=TXT_SOFT, bg=LIME_MIST,
                         anchor="w", padx=10,
                         wraplength=160).pack(anchor="w", pady=(0, 8))
            else:
                tk.Frame(card, bg=LIME_MIST, height=4).pack()

    # ── Section: Structured fields ─────────────────────────────────────────
    def _clf_section_fields(self, label, icon, fg_color, bg_color, fields: list):
        sec = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        sec.pack(fill="x", padx=16, pady=(14, 0))

        self._clf_section_header(
            sec, f"📋  EXTRACTED FIELDS — {label.upper()}", fg_color)

        card_outer = tk.Frame(sec, bg=BORDER_LIGHT, padx=1, pady=1)
        card_outer.pack(fill="x", pady=(6, 0))
        card = tk.Frame(card_outer, bg=CARD_WHITE)
        card.pack(fill="both", expand=True)

        for i, (ico, field_label, value) in enumerate(fields):
            row_bg = CARD_WHITE if i % 2 == 0 else OFF_WHITE
            row    = tk.Frame(card, bg=row_bg)
            row.pack(fill="x")

            found = value not in ("[not found]", "[see raw text]")
            tk.Frame(row, bg=LIME_DARK if found else BORDER_LIGHT,
                     width=3).pack(side="left", fill="y")

            tk.Label(row, text=ico,
                     font=("Segoe UI Emoji", 11),
                     fg=fg_color if found else TXT_MUTED,
                     bg=row_bg, width=3).pack(side="left", padx=(8, 0), pady=6)

            tk.Label(row, text=field_label,
                     font=self._F(8, "bold"),
                     fg=TXT_SOFT, bg=row_bg,
                     width=18, anchor="w").pack(side="left", padx=(6, 0))

            val_color = (
                LIME_DARK if found and any(c in value for c in "₱0123456789")
                else TXT_NAVY if found
                else TXT_MUTED
            )
            tk.Label(row, text=value,
                     font=self._F(9, "bold") if found else self._F(8),
                     fg=val_color, bg=row_bg,
                     anchor="w", wraplength=500).pack(
                side="left", padx=(8, 16), pady=6, fill="x", expand=True)

    # ── Section: Fields loading placeholder ───────────────────────────────
    def _clf_section_fields_loading(self, label, icon, fg_color, bg_color):
        sec = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        sec.pack(fill="x", padx=16, pady=(14, 0))
        self._clf_section_header(
            sec, f"📋  EXTRACTED FIELDS — {label.upper()} (AI extracting…)", fg_color)

        card_outer = tk.Frame(sec, bg=BORDER_LIGHT, padx=1, pady=1)
        card_outer.pack(fill="x", pady=(6, 0))
        card = tk.Frame(card_outer, bg=CARD_WHITE)
        card.pack(fill="both", expand=True)
        card_outer._is_fields_placeholder = True
        self._clf_fields_placeholder = card_outer

        loading_row = tk.Frame(card, bg=CARD_WHITE)
        loading_row.pack(fill="x", padx=16, pady=16)
        tk.Label(
            loading_row,
            text="🤖  Sending to Gemini VLM for accurate field extraction…",
            font=self._F(9), fg=NAVY_PALE, bg=CARD_WHITE,
        ).pack(anchor="w")
        tk.Label(
            loading_row,
            text="This uses Gemini Vision to read the image directly — much more accurate than regex.",
            font=self._F(8), fg=TXT_MUTED, bg=CARD_WHITE,
        ).pack(anchor="w", pady=(4, 0))

    def _clf_replace_fields_section(
        self, label, icon, fg_color, bg_color,
        fields, text, method_label="🤖 VLM",
        doc_type="GENERAL", file_path="",
    ):
        if hasattr(self, "_clf_fields_placeholder"):
            try:
                self._clf_fields_placeholder.destroy()
            except Exception:
                pass

        sec = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        sec.pack(fill="x", padx=16, pady=(14, 0))

        header_row = tk.Frame(sec, bg=OFF_WHITE)
        header_row.pack(fill="x", pady=(0, 2))
        tk.Frame(header_row, bg=fg_color, width=4, height=16).pack(
            side="left", padx=(0, 8))
        tk.Label(header_row,
                 text=f"📋  EXTRACTED FIELDS — {label.upper()}",
                 font=self._F(8, "bold"),
                 fg=fg_color, bg=OFF_WHITE).pack(side="left", anchor="w")

        badge_bg = LIME_MIST if "VLM" in method_label else NAVY_MIST
        badge_fg = LIME_DARK if "VLM" in method_label else NAVY_MID
        tk.Label(header_row, text=f"  {method_label}  ",
                 font=self._F(7, "bold"),
                 fg=badge_fg, bg=badge_bg,
                 padx=6, pady=2).pack(side="left", padx=(10, 0))

        card_outer = tk.Frame(sec, bg=BORDER_LIGHT, padx=1, pady=1)
        card_outer.pack(fill="x", pady=(6, 0))
        card = tk.Frame(card_outer, bg=CARD_WHITE)
        card.pack(fill="both", expand=True)

        for i, (ico, field_label, value) in enumerate(fields):
            row_bg = CARD_WHITE if i % 2 == 0 else OFF_WHITE
            row    = tk.Frame(card, bg=row_bg)
            row.pack(fill="x")
            found = value not in ("[not found]", "[see raw text]")
            tk.Frame(row, bg=LIME_DARK if found else BORDER_LIGHT,
                     width=3).pack(side="left", fill="y")
            tk.Label(row, text=ico,
                     font=("Segoe UI Emoji", 11),
                     fg=fg_color if found else TXT_MUTED,
                     bg=row_bg, width=3).pack(side="left", padx=(8, 0), pady=6)
            tk.Label(row, text=field_label,
                     font=self._F(8, "bold"),
                     fg=TXT_SOFT, bg=row_bg,
                     width=18, anchor="w").pack(side="left", padx=(6, 0))
            val_color = (
                LIME_DARK if found and any(c in value for c in "₱0123456789")
                else TXT_NAVY if found
                else TXT_MUTED
            )
            tk.Label(row, text=value,
                     font=self._F(9, "bold") if found else self._F(8),
                     fg=val_color, bg=row_bg,
                     anchor="w", wraplength=500).pack(
                side="left", padx=(8, 16), pady=6, fill="x", expand=True)

        self._clf_section_raw_toggle(text)

        _CIBI_OWNED = {"PAYSLIP", "ITR", "SALN", "CIC", "BANK_CI"}
        if "VLM" in method_label:
            if doc_type in _CIBI_OWNED:
                self._clf_section_cibi_redirect(doc_type, fields, file_path)
            elif doc_type == "BANK_STATEMENT":
                # Bank Statement → risk assessment instead of Excel populate
                self._clf_section_bank_risk_loading()
                def _risk_worker(
                    _text=text, _file=file_path, _fields=fields
                ):
                    api_key = os.environ.get("GEMINI_API_KEY", "")
                    if _HAS_BANK_RISK:
                        result = assess_bank_statement(
                            text=_text,
                            file_path=_file,
                            fields=_fields,
                            api_key=api_key,
                        )
                    else:
                        result = None
                    self.after(0, lambda r=result: self._clf_section_bank_risk_result(r))
                threading.Thread(target=_risk_worker, daemon=True).start()
            else:
                self._clf_section_populate_btn(doc_type, fields, file_path)

        self._clf_canvas.configure(scrollregion=self._clf_canvas.bbox("all"))
        self._clf_canvas.yview_moveto(0)

    # ── Section: Raw text toggle ───────────────────────────────────────────
    def _clf_section_raw_toggle(self, text: str):
        sec = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        sec.pack(fill="x", padx=16, pady=(14, 16))

        tk.Frame(sec, bg=BORDER_LIGHT, height=1).pack(fill="x", pady=(0, 8))

        info_row = tk.Frame(sec, bg=OFF_WHITE)
        info_row.pack(fill="x")

        tk.Label(info_row,
                 text=f"📄  {len(text):,} characters  ·  {text.count(chr(10)):,} lines",
                 font=self._F(8),
                 fg=TXT_MUTED, bg=OFF_WHITE).pack(side="left")
        tk.Label(info_row,
                 text="Use the  ✕ Clear  button above to reset the extracted panel",
                 font=self._F(7),
                 fg=TXT_MUTED, bg=OFF_WHITE).pack(side="right")

    # ── Section header helper ──────────────────────────────────────────────
    def _clf_section_header(self, parent, text: str, color=None):
        color = color or NAVY_PALE
        row   = tk.Frame(parent, bg=OFF_WHITE)
        row.pack(fill="x", pady=(0, 2))
        tk.Frame(row, bg=color, width=4, height=16).pack(side="left", padx=(0, 8))
        tk.Label(row, text=text,
                 font=self._F(8, "bold"),
                 fg=color, bg=OFF_WHITE).pack(side="left", anchor="w")

    # ── Clear panel ────────────────────────────────────────────────────────
    def _clf_clear_panel(self):
        """
        Clears the extracted panel back to the blank placeholder state.
        FIX: children are destroyed before _clf_show_placeholder() is called,
             and the call order is correct so the placeholder renders cleanly.
        """
        for w in self._clf_inner.winfo_children():
            w.destroy()

        try:
            self._textbox.config(state="normal")
            self._textbox.delete("1.0", "end")
            self._textbox.config(state="disabled")
        except Exception:
            pass

        try:
            self._clf_raw_frame.pack_forget()
        except Exception:
            pass
        try:
            self._clf_canvas.pack(side="left", fill="both", expand=True)
        except Exception:
            pass
        self._clf_show_raw = False

        self._clf_clear_tag_bar()

        # Show the full pill-grid placeholder
        self._clf_show_placeholder()

    # ── Auto-route to CIBI slot ────────────────────────────────────────────
    def _clf_auto_route(self, cibi_slot: str, file_path: str, text: str):
        try:
            if not hasattr(self, "_cibi_slots"):
                return
            if cibi_slot not in self._cibi_slots:
                return

            slot = self._cibi_slots[cibi_slot]
            if slot.get("path"):
                self._clf_route_banner(
                    f"ℹ  {cibi_slot} slot already filled — not auto-routed.",
                    bg=NAVY_MIST, fg=NAVY_MID,
                )
                return

            slot["path"] = file_path
            slot["text"] = text

            if hasattr(self, "_cibi_status_labels") and cibi_slot in self._cibi_status_labels:
                self._cibi_status_labels[cibi_slot].config(
                    text="✅", fg=ACCENT_SUCCESS)
            if hasattr(self, "_cibi_name_labels") and cibi_slot in self._cibi_name_labels:
                name  = Path(file_path).name
                short = name if len(name) <= 32 else name[:29] + "…"
                self._cibi_name_labels[cibi_slot].config(
                    text=short, fg=TXT_NAVY_MID)
            if hasattr(self, "_cibi_refresh_stage_buttons"):
                self._cibi_refresh_stage_buttons()

            self._clf_route_banner(
                f"✅  Auto-routed to CIBI slot: {cibi_slot}",
                bg=LIME_MIST, fg=LIME_DARK,
            )

        except Exception as e:
            self._clf_route_banner(
                f"⚠  Auto-route to {cibi_slot} failed: {e}",
                bg="#FFF8EC", fg=ACCENT_GOLD,
            )

    def _clf_route_banner(self, msg: str, bg=LIME_MIST, fg=LIME_DARK):
        tk.Label(
            self._clf_inner,
            text=msg,
            font=self._F(8, "bold"),
            fg=fg, bg=bg,
            padx=16, pady=8,
            anchor="w",
        ).pack(fill="x", padx=0, pady=(0, 4))

    # ── Section: Bank Statement Risk Assessment ──────────────────────────────

    def _clf_section_bank_risk_loading(self):
        """Show a loading card while the risk assessment runs in background."""
        sec = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        sec.pack(fill="x", padx=16, pady=(14, 4))
        self._clf_section_header(sec, "🏦  BANK STATEMENT RISK ASSESSMENT", "#1D4ED8")

        card_outer = tk.Frame(sec, bg=BORDER_LIGHT, padx=1, pady=1)
        card_outer.pack(fill="x", pady=(6, 0))
        card = tk.Frame(card_outer, bg=CARD_WHITE)
        card.pack(fill="both", expand=True)
        # Tag so _clf_section_bank_risk_result can find and destroy it
        self._clf_bank_risk_placeholder = card_outer

        row = tk.Frame(card, bg=CARD_WHITE)
        row.pack(fill="x", padx=16, pady=16)
        tk.Label(
            row,
            text="⏳  Analysing cash flow, balance, and transaction patterns…",
            font=self._F(9), fg=NAVY_PALE, bg=CARD_WHITE,
        ).pack(anchor="w")
        tk.Label(
            row,
            text="Structured analysis running — Gemini VLM will be consulted if result is uncertain.",
            font=self._F(8), fg=TXT_MUTED, bg=CARD_WHITE,
        ).pack(anchor="w", pady=(4, 0))

    def _clf_section_bank_risk_result(self, result):
        """
        Replace the loading placeholder with the full risk assessment panel.
        Mirrors the Bank CI verdict card layout from ui_cibi Stage 2.
        """
        # Destroy loading placeholder
        if hasattr(self, "_clf_bank_risk_placeholder"):
            try:
                self._clf_bank_risk_placeholder.destroy()
            except Exception:
                pass

        if not _HAS_BANK_RISK or result is None:
            # bank_statement_risk.py not available — show plain notice
            sec = tk.Frame(self._clf_inner, bg=OFF_WHITE)
            sec.pack(fill="x", padx=16, pady=(14, 4))
            self._clf_section_header(sec, "🏦  BANK STATEMENT RISK ASSESSMENT", "#1D4ED8")
            card_outer = tk.Frame(sec, bg=BORDER_LIGHT, padx=1, pady=1)
            card_outer.pack(fill="x", pady=(6, 0))
            tk.Label(
                tk.Frame(card_outer, bg=CARD_WHITE),
                text="⚠  bank_statement_risk.py not found — install it to enable risk assessment.",
                font=self._F(9), fg=ACCENT_RED, bg=CARD_WHITE, padx=16, pady=12,
            ).pack(anchor="w")
            self._clf_canvas.configure(scrollregion=self._clf_canvas.bbox("all"))
            return

        # ── Colour scheme mirrors Bank CI verdict cards ───────────────────
        verdict_cfg = {
            VERDICT_GOOD: {
                "bg":      "#F0FDF4",
                "fg":      "#166534",
                "acc":     LIME_DARK,
                "icon":    "✅",
                "badge":   ("LOW RISK",  LIME_DARK,  LIME_MIST),
            },
            VERDICT_BAD: {
                "bg":      "#FEF2F2",
                "fg":      "#991B1B",
                "acc":     ACCENT_RED,
                "icon":    "❌",
                "badge":   ("HIGH RISK", ACCENT_RED, "#FEE2E2"),
            },
            VERDICT_UNCERTAIN: {
                "bg":      "#FFFBEB",
                "fg":      "#92400E",
                "acc":     ACCENT_GOLD,
                "icon":    "⚠",
                "badge":   ("REVIEW",    ACCENT_GOLD, "#FEF3C7"),
            },
        }
        cfg = verdict_cfg.get(result.verdict, verdict_cfg[VERDICT_UNCERTAIN])

        sec = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        sec.pack(fill="x", padx=16, pady=(14, 4))

        # Section header
        header_row = tk.Frame(sec, bg=OFF_WHITE)
        header_row.pack(fill="x", pady=(0, 2))
        tk.Frame(header_row, bg=cfg["acc"], width=4, height=16).pack(
            side="left", padx=(0, 8))
        tk.Label(header_row,
                 text="🏦  BANK STATEMENT RISK ASSESSMENT",
                 font=self._F(8, "bold"),
                 fg=cfg["acc"], bg=OFF_WHITE).pack(side="left", anchor="w")
        badge_label, badge_fg, badge_bg = cfg["badge"]
        tk.Label(header_row, text=f"  {badge_label}  ",
                 font=self._F(7, "bold"),
                 fg=badge_fg, bg=badge_bg,
                 padx=6, pady=2).pack(side="left", padx=(10, 0))
        if result.vlm_used:
            tk.Label(header_row, text="  🤖 VLM  ",
                     font=self._F(7, "bold"),
                     fg=NAVY_MID, bg=NAVY_MIST,
                     padx=6, pady=2).pack(side="left", padx=(4, 0))

        # Verdict card — accent bar top + coloured background
        card_outer = tk.Frame(sec, bg=BORDER_LIGHT, padx=1, pady=1)
        card_outer.pack(fill="x", pady=(6, 0))
        card = tk.Frame(card_outer, bg=cfg["bg"])
        card.pack(fill="both", expand=True)

        tk.Canvas(card, height=3, bg=cfg["acc"],
                  highlightthickness=0).pack(fill="x")

        # Verdict + summary row
        verdict_row = tk.Frame(card, bg=cfg["bg"])
        verdict_row.pack(fill="x", padx=16, pady=(10, 4))
        tk.Label(verdict_row,
                 text=f"{cfg['icon']}  {result.verdict}",
                 font=self._F(13, "bold"),
                 fg=cfg["fg"], bg=cfg["bg"]).pack(side="left")
        tk.Label(verdict_row,
                 text="  Proceed" if result.proceed else "  Manual review required",
                 font=self._F(8),
                 fg=cfg["fg"], bg=cfg["bg"]).pack(side="left", padx=(8, 0))

        tk.Label(card,
                 text=result.summary,
                 font=self._F(9),
                 fg=cfg["fg"], bg=cfg["bg"],
                 padx=16, wraplength=700, justify="left", anchor="w",
                 ).pack(fill="x", pady=(0, 6))

        # Signals table
        if result.signals:
            self._clf_section_bank_risk_signals(card, result, cfg)

        # Key figures summary row (deposits / withdrawals / balance)
        self._clf_section_bank_risk_figures(card, result, cfg)

        tk.Frame(card, bg=cfg["bg"], height=6).pack()

        self._clf_canvas.configure(scrollregion=self._clf_canvas.bbox("all"))
        self._clf_canvas.yview_moveto(0)

    def _clf_section_bank_risk_signals(self, parent, result, cfg):
        """Renders the signals table inside the verdict card."""
        tbl = tk.Frame(parent, bg=cfg["bg"])
        tbl.pack(fill="x", padx=16, pady=(4, 0))

        tk.Frame(tbl, bg=BORDER_LIGHT, height=1).pack(fill="x", pady=(0, 6))
        tk.Label(tbl, text="CASH FLOW SIGNALS",
                 font=self._F(7, "bold"),
                 fg=TXT_MUTED, bg=cfg["bg"]).pack(anchor="w", pady=(0, 4))

        signal_bg = {
            "POSITIVE": (LIME_MIST,   LIME_DARK,  "✅"),
            "NEGATIVE": ("#FEF2F2",   ACCENT_RED, "❌"),
            "NEUTRAL":  (NAVY_MIST,   NAVY_MID,   "ℹ"),
        }

        for s in result.signals:
            s_bg, s_fg, s_icon = signal_bg.get(
                s.verdict, (OFF_WHITE, TXT_SOFT, "•"))
            row = tk.Frame(tbl, bg=s_bg)
            row.pack(fill="x", pady=2)

            # Left accent bar
            tk.Frame(row, bg=s_fg, width=3).pack(side="left", fill="y")

            # Icon
            tk.Label(row, text=s_icon,
                     font=("Segoe UI Emoji", 10),
                     fg=s_fg, bg=s_bg,
                     width=3).pack(side="left", padx=(6, 0), pady=5)

            # Label
            tk.Label(row, text=s.label,
                     font=self._F(8, "bold"),
                     fg=TXT_SOFT, bg=s_bg,
                     width=22, anchor="w").pack(side="left", padx=(4, 0))

            # Value
            tk.Label(row, text=s.value,
                     font=self._F(8, "bold"),
                     fg=s_fg, bg=s_bg,
                     anchor="w", wraplength=320).pack(
                side="left", padx=(6, 8), pady=5, fill="x", expand=True)

    def _clf_section_bank_risk_figures(self, parent, result, cfg):
        """Renders the key figures summary row (deposits / withdrawals / balance)."""
        has_any = any([
            result.deposits is not None,
            result.withdrawals is not None,
            result.closing_balance is not None,
        ])
        if not has_any:
            return

        tk.Frame(parent, bg=BORDER_LIGHT, height=1).pack(
            fill="x", padx=16, pady=(8, 4))

        fig_row = tk.Frame(parent, bg=cfg["bg"])
        fig_row.pack(fill="x", padx=16, pady=(0, 4))

        def _fig_card(label, value, accent):
            card = tk.Frame(fig_row, bg=CARD_WHITE,
                            highlightbackground=BORDER_LIGHT,
                            highlightthickness=1)
            card.pack(side="left", padx=(0, 8), pady=4, ipadx=10, ipady=6)
            tk.Label(card, text=label,
                     font=self._F(7),
                     fg=TXT_MUTED, bg=CARD_WHITE).pack(anchor="w", padx=6)
            tk.Label(card, text=value,
                     font=self._F(10, "bold"),
                     fg=accent, bg=CARD_WHITE).pack(anchor="w", padx=6)

        if result.deposits is not None:
            _fig_card("Total Deposits",
                      f"₱{result.deposits:,.2f}", LIME_DARK)
        if result.withdrawals is not None:
            _fig_card("Total Withdrawals",
                      f"₱{result.withdrawals:,.2f}", ACCENT_RED)
        if result.closing_balance is not None:
            bal_color = (
                LIME_DARK if result.closing_balance >= 5_000
                else ACCENT_RED
            )
            _fig_card("Closing Balance",
                      f"₱{result.closing_balance:,.2f}", bal_color)

    # ── Section: CIBI Mode redirect ────────────────────────────────────────
    def _clf_section_cibi_redirect(self, doc_type: str, fields: list = None, file_path: str = ""):
        _LABELS = {
            "PAYSLIP": "Payslip / Payroll",
            "ITR":     "Income Tax Return",
            "SALN":    "SALN / Net Worth",
            "CIC":     "CIC Credit Report",
            "BANK_CI": "Bank CI Certification",
            "SPOUSE_PAYSLIP": "Spouse Payslip / Payroll",
        }
        label = _LABELS.get(doc_type, doc_type)

        outer = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        outer.pack(fill="x", padx=16, pady=(12, 4))

        card = tk.Frame(outer, bg=NAVY_MIST,
                        highlightbackground=NAVY_PALE, highlightthickness=1)
        card.pack(fill="x")

        inner = tk.Frame(card, bg=NAVY_MIST)
        inner.pack(fill="x", padx=16, pady=12)

        tk.Label(
            inner,
            text=f"ℹ  {label} — handled by CIBI Mode",
            font=self._F(9, "bold"), fg=NAVY_MID, bg=NAVY_MIST,
        ).pack(anchor="w")
        tk.Label(
            inner,
            text=(
                "This document type is part of the CIBI workflow.\n"
                f"To populate the Excel template with this {label}, use:\n"
                "  📋  CIBI Mode tab → Upload to the correct slot → Extract All → Populate"
            ),
            font=self._F(8), fg=TXT_SOFT, bg=NAVY_MIST,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

        # ── Manual override: only shown for PAYSLIP → re-classify as SPOUSE_PAYSLIP ──
        if doc_type == "PAYSLIP" and fields is not None and _HAS_CTK:
            tk.Frame(inner, bg=BORDER_MID, height=1).pack(fill="x", pady=(10, 6))
            tk.Label(
                inner,
                text="Is this the spouse's payslip?",
                font=self._F(8, "bold"), fg=NAVY_MID, bg=NAVY_MIST,
            ).pack(anchor="w")
            tk.Label(
                inner,
                text="Click below to re-classify and get the Excel populate button.",
                font=self._F(8), fg=TXT_SOFT, bg=NAVY_MIST,
            ).pack(anchor="w", pady=(2, 6))

            override_status = tk.Label(
                inner, text="", font=self._F(8),
                fg=LIME_DARK, bg=NAVY_MIST,
            )

            def _reclassify_as_spouse(_fields=fields, _file=file_path, _lbl=override_status):
                """Destroy this card and replace with the Spouse Payslip populate button."""
                try:
                    outer.destroy()
                except Exception:
                    pass
                self._clf_section_populate_btn("SPOUSE_PAYSLIP", _fields, _file)
                self._clf_canvas.configure(
                    scrollregion=self._clf_canvas.bbox("all"))

            ctk.CTkButton(
                inner,
                text="👫  Re-classify as Spouse Payslip → Populate Excel",
                command=_reclassify_as_spouse,
                height=32, corner_radius=6,
                fg_color="#BE123C", hover_color="#9F1239",
                text_color=WHITE,
                font=ctk.CTkFont("Segoe UI", 9, weight="bold"),
                border_width=0,
            ).pack(anchor="w")
            override_status.pack(anchor="w", pady=(4, 0))

    # ── Section: Populate Excel button ────────────────────────────────────
    def _clf_section_populate_btn(self, doc_type: str, fields: list, file_path: str):
        if not _HAS_CTK:
            return  # graceful skip if customtkinter unavailable

        outer = tk.Frame(self._clf_inner, bg=OFF_WHITE)
        outer.pack(fill="x", padx=16, pady=(12, 4))

        card = tk.Frame(outer, bg=LIME_MIST,
                        highlightbackground=LIME_MID, highlightthickness=1)
        card.pack(fill="x")

        inner = tk.Frame(card, bg=LIME_MIST)
        inner.pack(fill="x", padx=16, pady=12)

        tk.Label(
            inner,
            text="✅  VLM extraction complete — ready to populate Excel",
            font=self._F(9, "bold"), fg=LIME_DARK, bg=LIME_MIST,
        ).pack(anchor="w")
        tk.Label(
            inner,
            text="Select your CIBI Excel template to auto-fill with the extracted fields.",
            font=self._F(8), fg=TXT_SOFT, bg=LIME_MIST,
        ).pack(anchor="w", pady=(2, 8))

        btn_row = tk.Frame(inner, bg=LIME_MIST)
        btn_row.pack(anchor="w")

        ctk.CTkButton(
            btn_row,
            text="📊  Populate → CIBI Excel",
            command=lambda: self._clf_pick_and_populate(doc_type, fields, file_path),
            height=38, corner_radius=8,
            fg_color=LIME_DARK, hover_color=LIME_MID, text_color=WHITE,
            font=ctk.CTkFont("Segoe UI", 10, weight="bold"),
            border_width=0,
        ).pack(side="left")

        self._clf_populate_status = tk.Label(
            btn_row, text="",
            font=self._F(8), fg=TXT_SOFT, bg=LIME_MIST,
        )
        self._clf_populate_status.pack(side="left", padx=(12, 0))

    def _clf_pick_and_populate(self, doc_type: str, fields: list, file_path: str):
        if not _HAS_FILEDIALOG:
            return

        template_path = _filedialog.askopenfilename(
            title="Select CIBI Excel Template",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files",   "*.*"),
            ],
        )
        if not template_path:
            return

        stem = "classified_doc"
        if file_path:
            stem = Path(file_path).stem
        else:
            for icon, label, value in fields:
                if label in (
                    "Account Name", "Employee Name", "Taxpayer Name",
                    "Borrower Name", "Declarant Name",
                ) and value not in ("[not found]", ""):
                    stem = value.replace(" ", "_").replace(",", "")[:30]
                    break

        if hasattr(self, "_clf_populate_status"):
            self._clf_populate_status.config(text="⏳ Populating…", fg=ACCENT_GOLD)

        def _worker():
            try:
                from doc_classifier_populator import populate_from_classifier

                def _cb(pct: int, msg: str = ""):
                    # FIX: use lambda so config() is called correctly by after()
                    if hasattr(self, "_clf_populate_status"):
                        self.after(
                            0,
                            lambda m=msg: self._clf_populate_status.config(
                                text=f"⏳ {m}", fg=ACCENT_GOLD
                            ),
                        )

                out_path = populate_from_classifier(
                    template_path=template_path,
                    doc_type=doc_type,
                    fields=fields,
                    output_stem=stem,
                    progress_cb=_cb,
                )
                self.after(0, lambda: self._clf_populate_done(out_path))

            except Exception as e:
                self.after(0, lambda err=str(e): self._clf_populate_error(err))

        threading.Thread(target=_worker, daemon=True).start()

    def _clf_populate_done(self, out_path: Path):
        if hasattr(self, "_clf_populate_status"):
            self._clf_populate_status.config(
                text=f"✅  Saved: {out_path.name}", fg=LIME_DARK)

        try:
            import sys
            if sys.platform == "win32":
                os.startfile(str(out_path))
            elif sys.platform == "darwin":
                os.system(f'open "{out_path}"')
            else:
                os.system(f'xdg-open "{out_path}"')
        except Exception:
            pass

        if hasattr(self, "_append_chat_bubble"):
            self._append_chat_bubble(
                "📊  Excel populated successfully!\n\n"
                f"📁  {out_path}\n\n"
                "File saved to Desktop → DocExtract_Files folder.",
                role="system",
            )

    def _clf_populate_error(self, error_msg: str):
        if hasattr(self, "_clf_populate_status"):
            self._clf_populate_status.config(
                text=f"❌  Failed: {error_msg[:60]}", fg=ACCENT_RED)

    # ── Legacy shim ────────────────────────────────────────────────────────
    def _write(self, txt, color=TXT_NAVY):
        """
        Legacy shim — updates the raw text widget only.
        The structured panel is updated separately via show_classified_result().
        """
        box = self._textbox
        box.config(state="normal", fg=color)
        box.delete("1.0", "end")
        box.insert("end", txt)
        box.config(state="disabled")
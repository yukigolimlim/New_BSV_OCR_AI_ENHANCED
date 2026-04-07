# CIBI Populator — Embedded Gemini Prompts

All prompts found in `cibi_populator.py`. Each is sent to the Gemini API
(`gemini-2.5-flash`) to extract structured JSON from Philippine loan documents.

---

## 1. `_PAYSLIP_PROMPT`
**Variable:** `_PAYSLIP_PROMPT`
**Function:** `extract_payslip()`
**Thinking budget:** 512 | **Max output tokens:** 16,000
**Text slice cap:** `payslip_text[:20,000]`

### Purpose
Extracts compensation and deduction fields from a Philippine payslip or
payroll document. Handles multiple pay periods by averaging scalar fields.

### Full Prompt
```
You are extracting structured data from a Philippine payslip or payroll document.
This document may contain MULTIPLE PAY PERIODS. Read the ENTIRE document carefully.
Return ONLY a valid JSON object — no explanation, no markdown, no extra text.
Use null for any field not found. Numbers only (no ₱, no commas).

CRITICAL RULES:
  1. Extract EVERY field present. Never return null for a field that exists.
  2. If the document contains MORE THAN ONE pay period:
       a. Fill "pay_periods" with one entry per period.
       b. Set "period_count" to the total number found.
       c. For ALL income/deduction scalar fields, return the AVERAGE across all periods.
  3. If only ONE period, set "period_count" to 1.
  4. net_pay = gross_pay minus total_deductions. Extract the ACTUAL net pay figure shown.
  5. date_of_birth — extract if shown on the payslip (some DepEd payslips show it).

{
  "employee_name": null, "employer_name": null, "position": null, "department": null,
  "period_from": null, "period_to": null, "pay_date": null, "period_count": null,
  "date_of_birth": null,
  "pay_periods": [
    {
      "period_from": null, "period_to": null, "pay_date": null,
      "basic_pay": null, "allowances": null, "gross_pay": null,
      "gsis_deduction": null, "sss_deduction": null, "philhealth_deduction": null,
      "pagibig_deduction": null, "tax_deduction": null, "other_deductions": null,
      "total_deductions": null, "net_pay": null
    }
  ],
  "monthly_rate": null, "basic_pay": null, "allowances": null, "gross_pay": null,
  "gsis_deduction": null, "sss_deduction": null, "philhealth_deduction": null,
  "pagibig_deduction": null, "tax_deduction": null, "other_deductions": null,
  "total_deductions": null, "net_pay": null,
  "tin": null, "sss_number": null, "philhealth_number": null, "pagibig_number": null,
  "business_income": null, "rental_income": null, "remittance_income": null,
  "other_income_label": null, "other_income_amount": null
}

--- PAYSLIP DOCUMENT ---
{text}
--- END ---
```

### Output Fields
| Field | Description |
|---|---|
| `employee_name` | Full name of the employee |
| `employer_name` | Name of the employer/agency |
| `position` | Job title |
| `department` | Department or office |
| `period_from / period_to` | Pay period date range |
| `pay_date` | Date of payment |
| `period_count` | Number of pay periods in document |
| `date_of_birth` | DOB if shown (e.g. DepEd payslips) |
| `pay_periods[]` | Array of individual period breakdowns |
| `basic_pay` | Base salary (averaged if multiple periods) |
| `allowances` | Total allowances |
| `gross_pay` | Total gross before deductions |
| `gsis_deduction` | GSIS contribution |
| `sss_deduction` | SSS contribution |
| `philhealth_deduction` | PhilHealth contribution |
| `pagibig_deduction` | Pag-IBIG contribution |
| `tax_deduction` | Withholding tax |
| `other_deductions` | Any other deductions |
| `total_deductions` | Sum of all deductions |
| `net_pay` | Actual take-home pay |
| `tin` | Tax Identification Number |
| `sss_number` | SSS number |
| `philhealth_number` | PhilHealth number |
| `pagibig_number` | Pag-IBIG number |
| `business_income` | Business income if declared |
| `rental_income` | Rental income if declared |
| `remittance_income` | OFW/remittance income if declared |
| `other_income_label` | Label for other income source |
| `other_income_amount` | Amount for other income source |

---

## 2. `_SALN_PROMPT`
**Variable:** `_SALN_PROMPT`
**Function:** `extract_saln()`
**Thinking budget:** 512 | **Max output tokens:** 16,000
**Text slice cap:** `saln_text[:30,000]`

### Purpose
Extracts assets, liabilities, net worth, and family data from a Philippine
SALN (Statement of Assets, Liabilities and Net Worth).

### Full Prompt
```
You are extracting structured data from a Philippine SALN
(Statement of Assets, Liabilities and Net Worth).
Return ONLY a valid JSON object — no explanation, no markdown, no extra text.
Use null for any field not found. Return [] for empty lists.
Numbers only (no ₱, no commas).

CRITICAL EXTRACTION RULES:
  1. Extract EVERY non-N/A row. Never stop after the first item.
  2. personal_properties — read the ENTIRE table top to bottom.
  3. Set current_value = acquisition_cost for every item (no separate current value column in SALN).
  4. real_properties — if all rows say "N/A", return [].
  5. liabilities — extract EVERY row: nature of loan, name of creditor, outstanding balance.
  6. children — name and age for every child listed. Return null for date_of_birth.
  7. net_worth may be negative. Return as a number.
  8. position — copy EXACTLY as printed.
  9. CASH FIELDS — always extract as TWO SEPARATE fields, never combine them:
       cash_on_hand  = the "Cash on Hand" amount only
       cash_in_bank  = the "Cash in Bank" amount only
       cash_on_hand_and_in_bank = null (leave null if separate fields exist)
  10. real_properties — use these EXACT key names:
       description, kind, location, area, assessed_value,
       current_fair_market_value (the "CURRENT FAIR MARKET VALUE" column),
       year_acquired, mode_of_acquisition, acquisition_cost
  11. spouse_position — extract the spouse's position/occupation from the SPOUSE section.
  12. real_properties column → JSON key mapping (use EXACTLY these key names):
        "DESCRIPTION" column            → "description"  (e.g. "Residential House and Lot")
        "KIND" column                   → "kind"          (e.g. "Residential")
        "EXACT LOCATION" column         → "location"
        "ASSESSED VALUE" column         → "assessed_value"
        "CURRENT FAIR MARKET VALUE" col → "current_fair_market_value"
        "ACQUISITION YEAR" column       → "year_acquired"
        "MODE OF ACQUISITION" column    → "mode_of_acquisition"
        "ACQUISITION COST" column       → "acquisition_cost"

{
  "declarant_name": null,
  "position": null,
  "agency": null,
  "office_address": null,
  "saln_year": null,
  "spouse_name": null,
  "spouse_position": null,
  "real_properties": [],
  "personal_properties": [
    {"description":null,"year_acquired":null,"acquisition_cost":null,"current_value":null}
  ],
  "cash_on_hand": null,
  "cash_in_bank": null,
  "cash_on_hand_and_in_bank": null,
  "receivables": null,
  "business_interests": null,
  "total_assets": null,
  "liabilities": [
    {"nature":null,"creditor":null,"outstanding_balance":null}
  ],
  "financial_liabilities": null,
  "personal_liabilities": null,
  "total_liabilities": null,
  "net_worth": null,
  "children": [
    {"name":null,"date_of_birth":null,"age":null}
  ]
}

--- SALN DOCUMENT ---
{text}
--- END ---
```

### Output Fields
| Field | Description |
|---|---|
| `declarant_name` | Full name of the SALN filer |
| `position` | Exact job position as printed |
| `agency` | Government agency or office |
| `office_address` | Office address |
| `saln_year` | Year covered by the SALN |
| `spouse_name` | Spouse's full name |
| `spouse_position` | Spouse's occupation/position |
| `real_properties[]` | List of real property assets |
| `personal_properties[]` | List of personal property assets |
| `cash_on_hand` | Cash on hand (separate field) |
| `cash_in_bank` | Cash in bank (separate field) |
| `cash_on_hand_and_in_bank` | Combined cash (only if not separated) |
| `receivables` | Accounts receivable |
| `business_interests` | Business interest value |
| `total_assets` | Total declared assets |
| `liabilities[]` | List of loans/liabilities |
| `financial_liabilities` | Total financial liabilities |
| `personal_liabilities` | Total personal liabilities |
| `total_liabilities` | Grand total liabilities |
| `net_worth` | Net worth (can be negative) |
| `children[]` | List of children with name and age |

---

## 3. `_ITR_PROMPT`
**Variable:** `_ITR_PROMPT`
**Function:** `extract_itr()`
**Thinking budget:** 512 | **Max output tokens:** 16,000
**Text slice cap:** `itr_text[:30,000]`

### Purpose
Extracts income, tax, and personal data from a Philippine ITR (Income Tax
Return) or BIR Form 2316 (Certificate of Compensation Payment / Tax Withheld).

### Full Prompt
```
You are extracting structured data from a Philippine ITR (Income Tax Return) or
BIR Form 2316 (Certificate of Compensation Payment / Tax Withheld).
Return ONLY a valid JSON object — no explanation, no markdown, no extra text.
Use null for any field not found. Numbers only (no ₱, no commas).

CRITICAL:
  1. Extract EVERY field present.
  2. date_of_birth — extract from Part I Employee Information if present (MM/DD/YYYY).
  3. registered_address — extract the employee's home address from Part I.
  4. gross_compensation_income — the total gross compensation before deductions.
  5. net_pay — if shown, the actual net take-home pay after all deductions.

{
  "taxpayer_name": null, "tin": null, "tax_year": null, "form_type": null,
  "registered_address": null, "zip_code": null, "taxpayer_type": null,
  "civil_status": null, "citizenship": null,
  "date_of_birth": null,
  "business_name": null, "business_address": null, "business_tin": null,
  "line_of_business": null,
  "gross_compensation_income": null, "gross_business_income": null,
  "gross_professional_income": null, "total_gross_income": null,
  "gross_annual_income": null, "gross_monthly_income": null,
  "net_taxable_income": null, "net_pay": null,
  "allowable_deductions": null,
  "tax_due": null, "tax_credits": null, "tax_paid": null,
  "tax_still_due": null, "surcharge": null, "interest": null,
  "compromise": null, "total_amount_payable": null,
  "spouse_name": null, "spouse_tin": null
}

--- ITR DOCUMENT ---
{text}
--- END ---
```

### Output Fields
| Field | Description |
|---|---|
| `taxpayer_name` | Full name of the taxpayer |
| `tin` | Tax Identification Number |
| `tax_year` | Year of the ITR filing |
| `form_type` | e.g. BIR Form 2316, 1700, 1701 |
| `registered_address` | Employee home address (Part I) |
| `zip_code` | ZIP code |
| `taxpayer_type` | Individual, corporate, etc. |
| `civil_status` | Civil status of taxpayer |
| `citizenship` | Citizenship |
| `date_of_birth` | DOB from Part I (MM/DD/YYYY) |
| `business_name` | Business name if applicable |
| `business_address` | Business address |
| `business_tin` | Employer/business TIN |
| `line_of_business` | Nature of business |
| `gross_compensation_income` | Total gross before deductions |
| `gross_business_income` | Business income |
| `gross_professional_income` | Professional income |
| `total_gross_income` | Total of all income sources |
| `gross_annual_income` | Annual gross income |
| `gross_monthly_income` | Monthly gross income |
| `net_taxable_income` | Net taxable income after deductions |
| `net_pay` | Actual take-home pay if shown |
| `allowable_deductions` | Total allowable deductions |
| `tax_due` | Tax due amount |
| `tax_credits` | Tax credits applied |
| `tax_paid` | Total tax withheld/paid |
| `tax_still_due` | Remaining tax payable |
| `surcharge` | Surcharge if any |
| `interest` | Interest penalty if any |
| `compromise` | Compromise penalty if any |
| `total_amount_payable` | Total amount payable |
| `spouse_name` | Spouse name |
| `spouse_tin` | Spouse TIN |

---

## 4. `_UNIFIED_PROMPT`
**Variable:** `_UNIFIED_PROMPT`
**Function:** `extract_all_unified()`
**Thinking budget:** 1,024 | **Max output tokens:** 24,576
**Text slice caps:** payslip `[:15,000]` · SALN `[:20,000]` · ITR `[:15,000]`

### Purpose
Combines the payslip, SALN, and ITR extractions into a **single Gemini call**,
saving API tokens. CIC is extracted separately and passed in as a placeholder.
This is the primary extraction path used in production.

### Full Prompt
```
You are extracting structured data from multiple Philippine loan documents.
Return ONLY a single valid JSON object with exactly four top-level keys:
"cic", "payslip", "saln", "itr".
Use null for any field not found. Numbers only (no ₱, no commas).
Return [] for empty lists.

CRITICAL RULES — apply to ALL sections:
  1. Extract EVERY field present. Never return null for a field that exists.
  2. For the payslip — if multiple pay periods exist, return AVERAGES for scalar fields
     and list each period in pay_periods[].
  3. net_pay in payslip = actual take-home after ALL deductions. Extract it precisely.
  4. For CIC installment tables — extract ALL rows from both tables.
  5. If a document section is marked "NOT PROVIDED", return an empty object {}
     for that key.

═══════════════════════════════════════════════════════
SALN-SPECIFIC RULES:
═══════════════════════════════════════════════════════
  S1. personal_properties — extract EVERY row. Never stop after the first item.
  S2. No separate current value column in SALN — set current_value = acquisition_cost.
  S3. liabilities — extract EVERY row: nature, creditor, outstanding balance.
  S4. children — name and age only; return null for date_of_birth.
  S5. net_worth may be negative.
  S6. position — copy EXACTLY as printed.
  S7. CASH — always extract as TWO SEPARATE fields:
        cash_on_hand = "Cash on Hand" amount only
        cash_in_bank = "Cash in Bank" amount only
        Leave cash_on_hand_and_in_bank as null when separate fields exist.
  S8. real_properties — use these EXACT field names:
        current_fair_market_value (from the "CURRENT FAIR MARKET VALUE" column)
        acquisition_cost, year_acquired, location, assessed_value
  S9. spouse_position — extract spouse's occupation/position from the SPOUSE section.
  S10. real_properties exact key names:
        "description" = DESCRIPTION column (e.g. "Residential House and Lot")
        "kind"        = KIND column (e.g. "Residential")
        "location"    = EXACT LOCATION column
        "current_fair_market_value" = CURRENT FAIR MARKET VALUE column
        "acquisition_cost" = ACQUISITION COST column
        "year_acquired"    = ACQUISITION YEAR column

═══════════════════════════════════════════════════════
ITR / BIR 2316-SPECIFIC RULES:
═══════════════════════════════════════════════════════
  I1. date_of_birth — extract from Part I Employee Information (MM/DD/YYYY).
  I2. registered_address — employee home address from Part I.
  I3. gross_compensation_income — total gross before any deductions.
  I4. net_pay — actual take-home pay after all deductions if shown.
  I5. tax_withheld / tax_paid — the total taxes withheld from compensation.

═══════════════════════════════════════════════════════
Return exactly this JSON structure:
═══════════════════════════════════════════════════════
{
  "cic": { ... },
  "payslip": { ... },
  "saln": { ... },
  "itr": { ... }
}

=== CIC DOCUMENT ===
{cic_text}
=== END CIC ===

=== PAYSLIP DOCUMENT ===
{payslip_text}
=== END PAYSLIP ===

=== SALN DOCUMENT ===
{saln_text}
=== END SALN ===

=== ITR DOCUMENT ===
{itr_text}
=== END ITR ===
```

> **Note:** The full JSON schema for each section is identical to the individual
> prompts above. The unified prompt wraps all four schemas under a single
> top-level object with keys `"cic"`, `"payslip"`, `"saln"`, `"itr"`.

### Fallback Behavior
If this unified call fails (API error or unparseable JSON), the system
automatically falls back to calling `extract_payslip()`, `extract_saln()`,
and `extract_itr()` individually — each with their own prompts above.

---

## 5. `_CIC_KEYWORD_PROMPT`
**Variable:** `_CIC_KEYWORD_PROMPT`
**Function:** `extract_cic()` / `extract_cic_chunked()`
**Thinking budget:** 0 (OFF) | **Max output tokens:** 8,192
**Text slice cap:** `filtered_text[:20,000]` (post keyword pre-filter)

### Purpose
Scans the CIC (Credit Information Corporation) credit report for accounts
flagged as **PAST DUE** or **WRITE OFF** only. All other CIC data (personal
info, employment, etc.) is intentionally ignored — personal data is sourced
from payslip, SALN, and ITR instead.

A pre-filter (`_cic_keyword_filter()`) runs **before** this prompt to strip
out irrelevant lines, keeping only lines containing target keywords plus
5 lines of surrounding context.

### Full Prompt
```
You are reviewing a Philippine CIC (Credit Information Corporation) credit report.
Your ONLY task is to find loan or credit accounts that are flagged as either:
  • PAST DUE
  • WRITE OFF  (also: "Write-Off", "Written Off")

Return ONLY a valid JSON object — no explanation, no markdown, no extra text.
Use null for any field not found. Numbers only (no ₱, no commas).
Return [] for empty lists.

RULES:
  1. Search the ENTIRE document — both the "Requested/Renounced/Refused" table
     AND the "Active/Closed" table.
  2. Include a row ONLY if its status, contract_phase, or any adjacent label
     contains "PAST DUE", "WRITE OFF", "WRITE-OFF", or "WRITTEN OFF"
     (case-insensitive).
  3. Do NOT include rows with other statuses (e.g. Active, Closed, Requested).
  4. If no matching rows are found, return empty lists.
  5. For each matching row extract ALL of the fields listed below.
     Use null for any field not present in that row.

Return exactly this structure:
{
  "past_due": [
    {
      "provider":                null,
      "contract_type":           null,
      "contract_phase":          null,
      "financed_amount":         null,
      "outstanding_balance":     null,
      "overdue_payments_amount": null,
      "monthly_payments_amount": null,
      "contract_start_date":     null,
      "contract_end_date":       null,
      "cic_contract_code":       null
    }
  ],
  "write_off": [
    {
      "provider":                null,
      "contract_type":           null,
      "contract_phase":          null,
      "financed_amount":         null,
      "outstanding_balance":     null,
      "overdue_payments_amount": null,
      "monthly_payments_amount": null,
      "contract_start_date":     null,
      "contract_end_date":       null,
      "cic_contract_code":       null
    }
  ]
}

--- CIC DOCUMENT ---
{text}
--- END ---
```

### Output Fields
| Field | Description |
|---|---|
| `past_due[]` | Accounts with PAST DUE status |
| `write_off[]` | Accounts with WRITE OFF status |
| `provider` | Lending institution name |
| `contract_type` | Type of loan/contract |
| `contract_phase` | Phase/status label from the report |
| `financed_amount` | Original loan amount |
| `outstanding_balance` | Remaining balance |
| `overdue_payments_amount` | Total overdue amount |
| `monthly_payments_amount` | Monthly amortization |
| `contract_start_date` | Loan start date |
| `contract_end_date` | Loan end/maturity date |
| `cic_contract_code` | CIC internal contract reference code |

---

## Summary Table

| # | Prompt Variable | Used By | Thinking Budget | Max Tokens | Slice Cap |
|---|---|---|---|---|---|
| 1 | `_PAYSLIP_PROMPT` | `extract_payslip()` | 512 | 16,000 | 20,000 chars |
| 2 | `_SALN_PROMPT` | `extract_saln()` | 512 | 16,000 | 30,000 chars |
| 3 | `_ITR_PROMPT` | `extract_itr()` | 512 | 16,000 | 30,000 chars |
| 4 | `_UNIFIED_PROMPT` | `extract_all_unified()` | 1,024 | 24,576 | 15k/20k/15k chars |
| 5 | `_CIC_KEYWORD_PROMPT` | `extract_cic()` | 0 (OFF) | 8,192 | 20,000 chars (post-filter) |

> **Note:** Prompts 1–3 are **fallback only**. In normal operation, Prompt 4
> (Unified) handles payslip + SALN + ITR in a single call. Prompts 1–3 are
> triggered only when the unified call fails.

---

## JSON Repair Prompt (inline, no variable)

There is also a **sixth inline prompt** inside `_gemini_extract_json()` used
when the primary response cannot be parsed as valid JSON:

```
The following text is supposed to be a valid JSON object but may have syntax
errors. Return ONLY the corrected, complete JSON object with no explanation,
no markdown, no extra text.

{raw response from failed call, truncated to 6,000 chars}
```

**Thinking budget:** 0 (OFF) | **Max output tokens:** 4,096
This repair call uses the same model that successfully responded in the
primary call.

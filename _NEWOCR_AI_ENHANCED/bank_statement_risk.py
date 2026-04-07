"""
bank_statement_risk.py — DocExtract Pro
=========================================
Risk assessment pipeline for bank statements.

Mirrors the Bank CI two-pass pattern from extraction.py:
  Pass 1 — Structured text analysis (instant, no API call)
            Analyses extracted text for:
              • Cash flow ratio  (deposits vs withdrawals)
              • Average monthly balance
              • Income consistency
              • Irregular / large transactions
  Pass 2 — Gemini VLM second opinion
            Only triggered when Pass 1 verdict is UNCERTAIN.
            Sends the original image + OCR reference to Gemini.

Public API
----------
  assess_bank_statement(text, file_path, api_key, progress_cb) -> BankStatementRiskResult
      Main entry point called from doc_classifier_tab.py after VLM field
      extraction completes for a BANK_STATEMENT document.

  bank_statement_risk_to_text(result) -> str
      Converts BankStatementRiskResult to a plain-text report suitable for
      display in the classifier panel.
"""
from __future__ import annotations

import re
import os
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Callable

logger = logging.getLogger(__name__)

# ── Verdict constants ─────────────────────────────────────────────────────────
VERDICT_GOOD      = "GOOD"
VERDICT_BAD       = "BAD"
VERDICT_UNCERTAIN = "UNCERTAIN"

# ── Risk thresholds ───────────────────────────────────────────────────────────
# Cash flow ratio = total_deposits / total_withdrawals
# < this → high withdrawal pressure → flag BAD
_CF_RATIO_BAD      = 0.80
# ≥ this → healthy surplus → flag GOOD
_CF_RATIO_GOOD     = 1.10

# Closing balance vs average monthly deposit
# < this fraction → low liquidity → contributes NEGATIVE signal
_BALANCE_LOW_RATIO = 0.20

# A single transaction is "large" if it exceeds this multiple of avg monthly
_LARGE_TXN_MULTIPLE = 2.5

# Minimum number of months expected in a bank statement
_MIN_MONTHS = 1

# Image extensions (mirrors extraction.py)
_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp", ".gif"}

# Gemini VLM models to try in order
_VLM_MODELS = ["gemini-2.5-flash", "gemini-2.5-flash-lite"]


# ── Result dataclass ──────────────────────────────────────────────────────────

@dataclass
class CashFlowSignal:
    """One analysed signal contributing to the overall verdict."""
    label:     str   = ""
    value:     str   = ""   # human-readable value
    verdict:   str   = ""   # POSITIVE / NEGATIVE / NEUTRAL
    reason:    str   = ""


@dataclass
class BankStatementRiskResult:
    verdict:         str                  = VERDICT_UNCERTAIN
    proceed:         bool                 = False
    summary:         str                  = ""
    details:         str                  = ""
    full_text:       str                  = ""
    signals:         list[CashFlowSignal] = field(default_factory=list)
    # Raw parsed figures (for display)
    deposits:        Optional[float]      = None
    withdrawals:     Optional[float]      = None
    closing_balance: Optional[float]      = None
    period:          str                  = ""
    bank_name:       str                  = ""
    account_name:    str                  = ""
    error:           Optional[str]        = None
    vlm_used:        bool                 = False


# ── Numeric helpers ───────────────────────────────────────────────────────────

def _parse_amount(raw: str) -> Optional[float]:
    """Parse a peso amount string to float. Returns None on failure."""
    if not raw:
        return None
    try:
        s = re.sub(r"(?i)^PHP", "", str(raw).strip())
        s = re.sub(r"^[P₱]", "", s)
        s = s.replace(",", "").strip()
        return float(s) if s else None
    except (ValueError, AttributeError):
        return None


def _extract_months(period: str) -> int:
    """
    Estimate the number of months covered by a period string.
    e.g. 'Sep 10 - Oct 25, 2023' → 2
         'Jan 2023 - Dec 2023'    → 12
    Returns 1 as fallback so division never fails.
    """
    if not period:
        return 1
    # Look for two month names
    months = re.findall(
        r"\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
        r"jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|"
        r"dec(?:ember)?)\b",
        period, re.I
    )
    month_order = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4,  "may": 5,  "jun": 6,
        "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    }
    if len(months) >= 2:
        a = month_order.get(months[0][:3].lower(), 1)
        b = month_order.get(months[-1][:3].lower(), 1)
        diff = (b - a) % 12
        return max(diff + 1, 1)
    return 1


def _extract_large_transactions(text: str, avg_monthly: float) -> list[float]:
    """
    Scan raw text for individual transaction amounts that exceed
    _LARGE_TXN_MULTIPLE × avg_monthly deposit.
    Returns list of suspicious amounts (deduped, descending).
    """
    if avg_monthly <= 0:
        return []

    threshold = avg_monthly * _LARGE_TXN_MULTIPLE
    found: set[float] = set()

    for m in re.finditer(
        r"[₱P]\s*([\d,]{4,}(?:\.\d{2})?)"
        r"|PHP\s*([\d,]{4,}(?:\.\d{2})?)",
        text, re.IGNORECASE
    ):
        raw = (m.group(1) or m.group(2) or "").replace(",", "")
        try:
            val = float(raw)
        except ValueError:
            continue
        if val >= threshold:
            found.add(val)

    return sorted(found, reverse=True)[:5]


# ── Pass 1: structured text analysis ─────────────────────────────────────────

def _analyse_text(
    text:            str,
    fields:          list[tuple[str, str, str]],   # (icon, label, value)
) -> BankStatementRiskResult:
    """
    Structured pass — works entirely from already-extracted text and VLM fields.
    Never makes an API call.
    """
    result  = BankStatementRiskResult()
    signals: list[CashFlowSignal] = []

    # ── Pull values from VLM fields dict ─────────────────────────────────
    field_map: dict[str, str] = {}
    for _, label, value in fields:
        if value not in ("[not found]", "[see raw text]", "", None):
            key = label.lower().replace(" ", "_").replace("/", "_")
            field_map[key] = value

    result.deposits        = _parse_amount(field_map.get("total_deposits", ""))
    result.withdrawals     = _parse_amount(field_map.get("total_withdrawals", ""))
    result.closing_balance = _parse_amount(field_map.get("closing_balance", ""))
    result.period          = field_map.get("period", "")
    result.bank_name       = field_map.get("bank_name", "")
    result.account_name    = field_map.get("account_name", "")

    months = _extract_months(result.period)

    # ── Signal 1: Cash flow ratio ─────────────────────────────────────────
    if result.deposits is not None and result.withdrawals is not None:
        if result.withdrawals > 0:
            ratio = result.deposits / result.withdrawals
            ratio_str = f"{ratio:.2f}x  (₱{result.deposits:,.2f} deposits ÷ ₱{result.withdrawals:,.2f} withdrawals)"
            if ratio < _CF_RATIO_BAD:
                signals.append(CashFlowSignal(
                    label   = "Cash Flow Ratio",
                    value   = ratio_str,
                    verdict = "NEGATIVE",
                    reason  = f"Withdrawals exceed deposits by more than 20% — cash flow pressure detected.",
                ))
            elif ratio >= _CF_RATIO_GOOD:
                signals.append(CashFlowSignal(
                    label   = "Cash Flow Ratio",
                    value   = ratio_str,
                    verdict = "POSITIVE",
                    reason  = f"Deposits exceed withdrawals by ≥10% — healthy surplus.",
                ))
            else:
                signals.append(CashFlowSignal(
                    label   = "Cash Flow Ratio",
                    value   = ratio_str,
                    verdict = "NEUTRAL",
                    reason  = "Deposits and withdrawals are roughly balanced.",
                ))
        else:
            # Withdrawals = 0 — unusual, flag as neutral
            signals.append(CashFlowSignal(
                label   = "Cash Flow Ratio",
                value   = f"₱{result.deposits:,.2f} deposits / ₱0 withdrawals",
                verdict = "NEUTRAL",
                reason  = "No withdrawal data — cannot compute ratio.",
            ))

    # ── Signal 2: Income consistency (avg monthly deposit) ────────────────
    avg_monthly: float = 0.0
    if result.deposits is not None and months >= _MIN_MONTHS:
        avg_monthly = result.deposits / months
        avg_str = f"₱{avg_monthly:,.2f}/month  (over {months} month{'s' if months > 1 else ''})"
        if avg_monthly >= 10_000:
            signals.append(CashFlowSignal(
                label   = "Average Monthly Deposit",
                value   = avg_str,
                verdict = "POSITIVE",
                reason  = "Monthly deposit average indicates consistent income flow.",
            ))
        elif avg_monthly >= 3_000:
            signals.append(CashFlowSignal(
                label   = "Average Monthly Deposit",
                value   = avg_str,
                verdict = "NEUTRAL",
                reason  = "Monthly deposit average is modest but present.",
            ))
        else:
            signals.append(CashFlowSignal(
                label   = "Average Monthly Deposit",
                value   = avg_str,
                verdict = "NEGATIVE",
                reason  = "Monthly deposit average is very low — income consistency concern.",
            ))

    # ── Signal 3: Average monthly balance / liquidity ─────────────────────
    if result.closing_balance is not None:
        bal_str = f"₱{result.closing_balance:,.2f}"
        if avg_monthly > 0:
            liquidity_ratio = result.closing_balance / avg_monthly
            bal_str += f"  ({liquidity_ratio:.1f}× avg monthly deposit)"
            if liquidity_ratio < _BALANCE_LOW_RATIO:
                signals.append(CashFlowSignal(
                    label   = "Closing Balance vs Income",
                    value   = bal_str,
                    verdict = "NEGATIVE",
                    reason  = f"Closing balance is only {liquidity_ratio:.0%} of avg monthly deposit — very low liquidity.",
                ))
            elif liquidity_ratio >= 1.0:
                signals.append(CashFlowSignal(
                    label   = "Closing Balance vs Income",
                    value   = bal_str,
                    verdict = "POSITIVE",
                    reason  = "Closing balance ≥ 1 month of deposits — adequate liquidity buffer.",
                ))
            else:
                signals.append(CashFlowSignal(
                    label   = "Closing Balance vs Income",
                    value   = bal_str,
                    verdict = "NEUTRAL",
                    reason  = "Closing balance is below one month of deposits but above minimum threshold.",
                ))
        else:
            # No avg to compare against — just report balance
            if result.closing_balance >= 5_000:
                signals.append(CashFlowSignal(
                    label   = "Closing Balance",
                    value   = bal_str,
                    verdict = "POSITIVE",
                    reason  = "Positive closing balance observed.",
                ))
            else:
                signals.append(CashFlowSignal(
                    label   = "Closing Balance",
                    value   = bal_str,
                    verdict = "NEUTRAL",
                    reason  = "Closing balance is low but deposit totals are unavailable for comparison.",
                ))

    # ── Signal 4: Irregular / large transactions ──────────────────────────
    large_txns = _extract_large_transactions(text, avg_monthly)
    if large_txns:
        large_str = ", ".join(f"₱{v:,.2f}" for v in large_txns[:3])
        if len(large_txns) >= 3:
            signals.append(CashFlowSignal(
                label   = "Large / Irregular Transactions",
                value   = large_str + (f" (+{len(large_txns)-3} more)" if len(large_txns) > 3 else ""),
                verdict = "NEGATIVE",
                reason  = f"Multiple ({len(large_txns)}) unusually large transactions detected — manual review recommended.",
            ))
        else:
            signals.append(CashFlowSignal(
                label   = "Large Transactions",
                value   = large_str,
                verdict = "NEUTRAL",
                reason  = "One or two large transactions detected — not unusual but noted.",
            ))
    else:
        if avg_monthly > 0:
            signals.append(CashFlowSignal(
                label   = "Large / Irregular Transactions",
                value   = "None detected",
                verdict = "POSITIVE",
                reason  = "No unusually large transactions above threshold.",
            ))

    result.signals = signals

    # ── Aggregate signals → verdict ───────────────────────────────────────
    neg_count  = sum(1 for s in signals if s.verdict == "NEGATIVE")
    pos_count  = sum(1 for s in signals if s.verdict == "POSITIVE")
    neut_count = sum(1 for s in signals if s.verdict == "NEUTRAL")

    if not signals:
        # No figures could be extracted at all
        result.verdict  = VERDICT_UNCERTAIN
        result.proceed  = False
        result.summary  = "Insufficient data — no financial figures could be extracted from this statement."
        result.details  = "Try re-extracting with a higher-quality scan or check the document manually."
        result.signals  = signals
        result.full_text = _build_full_text(result)
        return result

    if neg_count >= 2:
        result.verdict = VERDICT_BAD
        result.proceed = False
        result.summary = (
            f"HIGH RISK: {neg_count} negative cash flow signal(s) detected "
            f"across {len(signals)} indicator(s)."
        )
    elif neg_count == 1 and pos_count == 0:
        result.verdict = VERDICT_UNCERTAIN
        result.proceed = False
        result.summary = (
            "MODERATE RISK: 1 negative signal detected — manual review recommended."
        )
    elif pos_count >= 2 and neg_count == 0:
        result.verdict = VERDICT_GOOD
        result.proceed = True
        result.summary = (
            f"LOW RISK: {pos_count} positive cash flow signal(s), "
            f"no adverse indicators."
        )
    elif pos_count >= 1 and neg_count == 0:
        result.verdict = VERDICT_GOOD
        result.proceed = True
        result.summary = (
            "LOW RISK: Positive cash flow indicators — no adverse signals."
        )
    else:
        result.verdict = VERDICT_UNCERTAIN
        result.proceed = False
        result.summary = (
            f"MIXED SIGNALS: {pos_count} positive, {neg_count} negative, "
            f"{neut_count} neutral — manual review recommended."
        )

    detail_lines = []
    for s in signals:
        icon = {"POSITIVE": "✅", "NEGATIVE": "❌", "NEUTRAL": "ℹ"}.get(s.verdict, "•")
        detail_lines.append(f"  {icon}  {s.label}: {s.value}")
        detail_lines.append(f"       {s.reason}")

    result.details  = "\n".join(detail_lines)
    result.full_text = _build_full_text(result)
    return result


# ── Pass 2: Gemini VLM second opinion ─────────────────────────────────────────

def _vlm_second_opinion(
    result:    BankStatementRiskResult,
    file_path: str,
    api_key:   str,
    cb:        Callable,
) -> BankStatementRiskResult:
    """
    VLM pass — only called when Pass 1 verdict is UNCERTAIN.
    Sends the image + OCR reference to Gemini and overrides the verdict
    if Gemini is decisive.
    """
    cb(78, "Uncertain — sending to Gemini VLM for second opinion…")

    try:
        from google import genai as _genai
        from google.genai import types as _gtypes
        import PIL.Image as _PILImage

        client  = _genai.Client(api_key=api_key)
        pil_img = _PILImage.open(file_path).convert("RGB")

        ocr_ref = ""
        if result.full_text:
            ocr_ref = f"\nSTRUCTURED ANALYSIS REFERENCE:\n{result.full_text[:2000]}\n"

        figures_ref = ""
        if result.deposits is not None:
            figures_ref += f"\n  Total Deposits:     ₱{result.deposits:,.2f}"
        if result.withdrawals is not None:
            figures_ref += f"\n  Total Withdrawals:  ₱{result.withdrawals:,.2f}"
        if result.closing_balance is not None:
            figures_ref += f"\n  Closing Balance:    ₱{result.closing_balance:,.2f}"
        if result.period:
            figures_ref += f"\n  Period:             {result.period}"

        vlm_prompt = (
            "You are a credit officer at Banco San Vicente (BSV), a rural bank "
            "in the Philippines. You are reviewing a bank statement for a loan application.\n\n"
            "TASK: Assess the overall cash flow health and credit risk of this "
            "bank statement.\n\n"
            "EVALUATE:\n"
            "  1. Cash flow patterns — are deposits consistently higher than withdrawals?\n"
            "  2. Income consistency — are deposits regular or erratic?\n"
            "  3. Balance levels — is the closing balance healthy vs. monthly income?\n"
            "  4. Irregular transactions — any suspiciously large single transactions?\n\n"
            "VERDICT RULES:\n"
            "  GOOD      — Healthy cash flow, consistent income, adequate balance, no red flags\n"
            "  BAD       — Withdrawals exceed deposits, very low balance, erratic income, "
            "or multiple large irregular transactions\n"
            "  UNCERTAIN — Mixed signals, insufficient data, or cannot determine clearly\n\n"
            f"EXTRACTED FIGURES:{figures_ref if figures_ref else ' (see image)'}\n"
            + ocr_ref +
            "\nRespond in this EXACT format (no extra text):\n"
            "VERDICT: GOOD | BAD | UNCERTAIN\n"
            "REASON: <one sentence>\n"
            "DETAILS: <bullet list of specific findings, or 'None'>\n"
        )

        resp = None
        for model in _VLM_MODELS:
            try:
                resp = client.models.generate_content(
                    model    = model,
                    contents = [vlm_prompt, pil_img],
                    config   = _gtypes.GenerateContentConfig(
                        max_output_tokens=512,
                        temperature=0.0,
                    ),
                )
                logger.info("Bank statement risk VLM response from %s.", model)
                break
            except Exception as e:
                err = str(e).lower()
                if any(kw in err for kw in ("429", "quota", "resource_exhausted")):
                    logger.warning("VLM %s quota hit — trying next.", model)
                    continue
                logger.warning("VLM error for bank statement risk (%s): %s", model, e)
                break

        if resp is None:
            logger.warning("VLM quota exhausted — returning structured result.")
            cb(100, "VLM quota exhausted — using structured analysis.")
            return result

        # Parse response
        vlm_text = ""
        try:
            vlm_text = resp.text or ""
        except Exception:
            pass
        if not vlm_text:
            try:
                vlm_text = "".join(
                    p.text for p in resp.candidates[0].content.parts
                    if hasattr(p, "text") and p.text
                )
            except Exception:
                pass

        if not vlm_text.strip():
            cb(100, "VLM returned empty — using structured analysis.")
            return result

        # Extract verdict
        vlm_upper   = vlm_text.upper()
        vlm_verdict = VERDICT_UNCERTAIN
        if "VERDICT: GOOD" in vlm_upper or "VERDICT:GOOD" in vlm_upper:
            vlm_verdict = VERDICT_GOOD
        elif "VERDICT: BAD" in vlm_upper or "VERDICT:BAD" in vlm_upper:
            vlm_verdict = VERDICT_BAD

        vlm_reason = ""
        for line in vlm_text.splitlines():
            if line.strip().upper().startswith("REASON:"):
                vlm_reason = line.split(":", 1)[1].strip()
                break

        vlm_details = ""
        for line in vlm_text.splitlines():
            if line.strip().upper().startswith("DETAILS:"):
                vlm_details = line.split(":", 1)[1].strip()
                break

        if vlm_verdict in (VERDICT_GOOD, VERDICT_BAD):
            logger.info(
                "VLM overrides structured UNCERTAIN → %s: %s",
                vlm_verdict, vlm_reason,
            )
            result.verdict   = vlm_verdict
            result.proceed   = (vlm_verdict == VERDICT_GOOD)
            result.summary   = f"[VLM] {vlm_reason or vlm_verdict}"
            result.vlm_used  = True
            if vlm_details and vlm_details.lower() not in ("none", "n/a", ""):
                result.details = (
                    f"VLM findings:\n  {vlm_details}\n\n"
                    + (result.details or "")
                ).strip()
            result.full_text = _build_full_text(result)
        else:
            logger.info("VLM also UNCERTAIN — keeping structured result.")

    except ImportError:
        logger.warning("google-genai not available — skipping VLM pass.")
    except Exception as e:
        logger.warning("Bank statement VLM pass failed (non-fatal): %s", e, exc_info=True)

    return result


# ── Report builder ────────────────────────────────────────────────────────────

def _build_full_text(r: BankStatementRiskResult) -> str:
    verdict_icons = {
        VERDICT_GOOD:      "✅  GOOD",
        VERDICT_BAD:       "❌  BAD",
        VERDICT_UNCERTAIN: "⚠   UNCERTAIN",
    }
    lines = [
        "═" * 62,
        "  BANK STATEMENT RISK ASSESSMENT  (BSV — Banco San Vicente)",
        "═" * 62,
        f"  Verdict  : {verdict_icons.get(r.verdict, r.verdict)}",
        f"  Proceed  : {'YES' if r.proceed else 'NO'}",
        f"  Summary  : {r.summary}",
    ]
    if r.bank_name:
        lines.append(f"  Bank     : {r.bank_name}")
    if r.account_name:
        lines.append(f"  Account  : {r.account_name}")
    if r.period:
        lines.append(f"  Period   : {r.period}")
    if r.vlm_used:
        lines.append(f"  Method   : Structured analysis + Gemini VLM")
    else:
        lines.append(f"  Method   : Structured text analysis")
    lines.append("─" * 62)

    if r.signals:
        lines.append(f"  {'SIGNAL':<30} {'VALUE':<28} VERDICT")
        lines.append("  " + "─" * 60)
        for s in r.signals:
            icon = {"POSITIVE": "✅", "NEGATIVE": "❌", "NEUTRAL": "ℹ"}.get(s.verdict, "•")
            val_short = s.value[:27] if len(s.value) > 27 else s.value
            lines.append(
                f"  {s.label[:29]:<30} {val_short:<28} {icon} {s.verdict}"
            )
    else:
        lines.append("  (no signals — insufficient data)")

    if r.details:
        lines.append("─" * 62)
        lines.append("  DETAILS:")
        for dl in r.details.splitlines():
            lines.append(f"  {dl}")

    lines.append("═" * 62)
    return "\n".join(lines)


def bank_statement_risk_to_text(result: BankStatementRiskResult) -> str:
    """Public helper — returns full_text, building it if needed."""
    if result.full_text and result.full_text.strip():
        return result.full_text
    return _build_full_text(result)


# ── Main entry point ──────────────────────────────────────────────────────────

def assess_bank_statement(
    text:      str,
    file_path: str                   = "",
    fields:    list                  = None,
    api_key:   str | None            = None,
    progress_cb: Callable | None     = None,
) -> BankStatementRiskResult:
    """
    Main bank statement risk assessment pipeline.

    Parameters
    ----------
    text        : extracted text from the bank statement
    file_path   : original file path (used for VLM pass on images)
    fields      : VLM-extracted fields as [(icon, label, value), ...]
                  If None, falls back to regex extraction from text.
    api_key     : Gemini API key for VLM pass
    progress_cb : optional (pct: int, msg: str) -> None

    Returns
    -------
    BankStatementRiskResult with verdict, signals, and formatted report.
    """
    def _cb(pct: int, msg: str = ""):
        if progress_cb:
            try:
                progress_cb(pct, msg)
            except Exception:
                pass

    _cb(5, "Parsing bank statement figures…")

    # If no fields provided, build a minimal fields list from regex on text
    if fields is None:
        fields = _fields_from_text(text)

    _cb(20, "Analysing cash flow patterns…")
    result = _analyse_text(text, fields)

    _cb(70, f"Structured analysis complete — verdict: {result.verdict}")

    # Pass 2: VLM only for UNCERTAIN and only for image files
    is_image = file_path and Path(file_path).suffix.lower() in _IMAGE_EXTS
    api_key  = api_key or os.environ.get("GEMINI_API_KEY", "")
    vlm_ok   = bool(api_key and api_key != "YOUR_GEMINI_API_KEY_HERE")

    if result.verdict == VERDICT_UNCERTAIN and is_image and vlm_ok:
        result = _vlm_second_opinion(result, file_path, api_key, _cb)

    _cb(100, f"Risk assessment complete — {result.verdict}.")
    return result


def _fields_from_text(text: str) -> list[tuple[str, str, str]]:
    """
    Fallback: extract key figures directly from raw text using regex.
    Returns a minimal (icon, label, value) list compatible with _analyse_text.
    """
    patterns = {
        "total_deposits":    r"(?:deposits?|total\s+(?:credit|deposit)s?)\s*[P₱]\s*([\d,]+(?:\.\d+)?)",
        "total_withdrawals": r"(?:withdrawals?|total\s+(?:debit|withdrawal)s?)\s*[P₱]\s*([\d,]+(?:\.\d+)?)",
        "closing_balance":   r"(?:closing|ending|available)\s+balance[^\d\n]*([\d,]+\.\d{2})",
        "period":            r"(?:for\s+([A-Za-z]+\s+\d+\s*[-–]\s*[A-Za-z]*\s*\d+,?\s*\d{4})|(?:period|statement\s+period)\s*[:\-]\s*([^\n]+))",
        "bank_name":         r"(?i)\b(BDO|BPI|Metrobank|Landbank|UnionBank|Security\s*Bank|PNB|RCBC|Chinabank|EastWest|PSBank)\b",
        "account_name":      r"account\s+name\s*[:\-]\s*([^\n]{3,60})",
    }
    label_map = {
        "total_deposits":    ("Total Deposits",    "⬆"),
        "total_withdrawals": ("Total Withdrawals", "⬇"),
        "closing_balance":   ("Closing Balance",   "💰"),
        "period":            ("Period",            "📅"),
        "bank_name":         ("Bank Name",         "🏦"),
        "account_name":      ("Account Name",      "👤"),
    }
    results = []
    for key, pattern in patterns.items():
        m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if m:
            val = next((g.strip() for g in m.groups() if g), "").strip()[:80]
            if val:
                icon, lbl = label_map[key]
                results.append((icon, lbl, val))
    return results
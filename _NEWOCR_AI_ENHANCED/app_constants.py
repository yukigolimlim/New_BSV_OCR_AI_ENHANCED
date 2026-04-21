"""
app_constants.py — DocExtract Pro
Single source of truth for paths, colours, fonts, gemini_chat, rule helpers.
Every other module does:  from app_constants import *
"""
from __future__ import annotations
import os, re as _re
from pathlib import Path

# ── Paths ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR   = Path(__file__).parent
LOGO_PATH    = SCRIPT_DIR / "bsv_logotxt.png"
POPPLER_PATH = r"C:\poppler\Release-25.12.0-0\poppler-25.12.0\Library\bin"
IMAGE_EXTS   = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp", ".gif"}

_env_path = SCRIPT_DIR / ".env"
if _env_path.exists():
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

# ── API ───────────────────────────────────────────────────────────────────────
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY_HERE")
GEMINI_MODEL   = "gemini-2.5-flash"
GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models"
FALLBACK_MODEL = "gemini-2.0-flash"

# ── Colours ───────────────────────────────────────────────────────────────────
NAVY_DEEP="#0A1628"; NAVY="#112240"; NAVY_MID="#1B3A6B"; NAVY_LIGHT="#2D5FA6"
NAVY_PALE="#5B8FD4"; NAVY_GHOST="#C5D8F5"; NAVY_MIST="#EAF1FB"
WHITE="#FFFFFF"; OFF_WHITE="#F7F9FC"; CARD_WHITE="#FFFFFF"; PANEL_LEFT="#F0F4FA"
LIME_BRIGHT="#00E5A0"; LIME="#00C48A"; LIME_MID="#00A876"; LIME_DARK="#007A56"
LIME_PALE="#B3F5E2"; LIME_MIST="#E6FBF5"
BORDER_LIGHT="#DDE6F4"; BORDER_MID="#B8CCEA"
TXT_NAVY="#0A1628"; TXT_NAVY_MID="#1B3A6B"; TXT_SOFT="#5B7BAD"
TXT_MUTED="#96AFCC"; TXT_ON_LIME="#002E1F"
ACCENT_GOLD="#F59E0B"; ACCENT_SUCCESS="#00C48A"; ACCENT_RED="#EF4444"
BUBBLE_USER="#112240"; BUBBLE_USER_TXT="#FFFFFF"
BUBBLE_AI="#EAF1FB"; BUBBLE_AI_TXT="#0A1628"
BUBBLE_SYS="#FFF8EC"; BUBBLE_SYS_TXT="#78450A"
SIDEBAR_BG="#0D1E3A"; SIDEBAR_ITEM="#142B52"; SIDEBAR_HVR="#1B3A6B"

# ── Fonts ─────────────────────────────────────────────────────────────────────
_FONT_FAMILY: str | None = None
_UI_ZOOM: float = 1.0

def get_ui_zoom() -> float:
    return float(_UI_ZOOM)

def set_ui_zoom(v: float) -> float:
    global _UI_ZOOM
    try:
        z = float(v)
    except Exception:
        z = 1.0
    _UI_ZOOM = max(0.8, min(1.6, z))
    return _UI_ZOOM

def best_font() -> str:
    import tkinter.font as tkfont
    available = set(tkfont.families())
    for f in ("Nunito","Montserrat","Poppins","Segoe UI","Calibri","Arial"):
        if f in available: return f
    return "Arial"

def register_fonts() -> None:
    try:
        import pyglet
        for ttf in SCRIPT_DIR.glob("Montserrat*.ttf"):
            pyglet.font.add_file(str(ttf))
    except ImportError: pass

def F(size: int, weight: str = "normal") -> tuple:
    global _FONT_FAMILY
    if _FONT_FAMILY is None: _FONT_FAMILY = best_font()
    return (_FONT_FAMILY, max(6, int(round(size * _UI_ZOOM))), weight)

def FMONO(size: int, weight: str = "normal") -> tuple:
    import tkinter.font as tkfont
    available = set(tkfont.families())
    for f in ("JetBrains Mono","Cascadia Code","Consolas","Courier New"):
        if f in available: return (f, max(6, int(round(size * _UI_ZOOM))), weight)
    return ("Courier New", max(6, int(round(size * _UI_ZOOM))), weight)

def hex_blend(c1: str, c2: str, t: float) -> str:
    r1,g1,b1 = int(c1[1:3],16),int(c1[3:5],16),int(c1[5:7],16)
    r2,g2,b2 = int(c2[1:3],16),int(c2[3:5],16),int(c2[5:7],16)
    return f"#{int(r1+(r2-r1)*t):02x}{int(g1+(g2-g1)*t):02x}{int(b1+(b2-b1)*t):02x}"

# ══════════════════════════════════════════════════════════════════════════════
#  GEMINI CHAT
# ══════════════════════════════════════════════════════════════════════════════
def gemini_chat(messages: list, api_key: str, model: str = GEMINI_MODEL,
                max_tokens: int = 2048) -> tuple[str, str]:
    if not api_key or api_key == "YOUR_GEMINI_API_KEY_HERE":
        return ("⚠ ERROR: No Gemini API key configured.\n\n"
                "Set GEMINI_API_KEY in your .env file.\n\n"
                "Get a FREE key at: https://aistudio.google.com/app/apikey", "error")
    try:
        from google import genai as _genai
        from google.genai import types as _gtypes
        client = _genai.Client(api_key=api_key)
        system_parts, contents, last_user = [], [], ""
        for msg in messages:
            role, content = msg["role"], msg["content"]
            if role == "system": system_parts.append(content)
            elif role == "user":
                contents.append(_gtypes.Content(role="user", parts=[_gtypes.Part(text=content)]))
                last_user = content
            elif role == "assistant":
                contents.append(_gtypes.Content(role="model", parts=[_gtypes.Part(text=content)]))
        system_instruction = "\n\n".join(system_parts) if system_parts else None
        prior = contents[:-1] if contents and contents[-1].role == "user" else contents

        def _safe_text(resp) -> str:
            try: return resp.text or ""
            except Exception:
                try: return "".join(p.text for p in resp.candidates[0].content.parts if hasattr(p,"text") and p.text)
                except Exception: return ""

        def _call(m: str) -> str:
            cfg = {"max_output_tokens": max_tokens}
            if system_instruction: cfg["system_instruction"] = system_instruction
            all_c = prior + [_gtypes.Content(role="user", parts=[_gtypes.Part(text=last_user)])]
            return _safe_text(client.models.generate_content(
                model=m, contents=all_c, config=_gtypes.GenerateContentConfig(**cfg)))

        try:
            return _call(model), "primary"
        except Exception as e:
            err = str(e).lower()
            if any(kw in err for kw in ("quota","rate","429","resource_exhausted")):
                try: return _call(FALLBACK_MODEL), "fallback"
                except Exception as e2:
                    return (f"⚠ Both models quota-limited.\n\nDetails: {e2}", "error")
            if any(kw in err for kw in ("token","too long","context","size")):
                trimmed = (system_instruction or "")[:2_000] + "\n\n[… context trimmed …]"
                try:
                    cfg2 = {"max_output_tokens": max_tokens, "system_instruction": trimmed}
                    all_c2 = prior + [_gtypes.Content(role="user", parts=[_gtypes.Part(text=last_user)])]
                    return _safe_text(client.models.generate_content(
                        model=model, contents=all_c2,
                        config=_gtypes.GenerateContentConfig(**cfg2))), "trimmed"
                except Exception: pass
            return f"⚠ Gemini error: {type(e).__name__}: {e}", "error"
    except ImportError:
        return "⚠ google-genai not installed.\n\nRun: pip install google-genai", "error"

# ══════════════════════════════════════════════════════════════════════════════
#  RULE-BASED BANK CI HELPERS
# ══════════════════════════════════════════════════════════════════════════════
_RULE_POSITIVE_KW = {
    "ncd","no contrary data","no derogatory","no adverse",
    "good standing","current","no outstanding","compliant",
    "paid","closed","settled","regular",
}
_RULE_NEGATIVE_KW = {
    "dishonored","bounced","nsf","insufficient funds",
    "overdraft","overdrawn","past due","past-due","delinquent",
    "default","written off","write-off","chargeoff","charge-off",
    "legal action","collection","adverse","derogatory",
    "blacklist","unpaid","overdue",
}

def _rule_fuzzy_ncd(text: str) -> bool:
    t = text.upper().strip()
    cleaned = _re.sub(r"[^A-Z0-9]","",t)
    if cleaned in ("NCD","NOCONTRARYDATA","NODEROGATORY","NOADVERSE"): return True
    if _re.match(r"N.{0,2}C.{0,2}D",t): return True
    a,b = cleaned[:5],"NCD"
    if len(a)<len(b): a,b=b,a
    if not b: return len(a)<=1
    prev = list(range(len(b)+1))
    for ca in a:
        curr=[prev[0]+1]
        for j,cb in enumerate(b,1):
            curr.append(min(prev[j]+1,curr[j-1]+1,prev[j-1]+(ca!=cb)))
        prev=curr
    return prev[-1]<=1

def _rule_assess_text(bank_ci_text: str) -> tuple[str,bool,str,str]:
    text = (bank_ci_text or "").strip()
    if not text:
        return ("UNCERTAIN",False,"No text extracted from Bank CI — please review manually.","Document appears empty or unreadable.")
    positive_rows,negative_rows,inconclusive_rows=[],[],[]
    for line in text.splitlines():
        s=line.strip()
        if s.startswith("✅") or "POSITIVE" in s: positive_rows.append(s)
        elif s.startswith("❌") or "NEGATIVE" in s: negative_rows.append(s)
        elif s.startswith("⚠") or "INCONCLUSIVE" in s: inconclusive_rows.append(s)
    if negative_rows or positive_rows or inconclusive_rows:
        if negative_rows:
            details="\n".join(f"  ❌  {r}" for r in negative_rows)
            if positive_rows: details+="\n"+"\n".join(f"  ✅  {r}" for r in positive_rows)
            return "BAD",False,f"ADVERSE: {len(negative_rows)} negative record(s) detected.",details
        if positive_rows:
            details="\n".join(f"  ✅  {r}" for r in positive_rows)
            if inconclusive_rows: details+="\n"+"\n".join(f"  ⚠   {r}" for r in inconclusive_rows)
            return "GOOD",True,f"CLEAN: {len(positive_rows)} positive record(s), no adverse findings.",details
        return ("UNCERTAIN",False,
                f"INCONCLUSIVE: {len(inconclusive_rows)} row(s) could not be fully assessed. Manual review recommended.",
                "\n".join(f"  ⚠   {r}" for r in inconclusive_rows))
    t_lower=text.lower()
    def _wm(phrase,hay):
        if " " in phrase: return phrase in hay
        return bool(_re.search(r"\b"+_re.escape(phrase)+r"\b",hay))
    ncd_found=_rule_fuzzy_ncd(text) or _wm("ncd",t_lower)
    pos_hits=[kw for kw in _RULE_POSITIVE_KW if _wm(kw,t_lower)]
    covered={nk for pk in pos_hits for nk in _RULE_NEGATIVE_KW if nk in pk}
    neg_hits=[kw for kw in _RULE_NEGATIVE_KW if kw not in covered and _wm(kw,t_lower)]
    if ncd_found:
        d="  Detected: 'NCD'"
        if neg_hits: d+="\n  Note: derogatory overridden by NCD: "+", ".join(f"'{k}'" for k in neg_hits)
        return "GOOD",True,"CLEAN: NCD (No Contrary Data) confirmed.",d
    if neg_hits:
        return "BAD",False,"ADVERSE: derogatory keyword(s) found.","  Detected: "+", ".join(f"'{k}'" for k in neg_hits)
    if pos_hits:
        return "GOOD",True,"CLEAN: positive indicator(s) found — no adverse keywords.","  Detected: "+", ".join(f"'{k}'" for k in pos_hits)
    return "UNCERTAIN",False,"No clear indicators found. Manual review recommended.","  Tip: Ensure the Bank CI document is clear and legible."

def _rule_extract_cic_tier(cic_text: str, has_cic: bool) -> dict:
    if not has_cic or not (cic_text or "").strip():
        return {"loan_amount":None,"tier":"below_100k","has_cic":has_cic,
                "applicant_name":"","summary":"No CIC provided — loan tier assumed below ₱100,000."}
    tier,loan_amount,applicant,summary="below_100k",None,"",""
    pats=[r"[₱P]\s*([\d,]+(?:\.\d+)?)",r"PHP\s*([\d,]+(?:\.\d+)?)",
          r"(?:loan|amount|balance|outstanding)[^₱\d]{0,20}([\d,]{5,}(?:\.\d+)?)"]
    amounts=[]
    for pat in pats:
        for m in _re.finditer(pat,cic_text,_re.I):
            try:
                v=float(m.group(1).replace(",",""))
                if v>=1_000: amounts.append(v)
            except (ValueError,IndexError): pass
    if amounts:
        loan_amount=max(amounts)
        tier="above_100k" if loan_amount>100_000 else "below_100k"
        summary=(f"Largest amount detected: ₱{loan_amount:,.0f} → "
                 f"{'above' if tier=='above_100k' else 'at or below'} ₱100,000 threshold.")
    else:
        summary="No loan amounts detected in CIC — assumed below ₱100,000."
    nm=_re.search(r"(?:full\s+name|subject|borrower|applicant|name)\s*[:\-]\s*([^\n,]{3,50})",cic_text,_re.I)
    if nm: applicant=nm.group(1).strip()
    return {"loan_amount":loan_amount,"tier":tier,"has_cic":has_cic,"applicant_name":applicant,"summary":summary}

def _ai_check_stage1(bank_ci_text: str, cic_text: str, api_key: str) -> tuple[dict,dict]:
    """Rule-based Bank CI + CIC tier evaluation. No Gemini calls."""
    has_cic=bool((cic_text or "").strip())
    verdict,proceed,summary,details=_rule_assess_text(bank_ci_text)
    bank_ci_result={"verdict":verdict,"summary":summary,"details":details,"proceed":proceed}
    cic_tier_result=_rule_extract_cic_tier(cic_text,has_cic)
    return bank_ci_result,cic_tier_result

def FF(size: int, weight: str = "normal"):
    """
    Returns a ctk.CTkFont using the resolved _FONT_FAMILY.
    Use this instead of ctk.CTkFont(_FONT_FAMILY, ...) in ui_* files
    so the font family is read at call-time, not at import-time.
    """
    import customtkinter as ctk
    import app_constants as _ac
    if _ac._FONT_FAMILY is None:
        _ac._FONT_FAMILY = best_font()
    return ctk.CTkFont(_ac._FONT_FAMILY, size, weight=weight)
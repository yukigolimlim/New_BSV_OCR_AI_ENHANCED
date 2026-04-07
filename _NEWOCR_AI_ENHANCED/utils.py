"""
utils.py — DocExtract Pro
==========================
Shared constants (colours, sizes, paths) and pure helper functions.
No UI imports — safe to import from both app.py and widgets.py.
"""
from pathlib import Path

# ── Window ────────────────────────────────────────────────────────────────────
WIN_W = 1240
WIN_H = 780

# ── Paths ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR   = Path(__file__).parent
LOGO_PATH    = SCRIPT_DIR / "bsv_logotxt.png"
POPPLER_PATH = r"C:\poppler\Release-25.12.0-0\poppler-25.12.0\Library\bin"
IMAGE_EXTS   = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp", ".gif"}

# ── Navy scale ────────────────────────────────────────────────────────────────
NAVY_DEEP    = "#1C2E7A"
NAVY         = "#2B45A8"
NAVY_MID     = "#3858C8"
NAVY_LIGHT   = "#4E6EE0"
NAVY_PALE    = "#7A96F0"
NAVY_GHOST   = "#C8D4F8"
NAVY_MIST    = "#E8EEFF"

# ── Whites / backgrounds ──────────────────────────────────────────────────────
WHITE        = "#FFFFFF"
OFF_WHITE    = "#F5F7FF"
CARD_WHITE   = "#FFFFFF"
PANEL_LEFT   = "#F0F3FF"

# ── Lime / accent-green scale ─────────────────────────────────────────────────
LIME_BRIGHT  = "#C8F020"
LIME         = "#A8D818"
LIME_MID     = "#90C010"
LIME_DARK    = "#6A9408"
LIME_PALE    = "#E0F870"
LIME_MIST    = "#F2FFCC"

# ── Borders ───────────────────────────────────────────────────────────────────
BORDER_LIGHT = "#D4DCF8"
BORDER_MID   = "#B0BEEC"

# ── Text colours ──────────────────────────────────────────────────────────────
TXT_NAVY     = "#1C2E7A"
TXT_NAVY_MID = "#3858C8"
TXT_SOFT     = "#6878B8"
TXT_MUTED    = "#9AAACE"
TXT_ON_LIME  = "#1C2E7A"

# ── Semantic accents ──────────────────────────────────────────────────────────
ACCENT_GOLD    = "#F0A800"
ACCENT_SUCCESS = "#22C870"
ACCENT_RED     = "#E74C3C"

import re

def strip_thinking(text: str) -> str:
    """
    Strip internal reasoning blocks that some models (Gemini 2.5, Qwen, etc.)
    emit before their actual response.
    
    Removes:
      <think>...</think>
      <thinking>...</thinking>
      <reflection>...</reflection>
    Also strips any leading/trailing whitespace left behind.
    """
    if not text:
        return text
    # Remove <think>...</think> and variants (case-insensitive, multiline)
    cleaned = re.sub(
        r'<(think|thinking|reflection)>[\s\S]*?<\/\1>',
        '',
        text,
        flags=re.IGNORECASE
    )
    return cleaned.strip()
# ── Pure helpers ──────────────────────────────────────────────────────────────
def hex_blend(c1: str, c2: str, t: float) -> str:
    """Linear interpolation between two hex colours at position t (0–1)."""
    r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
    r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
    return (f"#{int(r1+(r2-r1)*t):02x}"
            f"{int(g1+(g2-g1)*t):02x}"
            f"{int(b1+(b2-b1)*t):02x}")
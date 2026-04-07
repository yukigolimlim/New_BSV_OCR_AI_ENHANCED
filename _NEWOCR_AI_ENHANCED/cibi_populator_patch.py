"""
cibi_populator_patch.py
=======================
Drop-in patch that makes cibi_populator.py load its five Gemini prompt
strings from ``cibi_prompts.md`` instead of using the hardcoded strings.

HOW TO APPLY
------------
Option A — Minimal one-liner at the top of cibi_populator.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Add these lines immediately AFTER the existing prompt variable definitions
(i.e. right after ``_CIC_KEYWORD_PROMPT = \"""...\"""``) in cibi_populator.py:

    # ── Load prompts from MD file (overrides hardcoded strings above) ──────
    try:
        from cibi_populator_patch import apply_prompt_patch
        apply_prompt_patch(globals())
    except ImportError:
        pass   # patch module not installed — hardcoded prompts are used

Option B — Full replacement of hardcoded blocks
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Replace the five hardcoded prompt string assignments in cibi_populator.py
with this single block:

    # ── Prompts loaded from cibi_prompts.md ────────────────────────────────
    from cibi_populator_patch import load_and_assign
    _PAYSLIP_PROMPT, _SALN_PROMPT, _ITR_PROMPT, _UNIFIED_PROMPT, \\
        _CIC_KEYWORD_PROMPT = load_and_assign()

FALLBACK SAFETY
---------------
Both options fall back silently to the hardcoded strings if:
  • cibi_populator_patch.py is not found
  • cibi_prompts.md is not found
  • A prompt block cannot be parsed from the MD

This means the application NEVER crashes because of the patch.

FILE PLACEMENT
--------------
Place cibi_populator_patch.py and cibi_prompts.md in the SAME directory as
cibi_populator.py.  The loader searches for the MD file relative to this
module's own location.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

_log = logging.getLogger("cibi_populator_patch")

# ── The five prompt variable names we manage ───────────────────────────────────
_PROMPT_NAMES: tuple[str, ...] = (
    "_PAYSLIP_PROMPT",
    "_SALN_PROMPT",
    "_ITR_PROMPT",
    "_UNIFIED_PROMPT",
    "_CIC_KEYWORD_PROMPT",
)


def apply_prompt_patch(
    caller_globals: dict,
    md_path:        Optional[str | Path] = None,
) -> dict[str, bool]:
    """
    Load prompts from the MD file and inject them into ``caller_globals``.

    This is the function you call from inside cibi_populator.py:

        from cibi_populator_patch import apply_prompt_patch
        apply_prompt_patch(globals())

    Parameters
    ----------
    caller_globals : dict
        The ``globals()`` dict from the calling module.  The five prompt
        variables will be set (or overwritten) in this namespace.
    md_path : str or Path, optional
        Path to cibi_prompts.md.  Defaults to the same directory as this file.

    Returns
    -------
    dict
        Mapping of prompt name → True (loaded from MD) / False (fallback used).
    """
    try:
        from prompt_loader import load_prompts
    except ImportError:
        _log.warning(
            "[patch] prompt_loader.py not found — "
            "all prompts will use hardcoded fallbacks."
        )
        return {name: False for name in _PROMPT_NAMES}

    prompts = load_prompts(md_path, strict=False)
    status:  dict[str, bool] = {}

    for name in _PROMPT_NAMES:
        md_text = prompts.get(name)
        if md_text:
            old = caller_globals.get(name, "")
            caller_globals[name] = md_text
            status[name] = True
            _log.info(
                f"[patch] {name} — replaced hardcoded "
                f"({len(old):,} chars) with MD version ({len(md_text):,} chars)."
            )
        else:
            status[name] = False
            _log.warning(
                f"[patch] {name} — MD block missing; "
                f"hardcoded string kept."
            )

    loaded  = sum(1 for v in status.values() if v)
    missing = len(status) - loaded
    _log.info(
        f"[patch] Prompt patch complete: "
        f"{loaded} loaded from MD, {missing} using hardcoded fallback."
    )
    return status


def load_and_assign(
    md_path: Optional[str | Path] = None,
) -> tuple[str, str, str, str, str]:
    """
    Load all five prompts and return them as a tuple for direct assignment.

    Usage in cibi_populator.py (Option B):

        from cibi_populator_patch import load_and_assign
        _PAYSLIP_PROMPT, _SALN_PROMPT, _ITR_PROMPT, _UNIFIED_PROMPT, \\
            _CIC_KEYWORD_PROMPT = load_and_assign()

    Falls back to empty string for any missing prompt — the calling module
    should define the hardcoded strings BEFORE this call so they can serve
    as the real fallback via the apply_prompt_patch() path.

    Parameters
    ----------
    md_path : str or Path, optional
        Path to cibi_prompts.md.

    Returns
    -------
    tuple of 5 str
        (_PAYSLIP_PROMPT, _SALN_PROMPT, _ITR_PROMPT,
         _UNIFIED_PROMPT, _CIC_KEYWORD_PROMPT)
    """
    try:
        from prompt_loader import load_prompts
        prompts = load_prompts(md_path, strict=False)
    except ImportError:
        _log.warning("[patch] prompt_loader.py not found; returning empty strings.")
        prompts = {}

    return tuple(
        prompts.get(name) or ""
        for name in _PROMPT_NAMES
    )  # type: ignore[return-value]


def diff_prompts(
    md_path: Optional[str | Path] = None,
) -> None:
    """
    Print a character-level diff summary between the MD prompts and any
    currently active prompts in the Python environment.

    Useful for verifying that the MD and the hardcoded strings are in sync.

    Usage (run from terminal):

        python -c "from cibi_populator_patch import diff_prompts; diff_prompts()"
    """
    try:
        from prompt_loader import load_prompts
        md_prompts = load_prompts(md_path, strict=False)
    except ImportError:
        print("prompt_loader.py not found.")
        return

    try:
        import cibi_populator as _cp  # noqa: F401
        hc: dict[str, str] = {
            name: getattr(_cp, name, "")
            for name in _PROMPT_NAMES
        }
    except ImportError:
        print("cibi_populator.py not importable — cannot compare.")
        return

    print()
    print("=" * 62)
    print("Prompt diff: MD file  vs  cibi_populator.py hardcoded strings")
    print("=" * 62)
    for name in _PROMPT_NAMES:
        md_text = md_prompts.get(name) or ""
        hc_text = hc.get(name) or ""
        same = md_text.strip() == hc_text.strip()
        icon = "✅ SAME" if same else "⚠  DIFF"
        print(
            f"  {icon}  {name:25s}  "
            f"MD={len(md_text):>6,}c  HC={len(hc_text):>6,}c"
        )
        if not same and md_text and hc_text:
            # Show first differing line
            md_lines = md_text.strip().splitlines()
            hc_lines = hc_text.strip().splitlines()
            for i, (a, b) in enumerate(zip(md_lines, hc_lines)):
                if a != b:
                    print(f"         First diff at line {i+1}:")
                    print(f"         MD : {a[:80]!r}")
                    print(f"         HC : {b[:80]!r}")
                    break
    print()


# ── Self-test ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")

    print()
    print("=" * 62)
    print("cibi_populator_patch — self-test")
    print("=" * 62)

    # Simulate a caller_globals dict with dummy hardcoded prompts
    fake_globals: dict = {
        name: f"HARDCODED_{name}" for name in _PROMPT_NAMES
    }

    print()
    print("Before patch:")
    for name in _PROMPT_NAMES:
        print(f"  {name:25s} = {fake_globals[name]!r}")

    print()
    md_arg = sys.argv[1] if len(sys.argv) > 1 else None
    status = apply_prompt_patch(fake_globals, md_path=md_arg)

    print()
    print("After patch:")
    all_ok = True
    for name in _PROMPT_NAMES:
        loaded = status.get(name, False)
        val    = fake_globals.get(name, "")
        icon   = "✅" if loaded else "⚠ "
        src    = "MD file" if loaded else "hardcoded fallback"
        print(
            f"  {icon}  {name:25s}  "
            f"source={src:18s}  "
            f"chars={len(val):>6,}  "
            f"preview={val[:50]!r}"
        )
        if not loaded:
            all_ok = False

    print()
    if all_ok:
        print("All prompts loaded from MD file. ✅")
    else:
        print("Some prompts used hardcoded fallback — check MD path/structure. ⚠")
        sys.exit(1)
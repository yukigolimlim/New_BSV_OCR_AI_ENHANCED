"""
prompt_loader.py
================
Loads Gemini prompt strings from cibi_prompts.md at runtime, replacing the
hardcoded string variables in cibi_populator.py.

HOW IT WORKS
------------
The MD file uses this repeating pattern for every prompt:

    ## N. `_VARIABLE_NAME`
    ...
    ### Full Prompt
    ```
    <prompt text here>
    ```

This loader:
  1. Finds each "### Full Prompt" section.
  2. Reads the variable name from the nearest preceding "## N. `_VAR`" heading.
  3. Extracts the fenced code block content between the ``` markers.
  4. Exposes each prompt as a module attribute AND returns them as a dict.

USAGE IN cibi_populator.py
--------------------------
Replace the five hardcoded prompt strings at the top of the file with:

    from prompt_loader import load_prompts
    _PROMPTS = load_prompts()          # loads from cibi_prompts.md by default
    _PAYSLIP_PROMPT     = _PROMPTS["_PAYSLIP_PROMPT"]
    _SALN_PROMPT        = _PROMPTS["_SALN_PROMPT"]
    _ITR_PROMPT         = _PROMPTS["_ITR_PROMPT"]
    _UNIFIED_PROMPT     = _PROMPTS["_UNIFIED_PROMPT"]
    _CIC_KEYWORD_PROMPT = _PROMPTS["_CIC_KEYWORD_PROMPT"]

FALLBACK BEHAVIOUR
------------------
If the MD file is missing, unreadable, or a particular prompt block cannot
be found, the loader logs a warning and returns None for that key.
cibi_populator.py should then fall back to its own hardcoded strings so
that the application never crashes due to a missing MD file.

MD FILE LOCATION
----------------
Default: same directory as prompt_loader.py, filename "cibi_prompts.md".
Override with the `md_path` argument to load_prompts().
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Optional

_log = logging.getLogger("prompt_loader")

# ── Public API ────────────────────────────────────────────────────────────────

# Known prompt variable names in the order they appear in the MD.
# Used for validation and ordering of the returned dict.
KNOWN_PROMPTS: tuple[str, ...] = (
    "_PAYSLIP_PROMPT",
    "_SALN_PROMPT",
    "_ITR_PROMPT",
    "_UNIFIED_PROMPT",
    "_CIC_KEYWORD_PROMPT",
)


def load_prompts(
    md_path: Optional[str | Path] = None,
    *,
    strict: bool = False,
) -> dict[str, str | None]:
    """
    Parse ``cibi_prompts.md`` and return a dict mapping variable name → prompt text.

    Parameters
    ----------
    md_path : str or Path, optional
        Path to the markdown file.  Defaults to ``cibi_prompts.md`` in the
        same directory as this module.
    strict : bool
        If True, raise FileNotFoundError / ValueError when the file or a
        prompt block is missing.  If False (default), log a warning and
        return None for missing entries.

    Returns
    -------
    dict
        Keys are the prompt variable names (e.g. ``"_PAYSLIP_PROMPT"``).
        Values are the extracted prompt strings, or None if not found.

    Example
    -------
    >>> from prompt_loader import load_prompts
    >>> prompts = load_prompts()
    >>> print(prompts["_PAYSLIP_PROMPT"][:80])
    You are extracting structured data from a Philippine payslip or payroll doc
    """
    path = _resolve_path(md_path, strict=strict)
    if path is None:
        return {k: None for k in KNOWN_PROMPTS}

    raw = _read_file(path, strict=strict)
    if raw is None:
        return {k: None for k in KNOWN_PROMPTS}

    extracted = _parse_prompts(raw)

    result: dict[str, str | None] = {}
    for name in KNOWN_PROMPTS:
        text = extracted.get(name)
        if text is None:
            msg = (
                f"[prompt_loader] Prompt block '{name}' not found in {path}. "
                f"cibi_populator.py will use its hardcoded fallback."
            )
            if strict:
                raise ValueError(msg)
            _log.warning(msg)
        else:
            _log.info(
                f"[prompt_loader] Loaded '{name}' "
                f"({len(text):,} chars) from {path.name}"
            )
        result[name] = text

    # Also include any extra prompts found in the MD that aren't in KNOWN_PROMPTS
    for name, text in extracted.items():
        if name not in result:
            result[name] = text
            _log.info(
                f"[prompt_loader] Extra prompt '{name}' "
                f"({len(text):,} chars) loaded from {path.name}"
            )

    return result


# ── Internal helpers ──────────────────────────────────────────────────────────

def _resolve_path(
    md_path: Optional[str | Path],
    *,
    strict: bool,
) -> Optional[Path]:
    """Resolve the MD file path, falling back to the default location."""
    if md_path is not None:
        p = Path(md_path)
    else:
        # Default: cibi_prompts.md next to this file
        p = Path(__file__).resolve().parent / "cibi_prompts.md"

    if not p.exists():
        msg = f"[prompt_loader] MD file not found: {p}"
        if strict:
            raise FileNotFoundError(msg)
        _log.warning(msg)
        return None

    return p


def _read_file(path: Path, *, strict: bool) -> Optional[str]:
    """Read the markdown file, returning None on error (unless strict)."""
    try:
        return path.read_text(encoding="utf-8")
    except Exception as exc:
        msg = f"[prompt_loader] Could not read {path}: {exc}"
        if strict:
            raise OSError(msg) from exc
        _log.warning(msg)
        return None


# Regex to find sections like:  ## 1. `_PAYSLIP_PROMPT`
_HEADING_RE = re.compile(
    r"^##\s+\d+\.\s+`(_[A-Z_]+)`",
    re.MULTILINE,
)

# Regex to find:  ### Full Prompt\n```\n<content>\n```
# Uses non-greedy match; handles optional language tag after the opening fence.
_FENCE_RE = re.compile(
    r"###\s+Full Prompt\s*\n```[^\n]*\n([\s\S]*?)```",
    re.MULTILINE,
)


def _parse_prompts(md_text: str) -> dict[str, str]:
    """
    Extract all prompt blocks from the markdown text.

    Strategy
    --------
    1. Find every "## N. `_VAR_NAME`" heading and its character position.
    2. Find every "### Full Prompt / ``` ... ```" block and its position.
    3. For each fence block, walk backwards to find the nearest heading — that
       heading's variable name owns the fence block.
    """
    # Step 1: collect all headings with their positions
    headings: list[tuple[int, str]] = [
        (m.start(), m.group(1))
        for m in _HEADING_RE.finditer(md_text)
    ]

    if not headings:
        _log.warning("[prompt_loader] No prompt headings found in the MD file.")
        return {}

    # Step 2: collect all fence blocks with their positions
    fences: list[tuple[int, str]] = [
        (m.start(), m.group(1).strip())
        for m in _FENCE_RE.finditer(md_text)
    ]

    if not fences:
        _log.warning("[prompt_loader] No '### Full Prompt' code blocks found.")
        return {}

    # Step 3: match each fence to its nearest preceding heading
    extracted: dict[str, str] = {}
    heading_positions = [pos for pos, _ in headings]

    for fence_pos, prompt_text in fences:
        # Find the last heading that starts before this fence
        owner_idx = None
        for i, hpos in enumerate(heading_positions):
            if hpos < fence_pos:
                owner_idx = i
            else:
                break

        if owner_idx is None:
            _log.warning(
                f"[prompt_loader] Found a fence block at position {fence_pos} "
                f"with no preceding heading — skipping."
            )
            continue

        var_name = headings[owner_idx][1]

        if var_name in extracted:
            _log.warning(
                f"[prompt_loader] Duplicate fence block for '{var_name}' — "
                f"keeping first occurrence."
            )
            continue

        extracted[var_name] = prompt_text
        _log.debug(
            f"[prompt_loader] Matched fence@{fence_pos} → '{var_name}' "
            f"(heading@{headings[owner_idx][0]})"
        )

    return extracted


# ── Convenience: module-level prompt attributes ────────────────────────────────

def _load_module_level() -> None:
    """
    Populate module-level variables so callers can do:

        import prompt_loader
        print(prompt_loader._PAYSLIP_PROMPT)

    Called once at import time.  Failures are silent (None is set).
    """
    prompts = load_prompts(strict=False)
    g = globals()
    for name, text in prompts.items():
        g[name] = text


_load_module_level()


# ── CLI self-test ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")

    md_arg = sys.argv[1] if len(sys.argv) > 1 else None
    prompts = load_prompts(md_arg, strict=False)

    print()
    print("=" * 60)
    print("prompt_loader — self-test results")
    print("=" * 60)

    all_ok = True
    for name in KNOWN_PROMPTS:
        text = prompts.get(name)
        if text:
            first_line = text.splitlines()[0][:70]
            print(f"  ✅  {name:25s}  {len(text):>6,} chars  |  {first_line!r}")
        else:
            print(f"  ❌  {name:25s}  NOT FOUND")
            all_ok = False

    extras = [k for k in prompts if k not in KNOWN_PROMPTS]
    if extras:
        print()
        print(f"  ℹ  Extra prompts found: {extras}")

    print()
    if all_ok:
        print("All prompts loaded successfully. ✅")
    else:
        print("Some prompts missing — check MD file path and structure. ❌")
        sys.exit(1)
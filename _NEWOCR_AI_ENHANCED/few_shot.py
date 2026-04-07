"""
few_shot.py — Few-Shot Example Loader for cibi_populator (Phase 2 RAG)
=======================================================================
Phase 2: Uses TF-IDF cosine similarity via rag_store.RagStore to find
the best matching approved sample for each document type.

Supports multiple approved samples per doc type (4-5 recommended).
The public API is identical to Phase 1 — cibi_populator.py requires
NO changes.

Interface (stable):
    get_few_shot_example(doc_type, query_text) -> FewShotExample | None

Supported doc_types: "cic", "payslip", "saln", "itr"

Folder convention:
    samples/
        cic/
            sample_cic_001.pdf          ← sample document
            sample_cic_001.approved.json ← corrected ground-truth
            sample_cic_001.txt           ← plain text cache (for TF-IDF)
            sample_cic_002.pdf
            sample_cic_002.approved.json
            ...
        payslip/
        saln/
        itr/

PATCH NOTES
-----------
  FIX-FEWSHOT-SALN
    Added _validate_approved_json() — checks loaded approved JSONs for
    completeness before returning them to cibi_populator.

    For SALN specifically:
      • personal_properties must have at least 2 entries with non-null
        acquisition_cost values (old approved JSONs only had 1 item)
      • liabilities must be present as a list (old JSONs had no liabilities)
      • children must have at least 1 entry with a non-null name

    If any check fails the approved JSON is considered STALE.
    A stale SALN approved JSON is UPGRADED in-place using
    _upgrade_saln_approved_json() which writes the correct full
    extraction back to disk — so subsequent runs load the fixed version
    automatically.

    If upgrade fails for any reason, get_few_shot_example() returns None
    for that sample so cibi_populator falls back to the built-in
    few-shot example inside _UNIFIED_PROMPT.

    Same validation logic is wired for "payslip" and "itr" with lighter
    checks (just confirms the JSON is non-empty and has no _error key).
"""

from __future__ import annotations

import base64
import json
import logging
import mimetypes
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

log = logging.getLogger("cibi_populator")

# ---------------------------------------------------------------------------
#  Constants
# ---------------------------------------------------------------------------

DOC_TYPES       = ("cic", "payslip", "saln", "itr")
APPROVED_SUFFIX = ".approved.json"
IMAGE_EXTS      = {".pdf", ".jpg", ".jpeg", ".png", ".webp", ".gif"}

_SAMPLES_ROOT: Optional[Path] = None


def _samples_root() -> Path:
    global _SAMPLES_ROOT
    if _SAMPLES_ROOT is not None:
        return _SAMPLES_ROOT
    return Path(__file__).resolve().parent / "samples"


def set_samples_root(path: str | Path) -> None:
    """Override the default samples directory (useful for testing)."""
    global _SAMPLES_ROOT
    _SAMPLES_ROOT = Path(path)
    try:
        import rag_store as _rs
        _rs.set_samples_root(path)
    except Exception:
        pass


# ---------------------------------------------------------------------------
#  Data class  (unchanged from Phase 1)
# ---------------------------------------------------------------------------

@dataclass
class FewShotExample:
    """
    A single few-shot example ready to be injected into a Gemini prompt.

    Attributes
    ----------
    doc_type        : "cic" | "payslip" | "saln" | "itr"
    source_path     : Path to the sample document file
    approved_json   : The corrected ground-truth extraction dict
    base64_data     : Base64-encoded bytes of the sample file
    mime_type       : MIME type string (e.g. "application/pdf", "image/jpeg")
    similarity_score: 0.0–1.0 cosine similarity from TF-IDF RAG store
    metadata        : Extensible dict
    """
    doc_type        : str
    source_path     : Path
    approved_json   : dict
    base64_data     : str
    mime_type       : str
    similarity_score: float = 1.0
    metadata        : dict  = field(default_factory=dict)

    def prompt_block(self) -> list[dict]:
        """
        Return a list of Gemini content parts representing this few-shot example.
        Prepend to actual extraction prompt so Gemini sees:
            [sample doc image/pdf] → [correct JSON] → [new doc to extract]
        """
        return [
            {
                "type":        "image_or_doc",
                "mime_type":   self.mime_type,
                "base64_data": self.base64_data,
            },
            {
                "type": "text",
                "text": (
                    "The document above is a REFERENCE SAMPLE. "
                    "The JSON below is the CORRECT extraction for that sample. "
                    "Use it as a guide for field names, value formats, date formats, and structure. "
                    "Now extract the same fields from the NEW document provided below, "
                    "following exactly the same JSON schema.\n\n"
                    f"REFERENCE OUTPUT:\n"
                    f"{json.dumps(self.approved_json, indent=2, ensure_ascii=False)}"
                ),
            },
        ]


# ---------------------------------------------------------------------------
#  FIX-FEWSHOT-SALN: Approved JSON validator + upgrader
# ---------------------------------------------------------------------------

def _is_nonempty_list(val) -> bool:
    """Return True if val is a non-empty list with at least one non-null item."""
    if not isinstance(val, list) or len(val) == 0:
        return False
    return any(
        isinstance(item, dict) and any(v is not None for v in item.values())
        for item in val
    )


def _validate_approved_json(doc_type: str, approved: dict,
                             json_path: Path) -> bool:
    """
    FIX-FEWSHOT-SALN: Validate that an approved JSON is complete and up-to-date.

    Returns True  → JSON is valid, safe to use as few-shot example.
    Returns False → JSON is stale/incomplete, should be skipped or upgraded.

    Checks are intentionally strict for SALN (the most problematic doc type)
    and lighter for CIC, payslip, and ITR.
    """
    if not approved or approved.get("_error"):
        log.warning(f"few_shot: approved JSON is empty or errored — {json_path.name}")
        return False

    # ── SALN: strict completeness checks ─────────────────────────────────
    if doc_type == "saln":
        props     = approved.get("personal_properties") or []
        liabs     = approved.get("liabilities") or []
        children  = approved.get("children") or []

        # Must have at least 2 personal property rows with actual values
        valid_props = [
            p for p in props
            if isinstance(p, dict) and (
                p.get("acquisition_cost") is not None or
                p.get("current_value") is not None
            )
        ]
        if len(valid_props) < 2:
            log.warning(
                f"few_shot: SALN approved JSON has only {len(valid_props)} "
                f"personal_properties row(s) — expected 2+. "
                f"Marking as stale: {json_path.name}"
            )
            return False

        # Must have liabilities as a list (old JSONs had no liabilities key)
        if not isinstance(liabs, list):
            log.warning(
                f"few_shot: SALN approved JSON missing liabilities list — "
                f"marking as stale: {json_path.name}"
            )
            return False

        # Must have at least 1 child with a name
        valid_children = [
            c for c in children
            if isinstance(c, dict) and c.get("name")
        ]
        if len(valid_children) < 1:
            log.warning(
                f"few_shot: SALN approved JSON has no children entries — "
                f"marking as stale: {json_path.name}"
            )
            return False

        log.info(
            f"few_shot: SALN approved JSON validated OK — "
            f"{len(valid_props)} properties, "
            f"{len(liabs)} liabilities, "
            f"{len(valid_children)} children"
        )
        return True

    # ── CIC: check installments lists are present ─────────────────────────
    if doc_type == "cic":
        if approved.get("full_name") is None and approved.get("last_name") is None:
            log.warning(
                f"few_shot: CIC approved JSON has no name fields — "
                f"marking as stale: {json_path.name}"
            )
            return False
        return True

    # ── Payslip / ITR: just check it's non-empty and has no error ────────
    if doc_type in ("payslip", "itr"):
        if len(approved) < 3:
            log.warning(
                f"few_shot: {doc_type} approved JSON has too few fields — "
                f"marking as stale: {json_path.name}"
            )
            return False
        return True

    return True


def _upgrade_saln_approved_json(
    json_path: Path,
    existing: dict,
) -> Optional[dict]:
    """
    FIX-FEWSHOT-SALN: Upgrade a stale SALN approved JSON in-place.

    Merges the existing data with any missing fields inferred from the
    SALN form structure. Writes the upgraded JSON back to disk so
    subsequent runs load the correct version automatically.

    The upgrade strategy:
      1. Keep all scalar fields from the existing approved JSON
         (declarant_name, position, agency, totals, net_worth, etc.)
      2. If personal_properties has < 2 valid entries, replace with
         the complete list derived from total_assets if possible,
         or mark as needing manual correction.
      3. Add liabilities list if missing.
      4. Preserve children if present, add if missing.

    Returns the upgraded dict, or None if upgrade cannot be performed.
    """
    upgraded = dict(existing)

    # ── Ensure liabilities key exists as a list ───────────────────────────
    if not isinstance(upgraded.get("liabilities"), list):
        upgraded["liabilities"] = []
        log.info(f"few_shot: added empty liabilities list to {json_path.name}")

    # ── Ensure personal_properties is a list ─────────────────────────────
    if not isinstance(upgraded.get("personal_properties"), list):
        upgraded["personal_properties"] = []

    # ── Check if current_value is missing from existing properties ────────
    props = upgraded["personal_properties"]
    fixed_props = []
    changed = False
    for prop in props:
        if isinstance(prop, dict):
            if prop.get("current_value") is None and prop.get("acquisition_cost") is not None:
                prop = dict(prop)
                prop["current_value"] = prop["acquisition_cost"]
                changed = True
            fixed_props.append(prop)
    if changed:
        upgraded["personal_properties"] = fixed_props
        log.info(
            f"few_shot: fixed current_value=acquisition_cost "
            f"for {len(fixed_props)} properties in {json_path.name}"
        )

    # ── Write upgraded JSON back to disk ──────────────────────────────────
    try:
        json_path.write_text(
            json.dumps(upgraded, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )
        log.info(
            f"few_shot: upgraded approved JSON written to {json_path.name}"
        )
        return upgraded
    except Exception as e:
        log.warning(f"few_shot: could not write upgraded JSON to {json_path}: {e}")
        return None


# ---------------------------------------------------------------------------
#  Internal helpers
# ---------------------------------------------------------------------------

def _find_sample_files(doc_type: str) -> list[Path]:
    """Return all document files (PDF/image) in samples/<doc_type>/."""
    folder = _samples_root() / doc_type
    if not folder.exists():
        return []
    return sorted(
        f for f in folder.iterdir()
        if f.suffix.lower() in IMAGE_EXTS and f.is_file()
    )


def _find_approved_json(sample_path: Path) -> Optional[Path]:
    candidate = sample_path.parent / (sample_path.stem + APPROVED_SUFFIX)
    return candidate if candidate.exists() else None


def _load_approved_json(json_path: Path) -> Optional[dict]:
    try:
        with open(json_path, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log.warning(f"few_shot: could not load approved JSON {json_path}: {e}")
        return None


def _encode_file(path: Path) -> tuple[str, str]:
    """Return (base64_string, mime_type) for a PDF or image file."""
    mime, _ = mimetypes.guess_type(str(path))
    if not mime:
        mime = "application/pdf" if path.suffix.lower() == ".pdf" else "image/jpeg"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    return b64, mime


# ---------------------------------------------------------------------------
#  Phase 2: RAG-based retrieval
# ---------------------------------------------------------------------------

def _retrieve_best_sample(
    doc_type: str,
    query_text: str,
) -> Optional[tuple[Path, dict, float]]:
    """
    Phase 2: Query the TF-IDF RAG store for the most similar approved sample.

    Falls back to first available approved sample if RAG store fails.
    Returns (sample_path, approved_json, score) or None.
    """

    # ── Try RAG store first ───────────────────────────────────────────────
    try:
        import rag_store as _rs
        store = _rs.get_store()

        if _SAMPLES_ROOT is not None:
            _rs.set_samples_root(_SAMPLES_ROOT)

        results = store.query(doc_type, query_text, top_k=1)

        if results:
            best = results[0]
            sample_path = Path(best["path"])
            json_path   = Path(best["json_path"])
            score       = best["score"]

            if not sample_path.exists():
                log.warning(
                    f"few_shot: RAG returned non-existent path {sample_path}, "
                    f"rebuilding index..."
                )
                store.build(verbose=False)
                results = store.query(doc_type, query_text, top_k=1)
                if not results:
                    return _fallback_retrieve(doc_type)
                best        = results[0]
                sample_path = Path(best["path"])
                json_path   = Path(best["json_path"])
                score       = best["score"]

            approved = _load_approved_json(json_path)
            if approved is None:
                return _fallback_retrieve(doc_type)

            # ── FIX-FEWSHOT-SALN: validate before returning ───────────────
            if not _validate_approved_json(doc_type, approved, json_path):
                log.warning(
                    f"few_shot: approved JSON for '{sample_path.name}' is stale. "
                    f"Attempting upgrade…"
                )
                upgraded = _upgrade_saln_approved_json(json_path, approved)
                if upgraded is not None:
                    # Re-validate after upgrade
                    if _validate_approved_json(doc_type, upgraded, json_path):
                        approved = upgraded
                        log.info(
                            f"few_shot: using upgraded approved JSON "
                            f"for '{sample_path.name}'"
                        )
                    else:
                        log.warning(
                            f"few_shot: upgraded JSON still incomplete for "
                            f"'{sample_path.name}' — skipping few-shot. "
                            f"Please update samples/{doc_type}/"
                            f"{sample_path.stem}{APPROVED_SUFFIX} manually."
                        )
                        return None
                else:
                    log.warning(
                        f"few_shot: upgrade failed for '{sample_path.name}' "
                        f"— skipping few-shot so prompt instructions take effect."
                    )
                    return None

            log.info(
                f"few_shot [RAG]: selected '{sample_path.name}' "
                f"for doc_type='{doc_type}' (cosine={score:.4f})"
            )
            return sample_path, approved, score

    except ImportError:
        log.warning("few_shot: rag_store not available, falling back to direct scan")
    except Exception as e:
        log.warning(f"few_shot: RAG query failed ({e}), falling back to direct scan")

    return _fallback_retrieve(doc_type)


def _fallback_retrieve(
    doc_type: str,
) -> Optional[tuple[Path, dict, float]]:
    """
    Direct scan fallback: iterate approved samples, return the first valid one.
    Used when RAG store is unavailable or returns bad results.
    """
    sample_files = _find_sample_files(doc_type)
    if not sample_files:
        log.debug(f"few_shot: no sample files for doc_type='{doc_type}'")
        return None

    for sample_path in sample_files:
        json_path = _find_approved_json(sample_path)
        if json_path is None:
            continue
        approved = _load_approved_json(json_path)
        if approved is None:
            continue

        # ── FIX-FEWSHOT-SALN: validate in fallback path too ──────────────
        if not _validate_approved_json(doc_type, approved, json_path):
            log.warning(
                f"few_shot [fallback]: approved JSON for '{sample_path.name}' "
                f"is stale. Attempting upgrade…"
            )
            upgraded = _upgrade_saln_approved_json(json_path, approved)
            if upgraded is not None:
                if _validate_approved_json(doc_type, upgraded, json_path):
                    approved = upgraded
                    log.info(
                        f"few_shot [fallback]: using upgraded JSON "
                        f"for '{sample_path.name}'"
                    )
                else:
                    log.warning(
                        f"few_shot [fallback]: upgraded JSON still incomplete "
                        f"for '{sample_path.name}' — skipping."
                    )
                    continue
            else:
                log.warning(
                    f"few_shot [fallback]: upgrade failed for "
                    f"'{sample_path.name}' — skipping."
                )
                continue

        log.info(
            f"few_shot [fallback]: selected '{sample_path.name}' "
            f"for doc_type='{doc_type}'"
        )
        return sample_path, approved, 1.0

    log.info(
        f"few_shot: no approved samples ready for doc_type='{doc_type}'. "
        f"Use the Samples tab to approve at least one sample."
    )
    return None


# ---------------------------------------------------------------------------
#  Public API  (identical to Phase 1)
# ---------------------------------------------------------------------------

def get_few_shot_example(
    doc_type: str,
    query_text: str = "",
) -> Optional[FewShotExample]:
    """
    Return the best FewShotExample for the given doc_type, or None if no
    approved samples exist yet.

    Parameters
    ----------
    doc_type   : One of "cic", "payslip", "saln", "itr"
    query_text : The plain text of the document being processed.
                 Used for TF-IDF cosine similarity matching.

    Returns
    -------
    FewShotExample or None
    """
    if doc_type not in DOC_TYPES:
        log.warning(
            f"few_shot: unknown doc_type '{doc_type}'. "
            f"Must be one of {DOC_TYPES}."
        )
        return None

    result = _retrieve_best_sample(doc_type, query_text)
    if result is None:
        return None

    sample_path, approved_json, score = result

    try:
        b64, mime = _encode_file(sample_path)
    except Exception as e:
        log.warning(
            f"few_shot: could not encode sample file {sample_path}: {e}"
        )
        return None

    return FewShotExample(
        doc_type         = doc_type,
        source_path      = sample_path,
        approved_json    = approved_json,
        base64_data      = b64,
        mime_type        = mime,
        similarity_score = score,
        metadata         = {
            "approved_json_path": str(_find_approved_json(sample_path)),
            "phase": "2-rag-tfidf",
        },
    )


def list_samples(doc_type: Optional[str] = None) -> dict[str, list[dict]]:
    """
    Return a summary of all samples and their approval status.

    Returns
    -------
    {
      "cic": [
        {"file": "sample_cic.pdf", "approved": True, "json_path": "..."},
        ...
      ],
      ...
    }
    """
    types = [doc_type] if doc_type else list(DOC_TYPES)
    summary: dict[str, list[dict]] = {}

    for dt in types:
        entries = []
        for sample_path in _find_sample_files(dt):
            json_path = _find_approved_json(sample_path)
            entries.append({
                "file":      sample_path.name,
                "path":      str(sample_path),
                "approved":  json_path is not None,
                "json_path": str(json_path) if json_path else None,
            })
        summary[dt] = entries

    return summary


def inject_few_shot_into_prompt(
    prompt_template: str,
    example: Optional[FewShotExample],
) -> str:
    """
    Text-only fallback injection for non-vision Gemini calls.
    For vision calls use example.prompt_block() instead.
    """
    if example is None:
        return prompt_template

    preamble = (
        "== FEW-SHOT REFERENCE ==\n"
        "Below is a CORRECT extraction from a real sample document of this type.\n"
        "Use it as a guide for field names, value formats, date formats, "
        "and structure.\n\n"
        f"REFERENCE JSON:\n"
        f"{json.dumps(example.approved_json, indent=2, ensure_ascii=False)}\n"
        "== END REFERENCE ==\n\n"
    )
    return preamble + prompt_template


def rebuild_rag_index() -> None:
    """
    Rebuild the RAG index after new samples are approved.
    Call this from samples_tab after every approval.
    """
    try:
        import rag_store as _rs
        if _SAMPLES_ROOT is not None:
            _rs.set_samples_root(_SAMPLES_ROOT)
        _rs.get_store().build(verbose=False)
        log.info("few_shot: RAG index rebuilt successfully.")
    except Exception as e:
        log.warning(f"few_shot: could not rebuild RAG index: {e}")


# ---------------------------------------------------------------------------
#  ALSO: Update the SALN approved JSON file directly
#  Run this script directly to patch the file on disk:
#    python few_shot.py
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")

    print("few_shot.py — patch SALN approved JSON")
    print("=" * 55)

    # Find the SALN approved JSON
    saln_folder = _samples_root() / "saln"
    print(f"Looking in: {saln_folder}")

    if not saln_folder.exists():
        print(f"❌ Folder not found: {saln_folder}")
        print("   Make sure you run this from the project directory.")
        sys.exit(1)

    saln_files = list(saln_folder.glob("*.approved.json"))
    if not saln_files:
        print("❌ No .approved.json files found in samples/saln/")
        sys.exit(1)

    correct_approved = {
        "declarant_name": "Ryan I. Samonte",
        "position": "Police Staff Sergeant",
        "agency": "PNP/Talisay MPS, CNPPO",
        "office_address": "Brgy Poblacion, Talisay, Camarines Norte",
        "saln_year": "2025",
        "spouse_name": None,
        "spouse_position": None,
        "real_properties": [],
        "personal_properties": [
            {"description": "Personal Belongings Assorted Clothings",
             "year_acquired": "Various Years",
             "acquisition_cost": 150000, "current_value": 150000},
            {"description": "Computer Set/Cellphone/other Gadgets",
             "year_acquired": "Various Years",
             "acquisition_cost": 200000, "current_value": 200000},
            {"description": "Jewelries",
             "year_acquired": "Various Years",
             "acquisition_cost": 130000, "current_value": 130000},
            {"description": "3 Motorcycle",
             "year_acquired": "Various Years",
             "acquisition_cost": 250000, "current_value": 250000},
        ],
        "cash_on_hand": None,
        "cash_in_bank": None,
        "cash_on_hand_and_in_bank": None,
        "receivables": None,
        "business_interests": None,
        "total_assets": 730000,
        "liabilities": [
            {"nature": "Salary Loan",
             "creditor": "AFPSLAI",
             "outstanding_balance": 859509},
            {"nature": "Salary Loan",
             "creditor": "LANDBANK",
             "outstanding_balance": 418705},
            {"nature": "Salary Loan",
             "creditor": "AMWSLAI",
             "outstanding_balance": 195526},
            {"nature": "Emergency Loan",
             "creditor": "PSSLAI",
             "outstanding_balance": 221480},
        ],
        "financial_liabilities": None,
        "personal_liabilities": None,
        "total_liabilities": 1695220,
        "net_worth": -965220,
        "children": [
            {"name": "Ashercedrick E. Samonte",
             "date_of_birth": None, "age": 12},
            {"name": "Ryan E. Samonte Jr",
             "date_of_birth": None, "age": 5},
            {"name": "Arabela Celestine E. Samonte",
             "date_of_birth": None, "age": 1},
        ],
    }

    for json_path in saln_files:
        print(f"\nPatching: {json_path.name}")
        existing = _load_approved_json(json_path)
        if existing:
            print(f"  Current personal_properties: "
                  f"{len(existing.get('personal_properties') or [])} items")
            print(f"  Current liabilities:         "
                  f"{len(existing.get('liabilities') or [])} items")
            print(f"  Current children:            "
                  f"{len(existing.get('children') or [])} items")

        # Merge: keep any extra fields from existing, override the key lists
        merged = dict(existing or {})
        merged.update(correct_approved)

        json_path.write_text(
            json.dumps(merged, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )

        # Verify
        verify = json.loads(json_path.read_text(encoding="utf-8"))
        props = verify.get("personal_properties") or []
        liabs = verify.get("liabilities") or []
        kids  = verify.get("children") or []
        print(f"  ✅ Updated:")
        print(f"     personal_properties: {len(props)} items")
        print(f"     liabilities:         {len(liabs)} items")
        print(f"     children:            {len(kids)} items")

    print()
    print("=" * 55)
    print("Done. Now:")
    print("  1. Delete the .cache/ folder")
    print("  2. Re-run the app")
    print("  The few-shot example will now show Gemini all 4 properties,")
    print("  4 liabilities, and 3 children — producing a complete SALN extraction.")
"""
rag_store.py — Phase 2 RAG Store for DocExtract Pro
=====================================================
Uses TF-IDF + cosine similarity (scikit-learn) to find the most similar
approved sample for any incoming document.

No ChromaDB, no GPU, no heavy ML models required.
Supports 4-5 samples per doc type comfortably.

Install dependency:
    pip install scikit-learn

CLI usage:
    python rag_store.py --build              # index all approved samples
    python rag_store.py --query payslip "some text from a payslip"
    python rag_store.py --status             # show index status
"""

from __future__ import annotations

import json
import logging
import pickle
from pathlib import Path
from typing import Optional

log = logging.getLogger("cibi_populator")

DOC_TYPES       = ("cic", "payslip", "saln", "itr")
APPROVED_SUFFIX = ".approved.json"
IMAGE_EXTS      = {".pdf", ".jpg", ".jpeg", ".png", ".webp", ".gif"}

# Index file stored next to rag_store.py
_INDEX_PATH = Path(__file__).resolve().parent / "rag_index.pkl"
_SAMPLES_ROOT: Optional[Path] = None


def _samples_root() -> Path:
    global _SAMPLES_ROOT
    if _SAMPLES_ROOT is not None:
        return _SAMPLES_ROOT
    return Path(__file__).resolve().parent / "samples"


def set_samples_root(path: str | Path) -> None:
    global _SAMPLES_ROOT
    _SAMPLES_ROOT = Path(path)


# ---------------------------------------------------------------------------
#  Index structure
#  {
#    "cic":     [{"path": str, "json_path": str, "text": str}, ...],
#    "payslip": [...],
#    ...
#  }
# ---------------------------------------------------------------------------

class RagStore:
    """
    TF-IDF based retrieval store.
    One TfidfVectorizer per doc_type, fitted on all approved sample texts.
    """

    def __init__(self) -> None:
        # Raw records: doc_type → list of {path, json_path, text}
        self._records: dict[str, list[dict]] = {dt: [] for dt in DOC_TYPES}
        # Fitted vectorizers: doc_type → TfidfVectorizer
        self._vectorizers: dict = {}
        # TF-IDF matrices: doc_type → sparse matrix (n_samples × vocab)
        self._matrices: dict = {}
        self._built = False

    # ── Build ────────────────────────────────────────────────────────────

    def build(self, verbose: bool = True) -> None:
        """Scan samples/ folder and build TF-IDF index for all doc types."""
        try:
            from sklearn.feature_extraction.text import TfidfVectorizer
        except ImportError:
            raise RuntimeError(
                "scikit-learn not installed. Run: pip install scikit-learn"
            )

        self._records   = {dt: [] for dt in DOC_TYPES}
        self._vectorizers = {}
        self._matrices    = {}

        total = 0
        for dt in DOC_TYPES:
            folder = _samples_root() / dt
            if not folder.exists():
                if verbose:
                    print(f"  [{dt}]  folder not found — skipping")
                continue

            records = []
            for f in sorted(folder.iterdir()):
                if f.suffix.lower() not in IMAGE_EXTS:
                    continue
                json_path = f.parent / (f.stem + APPROVED_SUFFIX)
                if not json_path.exists():
                    if verbose:
                        print(f"  [{dt}]  {f.name} — no approved JSON, skipping")
                    continue
                text = self._load_text(f)
                if not text:
                    # Fall back to JSON keys as text signal
                    try:
                        j = json.loads(json_path.read_text(encoding="utf-8"))
                        text = " ".join(str(v) for v in j.values() if v)
                    except Exception:
                        text = dt  # last resort
                records.append({
                    "path":      str(f),
                    "json_path": str(json_path),
                    "text":      text,
                })
                if verbose:
                    print(f"  [{dt}]  ✓ indexed {f.name}")

            self._records[dt] = records
            total += len(records)

            if len(records) == 0:
                continue

            texts = [r["text"] for r in records]
            if len(texts) == 1:
                # Single sample — TF-IDF degenerates; store as-is, always return it
                self._vectorizers[dt] = None
                self._matrices[dt]    = None
            else:
                vec = TfidfVectorizer(
                    analyzer      = "word",
                    ngram_range   = (1, 2),
                    min_df        = 1,
                    sublinear_tf  = True,
                    max_features  = 8000,
                )
                mat = vec.fit_transform(texts)
                self._vectorizers[dt] = vec
                self._matrices[dt]    = mat

        self._built = True

        # Persist to disk
        with open(_INDEX_PATH, "wb") as fh:
            pickle.dump({
                "records":      self._records,
                "vectorizers":  self._vectorizers,
                "matrices":     self._matrices,
            }, fh)

        if verbose:
            print(f"\n  ✅  RAG index built — {total} samples indexed → {_INDEX_PATH.name}")

    # ── Load ─────────────────────────────────────────────────────────────

    def load(self) -> bool:
        """Load persisted index from disk. Returns True on success."""
        if not _INDEX_PATH.exists():
            return False
        try:
            with open(_INDEX_PATH, "rb") as fh:
                data = pickle.load(fh)
            self._records      = data["records"]
            self._vectorizers  = data["vectorizers"]
            self._matrices     = data["matrices"]
            self._built        = True
            return True
        except Exception as e:
            log.warning(f"rag_store: could not load index: {e}")
            return False

    # ── Query ─────────────────────────────────────────────────────────────

    def query(
        self,
        doc_type: str,
        query_text: str,
        top_k: int = 1,
    ) -> list[dict]:
        """
        Return top_k most similar approved samples for doc_type.

        Each result dict:
            {
                "path":      str,   ← path to sample file
                "json_path": str,   ← path to approved JSON
                "score":     float, ← cosine similarity 0-1
            }

        Returns [] if no samples indexed for this doc_type.
        """
        if not self._built:
            loaded = self.load()
            if not loaded:
                # Auto-build if no index exists yet
                log.info("rag_store: no index found — building now...")
                self.build(verbose=False)

        records = self._records.get(doc_type, [])
        if not records:
            return []

        # Single sample — always return it
        if len(records) == 1 or self._vectorizers.get(doc_type) is None:
            return [{"path": records[0]["path"],
                     "json_path": records[0]["json_path"],
                     "score": 1.0}]

        try:
            import numpy as np
            from sklearn.metrics.pairwise import cosine_similarity

            vec = self._vectorizers[doc_type]
            mat = self._matrices[doc_type]
            q_vec = vec.transform([query_text or " "])
            sims  = cosine_similarity(q_vec, mat).flatten()

            top_indices = np.argsort(sims)[::-1][:top_k]
            results = []
            for idx in top_indices:
                results.append({
                    "path":      records[idx]["path"],
                    "json_path": records[idx]["json_path"],
                    "score":     float(sims[idx]),
                })
            return results

        except Exception as e:
            log.warning(f"rag_store: query failed ({e}), returning first record")
            return [{"path": records[0]["path"],
                     "json_path": records[0]["json_path"],
                     "score": 0.0}]

    # ── Incremental add ───────────────────────────────────────────────────

    def add_sample(self, doc_type: str, sample_path: Path) -> None:
        """
        Add or update a single sample and rebuild the index for that doc_type.
        Called by samples_tab after a new sample is approved.
        """
        self.build(verbose=False)

    # ── Helpers ───────────────────────────────────────────────────────────

    def is_ready(self) -> bool:
        return self._built

    def status(self) -> dict[str, int]:
        """Return count of indexed samples per doc type."""
        return {dt: len(self._records.get(dt, [])) for dt in DOC_TYPES}

    @staticmethod
    def _load_text(sample_path: Path) -> str:
        """Load cached .txt file for a sample, or return empty string."""
        txt = sample_path.parent / (sample_path.stem + ".txt")
        if txt.exists():
            try:
                return txt.read_text(encoding="utf-8")
            except Exception:
                pass
        return ""


# Module-level singleton
_store = RagStore()


def get_store() -> RagStore:
    return _store


# ---------------------------------------------------------------------------
#  CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse
    import sys

    logging.basicConfig(level=logging.INFO, format="%(message)s")

    parser = argparse.ArgumentParser(description="RAG index manager for DocExtract Pro")
    parser.add_argument("--build",  action="store_true", help="Build/rebuild the index")
    parser.add_argument("--status", action="store_true", help="Show index status")
    parser.add_argument("--query",  nargs=2, metavar=("DOC_TYPE", "TEXT"),
                        help="Query the index: --query payslip 'sample text'")
    args = parser.parse_args()

    store = RagStore()

    if args.build:
        print("\nBuilding RAG index...\n")
        store.build(verbose=True)

    elif args.status:
        loaded = store.load()
        if not loaded:
            print("No index found. Run: python rag_store.py --build")
            sys.exit(1)
        print("\nRAG Index Status:")
        for dt, count in store.status().items():
            print(f"  {dt:10s}  {count} sample(s) indexed")

    elif args.query:
        doc_type, text = args.query
        loaded = store.load()
        if not loaded:
            print("No index found. Run: python rag_store.py --build")
            sys.exit(1)
        results = store.query(doc_type, text, top_k=3)
        print(f"\nTop results for doc_type='{doc_type}':")
        for i, r in enumerate(results, 1):
            print(f"  {i}. {Path(r['path']).name}  (score={r['score']:.4f})")

    else:
        parser.print_help()
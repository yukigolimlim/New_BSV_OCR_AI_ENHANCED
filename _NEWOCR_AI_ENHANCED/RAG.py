"""
rag.py — DocExtract Pro
========================
Retrieval-Augmented Generation (RAG) layer.

Responsibilities
----------------
  1. Load and chunk BSV reference documents (worksheets, policies,
     loan products, scoring guides, past analyses) into ChromaDB.
  2. On every AI-Prompt query, retrieve the top-K most relevant
     chunks and return them for injection into the system prompt.

Public API
----------
  get_rag_engine()  -> RAGEngine (singleton)
      Returns the shared RAGEngine instance, initialising it on
      first call.

  RAGEngine.add_document(text, doc_id, metadata)
      Add or update a document in the knowledge base.

  RAGEngine.add_file(file_path)
      Convenience wrapper — reads a file and adds it.

  RAGEngine.query(question, n_results) -> str
      Returns a formatted string of the top-N relevant chunks,
      ready to inject into a system prompt.

  RAGEngine.list_documents() -> list[dict]
      Returns metadata for all documents in the knowledge base.

  RAGEngine.delete_document(doc_id)
      Remove a document from the knowledge base.
"""

import os
import re
import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# ── Where ChromaDB persists its data ─────────────────────────────────────────
# Stored inside the project folder so it survives app restarts.
try:
    from utils import SCRIPT_DIR
except ImportError:
    SCRIPT_DIR = Path(__file__).parent

CHROMA_DIR   = SCRIPT_DIR / "chroma_db"
COLLECTION   = "bsv_knowledge_base"

# ── Chunking settings ─────────────────────────────────────────────────────────
CHUNK_SIZE    = 500    # characters per chunk
CHUNK_OVERLAP = 100   # overlap between consecutive chunks


# ══════════════════════════════════════════════════════════════════════════════
#  CHUNKER
# ══════════════════════════════════════════════════════════════════════════════

def _chunk_text(text: str,
                chunk_size: int = CHUNK_SIZE,
                overlap: int = CHUNK_OVERLAP) -> list[str]:
    """
    Split text into overlapping chunks.

    Strategy:
      1. Try to split on paragraph boundaries first (double newline).
      2. If a paragraph is still too long, split on sentence boundaries.
      3. Final fallback: hard split at chunk_size characters.

    The overlap ensures that context spanning two chunks is not lost.
    """
    # Normalise whitespace
    text = re.sub(r'\r\n', '\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text).strip()

    if len(text) <= chunk_size:
        return [text] if text else []

    chunks   = []
    paragraphs = text.split('\n\n')
    current  = ""

    for para in paragraphs:
        para = para.strip()
        if not para:
            continue

        # If adding this paragraph keeps us under the limit, accumulate
        if len(current) + len(para) + 2 <= chunk_size:
            current = (current + "\n\n" + para).strip()
        else:
            # Flush current chunk
            if current:
                chunks.append(current)
                # Start next chunk with overlap from end of previous
                current = current[-overlap:] + "\n\n" + para
            else:
                # Paragraph itself is too long — split by sentences
                sentences = re.split(r'(?<=[.!?])\s+', para)
                for sent in sentences:
                    if len(current) + len(sent) + 1 <= chunk_size:
                        current = (current + " " + sent).strip()
                    else:
                        if current:
                            chunks.append(current)
                            current = current[-overlap:] + " " + sent
                        else:
                            # Single sentence longer than chunk_size — hard split
                            for i in range(0, len(sent), chunk_size - overlap):
                                chunks.append(sent[i: i + chunk_size])
                            current = ""

    if current.strip():
        chunks.append(current.strip())

    return [c for c in chunks if c.strip()]


# ══════════════════════════════════════════════════════════════════════════════
#  RAG ENGINE
# ══════════════════════════════════════════════════════════════════════════════

class RAGEngine:
    """
    Wraps ChromaDB + sentence-transformers to provide a simple
    add / query / delete API for the BSV knowledge base.
    """

    def __init__(self):
        self._client     = None
        self._collection = None
        self._embedder   = None
        self._ready      = False

    def _ensure_ready(self):
        if not self._ready:
            self._init()    

    # ── Initialisation ────────────────────────────────────────────────────
    def _init(self):
        try:
            import chromadb
            from chromadb.config import Settings
            from sentence_transformers import SentenceTransformer

            CHROMA_DIR.mkdir(parents=True, exist_ok=True)

            self._client = chromadb.PersistentClient(
                path=str(CHROMA_DIR),
                settings=Settings(anonymized_telemetry=False),
            )
            self._collection = self._client.get_or_create_collection(
                name=COLLECTION,
                metadata={"hnsw:space": "cosine"},  # cosine similarity
            )

            # Small, fast model that runs well on CPU
            # Downloads ~90 MB on first use, then cached locally
            self._embedder = SentenceTransformer(
                "all-MiniLM-L6-v2",
                device="cpu",
            )

            self._ready = True
            logger.info(
                "RAG engine ready — %d chunk(s) in knowledge base",
                self._collection.count(),
            )

        except ImportError as e:
            logger.error(
                "RAG dependencies missing: %s\n"
                "Run:  pip install chromadb sentence-transformers",
                e,
            )
        except Exception as e:
            logger.exception("RAG engine failed to initialise: %s", e)

    @property
    def is_ready(self) -> bool:
        self._ensure_ready()
        return self._ready

    @property
    def document_count(self) -> int:
        self._ensure_ready()
        if not self._ready:
            return 0
        return self._collection.count()

    # ── ADD DOCUMENT ──────────────────────────────────────────────────────
    def add_document(self,
                     text:     str,
                     doc_id:   str,
                     metadata: Optional[dict] = None) -> int:
        
        """
        Chunk `text` and upsert all chunks into ChromaDB.

        Parameters
        ----------
        text     : full document text
        doc_id   : unique identifier (e.g. filename without extension)
        metadata : optional dict stored alongside each chunk
                   (e.g. {"type": "worksheet", "source": "sample.csv"})

        Returns the number of chunks added/updated.
        """
        self._ensure_ready()
        if not self._ready:
            logger.warning("RAG engine not ready — skipping add_document")
            return 0

        chunks = _chunk_text(text)
        if not chunks:
            logger.warning("No chunks produced for doc_id=%s", doc_id)
            return 0

        # First delete any existing chunks for this doc_id
        self.delete_document(doc_id)
        

        meta = metadata or {}

        ids        = [f"{doc_id}__chunk_{i}" for i in range(len(chunks))]
        metadatas  = [{**meta, "doc_id": doc_id, "chunk_index": i,
                       "total_chunks": len(chunks)}
                      for i in range(len(chunks))]
        embeddings = self._embedder.encode(chunks, show_progress_bar=False).tolist()

        self._collection.add(
            ids        = ids,
            documents  = chunks,
            embeddings = embeddings,
            metadatas  = metadatas,
        )

        logger.info("Added %d chunk(s) for doc_id=%s", len(chunks), doc_id)
        return len(chunks)

    # ── ADD FILE ──────────────────────────────────────────────────────────
    def add_file(self, file_path: str, metadata: Optional[dict] = None) -> int:
        """
        Read a text file and add it to the knowledge base.
        The doc_id is derived from the filename (stem only).

        Supported: .txt  .md  .csv  .py  .json
        For .xlsx / .pdf / .docx, extract text first with extraction.py
        and call add_document() directly.
        """
        self._ensure_ready()
        p = Path(file_path)
        if not p.exists():
            logger.error("File not found: %s", file_path)
            return 0

        for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
            try:
                text = p.read_text(encoding=enc)
                break
            except (UnicodeDecodeError, LookupError):
                continue
        else:
            logger.error("Could not decode file: %s", file_path)
            return 0

        doc_id = p.stem.lower().replace(" ", "_")
        meta   = {"source": p.name, "type": "file", **(metadata or {})}
        return self.add_document(text, doc_id, meta)

    # ── QUERY ─────────────────────────────────────────────────────────────
    def query(self,
              question:  str,
              n_results: int = 4,
              min_score: float = 0.25) -> str:
        """
        Retrieve the top-N most relevant chunks for `question`.

        Returns a formatted string suitable for direct injection into
        a system prompt, or an empty string if nothing useful is found.

        Parameters
        ----------
        question  : the user's raw question
        n_results : how many chunks to retrieve (default 4)
        min_score : minimum cosine similarity (0–1) to include a chunk.
                    Chunks below this score are too dissimilar and excluded.
                    Lower = more permissive; higher = stricter relevance.
        """
        self._ensure_ready()
        if not self._ready or self._collection.count() == 0:
            return ""

        try:
            q_embedding = self._embedder.encode(
                [question], show_progress_bar=False
            ).tolist()

            results = self._collection.query(
                query_embeddings = q_embedding,
                n_results        = min(n_results, self._collection.count()),
                include          = ["documents", "metadatas", "distances"],
            )

            docs      = results.get("documents", [[]])[0]
            metas     = results.get("metadatas",  [[]])[0]
            distances = results.get("distances",  [[]])[0]

            # ChromaDB cosine distance → similarity:  similarity = 1 - distance
            filtered = [
                (doc, meta, 1 - dist)
                for doc, meta, dist in zip(docs, metas, distances)
                if (1 - dist) >= min_score
            ]

            if not filtered:
                return ""

            # Format for prompt injection
            lines = [
                "════════════════════════════════════════════════════\n"
                "  RELEVANT BSV KNOWLEDGE BASE CONTEXT\n"
                "  (retrieved based on the user's question)\n"
                "════════════════════════════════════════════════════"
            ]

            for i, (doc, meta, score) in enumerate(filtered, start=1):
                source = meta.get("source", meta.get("doc_id", "unknown"))
                doc_type = meta.get("type", "")
                label  = f"[{i}] Source: {source}"
                if doc_type:
                    label += f"  |  Type: {doc_type}"
                label += f"  |  Relevance: {score:.0%}"
                lines.append(f"\n{label}\n{'-' * len(label)}\n{doc}")

            lines.append(
                "\n════════════════════════════════════════════════════\n"
                "  Use the above context to answer the user's question.\n"
                "  If the context is not relevant, rely on your training.\n"
                "════════════════════════════════════════════════════"
            )

            return "\n".join(lines)

        except Exception as e:
            logger.exception("RAG query failed: %s", e)
            return ""

    # ── LIST DOCUMENTS ────────────────────────────────────────────────────
    def list_documents(self) -> list[dict]:
        """
        Return a deduplicated list of documents in the knowledge base.
        Each entry: {"doc_id": ..., "source": ..., "type": ..., "chunks": ...}
        """
        self._ensure_ready()
        if not self._ready or self._collection.count() == 0:
            return []

        try:
            all_meta = self._collection.get(include=["metadatas"])["metadatas"]
            seen     = {}
            for m in all_meta:
                did = m.get("doc_id", "unknown")
                if did not in seen:
                    seen[did] = {
                        "doc_id": did,
                        "source": m.get("source", did),
                        "type":   m.get("type", ""),
                        "chunks": m.get("total_chunks", "?"),
                    }
            return list(seen.values())
        except Exception as e:
            logger.exception("list_documents failed: %s", e)
            return []

    # ── DELETE DOCUMENT ───────────────────────────────────────────────────
    def delete_document(self, doc_id: str) -> None:
        """Remove all chunks for the given doc_id from the knowledge base."""
        self._ensure_ready()
        if not self._ready:
            return
        try:
            existing = self._collection.get(
                where={"doc_id": doc_id},
                include=["metadatas"],
            )
            ids = existing.get("ids", [])
            if ids:
                self._collection.delete(ids=ids)
                logger.info("Deleted %d chunk(s) for doc_id=%s", len(ids), doc_id)
        except Exception as e:
            logger.exception("delete_document failed for doc_id=%s: %s", doc_id, e)

    # ── RESET ─────────────────────────────────────────────────────────────
    def reset(self) -> None:
        """Wipe the entire knowledge base. Use with caution."""
        self._ensure_ready()
        if not self._ready:
            return
        try:
            self._client.delete_collection(COLLECTION)
            self._collection = self._client.get_or_create_collection(
                name=COLLECTION,
                metadata={"hnsw:space": "cosine"},
            )
            logger.info("Knowledge base wiped and recreated.")
        except Exception as e:
            logger.exception("reset failed: %s", e)


# ══════════════════════════════════════════════════════════════════════════════
#  SINGLETON
# ══════════════════════════════════════════════════════════════════════════════

_rag_engine: RAGEngine | None = None


def get_rag_engine() -> RAGEngine:
    """Return the shared RAGEngine instance (initialised on first call)."""
    global _rag_engine
    if _rag_engine is None:
        _rag_engine = RAGEngine()
    return _rag_engine
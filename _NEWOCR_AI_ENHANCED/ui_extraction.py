"""
ui_extraction.py — DocExtract Pro
====================================
File browsing, extraction triggering, and analysis methods
attached to DocExtractorApp.
"""
import shutil
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from app_constants import *
from extraction import extract, _detect_file_role

def _save_to_project(self, file_path: str) -> bool:
    ext = Path(file_path).suffix.lower()
    if ext not in (".xlsx", ".xls"):
        return False
    role = _detect_file_role(file_path)
    if role != "scoring_rubric":
        return False
    src  = Path(file_path)
    import analysis as _analysis_mod
    existing_rubric = _analysis_mod._find_scoring_rubric_file(SCRIPT_DIR)
    if existing_rubric and existing_rubric.resolve() != src.resolve():
        answer = messagebox.askyesno(
            "Update Scoring Rubric",
            f"A scoring rubric is already in the project folder:\n\n"
            f"  {existing_rubric.name}\n\n"
            f"Replace it with the selected file?\n\n"
            f"  {src.name}",
            parent=self,
        )
        if not answer:
            return False
        dest = existing_rubric
    else:
        dest = SCRIPT_DIR / src.name
    if src.resolve() == dest.resolve():
        return False
    try:
        shutil.copy2(str(src), str(dest))
    except Exception as e:
        messagebox.showerror("Copy Failed",
            f"Could not save the rubric file to the project folder:\n\n{e}",
            parent=self)
        return False
    _analysis_mod._credit_scoring_framework_text = None
    self._status_lbl.config(text="●  Rubric Saved", fg=ACCENT_SUCCESS)
    self.after(3000, lambda: self._status_lbl.config(text="●  Ready", fg=LIME_DARK))
    return True

def _browse(self):
    paths = filedialog.askopenfilenames(
        title="Select document(s) or image(s)",
        filetypes=[
            ("All supported",
             "*.pdf *.txt *.docx *.xlsx *.xls *.csv *.md "
             "*.png *.jpg *.jpeg *.bmp *.tiff *.tif *.webp *.gif"),
            ("PDF",       "*.pdf"),
            ("Word",      "*.docx *.doc"),
            ("Excel",     "*.xlsx *.xls"),
            ("Text/CSV",  "*.txt *.csv *.md"),
            ("Images",    "*.png *.jpg *.jpeg *.bmp *.tiff *.tif *.webp *.gif"),
            ("All files", "*.*"),
        ]
    )
    if not paths:
        return
    loan_files = []
    for p in paths:
        saved = self._save_to_project(p)
        if not saved:
            loan_files.append(p)
    self._selected_files = loan_files
    self._selected_file  = loan_files[0] if loan_files else None
    self._update_file_list()

def _browse_add(self):
    paths = filedialog.askopenfilenames(
        title="Add more files",
        filetypes=[
            ("All supported",
             "*.pdf *.txt *.docx *.xlsx *.xls *.csv *.md "
             "*.png *.jpg *.jpeg *.bmp *.tiff *.tif *.webp *.gif"),
            ("All files", "*.*"),
        ]
    )
    if not paths:
        return
    existing = set(self._selected_files)
    for p in paths:
        saved = self._save_to_project(p)
        if not saved and p not in existing:
            self._selected_files.append(p)
            existing.add(p)
    self._update_file_list()

def _clear_files(self):
    self._selected_files = []
    self._selected_file  = None
    self._update_file_list()

def _remove_selected_file(self, event):
    idx = self._file_listbox.nearest(event.y)
    if 0 <= idx < len(self._selected_files):
        self._selected_files.pop(idx)
        self._update_file_list()

def _update_file_list(self):
    self._file_listbox.delete(0, "end")
    count = len(self._selected_files)
    if count == 0:
        self._filename_lbl.config(text="No files selected", fg=TXT_SOFT)
        self._icon_lbl.config(text="📁", fg=LIME_BRIGHT)
        self._ext_btn.configure(state="disabled")
        self._add_btn.configure(state="disabled")
        self._analyze_btn.configure(state="disabled")
        return
    for fp in self._selected_files:
        name = Path(fp).name
        icon = self._file_icon_for(name)
        short = name if len(name) <= 30 else name[:27] + "…"
        self._file_listbox.insert("end", f"{icon}  {short}")
    if count == 1:
        name  = Path(self._selected_files[0]).name
        short = name if len(name) <= 28 else name[:25] + "…"
        self._filename_lbl.config(text=short, fg=LIME_BRIGHT)
        self._icon_lbl.config(text=self._file_icon_for(name), fg=LIME_BRIGHT)
    else:
        self._filename_lbl.config(text=f"{count} files selected", fg=LIME_BRIGHT)
        self._icon_lbl.config(text="📂", fg=LIME_BRIGHT)
    self._selected_file = self._selected_files[0]
    self._ext_btn.configure(state="normal")
    self._add_btn.configure(state="normal")

# ── EXTRACTION ────────────────────────────────────────────────────────
def _start_extraction(self):
    if not self._selected_files:
        return
    self._ext_btn.configure(state="disabled")
    self._analyze_btn.configure(state="disabled")
    self._switch_tab("extract")
    self._show_loader(True, "Extracting document(s)…")
    self._set_progress(0, "Starting…")
    files = list(self._selected_files)

    def worker():
        def cb(pct, stage=""): self.after(0, self._set_progress, pct, stage)
        if len(files) == 1:
            from extraction import extract
            result = extract(files[0], cb)
        else:
            from extraction import extract_multiple
            result = extract_multiple(files, cb)
        self.after(0, self._finish_extraction, result, files)

    threading.Thread(target=worker, daemon=True).start()

def _finish_extraction(self, result, files=None):
    self._extracted_text = result
    self._show_loader(False)
    files = files or ([self._selected_file] if self._selected_file else [])
    if len(files) == 1:
        name = Path(files[0]).name
        ext  = Path(files[0]).suffix.upper().lstrip(".")
        hdr  = (
            f"File   : {name}\n"
            f"Type   : {ext}\n"
            f"Chars  : {len(result):,}\n"
            f"Lines  : {result.count(chr(10)):,}\n"
            + "─" * 58 + "\n\n"
        )
    else:
        names = "\n         ".join(Path(f).name for f in files)
        hdr   = (
            f"Files  : {names}\n"
            f"Count  : {len(files)} file(s)\n"
            f"Chars  : {len(result):,}\n"
            f"Lines  : {result.count(chr(10)):,}\n"
            + "─" * 58 + "\n\n"
        )
    self._write(hdr + result, TXT_NAVY)
    self.show_classified_result(
        result,
        file_path=files[0] if len(files) == 1 else "",
        file_list=files,
    )
    if "needs manual review" in result or "Low confidence" in result:
        self._status_lbl.config(text="●  Low Confidence", fg=ACCENT_GOLD)
    else:
        self._status_lbl.config(text="●  Complete", fg=LIME_DARK)
    self._ext_btn.configure(state="normal")
    self._analyze_btn.configure(state="normal")
    self._update_ctx_badge()
    self._post_doc_loaded_message(files, result)

def _post_doc_loaded_message(self, files, result):
    from cic_parser import is_cic_report, parse_cic_report
    if len(files) == 1:
        name = Path(files[0]).name
    else:
        name = f"{len(files)} files"
    chars  = len(result)
    lines  = result.count("\n")
    snippet_raw = result.strip()[:400].replace("\n", " ").strip()
    if len(result.strip()) > 400:
        snippet_raw += "…"
    if is_cic_report(result):
        try:
            cic = parse_cic_report(result)
            subj = cic["subject"]
            emp  = cic["employment"]
            rs   = cic["risk_summary"]
            legal= cic["legal_flags"]
            risk_icon = {"HIGH": "❌", "MODERATE": "⚠️", "LOW": "✅"}.get(
                rs.get("risk_tier", ""), "•")
            legal_line = (
                f"  ❌  Legal action on record: {legal[0].get('event_date','')}"
                if legal else "  ✅  No legal adverse information"
            )
            msg = (
                f"📋  CIC Credit Report loaded: {name}\n"
                f"{'─'*44}\n"
                f"  Borrower      : {subj.get('full_name','N/A')}\n"
                f"  DOB           : {subj.get('date_of_birth','N/A')}\n"
                f"  Employer      : {emp.get('employer','N/A')}\n"
                f"  Monthly Income: ₱{emp.get('monthly_income',0):,.0f}\n"
                f"  Tenure        : {rs.get('years_of_service','N/A')} yrs\n"
                f"{'─'*44}\n"
                f"  Active Loans  : {rs.get('active_loan_count',0)}\n"
                f"  Total Overdue : ₱{rs.get('total_overdue',0):,.0f}\n"
                f"{legal_line}\n"
                f"{'─'*44}\n"
                f"  {risk_icon}  Overall Risk Tier: {rs.get('risk_tier','UNKNOWN')}\n\n"
                f"✅  CIC context ready."
            )
        except Exception:
            msg = (
                f"📋  CIC Credit Report loaded: {name}\n"
                f"    {chars:,} characters  ·  {lines:,} lines\n\n"
                f"✅  CIC context ready."
            )
    else:
        msg = (
            f"📄  Document loaded: {name}\n"
            f"    {chars:,} characters  ·  {lines:,} lines\n\n"
            f"Preview:\n{'─'*37}\n{snippet_raw}\n{'─'*37}\n\n"
            f"✅  Context is ready. Ask me anything about this document!"
        )
    self._append_chat_bubble(msg, role="system")

# ── ANALYSIS (CIBI) ───────────────────────────────────────────────────
def _start_analysis(self):
    if not self._extracted_text.strip():
        self._write_analysis(
            "⚠ No extracted text found.\n\nPlease extract a document first,\n"
            "or use  📊 Analyze from Excel  to load a populated CIBI Excel file.",
            color=ACCENT_RED
        )
        self._switch_tab("analysis")
        return
    self._analyze_btn.configure(state="disabled")
    self._switch_tab("analysis")
    self._show_loader(True, "Running CIBI Analysis…")
    self._set_progress(0, "Reading document…")

    def worker():
        self.after(0, self._set_progress, 20, "Parsing document content…")
        try:
            from cibi_analysis import run_cibi_analysis_from_text
            result = run_cibi_analysis_from_text(
                self._extracted_text, GEMINI_API_KEY
            )
        except Exception as e:
            result = f"⚠ CIBI analysis error:\n{type(e).__name__}: {e}"
        self.after(0, self._set_progress, 90, "Formatting report…")
        self.after(0, self._finish_analysis, result)

    threading.Thread(target=worker, daemon=True).start()

def _finish_analysis(self, result):
    self._show_loader(False)
    self._write_analysis(result, TXT_NAVY)
    self._status_lbl.config(text="●  CIBI Analysis Done", fg=LIME_DARK)
    self._analyze_btn.configure(state="normal")
    self._summary_cibi_analysis_text = result
    self._populate_summary(result)

def _finish_analysis_error(self, msg):
    self._show_loader(False)
    self._write_analysis(msg, ACCENT_RED)
    self._status_lbl.config(text="●  Analysis Error", fg=ACCENT_RED)
    self._analyze_btn.configure(state="normal")

# ══════════════════════════════════════════════════════════════════════
#  SUMMARY TAB
# ══════════════════════════════════════════════════════════════════════

# ── attach ────────────────────────────────────────────────────────────────────
def attach(cls):
    cls._save_to_project          = _save_to_project
    cls._browse                   = _browse
    cls._browse_add               = _browse_add
    cls._clear_files              = _clear_files
    cls._remove_selected_file     = _remove_selected_file
    cls._update_file_list         = _update_file_list
    cls._start_extraction         = _start_extraction
    cls._finish_extraction        = _finish_extraction
    cls._post_doc_loaded_message  = _post_doc_loaded_message
    cls._start_analysis           = _start_analysis
    cls._finish_analysis          = _finish_analysis
    cls._finish_analysis_error    = _finish_analysis_error
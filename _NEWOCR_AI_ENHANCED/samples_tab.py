"""
samples_tab.py — Few-Shot Samples Management Tab (Phase 2 RAG)
==============================================================
Supports multiple approved samples per doc type.
Triggers RAG index rebuild after every approval/deletion.

Integration into app.py — same 5 steps as Phase 1:
  1. from samples_tab import SamplesTabMixin
     class DocExtractorApp(SamplesTabMixin, ctk.CTk):
  2. Add tab button in _build_right()
  3. Call self._build_samples_panel(card) in _build_right()
  4. Wire _switch_tab() hide/show blocks
  5. Add _tab_style(self._tab_samples_btn, False)
"""
from __future__ import annotations
import json, shutil, threading
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
import customtkinter as ctk

# Palette
NAVY_DEEP="#0A1628";NAVY_MID="#1B3A6B";NAVY_LIGHT="#2D5FA6"
NAVY_PALE="#5B8FD4";NAVY_GHOST="#C5D8F5";NAVY_MIST="#EAF1FB"
WHITE="#FFFFFF";OFF_WHITE="#F7F9FC";CARD_WHITE="#FFFFFF"
LIME_BRIGHT="#00E5A0";LIME_MID="#00A876";LIME_DARK="#007A56"
LIME_PALE="#B3F5E2";LIME_MIST="#E6FBF5"
BORDER_LIGHT="#DDE6F4";BORDER_MID="#B8CCEA"
TXT_NAVY="#0A1628";TXT_NAVY_MID="#1B3A6B";TXT_SOFT="#5B7BAD";TXT_MUTED="#96AFCC"
ACCENT_GOLD="#F59E0B";ACCENT_SUCCESS="#00C48A";ACCENT_RED="#EF4444"
SIDEBAR_ITEM="#142B52";BUBBLE_SYS="#FFF8EC";BUBBLE_SYS_TXT="#78450A"

DOC_TYPES = ("cic", "payslip", "saln", "itr")
APPROVED_SUFFIX = ".approved.json"
IMAGE_EXTS = {".pdf", ".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tiff"}

_SCHEMAS = {
    "cic": """{
  "full_name": null, "first_name": null, "middle_name": null, "last_name": null,
  "suffix": null, "date_of_birth": null, "age": null, "gender": null,
  "civil_status": null, "nationality": null, "number_of_dependents": null,
  "tin": null, "sss": null, "drivers_license": null,
  "residence_address": null, "mailing_address": null, "contact_number": null,
  "spouse_first_name": null, "spouse_middle_name": null, "spouse_last_name": null,
  "spouse_age": null, "spouse_occupation": null, "spouse_income": null,
  "spouse_employment_status": null, "spouse_hired_from": null, "spouse_tin": null,
  "employer_name": null, "employer_address": null, "occupation": null,
  "employment_status": null, "gross_income": null, "income_frequency": null,
  "hired_from": null, "hired_to": null,
  "credit_accounts": [], "installments_requested": [], "installments_active": [],
  "total_monthly_amortization": null, "total_loan_balance": null, "total_overdue_payments": null
}""",
    "payslip": """{
  "employee_name": null, "employer_name": null, "position": null, "department": null,
  "period_from": null, "period_to": null, "pay_date": null, "period_count": null,
  "pay_periods": [],
  "monthly_rate": null, "basic_pay": null, "allowances": null, "gross_pay": null,
  "gsis_deduction": null, "sss_deduction": null, "philhealth_deduction": null,
  "pagibig_deduction": null, "tax_deduction": null, "other_deductions": null,
  "total_deductions": null, "net_pay": null,
  "tin": null, "sss_number": null, "philhealth_number": null, "pagibig_number": null
}""",
    "saln": """{
  "compliance_type": null, "compliance_date": null,
  "declarant_family_name": null, "declarant_first_name": null, "declarant_middle_initial": null,
  "declarant_position": null, "declarant_agency": null, "declarant_office_address": null,
  "spouse_family_name": null, "spouse_first_name": null, "spouse_middle_initial": null,
  "spouse_position": null, "spouse_agency": null, "spouse_office_address": null,
  "filing_type": null,
  "unmarried_children": [{"name": null, "age": null}],
  "real_properties": [{"description": null, "kind": null, "exact_location": null,
    "assessed_value": null, "current_fair_market_value": null,
    "acquisition_year": null, "acquisition_mode": null, "acquisition_cost": null}],
  "real_properties_subtotal": null,
  "personal_properties": [{"description": null, "acquisition_year": null, "acquisition_cost": null}],
  "personal_properties_subtotal": null,
  "total_assets": null,
  "liabilities": [{"nature": null, "creditor_name": null, "outstanding_balance": null}],
  "total_liabilities": null, "net_worth": null,
  "business_interests": [{"entity_name": null, "business_address": null,
    "nature_of_interest": null, "date_of_acquisition": null}],
  "no_business_interest": null,
  "relatives_in_government": [{"name": null, "relationship": null,
    "position": null, "agency_and_address": null}],
  "no_relatives_in_government": null,
  "date_signed": null, "government_id_type": null,
  "government_id_number": null, "government_id_date_issued": null
}""",
    "itr": """{
  "taxpayer_name": null, "tin": null, "tax_year": null, "form_type": null,
  "registered_address": null, "civil_status": null, "citizenship": null,
  "business_name": null, "line_of_business": null,
  "gross_compensation_income": null, "gross_business_income": null,
  "total_gross_income": null, "net_taxable_income": null,
  "tax_due": null, "tax_paid": null,
  "spouse_name": null, "spouse_tin": null
}""",
}

_PROMPTS = {
    "cic": (
        "Extract ALL fields from this Philippine CIC Credit Report. "
        "Return ONLY valid JSON.\n\n{schema}\n\n"
        "--- DOCUMENT ---\n[See attached image/PDF above]\n--- END ---"
    ),
    "payslip": (
        "Extract ALL fields from this Philippine payslip. "
        "If multiple periods exist, return AVERAGES and list each period in pay_periods[]. "
        "Return ONLY valid JSON.\n\n{schema}\n\n"
        "--- DOCUMENT ---\n[See attached image/PDF above]\n--- END ---"
    ),
    "saln": (
        "Extract ALL fields from this Philippine SALN document. "
        "Capture every real property, personal property, liability, business interest, "
        "and relative in government service as separate array items. "
        "For filing_type extract Joint/Separate/Not Applicable. "
        "For compliance_type extract Assumption/Annual/Exit and its corresponding date. "
        "Return ONLY valid JSON.\n\n{schema}\n\n"
        "--- DOCUMENT ---\n[See attached image/PDF above]\n--- END ---"
    ),
    "itr": (
        "Extract ALL fields from this Philippine ITR. "
        "Return ONLY valid JSON.\n\n{schema}\n\n"
        "--- DOCUMENT ---\n[See attached image/PDF above]\n--- END ---"
    ),
}

_DOC_META = {
    "cic":     {"icon": "📋", "label": "CIC Credit Report",      "color": NAVY_LIGHT, "hover": NAVY_PALE,  "bg": NAVY_MIST},
    "payslip": {"icon": "💵", "label": "Payslip / Payroll",       "color": "#D97706",  "hover": "#F59E0B",  "bg": "#FFF9F0"},
    "saln":    {"icon": "📄", "label": "SALN / Net Worth",         "color": LIME_DARK,  "hover": LIME_MID,   "bg": LIME_MIST},
    "itr":     {"icon": "📊", "label": "ITR / Income Tax Return",  "color": "#00796B",  "hover": "#00695C",  "bg": "#F0FDF4"},
}


class SamplesTabMixin:
    """Mixin adding the multi-sample Samples tab to DocExtractorApp (Phase 2 RAG)."""

    # ── State ─────────────────────────────────────────────────────────────

    def _samples_init_state(self) -> None:
        self._samples_draft:         dict[str, str | None]  = {k: None for k in DOC_TYPES}
        self._samples_active_type:   str                    = "cic"
        self._samples_active_file:   Path | None            = None
        self._samples_status_bar:    tk.Label | None        = None
        self._samples_json_box:      tk.Text | None         = None
        self._samples_type_header:   tk.Label | None        = None
        self._samples_approve_btn:   ctk.CTkButton | None   = None
        self._samples_extract_btn:   ctk.CTkButton | None   = None
        self._samples_card_frames:   dict[str, tk.Frame]    = {}
        self._samples_list_frames:   dict[str, tk.Frame]    = {}
        self._samples_status_badges: dict[str, tk.Label]    = {}
        self._samples_schema_visible: bool                  = False
        self._samples_current_file_lbl: tk.Label | None     = None
        self._samples_schema_frame:  tk.Frame | None        = None
        self._samples_schema_lbl:    tk.Label | None        = None

    # ── Font helpers ──────────────────────────────────────────────────────

    def _s_font(self, size: int, weight: str = "normal") -> tuple:
        try:
            return self.F(size, weight)
        except AttributeError:
            return ("Arial", size, weight)

    def _s_mono(self, size: int, weight: str = "normal") -> tuple:
        try:
            return self.FMONO(size, weight)
        except AttributeError:
            return ("Consolas", size, weight)

    # ── Path helpers ──────────────────────────────────────────────────────

    def _s_samples_root(self) -> Path:
        return Path(__file__).resolve().parent / "samples"

    def _s_approved_paths(self, doc_type: str) -> list[Path]:
        folder = self._s_samples_root() / doc_type
        if not folder.exists():
            return []
        return sorted(f for f in folder.iterdir()
                      if f.suffix == ".json" and ".approved" in f.name)

    def _s_sample_files(self, doc_type: str) -> list[Path]:
        folder = self._s_samples_root() / doc_type
        if not folder.exists():
            return []
        return sorted(f for f in folder.iterdir()
                      if f.suffix.lower() in IMAGE_EXTS and f.is_file())

    def _s_approved_for_file(self, sample_path: Path) -> Path | None:
        c = sample_path.parent / (sample_path.stem + APPROVED_SUFFIX)
        return c if c.exists() else None

    def _s_status_for(self, doc_type: str) -> tuple[str, str]:
        approved = self._s_approved_paths(doc_type)
        files    = self._s_sample_files(doc_type)
        if approved:
            return f"✅ {len(approved)} approved", ACCENT_SUCCESS
        if files:
            return "⏳ Pending", ACCENT_GOLD
        return "○ Empty", TXT_MUTED

    def _s_ensure_folders(self) -> None:
        for dt in DOC_TYPES:
            (self._s_samples_root() / dt).mkdir(parents=True, exist_ok=True)

    def _s_trigger_rag_rebuild(self) -> None:
        def _worker():
            try:
                import few_shot as _fs
                _fs.rebuild_rag_index()
            except Exception as e:
                import logging
                logging.getLogger("cibi_populator").warning(f"RAG rebuild: {e}")
        threading.Thread(target=_worker, daemon=True).start()

    # ── Build panel ───────────────────────────────────────────────────────

    def _build_samples_panel(self, parent: tk.Frame) -> None:
        if not hasattr(self, "_samples_draft"):
            self._samples_init_state()
        self._s_ensure_folders()

        self._samples_frame = tk.Frame(parent, bg=OFF_WHITE)

        # Header
        hdr = tk.Frame(self._samples_frame, bg=NAVY_DEEP)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🗂  Few-Shot Sample Manager",
                 font=self._s_font(12, "bold"),
                 fg=WHITE, bg=NAVY_DEEP, padx=20, pady=10).pack(side="left")
        tk.Label(hdr, text="  Phase 2 — TF-IDF RAG  ",
                 font=self._s_font(7, "bold"),
                 fg=LIME_BRIGHT, bg=NAVY_MID, padx=8, pady=4
                 ).pack(side="left", padx=(0, 12), pady=8)
        ctk.CTkButton(hdr, text="🔄  Rebuild Index",
                      command=self._samples_rebuild_index_ui,
                      width=120, height=28, corner_radius=6,
                      fg_color=NAVY_MID, hover_color=NAVY_LIGHT, text_color=WHITE,
                      font=ctk.CTkFont(self._s_font(8)[0], 8, weight="bold"),
                      border_width=0).pack(side="right", padx=(0, 8), pady=10)
        ctk.CTkButton(hdr, text="📁  Open Folder",
                      command=self._samples_open_folder,
                      width=110, height=28, corner_radius=6,
                      fg_color=NAVY_MID, hover_color=NAVY_LIGHT, text_color=WHITE,
                      font=ctk.CTkFont(self._s_font(8)[0], 8, weight="bold"),
                      border_width=0).pack(side="right", padx=(0, 4), pady=10)

        # Info strip
        info = tk.Frame(self._samples_frame, bg=BUBBLE_SYS,
                        highlightbackground=BORDER_MID, highlightthickness=1)
        info.pack(fill="x")
        tk.Label(info,
                 text=("ℹ  Add up to 4-5 sample files per document type. "
                       "Each approved sample is indexed — the RAG store automatically "
                       "picks the most similar one for every new extraction."),
                 font=self._s_font(8), fg=BUBBLE_SYS_TXT, bg=BUBBLE_SYS,
                 padx=16, pady=8, wraplength=860, justify="left", anchor="w"
                 ).pack(fill="x")

        # Body
        body = tk.Frame(self._samples_frame, bg=OFF_WHITE)
        body.pack(fill="both", expand=True)

        # Left column
        left_outer = tk.Frame(body, bg=OFF_WHITE, width=315)
        left_outer.pack(side="left", fill="y", padx=(16, 0), pady=12)
        left_outer.pack_propagate(False)
        tk.Label(left_outer, text="DOCUMENT TYPES",
                 font=self._s_font(7, "bold"), fg=TXT_MUTED, bg=OFF_WHITE
                 ).pack(anchor="w", pady=(0, 8))
        for dt in DOC_TYPES:
            self._samples_build_type_card(left_outer, dt)

        # Right column
        right = tk.Frame(body, bg=CARD_WHITE,
                         highlightbackground=BORDER_MID, highlightthickness=1)
        right.pack(side="left", fill="both", expand=True, padx=16, pady=12)

        # Editor header
        editor_hdr = tk.Frame(right, bg=NAVY_DEEP)
        editor_hdr.pack(fill="x")
        self._samples_type_header = tk.Label(
            editor_hdr, text="📋  CIC Credit Report — JSON Editor",
            font=self._s_font(9, "bold"), fg=WHITE, bg=NAVY_DEEP, padx=16, pady=8)
        self._samples_type_header.pack(side="left")
        self._samples_current_file_lbl = tk.Label(
            editor_hdr, text="No file selected",
            font=self._s_font(7), fg=NAVY_GHOST, bg=NAVY_MID, padx=10, pady=4)
        self._samples_current_file_lbl.pack(side="right", padx=12, pady=8)

        # Toolbar
        toolbar = tk.Frame(right, bg=NAVY_MIST,
                           highlightbackground=BORDER_LIGHT, highlightthickness=1)
        toolbar.pack(fill="x")
        self._samples_extract_btn = ctk.CTkButton(
            toolbar, text="⚡  Extract Draft",
            command=self._samples_run_extraction,
            width=130, height=30, corner_radius=6,
            fg_color=NAVY_DEEP, hover_color=NAVY_MID, text_color=WHITE,
            font=ctk.CTkFont(self._s_font(8)[0], 8, weight="bold"),
            border_width=0, state="disabled")
        self._samples_extract_btn.pack(side="left", padx=(12, 6), pady=7)
        self._samples_approve_btn = ctk.CTkButton(
            toolbar, text="✅  Approve as Ground Truth",
            command=self._samples_approve,
            width=180, height=30, corner_radius=6,
            fg_color=LIME_DARK, hover_color=LIME_MID, text_color=WHITE,
            font=ctk.CTkFont(self._s_font(8)[0], 8, weight="bold"),
            border_width=0, state="disabled")
        self._samples_approve_btn.pack(side="left", padx=(0, 6), pady=7)
        ctk.CTkButton(toolbar, text="🗑  Clear",
                      command=self._samples_clear_editor,
                      width=70, height=30, corner_radius=6,
                      fg_color=SIDEBAR_ITEM, hover_color="#1E3A5F", text_color=NAVY_GHOST,
                      font=ctk.CTkFont(self._s_font(8)[0], 8, weight="bold"),
                      border_width=0).pack(side="left", pady=7)
        ctk.CTkButton(toolbar, text="{ } Schema",
                      command=self._samples_toggle_schema,
                      width=80, height=30, corner_radius=6,
                      fg_color=NAVY_MIST, hover_color=NAVY_GHOST, text_color=NAVY_MID,
                      font=ctk.CTkFont(self._s_font(7)[0], 7, weight="bold"),
                      border_width=1, border_color=BORDER_MID
                      ).pack(side="right", padx=12, pady=7)

        # Schema panel (hidden by default)
        self._samples_schema_frame = tk.Frame(right, bg="#F8F4FF",
                                              highlightbackground=BORDER_MID,
                                              highlightthickness=1)
        self._samples_schema_lbl = tk.Label(
            self._samples_schema_frame, text="",
            font=self._s_mono(8), fg=NAVY_MID, bg="#F8F4FF",
            justify="left", anchor="w", padx=14, pady=8, wraplength=700)
        self._samples_schema_lbl.pack(fill="x")

        # JSON editor
        editor_wrap = tk.Frame(right, bg=CARD_WHITE)
        editor_wrap.pack(fill="both", expand=True)
        sb = tk.Scrollbar(editor_wrap, relief="flat",
                          troughcolor=OFF_WHITE, bg=BORDER_LIGHT, width=8, bd=0)
        sb.pack(side="right", fill="y")
        self._samples_json_box = tk.Text(
            editor_wrap, wrap="none", font=self._s_mono(10),
            fg=TXT_NAVY, bg=CARD_WHITE, relief="flat", bd=0,
            padx=16, pady=12, insertbackground=LIME_DARK,
            yscrollcommand=sb.set, undo=True,
            selectbackground=NAVY_GHOST, selectforeground=TXT_NAVY)
        self._samples_json_box.pack(side="left", fill="both", expand=True)
        sb.config(command=self._samples_json_box.yview)
        hsb = tk.Scrollbar(right, orient="horizontal", relief="flat",
                            troughcolor=OFF_WHITE, bg=BORDER_LIGHT, width=8, bd=0)
        hsb.pack(fill="x")
        self._samples_json_box.config(xscrollcommand=hsb.set)
        hsb.config(command=self._samples_json_box.xview)
        self._samples_json_box.bind("<<Modified>>", self._samples_on_edit)
        self._samples_json_box.bind("<Control-z>", lambda e: self._samples_json_box.edit_undo())
        self._samples_json_box.bind("<Control-y>", lambda e: self._samples_json_box.edit_redo())

        # Status bar
        self._samples_status_bar = tk.Label(
            self._samples_frame,
            text="  Select a document type on the left to begin.",
            font=self._s_font(8), fg=TXT_SOFT, bg=NAVY_MIST,
            anchor="w", padx=16, pady=6)
        self._samples_status_bar.pack(fill="x", side="bottom")

        self._samples_select_type("cic")

    # ── Card ──────────────────────────────────────────────────────────────

    def _samples_build_type_card(self, parent: tk.Frame, dt: str) -> None:
        meta = _DOC_META[dt]
        badge_txt, badge_color = self._s_status_for(dt)

        outer = tk.Frame(parent, bg=BORDER_LIGHT, padx=1, pady=1, cursor="hand2")
        outer.pack(fill="x", pady=(0, 8))
        inner = tk.Frame(outer, bg=meta["bg"])
        inner.pack(fill="both", expand=True)

        tk.Canvas(inner, height=3, bg=meta["color"], highlightthickness=0).pack(fill="x")

        row = tk.Frame(inner, bg=meta["bg"])
        row.pack(fill="x", padx=12, pady=(8, 4))
        tk.Label(row, text=meta["icon"], font=("Segoe UI Emoji", 15),
                 fg=TXT_NAVY, bg=meta["bg"]).pack(side="left")

        info_col = tk.Frame(row, bg=meta["bg"])
        info_col.pack(side="left", padx=(10, 0), fill="x", expand=True)
        tk.Label(info_col, text=meta["label"],
                 font=self._s_font(9, "bold"), fg=TXT_NAVY, bg=meta["bg"], anchor="w"
                 ).pack(anchor="w")

        approved = self._s_approved_paths(dt)
        files    = self._s_sample_files(dt)
        tk.Label(info_col, text=f"{len(approved)}/{len(files)} approved",
                 font=self._s_font(7), fg=TXT_SOFT, bg=meta["bg"], anchor="w"
                 ).pack(anchor="w")

        self._samples_status_badges[dt] = tk.Label(
            row, text=badge_txt, font=self._s_font(7, "bold"),
            fg=badge_color, bg=meta["bg"])
        self._samples_status_badges[dt].pack(side="right", padx=(0, 4))

        btn_row = tk.Frame(inner, bg=meta["bg"])
        btn_row.pack(fill="x", padx=12, pady=(0, 6))
        ctk.CTkButton(btn_row, text="➕ Add Sample",
                      command=lambda d=dt: self._samples_browse(d),
                      width=100, height=26, corner_radius=6,
                      fg_color=meta["color"], hover_color=meta["hover"],
                      text_color=WHITE,
                      font=ctk.CTkFont(self._s_font(7)[0], 7, weight="bold"),
                      border_width=0).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btn_row, text="Select",
                      command=lambda d=dt: self._samples_select_type(d),
                      width=55, height=26, corner_radius=6,
                      fg_color=NAVY_MIST, hover_color=NAVY_GHOST, text_color=NAVY_MID,
                      font=ctk.CTkFont(self._s_font(7)[0], 7, weight="bold"),
                      border_width=1, border_color=BORDER_MID).pack(side="left")

        list_frame = tk.Frame(inner, bg=meta["bg"])
        list_frame.pack(fill="x", padx=8, pady=(0, 8))
        self._samples_list_frames[dt] = list_frame
        self._samples_rebuild_file_list(dt)

        self._samples_card_frames[dt] = outer
        for w in (outer, inner, row, info_col):
            w.bind("<Button-1>", lambda e, d=dt: self._samples_select_type(d))

    def _samples_rebuild_file_list(self, dt: str) -> None:
        lf = self._samples_list_frames.get(dt)
        if not lf:
            return
        for w in lf.winfo_children():
            w.destroy()
        meta  = _DOC_META[dt]
        files = self._s_sample_files(dt)
        if not files:
            tk.Label(lf, text="No files yet — click ➕ Add Sample",
                     font=self._s_font(7), fg=TXT_MUTED, bg=meta["bg"], padx=4
                     ).pack(anchor="w")
            return
        for f in files:
            is_approved = self._s_approved_for_file(f) is not None
            row = tk.Frame(lf, bg=meta["bg"])
            row.pack(fill="x", pady=1)
            tk.Label(row, text="✅" if is_approved else "○ ",
                     font=self._s_font(8),
                     fg=ACCENT_SUCCESS if is_approved else TXT_MUTED,
                     bg=meta["bg"]).pack(side="left")
            name = f.name if len(f.name) <= 22 else f.name[:19] + "…"
            tk.Label(row, text=name, font=self._s_font(7),
                     fg=TXT_NAVY_MID, bg=meta["bg"], cursor="hand2"
                     ).pack(side="left", padx=(2, 6))
            ctk.CTkButton(row, text="Load",
                          command=lambda fp=f, d=dt: self._samples_load_file(fp, d),
                          width=38, height=20, corner_radius=4,
                          fg_color=NAVY_MIST, hover_color=NAVY_GHOST, text_color=NAVY_MID,
                          font=ctk.CTkFont(self._s_font(7)[0], 7),
                          border_width=1, border_color=BORDER_MID
                          ).pack(side="left", padx=(0, 2))
            ctk.CTkButton(row, text="✕",
                          command=lambda fp=f, d=dt: self._samples_delete_file(fp, d),
                          width=24, height=20, corner_radius=4,
                          fg_color="#FEE2E2", hover_color="#FECACA",
                          text_color=ACCENT_RED,
                          font=ctk.CTkFont(self._s_font(7)[0], 7),
                          border_width=0).pack(side="left")

    def _samples_rebuild_card(self, dt: str) -> None:
        frame = self._samples_card_frames.get(dt)
        if not frame:
            return
        parent = frame.master
        frame.destroy()
        self._samples_build_type_card(parent, dt)
        if self._samples_active_type == dt:
            self._samples_highlight_card(dt)

    def _samples_highlight_card(self, dt: str) -> None:
        for key, frame in self._samples_card_frames.items():
            try:
                if key == dt:
                    frame.config(highlightbackground=_DOC_META[dt]["color"],
                                 highlightthickness=2, padx=0, pady=0)
                else:
                    frame.config(highlightbackground=BORDER_LIGHT,
                                 highlightthickness=1, padx=1, pady=1)
            except Exception:
                pass

    # ── Actions ───────────────────────────────────────────────────────────

    def _samples_select_type(self, dt: str) -> None:
        self._samples_active_type = dt
        meta = _DOC_META[dt]
        if self._samples_type_header:
            self._samples_type_header.config(
                text=f"{meta['icon']}  {meta['label']} — JSON Editor")
        if self._samples_schema_lbl:
            self._samples_schema_lbl.config(text=_SCHEMAS.get(dt, ""))

        # If active file belongs to this type, keep it loaded
        if (self._samples_active_file and
                self._samples_active_file.parent.name == dt and
                self._samples_active_file.exists()):
            self._samples_load_file(self._samples_active_file, dt)
        else:
            draft = self._samples_draft.get(dt)
            if draft:
                self._samples_set_editor(draft, editable=True)
                self._samples_set_status(
                    f"Editing draft for {meta['label']}. Review then Approve.", ACCENT_GOLD)
                self._samples_approve_btn and self._samples_approve_btn.configure(state="normal")
            else:
                self._samples_show_placeholder(dt)

        has_file = bool(self._s_sample_files(dt))
        if self._samples_extract_btn:
            self._samples_extract_btn.configure(state="normal" if has_file else "disabled")
        self._samples_highlight_card(dt)

    def _samples_show_placeholder(self, dt: str) -> None:
        meta = _DOC_META[dt]
        ph = (
            f"// No file loaded for {meta['label']}.\n//\n"
            f"// Steps:\n"
            f"//   1. Click ➕ Add Sample on the left to upload a file.\n"
            f"//   2. Click Load next to a file to open it here.\n"
            f"//   3. Click ⚡ Extract Draft to send it to Gemini.\n"
            f"//   4. Review the JSON, then click ✅ Approve.\n"
            f"//\n"
            f"// Add up to 4-5 samples per type for best RAG matching.\n"
        )
        self._samples_set_editor(ph, editable=False)
        self._samples_set_status(f"No file loaded for {meta['label']}.", TXT_MUTED)
        if self._samples_approve_btn:
            self._samples_approve_btn.configure(state="disabled")
        if self._samples_current_file_lbl:
            self._samples_current_file_lbl.config(text="No file selected")

    def _samples_load_file(self, file_path: Path, dt: str) -> None:
        self._samples_active_type = dt
        self._samples_active_file = file_path
        if self._samples_current_file_lbl:
            n = file_path.name
            self._samples_current_file_lbl.config(
                text=n if len(n) <= 30 else n[:27] + "…")
        approved = self._s_approved_for_file(file_path)
        if approved:
            try:
                content = approved.read_text(encoding="utf-8")
                self._samples_set_editor(content, editable=True)
                self._samples_set_status(
                    f"✅ Loaded approved JSON for '{file_path.name}'. Edit and re-approve if needed.",
                    ACCENT_SUCCESS)
                self._samples_approve_btn and self._samples_approve_btn.configure(state="normal")
            except Exception as e:
                self._samples_set_editor(f"// Error:\n// {e}", editable=False)
        else:
            self._samples_show_placeholder(dt)
            self._samples_set_status(
                f"'{file_path.name}' selected. Click ⚡ Extract Draft to generate JSON.",
                ACCENT_GOLD)
        if self._samples_extract_btn:
            self._samples_extract_btn.configure(state="normal")
        self._samples_highlight_card(dt)

    def _samples_browse(self, dt: str) -> None:
        meta = _DOC_META[dt]
        path = filedialog.askopenfilename(
            title=f"Add sample {meta['label']} file",
            filetypes=[
                ("PDF / Images", "*.pdf *.jpg *.jpeg *.png *.webp *.gif *.bmp *.tiff"),
                ("PDF", "*.pdf"),
                ("Images", "*.jpg *.jpeg *.png *.webp *.gif *.bmp *.tiff"),
                ("All files", "*.*"),
            ])
        if not path:
            return
        dest_folder = self._s_samples_root() / dt
        dest_folder.mkdir(parents=True, exist_ok=True)
        dest = dest_folder / Path(path).name
        try:
            if Path(path).resolve() != dest.resolve():
                shutil.copy2(path, dest)
        except Exception as e:
            self._samples_set_status(f"❌ Could not copy file: {e}", ACCENT_RED)
            return
        self._samples_set_status(
            f"✅ '{Path(path).name}' added. Click Load to open it in the editor.",
            ACCENT_SUCCESS)
        self._samples_rebuild_card(dt)
        self._samples_load_file(dest, dt)

    def _samples_delete_file(self, file_path: Path, dt: str) -> None:
        if not messagebox.askyesno(
                "Delete Sample",
                f"Delete '{file_path.name}' and its approved JSON?\nThis cannot be undone."):
            return
        try:
            file_path.unlink(missing_ok=True)
            approved = self._s_approved_for_file(file_path)
            if approved and approved.exists():
                approved.unlink()
            txt = file_path.parent / (file_path.stem + ".txt")
            if txt.exists():
                txt.unlink()
        except Exception as e:
            self._samples_set_status(f"❌ Delete failed: {e}", ACCENT_RED)
            return
        if self._samples_active_file == file_path:
            self._samples_active_file = None
            self._samples_show_placeholder(dt)
        self._samples_set_status(f"🗑 '{file_path.name}' deleted.", TXT_SOFT)
        self._samples_rebuild_card(dt)
        self._s_trigger_rag_rebuild()

    def _samples_run_extraction(self) -> None:
        dt        = self._samples_active_type
        meta      = _DOC_META[dt]
        file_path = self._samples_active_file
        if not file_path:
            files = self._s_sample_files(dt)
            file_path = files[0] if files else None
        if not file_path or not file_path.exists():
            self._samples_set_status("⚠ No file loaded. Click Load next to a file first.", ACCENT_RED)
            return
        if self._samples_extract_btn:
            self._samples_extract_btn.configure(state="disabled", text="⏳  Extracting…")
        self._samples_set_status(f"⏳ Sending '{file_path.name}' to Gemini Vision…", ACCENT_GOLD)

        def _worker():
            try:
                result    = self._samples_gemini_vision_extract(str(file_path), dt)
                draft_str = json.dumps(result, indent=2, ensure_ascii=False)
                self._samples_draft[dt] = draft_str
                (file_path.parent / (file_path.stem + ".draft.json")
                 ).write_text(draft_str, encoding="utf-8")
                self.after(0, self._samples_finish_extraction, dt, draft_str, None)
            except Exception as e:
                self.after(0, self._samples_finish_extraction, dt, None, str(e))

        threading.Thread(target=_worker, daemon=True).start()

    def _samples_finish_extraction(self, dt, draft_str, error) -> None:
        if self._samples_extract_btn:
            self._samples_extract_btn.configure(state="normal", text="⚡  Extract Draft")
        if error:
            self._samples_set_editor(f"// ERROR:\n// {error}", editable=False)
            self._samples_set_status("❌ Extraction failed — see editor above for details.", ACCENT_RED)
            return
        self._samples_set_editor(draft_str, editable=True)
        self._samples_approve_btn and self._samples_approve_btn.configure(state="normal")
        self._samples_set_status(
            "✅ Draft ready. Review the JSON above, correct if needed, then click Approve.",
            ACCENT_SUCCESS)
        self._samples_rebuild_card(dt)

    def _samples_approve(self) -> None:
        dt   = self._samples_active_type
        raw  = self._samples_json_box.get("1.0", "end").strip() if self._samples_json_box else ""
        if not raw or raw.startswith("//"):
            self._samples_set_status("⚠ Nothing to approve — editor is empty.", ACCENT_GOLD)
            return
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as e:
            self._samples_set_status(f"❌ Invalid JSON — fix before approving: {e}", ACCENT_RED)
            return

        file_path = self._samples_active_file
        if not file_path:
            files = self._s_sample_files(dt)
            file_path = files[0] if files else None
        stem = file_path.stem if file_path else f"sample_{dt}"

        dest_folder = self._s_samples_root() / dt
        dest_folder.mkdir(parents=True, exist_ok=True)
        existing = dest_folder / (stem + APPROVED_SUFFIX)
        if existing.exists():
            try:
                shutil.copy2(existing, existing.with_suffix(".json.bak"))
            except Exception:
                pass

        approved_path = dest_folder / (stem + APPROVED_SUFFIX)
        approved_path.write_text(
            json.dumps(parsed, indent=2, ensure_ascii=False), encoding="utf-8")

        self._samples_draft[dt] = None
        self._samples_set_status(
            f"✅ Approved: '{approved_path.name}' — RAG index updating in background…",
            ACCENT_SUCCESS)
        self._samples_rebuild_card(dt)
        self._samples_highlight_card(dt)
        self._s_trigger_rag_rebuild()

    def _samples_rebuild_index_ui(self) -> None:
        self._samples_set_status("🔄 Rebuilding RAG index…", ACCENT_GOLD)
        def _worker():
            try:
                import rag_store as _rs
                _rs.get_store().build(verbose=False)
                self.after(0, self._samples_set_status,
                           "✅ RAG index rebuilt successfully.", ACCENT_SUCCESS)
            except Exception as e:
                self.after(0, self._samples_set_status,
                           f"❌ Index rebuild failed: {e}", ACCENT_RED)
        threading.Thread(target=_worker, daemon=True).start()

    def _samples_clear_editor(self) -> None:
        if self._samples_json_box:
            self._samples_json_box.config(state="normal")
            self._samples_json_box.delete("1.0", "end")
        self._samples_draft[self._samples_active_type] = None
        if self._samples_approve_btn:
            self._samples_approve_btn.configure(state="disabled")
        self._samples_set_status("Editor cleared.", TXT_SOFT)

    def _samples_toggle_schema(self) -> None:
        self._samples_schema_visible = not self._samples_schema_visible
        if self._samples_schema_visible:
            self._samples_schema_frame.pack(fill="x")
        else:
            self._samples_schema_frame.pack_forget()

    def _samples_open_folder(self) -> None:
        import sys, os
        folder = self._s_samples_root()
        folder.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform == "win32":    os.startfile(str(folder))
            elif sys.platform == "darwin": os.system(f'open "{folder}"')
            else:                          os.system(f'xdg-open "{folder}"')
        except Exception:
            pass

    def _samples_on_edit(self, _event=None) -> None:
        if self._samples_json_box:
            self._samples_json_box.edit_modified(False)
        raw = self._samples_json_box.get("1.0", "end").strip() if self._samples_json_box else ""
        if self._samples_approve_btn:
            self._samples_approve_btn.configure(
                state="normal" if (raw and not raw.startswith("//")) else "disabled")

    # ── Editor helpers ─────────────────────────────────────────────────────

    def _samples_set_editor(self, content: str, editable: bool = True) -> None:
        if not self._samples_json_box:
            return
        box = self._samples_json_box
        box.config(state="normal")
        box.delete("1.0", "end")
        box.insert("1.0", content or "")
        box.edit_modified(False)
        box.config(state="normal" if editable else "disabled",
                   fg=TXT_NAVY if editable else TXT_MUTED)

    def _samples_set_status(self, msg: str, color: str = TXT_SOFT) -> None:
        if self._samples_status_bar:
            self._samples_status_bar.config(text=f"  {msg}", fg=color)

    # ── Gemini Vision ──────────────────────────────────────────────────────

    def _samples_gemini_vision_extract(self, file_path: str, doc_type: str) -> dict:
        import re as _re, mimetypes, os
        try:
            from google import genai as _genai
            from google.genai import types as _gtypes
        except ImportError:
            raise RuntimeError("google-genai not installed — run: pip install google-genai")

        api_key = (
            getattr(self, "GEMINI_API_KEY", "") or
            os.environ.get("GEMINI_API_KEY", "") or ""
        )
        if not api_key or api_key == "YOUR_GEMINI_API_KEY_HERE":
            raise RuntimeError("No Gemini API key found. Set GEMINI_API_KEY in your .env file.")

        mime, _ = mimetypes.guess_type(file_path)
        if not mime:
            mime = "application/pdf" if file_path.lower().endswith(".pdf") else "image/jpeg"
        with open(file_path, "rb") as fh:
            raw_bytes = fh.read()

        prompt_text = _PROMPTS[doc_type].format(schema=_SCHEMAS[doc_type])
        client    = _genai.Client(api_key=api_key)
        doc_part  = _gtypes.Part(inline_data=_gtypes.Blob(mime_type=mime, data=raw_bytes))
        text_part = _gtypes.Part(text=prompt_text)

        response_text = ""
        for model in ("gemini-2.5-flash", "gemini-2.0-flash"):
            try:
                resp = client.models.generate_content(
                    model    = model,
                    contents = [_gtypes.Content(role="user", parts=[doc_part, text_part])],
                    config   = _gtypes.GenerateContentConfig(
                        max_output_tokens=16_000, temperature=0.0),
                )
                try:
                    response_text = resp.text or ""
                except Exception:
                    response_text = "".join(
                        p.text for p in resp.candidates[0].content.parts
                        if hasattr(p, "text") and p.text)
                if response_text:
                    break
            except Exception as e:
                if model == "gemini-2.0-flash":
                    raise RuntimeError(f"All Gemini models failed: {e}") from e
                continue

        cleaned = _re.sub(r"^```(?:json)?\s*", "", response_text.strip(), flags=_re.I)
        cleaned = _re.sub(r"\s*```$", "", cleaned.strip())
        try:
            return json.loads(cleaned)
        except json.JSONDecodeError:
            pass
        m = _re.search(r"\{[\s\S]+\}", cleaned)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
        raise RuntimeError(f"JSON parse failed. Gemini returned:\n{response_text[:400]}")
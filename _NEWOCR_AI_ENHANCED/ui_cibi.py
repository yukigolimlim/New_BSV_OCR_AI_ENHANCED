"""
ui_cibi.py — DocExtract Pro
==============================
All 4 CIBI stage builders and workflow logic
attached to DocExtractorApp.
"""
import os
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from app_constants import *
from app_constants import _ai_check_stage1
from extraction import extract as _extract_fn, extract_bank_ci_vlm, bank_ci_to_ai_context
from Cibi_populator import populate_cibi_form
from cibi_analysis import run_cibi_analysis, run_cibi_analysis_from_text
from admin_logs import insert_log


def _cibi_log(self, action: str, description: str):
    """Best-effort audit log writer for CIBI workflow actions."""
    try:
        insert_log(self, action, description)
    except Exception:
        pass

def _build_cibi_output_panel(self, parent):
    self._cibi_output_frame = tk.Frame(parent, bg=CARD_WHITE)

    hdr = tk.Frame(self._cibi_output_frame, bg=NAVY_DEEP)
    hdr.pack(fill="x")
    tk.Label(hdr, text="📋  CIBI Mode — Intelligent Loan Workflow",
             font=F(12, "bold"), fg=WHITE, bg=NAVY_DEEP, padx=20, pady=10).pack(side="left")
    tk.Label(hdr, text="  Gemini 2.5 Flash  ",
             font=F(7, "bold"), fg=LIME_BRIGHT, bg=NAVY_MID, padx=8, pady=4
             ).pack(side="left", padx=(0, 12), pady=8)

    ctrl_row = tk.Frame(hdr, bg=NAVY_DEEP)
    ctrl_row.pack(side="right", padx=12, pady=8)
    ctk.CTkButton(ctrl_row, text="📁  Open Folder",
                  command=self._open_cibi_output_folder,
                  width=110, height=28, corner_radius=6,
                  fg_color=NAVY_MID, hover_color=NAVY_LIGHT, text_color=WHITE,
                  font=FF(8, "bold"),
                  border_width=0).pack(side="left", padx=(0, 6))
    ctk.CTkButton(ctrl_row, text="↺  Reset",
                  command=self._cibi_full_reset,
                  width=80, height=28, corner_radius=6,
                  fg_color=SIDEBAR_ITEM, hover_color=SIDEBAR_HVR,
                  text_color="#C5D8F5",
                  font=FF(8, "bold"),
                  border_width=0).pack(side="left")

    self._stage_strip = tk.Frame(self._cibi_output_frame, bg=NAVY_MIST,
                                  highlightbackground=BORDER_MID, highlightthickness=1)
    self._stage_strip.pack(fill="x")
    self._stage_indicators = {}
    stage_defs = [
        ("1", "CIC + Bank CI"),
        ("2", "CI Gate"),
        ("3", "Other Docs"),
        ("4", "Populate CIBI"),
    ]
    for num, label in stage_defs:
        col = tk.Frame(self._stage_strip, bg=NAVY_MIST)
        col.pack(side="left", padx=12, pady=6)
        badge = tk.Canvas(col, width=22, height=22, bg=NAVY_MIST, highlightthickness=0)
        badge.pack(side="left", padx=(0, 5))
        badge.create_oval(1, 1, 21, 21, fill=BORDER_MID, outline="")
        badge.create_text(11, 11, text=num, font=F(8, "bold"), fill=WHITE, anchor="center")
        lbl = tk.Label(col, text=label, font=F(8), fg=TXT_MUTED, bg=NAVY_MIST)
        lbl.pack(side="left")
        self._stage_indicators[num] = (badge, lbl)

    body_sb = tk.Scrollbar(self._cibi_output_frame, relief="flat",
                           troughcolor=OFF_WHITE, bg=BORDER_LIGHT, width=8, bd=0)
    body_sb.pack(side="right", fill="y")
    body_canvas = tk.Canvas(self._cibi_output_frame, bg=OFF_WHITE,
                            highlightthickness=0, yscrollcommand=body_sb.set)
    body_canvas.pack(side="left", fill="both", expand=True)
    body_sb.config(command=body_canvas.yview)
    self._cibi_body = tk.Frame(body_canvas, bg=OFF_WHITE)
    body_win = body_canvas.create_window((0, 0), window=self._cibi_body, anchor="nw")
    self._cibi_body.bind("<Configure>",
        lambda e: body_canvas.configure(scrollregion=body_canvas.bbox("all")))
    body_canvas.bind("<Configure>",
        lambda e: body_canvas.itemconfig(body_win, width=e.width))
    body_canvas.bind("<MouseWheel>",
        lambda e: body_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
    self._cibi_body.bind("<MouseWheel>",
        lambda e: body_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

    self._build_cibi_stage1()
    self._build_cibi_stage2()
    self._build_cibi_stage3()
    self._build_cibi_stage4()

    self._cibi_show_stage("1")

# ── STAGE INDICATOR HELPERS ───────────────────────────────────────────
def _cibi_show_stage(self, active_num: str):
    colors = {"done": LIME_DARK, "active": NAVY_DEEP, "inactive": BORDER_MID}
    stage_order = ["1", "2", "3", "4"]
    active_idx  = stage_order.index(active_num) if active_num in stage_order else 0

    for i, num in enumerate(stage_order):
        badge, lbl = self._stage_indicators[num]
        badge.delete("all")
        if i < active_idx:
            fill = colors["done"]; txt_color = WHITE; lbl_color = LIME_DARK
        elif i == active_idx:
            fill = colors["active"]; txt_color = WHITE; lbl_color = NAVY_DEEP
        else:
            fill = colors["inactive"]; txt_color = WHITE; lbl_color = TXT_MUTED
        badge.create_oval(1, 1, 21, 21, fill=fill, outline="")
        if i < active_idx:
            badge.create_text(11, 11, text="✓", font=F(8, "bold"), fill=txt_color)
        else:
            badge.create_text(11, 11, text=num, font=F(8, "bold"), fill=txt_color)
        lbl.config(fg=lbl_color)

# ── STAGE 1 PANEL ─────────────────────────────────────────────────────
def _build_cibi_stage1(self):
    self._cibi_s1_frame = tk.Frame(self._cibi_body, bg=OFF_WHITE)
    self._cibi_s1_frame.pack(fill="x", padx=24, pady=(16, 0))

    hdr = tk.Frame(self._cibi_s1_frame, bg=OFF_WHITE)
    hdr.pack(fill="x", pady=(0, 10))
    self._step_number_lbl(hdr, "1")
    tk.Label(hdr, text="UPLOAD CIC & BANK CI",
             font=F(9, "bold"), fg=NAVY_DEEP, bg=OFF_WHITE).pack(side="left", padx=(8, 0))

    tier_box = tk.Frame(self._cibi_s1_frame, bg=NAVY_MIST,
                         highlightbackground=BORDER_MID, highlightthickness=1)
    tier_box.pack(fill="x", pady=(0, 12))
    tk.Label(tier_box,
             text="ℹ  Loan Tier Logic:  CIC present → loan likely > ₱100,000  |  "
                  "No CIC → loan < ₱100,000 (Bank CI only required)",
             font=F(8), fg=NAVY_MID, bg=NAVY_MIST, padx=14, pady=8,
             wraplength=800, justify="left").pack(fill="x")

    slots_row = tk.Frame(self._cibi_s1_frame, bg=OFF_WHITE)
    slots_row.pack(fill="x")

    slot_defs = {
        "CIC":     ("📋", "CIC Credit Report",   "(Optional — for loans > ₱100k)", NAVY_MIST,  NAVY_LIGHT,  NAVY_PALE),
        "BANK_CI": ("🏦", "Bank CI / Certification", "(Required for all loans)",   LIME_MIST,  LIME_DARK,   LIME_MID),
    }
    for key, (icon, label, hint, slot_bg, btn_color, hover_color) in slot_defs.items():
        slot = tk.Frame(slots_row, bg=BORDER_LIGHT, padx=1, pady=1)
        slot.pack(side="left", fill="both", expand=True, padx=(0, 8))
        inner = tk.Frame(slot, bg=slot_bg)
        inner.pack(fill="both", expand=True)

        acc = tk.Canvas(inner, height=3, bg=btn_color, highlightthickness=0)
        acc.pack(fill="x")
        if key == "CIC":
            acc.bind("<Configure>", lambda e, c=acc: self._hbar(c, NAVY_LIGHT, NAVY_PALE, 30))
        else:
            acc.bind("<Configure>", lambda e, c=acc: self._hbar(c, LIME_MID, LIME_PALE, 30))

        sh = tk.Frame(inner, bg=slot_bg)
        sh.pack(fill="x", padx=12, pady=(10, 2))
        tk.Label(sh, text=icon, font=("Segoe UI Emoji", 18), fg=NAVY_MID, bg=slot_bg).pack(side="left")
        title_col = tk.Frame(sh, bg=slot_bg)
        title_col.pack(side="left", padx=(8, 0))
        tk.Label(title_col, text=label, font=F(9, "bold"), fg=TXT_NAVY, bg=slot_bg, anchor="w").pack(anchor="w")
        tk.Label(title_col, text=hint, font=F(7), fg=TXT_SOFT, bg=slot_bg, anchor="w").pack(anchor="w")
        self._cibi_status_labels[key] = tk.Label(sh, text="❌", font=F(11), fg=ACCENT_RED, bg=slot_bg)
        self._cibi_status_labels[key].pack(side="right")

        name_lbl = tk.Label(inner, text="No file selected",
                             font=F(8), fg=TXT_MUTED, bg=slot_bg,
                             wraplength=240, anchor="w", justify="left")
        name_lbl.pack(fill="x", padx=12, pady=(0, 6))
        self._cibi_name_labels[key] = name_lbl

        ctk.CTkButton(
            inner, text="Browse…",
            command=lambda k=key: self._cibi_browse_slot(k),
            height=30, corner_radius=6,
            fg_color=btn_color, hover_color=hover_color,
            text_color=WHITE,
            font=FF(8, "bold"),
            border_width=0,
        ).pack(fill="x", padx=12, pady=(0, 12))

    tk.Frame(self._cibi_s1_frame, bg=OFF_WHITE, height=12).pack()

    self._cibi_check_btn = ctk.CTkButton(
        self._cibi_s1_frame,
        text="🔍  Check CIC & Evaluate Bank CI",
        command=self._cibi_run_stage1,
        height=42, corner_radius=8,
        fg_color=NAVY_DEEP, hover_color=NAVY_MID, text_color=WHITE,
        font=FF(10, "bold"),
        state="disabled", border_width=0
    )
    self._cibi_check_btn.pack(fill="x")

    self._cibi_s1_result = tk.Label(
        self._cibi_s1_frame, text="",
        font=F(9), fg=TXT_SOFT, bg=OFF_WHITE,
        wraplength=700, justify="left", anchor="w"
    )
    self._cibi_s1_result.pack(fill="x", pady=(8, 0))

# ── STAGE 2 PANEL — Bank CI Gate ──────────────────────────────────────
def _build_cibi_stage2(self):
    self._cibi_s2_frame = tk.Frame(self._cibi_body, bg=OFF_WHITE)

    tk.Frame(self._cibi_s2_frame, bg=BORDER_LIGHT, height=1).pack(fill="x", pady=(16, 0))

    hdr = tk.Frame(self._cibi_s2_frame, bg=OFF_WHITE)
    hdr.pack(fill="x", padx=24, pady=(10, 8))
    self._step_number_lbl(hdr, "2")
    tk.Label(hdr, text="BANK CI EVALUATION — PROCEED / STOP",
             font=F(9, "bold"), fg=NAVY_DEEP, bg=OFF_WHITE).pack(side="left", padx=(8, 0))

    self._cibi_ci_card = tk.Frame(self._cibi_s2_frame, bg=CARD_WHITE,
                                   highlightbackground=BORDER_MID, highlightthickness=1)
    self._cibi_ci_card.pack(fill="x", padx=24)

    self._cibi_ci_verdict_lbl = tk.Label(
        self._cibi_ci_card, text="",
        font=F(13, "bold"), fg=TXT_NAVY,
        bg=CARD_WHITE, padx=16
    )
    self._cibi_ci_verdict_lbl.pack(fill="x", pady=(8, 4))

    self._cibi_ci_summary_lbl = tk.Label(
        self._cibi_ci_card, text="",
        font=F(9), fg=TXT_SOFT,
        bg=CARD_WHITE, padx=16,
        wraplength=740, justify="left", anchor="w"
    )
    self._cibi_ci_summary_lbl.pack(fill="x", pady=(0, 4))

    self._cibi_ci_details_lbl = tk.Label(
        self._cibi_ci_card, text="",
        font=FMONO(9), fg=TXT_NAVY,
        bg=CARD_WHITE, padx=16,
        wraplength=740, justify="left", anchor="w"
    )
    self._cibi_ci_details_lbl.pack(fill="x", pady=(0, 10))

    btn_row = tk.Frame(self._cibi_s2_frame, bg=OFF_WHITE)
    btn_row.pack(fill="x", padx=24, pady=(10, 4))

    self._cibi_proceed_btn = ctk.CTkButton(
        btn_row, text="✅  PROCEED — Continue to Document Upload",
        command=self._cibi_proceed_to_stage3,
        height=44, corner_radius=8,
        fg_color=LIME_DARK, hover_color=LIME_MID, text_color=WHITE,
        font=FF(10, "bold"),
        state="disabled", border_width=0
    )
    self._cibi_proceed_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))

    self._cibi_stop_btn = ctk.CTkButton(
        btn_row, text="🛑  STOP — End Process",
        command=self._cibi_stop_process,
        height=44, corner_radius=8,
        fg_color=ACCENT_RED, hover_color="#DC2626", text_color=WHITE,
        font=FF(10, "bold"),
        state="disabled", border_width=0
    )
    self._cibi_stop_btn.pack(side="left", fill="x", expand=True)

    self._cibi_override_lbl = tk.Label(
        self._cibi_s2_frame,
        text="",
        font=F(8), fg=TXT_MUTED, bg=OFF_WHITE,
        padx=24, anchor="w"
    )
    self._cibi_override_lbl.pack(fill="x", pady=(4, 8))

# ── STAGE 3 PANEL — Upload remaining docs ─────────────────────────────
def _build_cibi_stage3(self):
    self._cibi_s3_frame = tk.Frame(self._cibi_body, bg=OFF_WHITE)

    tk.Frame(self._cibi_s3_frame, bg=BORDER_LIGHT, height=1).pack(fill="x", pady=(16, 0))

    hdr = tk.Frame(self._cibi_s3_frame, bg=OFF_WHITE)
    hdr.pack(fill="x", padx=24, pady=(10, 8))
    self._step_number_lbl(hdr, "3")

    self._cibi_tier_badge = tk.Label(
        hdr, text="",
        font=F(8, "bold"), fg=WHITE, bg=BORDER_MID, padx=10, pady=3
    )
    self._cibi_tier_badge.pack(side="right", padx=(0, 4))

    tk.Label(hdr, text="UPLOAD REMAINING DOCUMENTS",
             font=F(9, "bold"), fg=NAVY_DEEP, bg=OFF_WHITE).pack(side="left", padx=(8, 0))

    slot_defs = {
        "PAYSLIP": ("💵", "Payslip / Payroll",      "Required",  "#FFF9F0", "#D97706", "#F59E0B"),
        "ITR":     ("📊", "ITR / Income Tax Return", "Optional",  "#F0FDF4", "#00796B", "#00695C"),
        "SALN":    ("📄", "SALN / Net Worth",        "Optional",  WHITE,     NAVY_LIGHT, NAVY_PALE),
    }

    slots_row = tk.Frame(self._cibi_s3_frame, bg=OFF_WHITE)
    slots_row.pack(fill="x", padx=24)

    for key, (icon, label, required_txt, slot_bg, btn_color, hover_color) in slot_defs.items():
        slot = tk.Frame(slots_row, bg=BORDER_LIGHT, padx=1, pady=1)
        slot.pack(side="left", fill="both", expand=True, padx=(0, 8))
        inner = tk.Frame(slot, bg=slot_bg)
        inner.pack(fill="both", expand=True)

        acc = tk.Canvas(inner, height=3, bg=btn_color, highlightthickness=0)
        acc.pack(fill="x")

        sh = tk.Frame(inner, bg=slot_bg)
        sh.pack(fill="x", padx=12, pady=(10, 2))
        tk.Label(sh, text=icon, font=("Segoe UI Emoji", 18), fg=NAVY_MID, bg=slot_bg).pack(side="left")
        title_col = tk.Frame(sh, bg=slot_bg)
        title_col.pack(side="left", padx=(8, 0))
        tk.Label(title_col, text=label, font=F(9, "bold"), fg=TXT_NAVY, bg=slot_bg, anchor="w").pack(anchor="w")
        req_color = ACCENT_RED if required_txt == "Required" else TXT_MUTED
        tk.Label(title_col, text=required_txt, font=F(7, "bold"), fg=req_color, bg=slot_bg, anchor="w").pack(anchor="w")
        self._cibi_status_labels[key] = tk.Label(sh, text="○", font=F(11), fg=TXT_MUTED, bg=slot_bg)
        self._cibi_status_labels[key].pack(side="right")

        name_lbl = tk.Label(inner, text="No file selected",
                             font=F(8), fg=TXT_MUTED, bg=slot_bg,
                             wraplength=200, anchor="w", justify="left")
        name_lbl.pack(fill="x", padx=12, pady=(0, 6))
        self._cibi_name_labels[key] = name_lbl

        ctk.CTkButton(
            inner, text="Browse…",
            command=lambda k=key: self._cibi_browse_slot(k),
            height=30, corner_radius=6,
            fg_color=btn_color, hover_color=hover_color,
            text_color=WHITE,
            font=FF(8, "bold"),
            border_width=0,
        ).pack(fill="x", padx=12, pady=(0, 12))

    tk.Frame(self._cibi_s3_frame, bg=OFF_WHITE, height=10).pack()

    # ══════════════════════════════════════════════════════════════════
    #  CHANGE 2 of 3 — _cibi_extract_all() updated with parallel extraction
    # ══════════════════════════════════════════════════════════════════
    self._cibi_extract_all_btn = ctk.CTkButton(
        self._cibi_s3_frame,
        text="⚡  Extract All Documents",
        command=self._cibi_extract_all,
        height=42, corner_radius=8,
        fg_color=NAVY_DEEP, hover_color=NAVY_MID, text_color=WHITE,
        font=FF(10, "bold"),
        state="disabled", border_width=0
    )
    self._cibi_extract_all_btn.pack(fill="x", padx=24)

    self._cibi_s3_progress_lbl = tk.Label(
        self._cibi_s3_frame, text="",
        font=F(9), fg=TXT_SOFT, bg=OFF_WHITE,
        wraplength=700, justify="left", anchor="w"
    )
    self._cibi_s3_progress_lbl.pack(fill="x", padx=24, pady=(8, 4))

# ── STAGE 4 PANEL — Template + Populate ───────────────────────────────
def _build_cibi_stage4(self):
    self._cibi_s4_frame = tk.Frame(self._cibi_body, bg=OFF_WHITE)

    tk.Frame(self._cibi_s4_frame, bg=BORDER_LIGHT, height=1).pack(fill="x", pady=(16, 0))

    hdr = tk.Frame(self._cibi_s4_frame, bg=OFF_WHITE)
    hdr.pack(fill="x", padx=24, pady=(10, 8))
    self._step_number_lbl(hdr, "4")
    tk.Label(hdr, text="SELECT TEMPLATE & POPULATE CIBI EXCEL",
             font=F(9, "bold"), fg=NAVY_DEEP, bg=OFF_WHITE).pack(side="left", padx=(8, 0))

    tmpl_card = tk.Frame(self._cibi_s4_frame, bg=WHITE,
                          highlightbackground=BORDER_MID, highlightthickness=1)
    tmpl_card.pack(fill="x", padx=24)
    tk.Label(tmpl_card, text="Template:", font=F(8, "bold"),
              fg=NAVY_MID, bg=WHITE, padx=14, pady=10).pack(side="left")
    self._cibi_tmpl_lbl = tk.Label(tmpl_card, text="None selected",
                                    font=F(9), fg=TXT_MUTED, bg=WHITE, pady=10)
    self._cibi_tmpl_lbl.pack(side="left", fill="x", expand=True)
    ctk.CTkButton(
        tmpl_card, text="Browse Template…",
        command=self._cibi_pick_template,
        width=140, height=28, corner_radius=6,
        fg_color=NAVY_MIST, hover_color=NAVY_GHOST, text_color=NAVY_MID,
        font=FF(8, "bold"),
        border_width=1, border_color=BORDER_MID
    ).pack(side="right", padx=10, pady=8)

    tk.Frame(self._cibi_s4_frame, bg=OFF_WHITE, height=10).pack()

    self._cibi_populate_btn = ctk.CTkButton(
        self._cibi_s4_frame,
        text="✅  Populate → CIBI Excel",
        command=self._cibi_populate,
        height=44, corner_radius=8,
        fg_color=LIME_DARK, hover_color=LIME_MID, text_color=WHITE,
        font=FF(10, "bold"),
        state="disabled", border_width=0
    )
    self._cibi_populate_btn.pack(fill="x", padx=24)

    self._cibi_s4_progress_lbl = tk.Label(
        self._cibi_s4_frame, text="",
        font=F(9), fg=TXT_SOFT, bg=OFF_WHITE,
        wraplength=700, justify="left", anchor="w"
    )
    self._cibi_s4_progress_lbl.pack(fill="x", padx=24, pady=(8, 4))

    tk.Frame(self._cibi_s4_frame, bg=BORDER_LIGHT, height=1).pack(fill="x", padx=24, pady=(8, 0))
    self._cibi_result_frame = tk.Frame(self._cibi_s4_frame, bg=LIME_MIST,
                                        highlightbackground=LIME_MID, highlightthickness=1)
    self._cibi_result_lbl = tk.Label(
        self._cibi_result_frame, text="",
        font=F(10, "bold"), fg=LIME_DARK, bg=LIME_MIST,
        padx=20, pady=10, anchor="w", justify="left"
    )
    self._cibi_result_lbl.pack(fill="x")
    self._cibi_open_btn = ctk.CTkButton(
        self._cibi_result_frame, text="📂  Open File",
        command=self._open_last_cibi_file,
        width=130, height=32, corner_radius=6,
        fg_color=LIME_DARK, hover_color=LIME_MID, text_color=WHITE,
        font=FF(9, "bold"), border_width=0
    )
    self._cibi_open_btn.pack(anchor="w", padx=20, pady=(0, 12))

# ══════════════════════════════════════════════════════════════════════
#  CIBI WORKFLOW LOGIC
# ══════════════════════════════════════════════════════════════════════

def _step_number_lbl(self, parent, number):
    badge = tk.Canvas(parent, width=24, height=24,
                      bg=parent.cget("bg"), highlightthickness=0)
    badge.pack(side="left")
    badge.create_oval(1, 1, 23, 23, fill=NAVY_DEEP, outline="")
    badge.create_text(12, 12, text=number, font=F(9, "bold"),
                      fill=LIME_BRIGHT, anchor="center")

def _cibi_browse_slot(self, key: str):
    path = filedialog.askopenfilename(
        title=f"Select {key} file",
        filetypes=[
            ("All supported",
             "*.pdf *.txt *.docx *.xlsx *.xls *.csv *.md "
             "*.png *.jpg *.jpeg *.bmp *.tiff *.tif *.webp *.gif"),
            ("PDF",    "*.pdf"),
            ("Word",   "*.docx *.doc"),
            ("Images", "*.png *.jpg *.jpeg *.bmp *.tiff *.tif *.webp *.gif"),
            ("All files", "*.*"),
        ]
    )
    if not path:
        return
    self._cibi_slots[key]["path"] = path
    self._cibi_slots[key]["text"] = None
    name  = Path(path).name
    short = name if len(name) <= 32 else name[:29] + "…"
    self._cibi_name_labels[key].config(text=short, fg=TXT_NAVY_MID)
    self._cibi_status_labels[key].config(text="⏳", fg=ACCENT_GOLD)
    self._cibi_refresh_stage_buttons()

def _cibi_refresh_stage_buttons(self):
    bank_ci_ready = bool(self._cibi_slots["BANK_CI"]["path"])
    self._cibi_check_btn.configure(state="normal" if bank_ci_ready else "disabled")

    payslip_ready = bool(self._cibi_slots["PAYSLIP"]["path"])
    self._cibi_extract_all_btn.configure(state="normal" if payslip_ready else "disabled")

    payslip_extracted = self._cibi_slots["PAYSLIP"]["text"] is not None
    has_template = self._cibi_template_path is not None
    self._cibi_populate_btn.configure(
        state="normal" if (payslip_extracted and has_template) else "disabled"
    )

def _cibi_full_reset(self):
    for key in self._cibi_slots:
        self._cibi_slots[key]["path"] = None
        self._cibi_slots[key]["text"] = None
        if key in self._cibi_status_labels:
            icon = "❌" if key in ("CIC", "BANK_CI") else "○"
            self._cibi_status_labels[key].config(text=icon, fg=ACCENT_RED if key == "BANK_CI" else TXT_MUTED)
        if key in self._cibi_name_labels:
            self._cibi_name_labels[key].config(text="No file selected", fg=TXT_MUTED)

    self._cibi_template_path = None
    self._cibi_loan_tier     = "unknown"
    self._cibi_bank_ci_result = {}
    self._cibi_stage         = "idle"
    self._cibi_has_cic       = False
    self._last_cibi_path     = None

    self._cibi_s2_frame.pack_forget()
    self._cibi_s3_frame.pack_forget()
    self._cibi_s4_frame.pack_forget()
    self._cibi_s1_result.config(text="")
    self._cibi_check_btn.configure(state="disabled")
    self._cibi_show_stage("1")
    self._switch_tab("cibi")

# ── STAGE 1 ACTION (FIX-2: updated with single AI call) ───────────────────
def _cibi_run_stage1(self):
    bank_ci_path = self._cibi_slots["BANK_CI"]["path"]
    if not bank_ci_path:
        self._cibi_s1_result.config(
            text="⚠ Bank CI file is required. Please browse and select it.",
            fg=ACCENT_RED)
        return

    self._cibi_check_btn.configure(state="disabled", text="Checking…")
    self._cibi_s1_result.config(text="⏳ Extracting files…", fg=ACCENT_GOLD)
    self._cibi_show_stage("1")

    cic_path = self._cibi_slots["CIC"]["path"]

    def worker():
        # ── Extract Bank CI (PaddleOCR + TrOCR — handwriting-aware) ──────
        self.after(0, self._cibi_s1_result.config,
                   {"text": "⏳ Extracting Bank CI (PaddleOCR + TrOCR)…",
                    "fg": ACCENT_GOLD})
        try:
            def _bank_ci_cb(pct, msg=""):
                self.after(0, self._cibi_s1_result.config,
                           {"text": f"Bank CI — {msg or 'Processing…'} ({pct}%)",
                            "fg": ACCENT_GOLD})

            _bc_ocr_result = extract_bank_ci_vlm(bank_ci_path, _bank_ci_cb)

            # Store plain text for AI prompt
            bank_ci_text = bank_ci_to_ai_context(_bc_ocr_result)
            self._cibi_slots["BANK_CI"]["text"]         = bank_ci_text
            # Stash structured result for the override check in Stage 2
            self._cibi_slots["BANK_CI"]["_ocr_result"]  = _bc_ocr_result

            # Status icon reflects OCR-level pre-verdict
            _bc_icon, _bc_color = {
                "GOOD":      ("✅", ACCENT_SUCCESS),
                "BAD":       ("❌", ACCENT_RED),
                "UNCERTAIN": ("⚠️", ACCENT_GOLD),
            }.get(_bc_ocr_result.verdict, ("⚠️", ACCENT_GOLD))
            self.after(0, self._cibi_status_labels["BANK_CI"].config,
                       {"text": _bc_icon, "fg": _bc_color})

        except Exception as e:
            self.after(0, self._cibi_s1_result.config,
                       {"text": f"❌ Bank CI extraction failed: {e}",
                        "fg": ACCENT_RED})
            self.after(0, self._cibi_check_btn.configure,
                       {"state": "normal",
                        "text": "🔍  Check CIC & Evaluate Bank CI"})
            return

        # ── Extract CIC (if provided) ─────────────────────────────────────
        cic_text = ""
        if cic_path:
            self.after(0, self._cibi_s1_result.config,
                       {"text": "⏳ Extracting CIC…", "fg": ACCENT_GOLD})
            try:
                cic_text = _extract_fn(cic_path)
                self._cibi_slots["CIC"]["text"] = cic_text
                self.after(0, self._cibi_status_labels["CIC"].config,
                           {"text": "✅", "fg": ACCENT_SUCCESS})
            except Exception as e:
                self.after(0, self._cibi_s1_result.config,
                           {"text": f"⚠ CIC extraction warning: {e}",
                            "fg": ACCENT_GOLD})
        else:
            self.after(0, self._cibi_status_labels["CIC"].config,
                       {"text": "—", "fg": TXT_MUTED})

        # ══════════════════════════════════════════════════════════════════
        #  CHANGE 3 of 3 — Prime CIC cache
        # ══════════════════════════════════════════════════════════════════
        # ── FIX-3: Prime the CIC cache so Stage 4 gets an instant hit ────
        # extract_cic() (called inside prime_cic_cache) stores the Gemini
        # JSON result in .cache/ keyed by the first 5,000 chars of cic_text.
        # When populate_cibi_form() calls extract_cic() with the same text
        # in Stage 4, it finds the entry and returns immediately — 0 API calls.
        if cic_text:
            try:
                from extraction import prime_cic_cache
                prime_cic_cache(cic_text, GEMINI_API_KEY)
            except Exception:
                pass   # non-fatal — Stage 4 will re-extract normally

        # ── FIX-2: Single combined AI check ───────────────────────────────
        self.after(0, self._cibi_s1_result.config,
                   {"text": "🤖 AI evaluating Bank CI & CIC (combined)…",
                    "fg": ACCENT_GOLD})

        bank_ci_result, cic_tier_result = _ai_check_stage1(
            bank_ci_text, cic_text, GEMINI_API_KEY
        )

        # Add has_cic flag for downstream use
        cic_tier_result["has_cic"] = bool(cic_path)

        self.after(0, self._cibi_finish_stage1, cic_tier_result, bank_ci_result)

    threading.Thread(target=worker, daemon=True).start()

def _cibi_finish_stage1(self, cic_result: dict, bank_ci_result: dict):
    self._cibi_has_cic   = cic_result.get("has_cic", False)
    self._cibi_loan_tier = cic_result.get("tier", "unknown")

    # Cross-check Gemini verdict against pixel-level OCR verdict.
    # If PaddleOCR+TrOCR found hard evidence (derogatory keywords confirmed
    # at pixel level) that Gemini missed, respect the OCR finding.
    # If OCR clearly confirmed NCD/signatures but Gemini returned UNCERTAIN,
    # upgrade to GOOD so the officer isn't unnecessarily blocked.
    _ocr_check = self._cibi_slots["BANK_CI"].get("_ocr_result")
    if _ocr_check is not None:
        if (_ocr_check.verdict == "BAD"
                and bank_ci_result.get("verdict") != "BAD"):
            bank_ci_result["verdict"] = "BAD"
            bank_ci_result["proceed"] = False
            bank_ci_result["summary"] = (
                "[OCR override — derogatory found at pixel level] "
                + _ocr_check.summary
            )
            bank_ci_result["details"] = _ocr_check.details
        elif (_ocr_check.verdict == "GOOD"
                  and bank_ci_result.get("verdict") == "UNCERTAIN"):
            bank_ci_result["verdict"] = "GOOD"
            bank_ci_result["proceed"] = True
            bank_ci_result["summary"] = (
                "[OCR confirmed — NCD/signatures detected] "
                + _ocr_check.summary
            )

    self._cibi_bank_ci_result = bank_ci_result
    self._cibi_stage          = "cic_checked"

    tier = self._cibi_loan_tier
    if tier == "above_100k":
        tier_txt = "⬆ Loan > ₱100,000 (CIC required)"
    elif tier == "below_100k":
        tier_txt = "⬇ Loan < ₱100,000"
    else:
        tier_txt = "Loan tier unknown"

    applicant = cic_result.get("applicant_name", "")
    loan_amt  = cic_result.get("loan_amount")
    cic_summary = cic_result.get("summary", "")
    amt_str   = f"₱{loan_amt:,.0f}" if loan_amt else "N/A"

    s1_txt = (
        f"{'✅' if self._cibi_has_cic else 'ℹ'}  CIC: "
        f"{'Uploaded — Applicant: ' + applicant if self._cibi_has_cic else 'Not provided'}"
        + (f"   |   Loan Amount: {amt_str}" if loan_amt else "")
        + f"\n📊  Loan Tier: {tier_txt}"
        + (f"\n\nCIC Summary: {cic_summary[:200]}" if cic_summary and "failed" not in cic_summary.lower() else "")
    )
    self._cibi_s1_result.config(text=s1_txt, fg=LIME_DARK)
    self._cibi_check_btn.configure(state="normal", text="🔍  Re-check CIC & Bank CI")

    verdict  = bank_ci_result.get("verdict", "UNCERTAIN")
    summary  = bank_ci_result.get("summary", "")
    details  = bank_ci_result.get("details", "")
    proceed  = bank_ci_result.get("proceed", False)

    verdict_colors = {"GOOD": ACCENT_SUCCESS, "BAD": ACCENT_RED, "UNCERTAIN": ACCENT_GOLD}
    verdict_icons  = {"GOOD": "✅", "BAD": "❌", "UNCERTAIN": "⚠️"}
    v_color = verdict_colors.get(verdict, TXT_SOFT)
    v_icon  = verdict_icons.get(verdict, "•")

    self._cibi_ci_verdict_lbl.config(
        text=f"{v_icon}  Bank CI Verdict: {verdict}",
        fg=v_color
    )
    self._cibi_ci_summary_lbl.config(text=summary)
    self._cibi_ci_details_lbl.config(text=details)

    self._cibi_proceed_btn.configure(
        state="normal",
        text="✅  PROCEED — Continue to Document Upload"
    )
    self._cibi_stop_btn.configure(state="normal")

    if not proceed:
        self._cibi_override_lbl.config(
            text="⚠ AI recommends stopping. You may still PROCEED manually if you have additional justification.",
            fg=ACCENT_GOLD
        )
        self._cibi_stop_btn.configure(fg_color="#DC2626")
    else:
        self._cibi_override_lbl.config(text="")

    self._cibi_s2_frame.pack(fill="x")
    self._cibi_show_stage("2")
    _cibi_log(
        self,
        "Bank CI Extraction",
        f"Re-check CIC and Bank CI completed. Verdict={verdict}, proceed={proceed}, tier={tier}."
    )

def _cibi_proceed_to_stage3(self):
    self._cibi_stage = "bank_ci_reviewed"
    self._cibi_proceed_btn.configure(state="disabled")
    self._cibi_stop_btn.configure(state="disabled")

    tier = self._cibi_loan_tier
    if tier == "above_100k":
        self._cibi_tier_badge.config(
            text="⬆ ABOVE ₱100,000", bg=NAVY_DEEP)
    elif tier == "below_100k":
        self._cibi_tier_badge.config(
            text="⬇ BELOW ₱100,000", bg=LIME_DARK)
    else:
        self._cibi_tier_badge.config(
            text="Tier Unknown", bg=TXT_SOFT)

    self._cibi_s3_frame.pack(fill="x")
    self._cibi_show_stage("3")

    tier_msg = (
        f"✅  Bank CI: GOOD — Proceeding.\n"
        f"📊  Loan Tier: {'Above ₱100,000' if tier == 'above_100k' else 'Below ₱100,000' if tier == 'below_100k' else 'Unknown'}\n\n"
        f"Please upload the remaining documents:\n"
        f"  • Payslip / Payroll (required)\n"
        f"  • ITR / Income Tax Return (optional)\n"
        f"  • SALN / Net Worth (optional)"
    )
    self._append_chat_bubble(tier_msg, role="system")

def _cibi_stop_process(self):
    self._cibi_stage = "stopped"
    self._cibi_proceed_btn.configure(state="disabled")
    self._cibi_stop_btn.configure(state="disabled")
    self._cibi_override_lbl.config(
        text="🛑  Process stopped. Bank CI record did not meet requirements. "
             "Click  ↺ Reset  to start over.",
        fg=ACCENT_RED
    )
    self._cibi_show_stage("2")
    self._append_chat_bubble(
        "🛑  CIBI Process STOPPED.\n\n"
        "Bank CI record does not meet BSV requirements.\n"
        "The loan application cannot proceed at this time.\n\n"
        "Click  ↺ Reset  in the CIBI Mode tab to start a new application.",
        role="system"
    )

# ══════════════════════════════════════════════════════════════════════
#  CHANGE 2 of 3 — _cibi_extract_all() (parallel extraction)
# ══════════════════════════════════════════════════════════════════════
def _cibi_extract_all(self):
    payslip_path = self._cibi_slots["PAYSLIP"]["path"]
    if not payslip_path:
        self._cibi_s3_progress_lbl.config(
            text="⚠ Payslip is required. Please browse and select it.",
            fg=ACCENT_RED)
        return

    self._cibi_extract_all_btn.configure(state="disabled")
    self._cibi_s3_progress_lbl.config(
        text="⏳ Starting parallel extraction…", fg=ACCENT_GOLD)

    extract_keys = [k for k in ("PAYSLIP", "ITR", "SALN")
                    if self._cibi_slots[k]["path"] is not None]
    for key in extract_keys:
        self._cibi_status_labels[key].config(text="⏳", fg=ACCENT_GOLD)
        self._cibi_slots[key]["text"] = None

    # FIX-2: Collect file paths in order
    file_paths = [self._cibi_slots[k]["path"] for k in extract_keys]

    def worker():
        from extraction import extract_parallel

        # FIX-2: All files extracted concurrently — one thread per file.
        # Progress updates are routed through self.after(0, ...) which
        # is the only thread-safe way to update Tkinter widgets.
        total = len(extract_keys)

        def _combined_cb(pct: int, stage: str = ""):
            self.after(0, self._cibi_s3_progress_lbl.config,
                       {"text": stage or f"Extracting… {pct}%",
                        "fg": ACCENT_GOLD})

        results = extract_parallel(
            file_paths,
            progress_cb=_combined_cb,
            max_workers=min(4, total),
        )

        # Write results back and update status icons
        all_ok = True
        for i, (key, text) in enumerate(zip(extract_keys, results)):
            if text.startswith("[Extraction error") or text.startswith("⚠"):
                self.after(0, self._cibi_status_labels[key].config,
                           {"text": "❌", "fg": ACCENT_RED})
                self.after(0, self._cibi_s3_progress_lbl.config,
                           {"text": f"❌ {key} failed: {text[:80]}",
                            "fg": ACCENT_RED})
                all_ok = False
            else:
                self._cibi_slots[key]["text"] = text
                self.after(0, self._cibi_status_labels[key].config,
                           {"text": "✅", "fg": ACCENT_SUCCESS})

        if not all_ok:
            self.after(0, self._cibi_extract_all_btn.configure,
                       {"state": "normal"})
            return

        self.after(0, self._cibi_all_extracted, extract_keys)

    threading.Thread(target=worker, daemon=True).start()

def _cibi_all_extracted(self, extract_keys: list):
    self._cibi_stage = "docs_extracted"
    self._cibi_extract_all_btn.configure(state="normal")

    extracted_summary = "\n".join(
        f"  {k:8}: {Path(self._cibi_slots[k]['path']).name}  "
        f"({len(self._cibi_slots[k]['text']):,} chars)"
        for k in extract_keys
    )
    self._cibi_s3_progress_lbl.config(
        text=f"✅ All documents extracted.\n{extracted_summary}",
        fg=ACCENT_SUCCESS
    )

    self._cibi_s4_frame.pack(fill="x")
    self._cibi_show_stage("4")
    self._cibi_refresh_stage_buttons()

    self._append_chat_bubble(
        "📋  All documents extracted:\n"
        + extracted_summary
        + "\n\n✅  Select your CIBI Excel template, then click  Populate → CIBI Excel.",
        role="system"
    )
    _cibi_log(
        self,
        "Extracted CIBI Documents",
        f"Extract all documents completed for: {', '.join(extract_keys)}."
    )

# ── STAGE 4 ACTIONS ───────────────────────────────────────────────────
def _cibi_pick_template(self):
    path = filedialog.askopenfilename(
        title="Select CIBI Excel Template",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if path:
        self._cibi_template_path = path
        name  = Path(path).name
        short = name if len(name) <= 28 else name[:25] + "…"
        self._cibi_tmpl_lbl.config(text=short, fg=LIME_DARK)
    else:
        self._cibi_template_path = None
        self._cibi_tmpl_lbl.config(text="No CIBI template selected", fg=TXT_MUTED)
    self._cibi_refresh_stage_buttons()

def _cibi_populate(self):
    if not self._cibi_template_path:
        self._cibi_s4_progress_lbl.config(
            text="⚠ No CIBI template selected.", fg=ACCENT_RED)
        return
    if not self._cibi_slots["PAYSLIP"]["text"]:
        self._cibi_s4_progress_lbl.config(
            text="⚠ Payslip not yet extracted. Please extract all documents first.",
            fg=ACCENT_RED)
        return

    self._cibi_populate_btn.configure(state="disabled")
    self._cibi_s4_progress_lbl.config(text="⏳ Populating CIBI template…", fg=ACCENT_GOLD)
    self._switch_tab("cibi")
    self._show_loader(True, "Populating CIBI Excel template…")

    cic_text     = self._cibi_slots["CIC"]["text"]     or ""
    payslip_text = self._cibi_slots["PAYSLIP"]["text"] or ""
    saln_text    = self._cibi_slots["SALN"]["text"]    or ""
    itr_text     = self._cibi_slots["ITR"]["text"]     or ""
    bank_ci_text = self._cibi_slots["BANK_CI"]["text"] or ""

    cic_path     = self._cibi_slots["CIC"]["path"]
    payslip_path = self._cibi_slots["PAYSLIP"]["path"]
    stem = Path(cic_path or payslip_path).stem if (cic_path or payslip_path) else "CIBI"

    def worker():
        try:
            def _cb(pct, msg=""):
                self.after(0, self._cibi_s4_progress_lbl.config,
                           {"text": msg or f"Populating… {pct}%", "fg": ACCENT_GOLD})
                self.after(0, self._set_progress, pct, msg)

            out_path = populate_cibi_form(
                template_path = self._cibi_template_path,
                api_key       = GEMINI_API_KEY,
                cic_text      = cic_text,
                payslip_text  = payslip_text,
                saln_text     = saln_text,
                itr_text      = itr_text,
                output_stem   = stem,
                progress_cb   = _cb,
            )
            self.after(0, self._cibi_finish_populate, out_path)
        except Exception as e:
            import traceback
            self.after(0, self._cibi_finish_error,
                       f"Population failed:\n{e}\n\n{traceback.format_exc()}")

    threading.Thread(target=worker, daemon=True).start()

def _cibi_finish_populate(self, out_path: Path):
    self._show_loader(False)
    self._cibi_populate_btn.configure(state="normal")
    self._cibi_s4_progress_lbl.config(
        text=f"✅ Done! Saved: {out_path.name}", fg=ACCENT_SUCCESS)
    self._status_lbl.config(text="●  CIBI Done", fg=LIME_DARK)
    self._last_cibi_path         = out_path
    self._summary_cibi_populated = True
    self._summary_cibi_excel_path = str(out_path)
    self._offer_cibi_analysis(out_path)

    # ── offer auto-analysis ───────────────────────────────────────────
    self._offer_cibi_analysis(out_path)

    self._cibi_result_lbl.config(
        text=f"✅  CIBI Excel populated!\n\n📁  {out_path}")
    self._cibi_result_frame.pack(fill="x", padx=24, pady=(8, 16))
    self._switch_tab("cibi")

    slots_used = [k for k, s in self._cibi_slots.items() if s["path"]]
    self._append_chat_bubble(
        f"✅  CIBI Excel populated successfully!\n\n"
        f"📁  {out_path}\n\n"
        f"Documents used:\n"
        + "\n".join(f"  {k:8}: {Path(self._cibi_slots[k]['path']).name}"
                    for k in slots_used)
        + "\n\nFile saved to Desktop → DocExtract_Files folder.",
        role="system"
    )
    try:
        folder = str(out_path.parent)
        if sys.platform == "win32":    os.startfile(folder)
        elif sys.platform == "darwin": os.system(f'open "{folder}"')
        else:                          os.system(f'xdg-open "{folder}"')
    except Exception:
        pass
    _cibi_log(self, "CIBI Population", f"Populate CIBI Excel completed: {out_path.name}")

def _cibi_finish_error(self, msg: str):
    self._show_loader(False)
    self._cibi_populate_btn.configure(state="normal")
    self._cibi_s4_progress_lbl.config(
        text=f"❌ Population failed: {msg[:120]}", fg=ACCENT_RED)
    self._status_lbl.config(text="●  Error", fg=ACCENT_RED)
    self._cibi_result_lbl.config(text=f"❌  Population failed\n\n{msg[:400]}")
    self._cibi_result_frame.pack(fill="x", padx=24, pady=(8, 16))
    self._switch_tab("cibi")

# ── CIBI UTILITY METHODS ──────────────────────────────────────────────
def _open_cibi_output_folder(self):
    from pathlib import Path as _Path
    desktop = _Path.home() / "Desktop"
    folder  = (desktop if desktop.exists() else _Path.home()) / "DocExtract_Files"
    folder.mkdir(parents=True, exist_ok=True)
    try:
        if sys.platform == "win32":    os.startfile(str(folder))
        elif sys.platform == "darwin": os.system(f'open "{folder}"')
        else:                          os.system(f'xdg-open "{folder}"')
    except Exception:
        pass

def _open_last_cibi_file(self):
    if self._last_cibi_path and Path(self._last_cibi_path).exists():
        try:
            if sys.platform == "win32":    os.startfile(str(self._last_cibi_path))
            elif sys.platform == "darwin": os.system(f'open "{self._last_cibi_path}"')
            else:                          os.system(f'xdg-open "{self._last_cibi_path}"')
        except Exception:
            pass

# ── NEW METHOD FOR OFFERING CIBI ANALYSIS AFTER POPULATION ─────────────
def _offer_cibi_analysis(self, excel_path):
    """Show a chat prompt offering to run CIBI analysis on the freshly populated file."""
    def _run():
        self._switch_tab("analysis")
        self._show_loader(True, "Reading CIBI Excel…")
        self._set_progress(0, "Opening workbook…")
        def worker():
            self.after(0, self._set_progress, 25, "Parsing cash-flow…")
            try:
                from cibi_analysis import run_cibi_analysis
                result = run_cibi_analysis(str(excel_path), GEMINI_API_KEY)
            except Exception as e:
                result = f"⚠ Analysis error: {e}"
            self.after(0, self._set_progress, 95, "Formatting…")
            self.after(0, self._finish_cibi_excel_analysis, result, str(excel_path))
        threading.Thread(target=worker, daemon=True).start()

    self._append_chat_bubble(
        f"✅  CIBI Excel is ready.\n\n"
        f"📊  Click below to run the full CIBI Analysis now,\n"
        f"     or switch to the  🏦 CIBI Analysis  tab manually.",
        role="system"
    )
    # inject a quick-action button into the chat
    btn_frame = tk.Frame(self._chat_inner, bg=OFF_WHITE)
    btn_frame.pack(fill="x", padx=14, pady=(0, 10))
    ctk.CTkButton(
        btn_frame,
        text="📊  Run CIBI Analysis Now",
        command=_run,
        height=36, corner_radius=8,
        fg_color=LIME_DARK, hover_color=LIME_MID, text_color=WHITE,
        font=FF(9, "bold"),
        border_width=0
    ).pack(side="left")
    self._chat_canvas.update_idletasks()
    self._chat_canvas.configure(scrollregion=self._chat_canvas.bbox("all"))
    self._chat_canvas.yview_moveto(1.0)

# ── NEW METHOD FOR ANALYZING FROM EXCEL ───────────────────────────────
def _start_cibi_analysis_from_excel(self):
    """Browse for a populated CIBI Excel file and run full CIBI analysis."""
    from tkinter import filedialog
    path = filedialog.askopenfilename(
        title="Select Populated CIBI Excel File",
        filetypes=[
            ("Excel files", "*.xlsx *.xls"),
            ("All files",   "*.*"),
        ]
    )
    if not path:
        return

    self._switch_tab("analysis")
    self._show_loader(True, "Reading CIBI Excel…")
    self._set_progress(0, "Opening workbook…")

    def worker():
        self.after(0, self._set_progress, 15, "Parsing cash-flow sheet…")
        try:
            from cibi_analysis import run_cibi_analysis
            self.after(0, self._set_progress, 35, "Parsing CIBI sheet…")
            result = run_cibi_analysis(path, GEMINI_API_KEY)
            self.after(0, self._set_progress, 85, "Formatting report…")
        except Exception as e:
            import traceback
            result = (
                f"⚠ CIBI Excel analysis failed:\n"
                f"{type(e).__name__}: {e}\n\n"
                f"{traceback.format_exc()}"
            )
        self.after(0, self._set_progress, 100, "Done")
        self.after(0, self._finish_cibi_excel_analysis, result, path)

    threading.Thread(target=worker, daemon=True).start()

# ══════════════════════════════════════════════════════════════════════
#  UPDATED _finish_cibi_excel_analysis (Change 3)
# ══════════════════════════════════════════════════════════════════════
def _finish_cibi_excel_analysis(self, result: str, path: str):
    from pathlib import Path
    self._show_loader(False)
    fname = Path(path).name
    header = (
        f"Source : {fname}\n"
        f"Type   : CIBI Excel Workbook\n"
        f"{'─' * 58}\n\n"
    )
    self._write_analysis(header + result, TXT_NAVY)
    self._status_lbl.config(text="●  CIBI Analysis Done", fg=LIME_DARK)
    self._summary_cibi_analysis_text = result
    self._summary_cibi_excel_path    = path
    self._populate_summary(result)
    self._append_chat_bubble(
        f"📊  CIBI Excel Analysis complete: {fname}\n\n"
        f"Results are now shown in the  🏦 CIBI Analysis  tab.\n"
        f"You can also ask me questions about this analysis here.",
        role="system"
    )



# ── attach ────────────────────────────────────────────────────────────────────
def attach(cls):
    cls._build_cibi_output_panel       = _build_cibi_output_panel
    cls._cibi_show_stage               = _cibi_show_stage
    cls._build_cibi_stage1             = _build_cibi_stage1
    cls._build_cibi_stage2             = _build_cibi_stage2
    cls._build_cibi_stage3             = _build_cibi_stage3
    cls._build_cibi_stage4             = _build_cibi_stage4
    cls._step_number_lbl               = _step_number_lbl
    cls._cibi_browse_slot              = _cibi_browse_slot
    cls._cibi_refresh_stage_buttons    = _cibi_refresh_stage_buttons
    cls._cibi_full_reset               = _cibi_full_reset
    cls._cibi_run_stage1               = _cibi_run_stage1
    cls._cibi_finish_stage1            = _cibi_finish_stage1
    cls._cibi_proceed_to_stage3        = _cibi_proceed_to_stage3
    cls._cibi_stop_process             = _cibi_stop_process
    cls._cibi_extract_all              = _cibi_extract_all
    cls._cibi_all_extracted            = _cibi_all_extracted
    cls._cibi_pick_template            = _cibi_pick_template
    cls._cibi_populate                 = _cibi_populate
    cls._cibi_finish_populate          = _cibi_finish_populate
    cls._cibi_finish_error             = _cibi_finish_error
    cls._open_cibi_output_folder       = _open_cibi_output_folder
    cls._open_last_cibi_file           = _open_last_cibi_file
    cls._offer_cibi_analysis           = _offer_cibi_analysis
    cls._start_cibi_analysis_from_excel = _start_cibi_analysis_from_excel
    cls._finish_cibi_excel_analysis    = _finish_cibi_excel_analysis
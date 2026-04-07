"""
ui_panels.py — DocExtract Pro
================================
UI building methods attached to DocExtractorApp:
  _build_topbar, _build_left, _build_right,
  _build_classifier_panel, show_classified_result, all _clf_* methods,
  _configure_analysis_tags, _write, _write_analysis, _insert_*,
  _put_placeholder, _switch_tab, _show_loader, _set_progress,
  gradient helpers, _file_icon_for, search methods, _copy
"""
import re
import tkinter as tk
import customtkinter as ctk
from pathlib import Path
from widgets import GradientCanvas, Spinner
from app_constants import *
from doc_classifier_tab import (
    classify_document, extract_fields, extract_financial_figures, DOC_TYPES
)
from lookup_tab import attach as _attach_lookup
from summary_tab import attach as _attach_summary, _refresh_summary
from lu_analysis_tab import attach as _attach_lu_analysis
from general_lookup import attach as _attach_general_lookup
from general_summary import attach as _attach_general_summary

def _build_topbar(self):
    bar = tk.Frame(self, bg=NAVY_DEEP, height=52)
    bar.pack(fill="x")
    bar.pack_propagate(False)
    bar.bind("<ButtonPress-1>", self._drag_start)
    bar.bind("<B1-Motion>",     self._drag_move)

    ctrl_frame = tk.Frame(bar, bg=NAVY_DEEP)
    ctrl_frame.pack(side="right", fill="y", padx=8)

    close_btn = tk.Label(ctrl_frame, text="✕", font=F(10, "bold"),
                          fg="#9AAACE", bg=NAVY_DEEP, cursor="hand2",
                          width=3, anchor="center")
    close_btn.bind("<Enter>",    lambda e: close_btn.config(fg=WHITE, bg=ACCENT_RED))
    close_btn.bind("<Leave>",    lambda e: close_btn.config(fg="#9AAACE", bg=NAVY_DEEP))
    close_btn.bind("<Button-1>", lambda e: self.destroy())
    close_btn.pack(side="right", padx=(4, 2), fill="y")

    min_btn = tk.Label(ctrl_frame, text="─", font=F(10, "bold"),
                        fg="#9AAACE", bg=NAVY_DEEP, cursor="hand2",
                        width=3, anchor="center")
    min_btn.bind("<Enter>",    lambda e: min_btn.config(fg=WHITE, bg=SIDEBAR_HVR))
    min_btn.bind("<Leave>",    lambda e: min_btn.config(fg="#9AAACE", bg=NAVY_DEEP))
    min_btn.bind("<Button-1>", lambda e: self._do_minimize())
    min_btn.pack(side="right", padx=(4, 2), fill="y")

    brand = tk.Frame(bar, bg=NAVY_DEEP)
    brand.pack(side="left", fill="y", padx=(16, 0))

    mark = tk.Canvas(brand, width=28, height=28, bg=NAVY_DEEP, highlightthickness=0)
    mark.pack(side="left", padx=(0, 10), pady=12)
    mark.create_rectangle(0, 8, 17, 28, fill=NAVY_LIGHT, outline="")
    mark.create_rectangle(11, 0, 28, 20, fill=LIME_BRIGHT, outline="")
    mark.create_rectangle(13, 10, 15, 12, fill=NAVY_DEEP, outline="")

    tk.Label(brand, text="DocExtract Pro", font=F(13, "bold"),
             fg=WHITE, bg=NAVY_DEEP).pack(side="left")
    tk.Label(brand, text=" — Banco San Vicente",
             font=F(10), fg=TXT_SOFT, bg=NAVY_DEEP).pack(side="left")

    self._topbar_status = tk.Label(bar, text="● Ready",
                                    font=F(8, "bold"), fg=LIME_BRIGHT,
                                    bg=SIDEBAR_ITEM, padx=12, pady=4,
                                    relief="flat")
    self._topbar_status.pack(side="right", padx=(0, 20), pady=14)

# ── LEFT PANEL ────────────────────────────────────────────────────────
def _build_left(self, p):
    wrap = tk.Frame(p, bg=SIDEBAR_BG)
    wrap.pack(fill="both", expand=True, padx=0, pady=0)

    ident = tk.Frame(wrap, bg=NAVY_DEEP)
    ident.pack(fill="x")
    tk.Frame(ident, bg=NAVY_DEEP, height=16).pack()

    logo_loaded = False
    if LOGO_PATH.exists():
        try:
            from PIL import Image
            img    = Image.open(LOGO_PATH).convert("RGBA")
            bg_pil = Image.new(
                "RGBA", img.size,
                tuple(int(NAVY_DEEP[i:i+2], 16) for i in (1, 3, 5)) + (255,)
            )
            bg_pil.paste(img, mask=img.split()[3])
            bg_pil = bg_pil.convert("RGB")
            bg_pil.thumbnail((180, 44), Image.LANCZOS)
            self._logo_img = ctk.CTkImage(
                light_image=bg_pil, dark_image=bg_pil,
                size=(bg_pil.width, bg_pil.height)
            )
            lf = ctk.CTkFrame(ident, fg_color="transparent")
            lf.pack(padx=18)
            ctk.CTkLabel(lf, image=self._logo_img, text="",
                        fg_color="transparent").pack()
            logo_loaded = True
        except Exception:
            pass

    if not logo_loaded:
        mark_frame = tk.Frame(ident, bg=NAVY_DEEP)
        mark_frame.pack(padx=18, anchor="w")
        mark = tk.Canvas(mark_frame, width=26, height=30,
                        bg=NAVY_DEEP, highlightthickness=0)
        mark.pack(side="left", padx=(0, 8))
        mark.create_rectangle(0, 8, 15, 30, fill=NAVY_LIGHT,  outline="")
        mark.create_rectangle(11, 0, 26, 20, fill=LIME_BRIGHT, outline="")
        tk.Label(mark_frame, text="Banco San Vicente",
                font=F(11, "bold"), fg=WHITE, bg=NAVY_DEEP).pack(side="left")

    tk.Frame(ident, bg=NAVY_DEEP, height=10).pack()

    ul = tk.Canvas(ident, height=2, bg=NAVY_DEEP, highlightthickness=0)
    ul.pack(fill="x")
    ul.bind("<Configure>", lambda e, c=ul: self._hbar(c, LIME_BRIGHT, LIME_DARK, 60))

    tk.Frame(ident, bg=NAVY_DEEP, height=16).pack()

    sub = tk.Frame(ident, bg=NAVY_DEEP)
    sub.pack(fill="x", padx=18)
    tk.Label(sub, text="BSV AI-OCR", font=F(18, "bold"),
             fg=WHITE, bg=NAVY_DEEP, anchor="w").pack(anchor="w")
    tk.Label(sub, text="Document & Image Extraction",
             font=F(9), fg=TXT_SOFT, bg=NAVY_DEEP, anchor="w").pack(anchor="w", pady=(2, 0))
    tk.Frame(ident, bg=NAVY_DEEP, height=20).pack()

    sec_wrap = tk.Frame(wrap, bg=SIDEBAR_BG)
    sec_wrap.pack(fill="x", padx=16, pady=(14, 0))

    self._sidebar_sec(sec_wrap, "UPLOAD FILES")

    drop = tk.Frame(sec_wrap, bg=SIDEBAR_ITEM,
                    highlightbackground=LIME_DARK, highlightthickness=1, height=64)
    drop.pack(fill="x", pady=(4, 0))
    drop.pack_propagate(False)

    inner = tk.Frame(drop, bg=SIDEBAR_ITEM)
    inner.place(relx=0.5, rely=0.5, anchor="center")
    self._icon_lbl = tk.Label(inner, text="📁",
                            font=("Segoe UI Emoji", 18),
                            fg=LIME_BRIGHT, bg=SIDEBAR_ITEM)
    self._icon_lbl.pack()
    self._filename_lbl = tk.Label(inner, text="No files selected",
                                font=F(8), fg=TXT_SOFT, bg=SIDEBAR_ITEM,
                                wraplength=230, justify="center")
    self._filename_lbl.pack(pady=(1, 0))

    tk.Frame(sec_wrap, bg=SIDEBAR_BG, height=6).pack()

    list_frame = tk.Frame(sec_wrap, bg=SIDEBAR_ITEM,
                        highlightbackground="#1E3A5F", highlightthickness=1)
    list_frame.pack(fill="x")
    list_header = tk.Frame(list_frame, bg=SIDEBAR_ITEM)
    list_header.pack(fill="x", padx=10, pady=(6, 2))
    tk.Label(list_header, text="SELECTED FILES",
            font=F(7, "bold"), fg=TXT_SOFT, bg=SIDEBAR_ITEM).pack(side="left")
    self._clear_btn = tk.Label(list_header, text="✕ Clear",
                                font=F(7, "bold"), fg=ACCENT_RED, bg=SIDEBAR_ITEM,
                                cursor="hand2")
    self._clear_btn.pack(side="right")
    self._clear_btn.bind("<Button-1>", lambda e: self._clear_files())

    list_scroll_frame = tk.Frame(list_frame, bg=SIDEBAR_ITEM, height=56)
    list_scroll_frame.pack(fill="x", padx=6, pady=(0, 4))
    list_scroll_frame.pack_propagate(False)
    list_sb = tk.Scrollbar(list_scroll_frame, relief="flat",
                            troughcolor=SIDEBAR_ITEM, bg=SIDEBAR_HVR, width=5, bd=0)
    list_sb.pack(side="right", fill="y")
    self._file_listbox = tk.Listbox(
        list_scroll_frame,
        font=F(8), fg="#C5D8F5", bg=SIDEBAR_ITEM,
        relief="flat", bd=0,
        selectbackground=SIDEBAR_HVR, selectforeground=WHITE,
        activestyle="none",
        yscrollcommand=list_sb.set,
        height=3,
    )
    self._file_listbox.pack(side="left", fill="both", expand=True)
    list_sb.config(command=self._file_listbox.yview)
    self._file_listbox.bind("<Button-3>", self._remove_selected_file)

    tk.Frame(sec_wrap, bg=SIDEBAR_BG, height=10).pack()

    self._browse_btn = ctk.CTkButton(
        sec_wrap, text="Browse File(s)", command=self._browse,
        height=38, corner_radius=8,
        fg_color=LIME_MID, hover_color=LIME_BRIGHT,
        text_color=TXT_ON_LIME,
        font=FF(10, "bold"),
        border_width=0
    )
    self._browse_btn.pack(fill="x")

    tk.Frame(sec_wrap, bg=SIDEBAR_BG, height=5).pack()

    self._add_btn = ctk.CTkButton(
        sec_wrap, text="Add More Files", command=self._browse_add,
        height=34, corner_radius=8,
        fg_color=SIDEBAR_ITEM, hover_color=SIDEBAR_HVR,
        text_color="#C5D8F5",
        font=FF(9),
        border_width=1, border_color="#1E3A5F",
        state="disabled"
    )
    self._add_btn.pack(fill="x")

    tk.Frame(sec_wrap, bg=SIDEBAR_BG, height=5).pack()

    self._ext_btn = ctk.CTkButton(
        sec_wrap, text="⚡  Extract Text", command=self._start_extraction,
        height=38, corner_radius=8,
        fg_color=NAVY_LIGHT, hover_color=NAVY_PALE,
        text_color=WHITE,
        font=FF(10, "bold"),
        state="disabled", border_width=0
    )
    self._ext_btn.pack(fill="x")

    tk.Frame(sec_wrap, bg=SIDEBAR_BG, height=5).pack()

    self._analyze_btn = ctk.CTkButton(
        sec_wrap, text="🏦  Analyze CIBI", command=self._start_analysis,
        height=38, corner_radius=8,
        fg_color=SIDEBAR_ITEM, hover_color=SIDEBAR_HVR,
        text_color=LIME_BRIGHT,
        font=FF(10, "bold"),
        state="disabled",
        border_width=1, border_color=LIME_DARK
    )
    self._analyze_btn.pack(fill="x")

    tk.Frame(sec_wrap, bg=SIDEBAR_BG, height=5).pack()

    self._analyze_excel_btn = ctk.CTkButton(
        sec_wrap, text="📊  Analyze from Excel",
        command=self._start_cibi_analysis_from_excel,
        height=38, corner_radius=8,
        fg_color=SIDEBAR_ITEM, hover_color=SIDEBAR_HVR,
        text_color=LIME_PALE,
        font=FF(10, "bold"),
        border_width=1, border_color=LIME_DARK
    )
    self._analyze_excel_btn.pack(fill="x")

    tk.Label(
        sec_wrap,
        text="Browse a populated CIBI Excel\nfile for direct analysis",
        font=F(7), fg="#2D4F7A", bg=SIDEBAR_BG,
        justify="center"
    ).pack(fill="x", pady=(3, 0))

    tk.Frame(sec_wrap, bg="#1E3A5F", height=1).pack(fill="x", pady=(16, 10))

    self._sidebar_sec(sec_wrap, "SUPPORTED FORMATS")

    formats = [
        ("📄", "PDF (.pdf)"),
        ("📝", "Word (.docx)"),
        ("📊", "Excel (.xlsx)"),
        ("📃", "Text / CSV / MD"),
        ("🖼", "Images (.png .jpg .bmp .tiff .webp)"),
    ]
    for icon, label in formats:
        row = tk.Frame(sec_wrap, bg=SIDEBAR_BG)
        row.pack(fill="x", pady=3)
        tk.Label(row, text=icon, font=("Segoe UI Emoji", 10),
                fg=LIME_MID, bg=SIDEBAR_BG).pack(side="left", padx=(4, 0))
        tk.Label(row, text=label, font=F(8), fg=TXT_SOFT,
                bg=SIDEBAR_BG).pack(side="left", padx=(8, 0))

    tk.Frame(sec_wrap, bg="#1E3A5F", height=1).pack(fill="x", pady=(16, 10))

    ver_frame = tk.Frame(sec_wrap, bg=SIDEBAR_BG)
    ver_frame.pack(fill="x", pady=(4, 16))
    tk.Label(ver_frame, text="Gemini 2.5 Flash · PaddleOCR",
             font=F(7), fg="#2D4F7A", bg=SIDEBAR_BG, anchor="center").pack(fill="x")
    tk.Label(ver_frame, text="BSV AI-OCR v2.0",
             font=F(7), fg="#2D4F7A", bg=SIDEBAR_BG, anchor="center").pack(fill="x")

# ── RIGHT PANEL ───────────────────────────────────────────────────────
def _build_right(self, p):
    hdr = tk.Frame(p, bg=WHITE, height=64)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)

    tk.Frame(p, bg=BORDER_LIGHT, height=1).pack(fill="x")

    hdr_inner = tk.Frame(hdr, bg=WHITE)
    hdr_inner.pack(side="left", fill="y", padx=(28, 0))

    tk.Label(hdr_inner, text="Extracted Content",
              font=F(17, "bold"), fg=NAVY_DEEP, bg=WHITE).pack(side="left", anchor="center")

    status_pill = tk.Frame(hdr, bg=LIME_MIST,
                            highlightbackground=LIME_MID, highlightthickness=1)
    status_pill.pack(side="left", padx=(16, 0), pady=18)
    self._status_lbl = tk.Label(status_pill, text="●  Ready",
                                 font=F(8, "bold"), fg=LIME_DARK,
                                 bg=LIME_MIST, padx=12, pady=4)
    self._status_lbl.pack()

    self._copy_btn = ctk.CTkButton(
        hdr, text="⎘  Copy", command=self._copy,
        width=100, height=32, corner_radius=6,
        fg_color=NAVY_MIST, hover_color=NAVY_GHOST, text_color=NAVY_MID,
        font=FF(9, "bold"),
        border_width=1, border_color=BORDER_MID
    )
    self._copy_btn.pack(side="right", padx=(0, 24))

    # ── Scrollable tab row with arrow indicators ─────────────────────────────
    tab_outer = tk.Frame(p, bg=WHITE)
    tab_outer.pack(fill="x", padx=24, pady=(10, 0))
    
    # Create a frame to hold the canvas and scroll buttons
    tab_container = tk.Frame(tab_outer, bg=WHITE)
    tab_container.pack(fill="x")

    # Canvas for tabs — expands to fill all available space
    tab_canvas = tk.Canvas(tab_container, bg=WHITE, highlightthickness=0, height=36)
    tab_canvas.pack(side="left", fill="x", expand=True)

    tab_hsb = tk.Scrollbar(tab_container, orient="horizontal",
                            command=tab_canvas.xview,
                            relief="flat", troughcolor=WHITE,
                            bg=BORDER_LIGHT, width=4, bd=0)
    tab_hsb.pack(side="bottom", fill="x", pady=(2, 0))

    # Both arrows grouped together on the right, always occupying space
    arrow_frame = tk.Frame(tab_container, bg=WHITE)
    arrow_frame.pack(side="left", padx=(4, 0))

    self._tab_left_arrow = tk.Label(
        arrow_frame, text="◀", font=("Segoe UI", 11, "bold"),
        fg=WHITE, bg=WHITE, cursor="arrow",
        width=2, anchor="center"
    )
    self._tab_left_arrow.pack(side="left")
    self._tab_left_arrow.bind("<Button-1>", lambda e: self._scroll_tabs(-1))

    self._tab_right_arrow = tk.Label(
        arrow_frame, text="▶", font=("Segoe UI", 11, "bold"),
        fg=WHITE, bg=WHITE, cursor="arrow",
        width=2, anchor="center"
    )
    self._tab_right_arrow.pack(side="left")
    self._tab_right_arrow.bind("<Button-1>", lambda e: self._scroll_tabs(1))

    tab_row = tk.Frame(tab_canvas, bg=WHITE)
    _tab_win = tab_canvas.create_window((0, 0), window=tab_row, anchor="nw")
    
    def _update_scroll_arrows():
        """Toggle arrow visibility via colour — never pack/pack_forget"""
        if hasattr(self, '_tab_left_arrow') and hasattr(self, '_tab_right_arrow'):
            xview = tab_canvas.xview()
            can_scroll_left  = xview[0] > 0.001
            can_scroll_right = xview[1] < 0.999

            self._tab_left_arrow.config(
                fg=LIME_DARK if can_scroll_left  else WHITE,
                cursor="hand2" if can_scroll_left  else "arrow"
            )
            self._tab_right_arrow.config(
                fg=LIME_DARK if can_scroll_right else WHITE,
                cursor="hand2" if can_scroll_right else "arrow"
            )
    
    # Update scroll arrows when tab row size changes
    def _on_tab_row_configure(e):
        tab_canvas.configure(scrollregion=tab_canvas.bbox("all"))
        _update_scroll_arrows()
    
    tab_row.bind("<Configure>", _on_tab_row_configure)
    
    # Update scroll arrows when canvas is scrolled
    def _on_canvas_scroll(*args):
        tab_hsb.set(*args)
        _update_scroll_arrows()
    
    tab_canvas.configure(xscrollcommand=_on_canvas_scroll)
    
    # Also update when canvas is resized
    tab_canvas.bind("<Configure>", lambda e: (tab_canvas.itemconfig(_tab_win, height=e.height), _update_scroll_arrows()))

    self._active_tab = tk.StringVar(value="extract")

    def _tab_style(btn, active):
        if active:
            btn.configure(fg_color=NAVY_DEEP, text_color=WHITE,
                           hover_color=NAVY_MID, border_width=0)
        else:
            btn.configure(fg_color=WHITE, text_color=TXT_SOFT,
                           hover_color=NAVY_MIST, border_width=1,
                           border_color=BORDER_LIGHT)

    self._tab_cibi_btn = ctk.CTkButton(
        tab_row, text="📋  CIBI Mode", width=130, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("cibi")
    )
    self._tab_cibi_btn.pack(side="left", padx=(0, 4))

    self._tab_extract_btn = ctk.CTkButton(
        tab_row, text="📄  Extracted", width=130, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("extract")
    )
    self._tab_extract_btn.pack(side="left", padx=(0, 4))

    self._tab_analysis_btn = ctk.CTkButton(
        tab_row, text="🏦  CIBI Analysis", width=140, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("analysis")
    )
    self._tab_analysis_btn.pack(side="left", padx=(0, 4))

    self._tab_summary_btn = ctk.CTkButton(
        tab_row, text="📊  Summary", width=120, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("summary")
    )
    self._tab_summary_btn.pack(side="left", padx=(0, 4))

    self._tab_aiprompt_btn = ctk.CTkButton(
        tab_row, text="🤖  AI Chat", width=110, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("aiprompt")
    )
    self._tab_aiprompt_btn.pack(side="left", padx=(0, 4))

    self._tab_samples_btn = ctk.CTkButton(
        tab_row, text="🗂  Samples", width=110, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("samples")
    )
    self._tab_samples_btn.pack(side="left", padx=(0, 4))

    self._tab_lookup_btn = ctk.CTkButton(
        tab_row, text="🔎  Look-Up", width=120, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("lookup")
    )
    self._tab_lookup_btn.pack(side="left", padx=(0, 4))

    self._tab_lookup_summary_btn = ctk.CTkButton(
        tab_row, text="📋  LU Summary", width=130, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("lookup_summary")
    )
    self._tab_lookup_summary_btn.pack(side="left", padx=(0, 4))

    self._tab_lu_analysis_btn = ctk.CTkButton(
        tab_row, text="📈  LU Analysis", width=130, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("lu_analysis")
    )
    self._tab_lu_analysis_btn.pack(side="left", padx=(0, 4))

    # ── NEW TABS ───────────────────────────────────────────────────────
    self._tab_general_lookup_btn = ctk.CTkButton(
        tab_row, text="📂  General Look-Up", width=150, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("general_lookup")
    )
    self._tab_general_lookup_btn.pack(side="left", padx=(0, 4))

    self._tab_general_summary_btn = ctk.CTkButton(
        tab_row, text="📋  General Summary", width=155, height=30,
        corner_radius=6, font=FF(9, "bold"),
        command=lambda: self._switch_tab("general_summary")
    )
    self._tab_general_summary_btn.pack(side="left", padx=(0, 4))
    # ──────────────────────────────────────────────────────────────────

    _tab_style(self._tab_extract_btn,          False)
    _tab_style(self._tab_cibi_btn,             True)
    _tab_style(self._tab_analysis_btn,         False)
    _tab_style(self._tab_summary_btn,          False)
    _tab_style(self._tab_aiprompt_btn,         False)
    _tab_style(self._tab_samples_btn,          False)
    _tab_style(self._tab_lookup_btn,           False)
    _tab_style(self._tab_lookup_summary_btn,   False)
    _tab_style(self._tab_lu_analysis_btn,      False)
    _tab_style(self._tab_general_lookup_btn,   False)
    _tab_style(self._tab_general_summary_btn,  False)
    self._tab_style_fn = _tab_style
    
    # Store canvas reference for scrolling
    self._tab_canvas = tab_canvas
    self._tab_hsb = tab_hsb

    pipeline_row = tk.Frame(p, bg=OFF_WHITE)
    pipeline_row.pack(fill="x", padx=24, pady=(8, 0))
    tk.Label(
        pipeline_row,
        text="Pipeline:  PaddleOCR  →  Gemini 2.5 Flash VLM  →  Confidence Scoring  →  CIBI Credit Analysis",
        font=F(8), fg=TXT_MUTED, bg=OFF_WHITE
    ).pack(anchor="w")

    search_row = tk.Frame(p, bg=OFF_WHITE)
    search_row.pack(fill="x", padx=24, pady=(8, 10))

    search_wrap = tk.Frame(search_row, bg=WHITE,
                            highlightbackground=BORDER_MID, highlightthickness=1)
    search_wrap.pack(side="left", fill="x", expand=True)

    tk.Label(search_wrap, text="🔍", font=("Segoe UI Emoji", 10),
              bg=WHITE, fg=NAVY_PALE).pack(side="left", padx=(10, 2))

    self._search_var = tk.StringVar()
    self._search_entry = tk.Entry(
        search_wrap, textvariable=self._search_var,
        font=F(10), fg=TXT_NAVY, bg=WHITE,
        relief="flat", bd=0, insertbackground=NAVY_MID, width=28
    )
    self._search_entry.pack(side="left", fill="x", expand=True, pady=7)
    self._search_entry.bind("<Return>",   lambda e: self._do_search())
    self._search_entry.bind("<KP_Enter>", lambda e: self._do_search())
    self._search_entry.bind("<Escape>",   lambda e: self._clear_search())
    self._search_var.trace_add("write",   lambda *a: self._do_search())

    self._match_lbl = tk.Label(search_wrap, text="",
                                font=F(8), fg=TXT_SOFT, bg=WHITE, padx=8)
    self._match_lbl.pack(side="left")

    nav_frame = tk.Frame(search_row, bg=OFF_WHITE)
    nav_frame.pack(side="left", padx=(6, 0))

    for symbol, cmd in [("▲", self._search_prev), ("▼", self._search_next)]:
        ctk.CTkButton(
            nav_frame, text=symbol, width=30, height=30, corner_radius=6,
            fg_color=WHITE, hover_color=NAVY_MIST, text_color=NAVY_MID,
            font=FF(9, "bold"),
            border_width=1, border_color=BORDER_MID,
            command=cmd
        ).pack(side="left", padx=(0, 4))

    ctk.CTkButton(
        nav_frame, text="✕", width=30, height=30, corner_radius=6,
        fg_color=WHITE, hover_color="#FFE8E8", text_color=ACCENT_RED,
        font=FF(9, "bold"),
        border_width=1, border_color=BORDER_MID,
        command=self._clear_search
    ).pack(side="left")

    self._prev_btn = nav_frame.winfo_children()[0]
    self._next_btn = nav_frame.winfo_children()[1]

    card_outer = tk.Frame(p, bg=BORDER_LIGHT, padx=1, pady=1)
    card_outer.pack(fill="both", expand=True, padx=24, pady=(0, 20))
    card = tk.Frame(card_outer, bg=CARD_WHITE)
    card.pack(fill="both", expand=True)

    top_acc = tk.Canvas(card, height=4, bg=CARD_WHITE, highlightthickness=0)
    top_acc.pack(fill="x")
    top_acc.bind("<Configure>",
                 lambda e, c=top_acc: self._hbar(c, NAVY_DEEP, LIME_BRIGHT, 80))

    self._loader_frame = tk.Frame(card, bg=CARD_WHITE)
    tk.Frame(self._loader_frame, bg=CARD_WHITE).pack(expand=True, fill="both")
    center = tk.Frame(self._loader_frame, bg=CARD_WHITE)
    center.pack()
    self._spinner = Spinner(center, size=96, bg=CARD_WHITE)
    self._spinner.pack(pady=(0, 20))
    tk.Label(center, text="Processing…",
              font=F(15, "bold"), fg=NAVY_DEEP, bg=CARD_WHITE).pack()
    self._stage_lbl = tk.Label(center, text="Initialising…",
                                font=F(10), fg=TXT_SOFT, bg=CARD_WHITE)
    self._stage_lbl.pack(pady=(6, 2))
    self._pct_lbl = tk.Label(center, text="0%",
                              font=F(13, "bold"), fg=LIME_DARK, bg=CARD_WHITE)
    self._pct_lbl.pack(pady=(0, 16))
    self._prog_bar = ctk.CTkProgressBar(
        center, width=300, height=8, corner_radius=4,
        fg_color=NAVY_MIST, progress_color=LIME_MID, border_width=0
    )
    self._prog_bar.set(0)
    self._prog_bar.pack()
    tk.Frame(self._loader_frame, bg=CARD_WHITE).pack(expand=True, fill="both")

    # Smart classifier panel — replaces plain _txt_frame/_textbox block.
    # Creates self._txt_frame, self._textbox, and all _clf_* widgets.
    self._build_classifier_panel(card)

    self._analysis_frame = tk.Frame(card, bg=CARD_WHITE)
    sb2 = tk.Scrollbar(self._analysis_frame, relief="flat", troughcolor=OFF_WHITE,
                       bg=BORDER_LIGHT, width=8, bd=0)
    sb2.pack(side="right", fill="y")
    self._analysis_box = tk.Text(
        self._analysis_frame, wrap="word", font=FMONO(11),
        fg=TXT_NAVY, bg=CARD_WHITE, relief="flat", bd=0,
        padx=32, pady=24, spacing1=6, spacing2=3, spacing3=8,
        insertbackground=LIME_DARK, yscrollcommand=sb2.set,
        state="disabled", cursor="arrow",
        selectbackground=NAVY_GHOST, selectforeground=TXT_NAVY
    )
    self._analysis_box.pack(side="left", fill="both", expand=True)
    sb2.config(command=self._analysis_box.yview)
    self._configure_analysis_tags(self._analysis_box)

    self._build_summary_panel(card)
    self._build_ai_prompt_panel(card)
    self._build_cibi_output_panel(card)
    self._build_samples_panel(card)
    self._build_lookup_panel(card)
    self._build_lookup_summary_panel(card)
    self._build_lu_analysis_panel(card)         # built by lu_analysis_tab
    self._build_general_lookup_panel(card)      # built by general_lookup
    self._build_general_summary_panel(card)     # built by general_summary

    self._put_placeholder()
    self._current_tab = "cibi"
    self._switch_tab("cibi")

def _scroll_tabs(self, direction):
    """Scroll tabs left (-1) or right (1)"""
    if hasattr(self, '_tab_canvas'):
        # Get current view
        xview = self._tab_canvas.xview()
        # Calculate scroll amount (about 200 pixels per click)
        canvas_width = self._tab_canvas.winfo_width()
        if canvas_width > 0:
            scroll_amount = 200 / canvas_width
            if direction == -1:
                new_pos = max(0, xview[0] - scroll_amount)
            else:
                new_pos = min(1, xview[0] + scroll_amount)
            self._tab_canvas.xview_moveto(new_pos)

# ── TAG CONFIGURATION ─────────────────────────────────────────────────
def _configure_analysis_tags(self, box):
    sz = 11
    ff = FMONO(11)[0]
    box.tag_configure("search_match",    background=LIME_PALE,   foreground=TXT_NAVY)
    box.tag_configure("search_current",  background=LIME_BRIGHT, foreground=NAVY_DEEP)
    box.tag_configure("sec_header",      font=(ff, sz+2, "bold"), foreground=NAVY_DEEP,
                    spacing1=18, spacing3=6,  lmargin1=0,  lmargin2=0)
    box.tag_configure("sub_header",      font=(ff, sz,   "bold"), foreground=NAVY_MID,
                    spacing1=10, spacing3=4,  lmargin1=10, lmargin2=10)
    box.tag_configure("bullet",          font=(ff, sz),           foreground=TXT_NAVY,
                    lmargin1=24, lmargin2=36, spacing1=3,  spacing3=3)
    box.tag_configure("body",            font=(ff, sz),           foreground=TXT_NAVY,
                    lmargin1=10, lmargin2=10, spacing1=3,  spacing3=3)
    box.tag_configure("verdict_approve", font=(ff, sz, "bold"),   foreground=ACCENT_SUCCESS)
    box.tag_configure("verdict_cond",    font=(ff, sz, "bold"),   foreground=ACCENT_GOLD)
    box.tag_configure("verdict_decline", font=(ff, sz, "bold"),   foreground=ACCENT_RED)
    box.tag_configure("sym_ok",          font=(ff, sz),           foreground=ACCENT_SUCCESS)
    box.tag_configure("sym_warn",        font=(ff, sz),           foreground=ACCENT_GOLD)
    box.tag_configure("sym_bad",         font=(ff, sz),           foreground=ACCENT_RED)
    box.tag_configure("peso",            font=(ff, sz, "bold"),   foreground=LIME_DARK)
    box.tag_configure("risk_low",        font=(ff, sz, "bold"),   foreground=ACCENT_SUCCESS)
    box.tag_configure("risk_mod",        font=(ff, sz, "bold"),   foreground=ACCENT_GOLD)
    box.tag_configure("risk_high",       font=(ff, sz, "bold"),   foreground=ACCENT_RED)
    box.tag_configure("divider",         font=(ff, sz-2),         foreground=BORDER_MID)

# ── WRITE HELPERS ─────────────────────────────────────────────────────
def _write(self, txt, color=TXT_NAVY):
    box = self._textbox
    box.config(state="normal", fg=color)
    box.delete("1.0", "end")
    box.insert("end", txt)
    box.config(state="disabled")
    if self._search_var.get().strip() and self._current_tab == "extract":
        self._do_search()

def _write_analysis(self, txt, color=TXT_NAVY):
    box = self._analysis_box
    box.config(state="normal")
    box.delete("1.0", "end")

    if color != TXT_NAVY:
        box.insert("end", txt)
        box.config(state="disabled")
        return

    for line in txt.splitlines(keepends=True):
        s = line.rstrip("\n")
        if re.match(r'^#{1,4}\s+', s):
            clean = re.sub(r'^#{1,4}\s+', '', s)
            box.insert("end", clean + "\n", "sec_header")
        elif re.match(r'^\d{1,2}\.\s+[A-Z]', s):
            box.insert("end", "\n" + s + "\n", "sec_header")
        elif re.match(r'^[A-C]\)\s', s):
            box.insert("end", s + "\n", "sub_header")
        elif re.search(r'\bAPPROVE\b', s) and not re.search(r'CONDITIONALLY', s):
            self._insert_with_peso(box, s + "\n", "verdict_approve")
        elif re.search(r'CONDITIONALLY APPROVE', s):
            self._insert_with_peso(box, s + "\n", "verdict_cond")
        elif re.search(r'\bDECLINE\b', s):
            self._insert_with_peso(box, s + "\n", "verdict_decline")
        elif re.search(r'\b(Low Risk|LOW RISK|Low)\b', s) and "risk" in s.lower():
            self._insert_with_peso(box, s + "\n", "risk_low")
        elif re.search(r'\b(Moderate Risk|MODERATE|Moderate)\b', s) and "risk" in s.lower():
            self._insert_with_peso(box, s + "\n", "risk_mod")
        elif re.search(r'\b(High Risk|HIGH RISK|High)\b', s) and "risk" in s.lower():
            self._insert_with_peso(box, s + "\n", "risk_high")
        elif set(s.strip()).issubset({"─","—","-","=","_","*"}) and len(s.strip()) > 4:
            box.insert("end", s + "\n", "divider")
        elif s.strip().startswith(("•", "-", "*", "–")) or re.match(r'^\s+[•\-\*]', s):
            self._insert_with_peso(box, s + "\n", "bullet")
        elif "✅" in s:
            self._insert_sym_line(box, s + "\n", "sym_ok")
        elif "⚠" in s:
            self._insert_sym_line(box, s + "\n", "sym_warn")
        elif "❌" in s:
            self._insert_sym_line(box, s + "\n", "sym_bad")
        else:
            self._insert_with_peso(box, s + "\n", "body")

    box.config(state="disabled")
    if self._search_var.get().strip() and self._current_tab == "analysis":
        self._do_search()

def _insert_with_peso(self, box, text, base_tag):
    for part in re.split(r'(₱[\d,]+(?:\.\d+)?)', text):
        box.insert("end", part, "peso" if re.match(r'₱[\d,]+', part) else base_tag)

def _insert_sym_line(self, box, text, sym_tag):
    m = re.match(r'^(\s*[✅⚠❌]\s*)', text)
    if m:
        box.insert("end", m.group(1), sym_tag)
        self._insert_with_peso(box, text[m.end():], "body")
    else:
        self._insert_with_peso(box, text, "body")

def _put_placeholder(self):
    self._write(
        "Results will appear here after extraction.\n\n"
        "←  Choose a file from the left panel, then click  Extract Text  to begin.\n\n"
        "Engine:  PaddleOCR  →  Gemini 2.5 Flash VLM  →  Confidence Check",
        color=TXT_MUTED
    )
    self._write_analysis(
        "CIBI Analysis will appear here.\n\n"
        "Option A:  Extract a document, then click  Analyze CIBI.\n"
        "Option B:  After populating a CIBI Excel in CIBI Mode,\n"
        "           use  Analyze from Excel  to read cash-flow\n"
        "           and CIBI data directly from the workbook.\n\n"
        "The AI will assess:\n"
        "  • Cash-flow adequacy & DSR\n"
        "  • CIBI credit background\n"
        "  • Risk flags & final recommendation",
        color=TXT_MUTED
    )
    self._summary_placeholder()

def _switch_tab(self, tab):
    self._current_tab = tab

    # ── Style all tab buttons ──────────────────────────────────────────
    self._tab_style_fn(self._tab_extract_btn,          tab == "extract")
    self._tab_style_fn(self._tab_cibi_btn,             tab == "cibi")
    self._tab_style_fn(self._tab_analysis_btn,         tab == "analysis")
    self._tab_style_fn(self._tab_summary_btn,          tab == "summary")
    self._tab_style_fn(self._tab_aiprompt_btn,         tab == "aiprompt")
    self._tab_style_fn(self._tab_samples_btn,          tab == "samples")
    self._tab_style_fn(self._tab_lookup_btn,           tab == "lookup")
    self._tab_style_fn(self._tab_lookup_summary_btn,   tab == "lookup_summary")
    self._tab_style_fn(self._tab_lu_analysis_btn,      tab == "lu_analysis")
    self._tab_style_fn(self._tab_general_lookup_btn,   tab == "general_lookup")
    self._tab_style_fn(self._tab_general_summary_btn,  tab == "general_summary")

    # ── Hide all content frames ────────────────────────────────────────
    self._loader_frame.pack_forget()
    self._txt_frame.pack_forget()
    self._analysis_frame.pack_forget()
    self._summary_frame.pack_forget()
    self._aiprompt_frame.pack_forget()
    self._cibi_output_frame.pack_forget()
    self._samples_frame.pack_forget()
    self._lookup_frame.pack_forget()
    self._lookup_summary_frame.pack_forget()
    self._lu_analysis_frame.pack_forget()
    self._general_lookup_frame.pack_forget()
    self._general_summary_frame.pack_forget()

    # ── Show the requested frame ───────────────────────────────────────
    if tab == "extract":
        self._txt_frame.pack(fill="both", expand=True)
    elif tab == "analysis":
        self._analysis_frame.pack(fill="both", expand=True)
    elif tab == "summary":
        self._summary_frame.pack(fill="both", expand=True)
    elif tab == "cibi":
        self._cibi_output_frame.pack(fill="both", expand=True)
    elif tab == "samples":
        self._samples_frame.pack(fill="both", expand=True)
    elif tab == "lookup":
        self._lookup_frame.pack(fill="both", expand=True)
    elif tab == "lookup_summary":
        self._lookup_summary_frame.pack(fill="both", expand=True)
        _refresh_summary(self)
    elif tab == "lu_analysis":
        self._lu_analysis_frame.pack(fill="both", expand=True)
    elif tab == "general_lookup":
        self._general_lookup_frame.pack(fill="both", expand=True)
    elif tab == "general_summary":
        self._general_summary_frame.pack(fill="both", expand=True)
    else:
        self._aiprompt_frame.pack(fill="both", expand=True)
        self.after(50, self._chat_input.focus_set)

    if self._search_var.get().strip() and tab in ("extract", "analysis"):
        self._do_search()

def _show_loader(self, show, stage_text="Processing…"):
    if show:
        self._txt_frame.pack_forget()
        self._analysis_frame.pack_forget()
        self._summary_frame.pack_forget()
        self._aiprompt_frame.pack_forget()
        self._cibi_output_frame.pack_forget()
        self._lookup_frame.pack_forget()
        self._lookup_summary_frame.pack_forget()
        self._lu_analysis_frame.pack_forget()
        self._general_lookup_frame.pack_forget()
        self._general_summary_frame.pack_forget()
        self._loader_frame.pack(fill="both", expand=True)
        self._stage_lbl.config(text=stage_text)
        self._spinner.start()
        self._status_lbl.config(text="●  Processing…", fg=ACCENT_GOLD)
        self._topbar_status.config(text="● Processing…", fg=ACCENT_GOLD, bg=SIDEBAR_ITEM)
    else:
        self._spinner.stop()
        self._loader_frame.pack_forget()
        if self._current_tab == "extract":
            self._txt_frame.pack(fill="both", expand=True)
        elif self._current_tab == "analysis":
            self._analysis_frame.pack(fill="both", expand=True)
        elif self._current_tab == "summary":
            self._summary_frame.pack(fill="both", expand=True)
        elif self._current_tab == "cibi":
            self._cibi_output_frame.pack(fill="both", expand=True)
        elif self._current_tab == "lookup":
            self._lookup_frame.pack(fill="both", expand=True)
        elif self._current_tab == "lookup_summary":
            self._lookup_summary_frame.pack(fill="both", expand=True)
        elif self._current_tab == "general_lookup":
            self._general_lookup_frame.pack(fill="both", expand=True)
        elif self._current_tab == "general_summary":
            self._general_summary_frame.pack(fill="both", expand=True)
        else:
            self._aiprompt_frame.pack(fill="both", expand=True)
        self._status_lbl.config(text="●  Ready", fg=LIME_DARK)
        self._topbar_status.config(text="● Ready", fg=LIME_BRIGHT, bg=SIDEBAR_ITEM)

def _set_progress(self, pct, stage=""):
    self._pct_lbl.config(text=f"{pct}%")
    if stage:
        self._stage_lbl.config(text=stage)
    self._prog_bar.set(pct / 100)

# ── GRADIENT HELPERS ──────────────────────────────────────────────────
def _hbar(self, canvas, c1, c2, steps=40):
    canvas.delete("all")
    w = canvas.winfo_width()
    h = canvas.winfo_height()
    if w < 2: return
    for i in range(steps):
        c = hex_blend(c1, c2, i / steps)
        canvas.create_rectangle(
            int(w*i/steps), 0, int(w*(i+1)/steps)+1, h,
            fill=c, outline=""
        )

def _vbar(self, canvas, c1, c2, steps=20):
    canvas.delete("all")
    h = canvas.winfo_height()
    if h < 2: return
    for i in range(steps):
        c = hex_blend(c1, c2, i / steps)
        canvas.create_rectangle(
            0, int(h*i/steps), 5, int(h*(i+1)/steps)+1,
            fill=c, outline=""
        )

def _vbar_full(self, canvas, c1, c2, steps=30):
    canvas.delete("all")
    h = canvas.winfo_height()
    w = canvas.winfo_width()
    if h < 2: return
    for i in range(steps):
        c = hex_blend(c1, c2, i / steps)
        canvas.create_rectangle(
            0, int(h*i/steps), w, int(h*(i+1)/steps)+1,
            fill=c, outline=""
        )

def _div(self, parent, pady=(14, 10)):
    tk.Frame(parent, bg=BORDER_LIGHT, height=1).pack(fill="x", pady=pady)

def _sec(self, parent, text, pady=(0, 6)):
    tk.Label(parent, text=text, font=F(8, "bold"),
          fg=NAVY_PALE, bg=PANEL_LEFT).pack(anchor="w", pady=pady)

def _sidebar_sec(self, parent, text, pady=(0, 6)):
    row = tk.Frame(parent, bg=SIDEBAR_BG)
    row.pack(fill="x", pady=pady)
    tk.Frame(row, bg=LIME_BRIGHT, width=3, height=14).pack(side="left", padx=(0, 8))
    tk.Label(row, text=text, font=F(7, "bold"),
             fg=TXT_SOFT, bg=SIDEBAR_BG).pack(side="left")

def _file_icon_for(self, name):
    ext = Path(name).suffix.lower()
    if ext == ".pdf":             return "📄"
    if ext in (".docx", ".doc"): return "📝"
    if ext in (".xlsx", ".xls"): return "📊"
    if ext in IMAGE_EXTS:         return "🖼"
    return "📃"

# ── SEARCH ────────────────────────────────────────────────────────────
def _active_textbox(self):
    return self._textbox if self._current_tab == "extract" else self._analysis_box

def _do_search(self, *_):
    if self._current_tab not in ("extract", "analysis"):
        return
    query = self._search_var.get().strip()
    box   = self._active_textbox()
    box.tag_remove("search_match",   "1.0", "end")
    box.tag_remove("search_current", "1.0", "end")
    self._search_matches = []
    self._search_cursor  = -1
    if not query:
        self._match_lbl.config(text="", fg=TXT_SOFT)
        return
    start = "1.0"
    while True:
        pos = box.search(query, start, stopindex="end", nocase=True)
        if not pos: break
        end = f"{pos}+{len(query)}c"
        self._search_matches.append((pos, end))
        box.tag_add("search_match", pos, end)
        start = end
    count = len(self._search_matches)
    if count == 0:
        self._match_lbl.config(text="No results", fg=ACCENT_RED)
        return
    self._search_cursor = 0
    self._highlight_current()
    self._match_lbl.config(text=f"1 / {count}", fg=LIME_DARK)

def _highlight_current(self):
    if not self._search_matches: return
    box = self._active_textbox()
    box.tag_remove("search_current", "1.0", "end")
    pos, end = self._search_matches[self._search_cursor]
    box.tag_add("search_current", pos, end)
    box.see(pos)
    self._match_lbl.config(
        text=f"{self._search_cursor + 1} / {len(self._search_matches)}",
        fg=LIME_DARK
    )

def _search_next(self):
    if not self._search_matches: return
    self._search_cursor = (self._search_cursor + 1) % len(self._search_matches)
    self._highlight_current()

def _search_prev(self):
    if not self._search_matches: return
    self._search_cursor = (self._search_cursor - 1) % len(self._search_matches)
    self._highlight_current()

def _clear_search(self):
    self._search_var.set("")
    if self._current_tab in ("extract", "analysis"):
        box = self._active_textbox()
        box.tag_remove("search_match",   "1.0", "end")
        box.tag_remove("search_current", "1.0", "end")
    self._search_matches = []
    self._search_cursor  = -1
    self._match_lbl.config(text="", fg=TXT_SOFT)
    self._search_entry.focus_set()

# ── COPY ──────────────────────────────────────────────────────────────
def _copy(self):
    if self._current_tab == "extract":
        content = self._textbox.get("1.0", "end").strip()
    else:
        content = self._analysis_box.get("1.0", "end").strip()
    skip = ("Results will appear here", "Loan analysis will appear here")
    if content and not any(s in content for s in skip):
        self.clipboard_clear()
        self.clipboard_append(content)
        self._copy_btn.configure(text="✓  Copied!")
        self.after(2200, lambda: self._copy_btn.configure(text="⎘  Copy"))

# ══════════════════════════════════════════════════════════════════════
#  FILE BROWSING
# ══════════════════════════════════════════════════════════════════════

# ── attach all methods to DocExtractorApp ─────────────────────────────────────
def attach(cls):
    # ── Core UI panels ────────────────────────────────────────────────
    cls._build_topbar            = _build_topbar
    cls._build_left              = _build_left
    cls._build_right             = _build_right

    # ── Text / analysis tag config ────────────────────────────────────
    cls._configure_analysis_tags = _configure_analysis_tags

    # ── Write helpers ─────────────────────────────────────────────────
    cls._write                   = _write
    cls._write_analysis          = _write_analysis
    cls._insert_with_peso        = _insert_with_peso
    cls._insert_sym_line         = _insert_sym_line
    cls._put_placeholder         = _put_placeholder

    # ── Tab switching & loader ────────────────────────────────────────
    cls._switch_tab              = _switch_tab
    cls._show_loader             = _show_loader
    cls._set_progress            = _set_progress
    cls._scroll_tabs             = _scroll_tabs  # Add scroll method

    # ── Gradient / layout helpers ─────────────────────────────────────
    cls._hbar                    = _hbar
    cls._vbar                    = _vbar
    cls._vbar_full               = _vbar_full
    cls._div                     = _div
    cls._sec                     = _sec
    cls._sidebar_sec             = _sidebar_sec
    cls._file_icon_for           = _file_icon_for

    # ── Search ────────────────────────────────────────────────────────
    cls._active_textbox          = _active_textbox
    cls._do_search               = _do_search
    cls._highlight_current       = _highlight_current
    cls._search_next             = _search_next
    cls._search_prev             = _search_prev
    cls._clear_search            = _clear_search

    # ── Copy ──────────────────────────────────────────────────────────
    cls._copy                    = _copy

    # ── External tab modules ──────────────────────────────────────────
    _attach_lookup(cls)                # Look-Up tab
    _attach_summary(cls)               # LU Summary tab
    _attach_lu_analysis(cls)           # LU Analysis tab
    _attach_general_lookup(cls)        # General Look-Up tab
    _attach_general_summary(cls)       # General Summary tab
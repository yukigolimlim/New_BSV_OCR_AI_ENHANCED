"""
ui_panels.py — DocExtract Pro
================================
UI building methods attached to DocExtractorApp.

REDESIGN (FinSet-inspired):
  - Sidebar now hosts all navigation tabs as icon+label menu items,
    matching the FinSet dashboard style (rounded active pill, subtle
    hover states, branding at top, utility links at bottom).
  - Right panel is a clean content area — no tab row, no toolbar clutter.
    The extract toolbar is injected inside the Extracted content frame.
  - Color palette: deep navy sidebar (#0F1B2D), white card content area,
    lime/green accent (#A8E063 → #6BBF3E) for active state.
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


# ─────────────────────────────────────────────────────────────────────
#  DESIGN TOKENS  (FinSet-inspired palette on top of existing constants)
# ─────────────────────────────────────────────────────────────────────
_SB_BG        = "#0F1B2D"   # sidebar background  — deep navy
_SB_ACTIVE_BG = "#1A2F47"   # active nav pill background
_SB_HOVER_BG  = "#162338"   # hover background
_SB_ACCENT    = "#6BBF3E"   # lime-green accent (active indicator)
_SB_TXT       = "#7A94B0"   # inactive label colour
_SB_TXT_ACT   = "#FFFFFF"   # active label colour
_SB_BORDER    = "#1E3A5F"   # subtle divider
_SB_WIDTH     = 220         # sidebar width in pixels

_CARD_BG      = "#FFFFFF"
_PAGE_BG      = "#F4F6FA"   # outer page / window background
_HDR_BG       = "#FFFFFF"
_HDR_BORDER   = "#E8ECF2"

_PILL_R       = 10          # nav-pill corner radius


# ─────────────────────────────────────────────────────────────────────
#  TOP BAR  (unchanged logic, refreshed palette)
# ─────────────────────────────────────────────────────────────────────
def _build_topbar(self):
    bar = tk.Frame(self, bg=_SB_BG, height=52)
    bar.pack(fill="x")
    bar.pack_propagate(False)
    bar.bind("<ButtonPress-1>", self._drag_start)
    bar.bind("<B1-Motion>",     self._drag_move)

    # Window controls
    ctrl = tk.Frame(bar, bg=_SB_BG)
    ctrl.pack(side="right", fill="y", padx=8)

    close_btn = tk.Label(ctrl, text="✕", font=F(10, "bold"),
                         fg="#7A94B0", bg=_SB_BG, cursor="hand2",
                         width=3, anchor="center")
    close_btn.bind("<Enter>",    lambda e: close_btn.config(fg=WHITE, bg=ACCENT_RED))
    close_btn.bind("<Leave>",    lambda e: close_btn.config(fg="#7A94B0", bg=_SB_BG))
    close_btn.bind("<Button-1>", lambda e: self.destroy())
    close_btn.pack(side="right", padx=(4, 2), fill="y")

    min_btn = tk.Label(ctrl, text="─", font=F(10, "bold"),
                       fg="#7A94B0", bg=_SB_BG, cursor="hand2",
                       width=3, anchor="center")
    min_btn.bind("<Enter>",    lambda e: min_btn.config(fg=WHITE, bg=_SB_HOVER_BG))
    min_btn.bind("<Leave>",    lambda e: min_btn.config(fg="#7A94B0", bg=_SB_BG))
    min_btn.bind("<Button-1>", lambda e: self._do_minimize())
    min_btn.pack(side="right", padx=(4, 2), fill="y")

    # Branding
    brand = tk.Frame(bar, bg=_SB_BG)
    brand.pack(side="left", fill="y", padx=(16, 0))

    mark = tk.Canvas(brand, width=28, height=28, bg=_SB_BG, highlightthickness=0)
    mark.pack(side="left", padx=(0, 10), pady=12)
    mark.create_rectangle(0, 8, 17, 28, fill=NAVY_LIGHT, outline="")
    mark.create_rectangle(11, 0, 28, 20, fill=_SB_ACCENT,  outline="")
    mark.create_rectangle(13, 10, 15, 12, fill=_SB_BG,     outline="")

    tk.Label(brand, text="DocExtract Pro",
             font=F(13, "bold"), fg=WHITE, bg=_SB_BG).pack(side="left")
    tk.Label(brand, text=" — Banco San Vicente",
             font=F(10), fg=_SB_TXT, bg=_SB_BG).pack(side="left")

    self._topbar_status = tk.Label(
        bar, text="● Ready",
        font=F(8, "bold"), fg=_SB_ACCENT,
        bg="#1A2F47", padx=12, pady=4, relief="flat"
    )
    self._topbar_status.pack(side="right", padx=(0, 20), pady=14)


# ─────────────────────────────────────────────────────────────────────
#  LEFT SIDEBAR  — FinSet-style icon+label navigation
# ─────────────────────────────────────────────────────────────────────
_NAV_ITEMS = [
    # (tab_key,         icon,  label)
    ("cibi",            "📋",  "CIBI Mode"),
    ("extract",         "📄",  "Extracted"),
    ("analysis",        "🏦",  "CIBI Analysis"),
    ("summary",         "📊",  "Summary"),
    ("aiprompt",        "🤖",  "AI Chat"),
    ("samples",         "🗂",  "Samples"),
    ("lookup",          "🔎",  "Look-Up"),
    ("lookup_summary",  "📋",  "LU Summary"),
    ("lu_analysis",     "📈",  "LU Analysis"),
    ("general_lookup",  "📂",  "General Look-Up"),
    ("general_summary", "📋",  "General Summary"),
]

def _build_left(self, p):
    # p is the left pane passed in from the main app — fill it directly
    # so the sidebar colour and layout take full ownership of that pane.
    p.configure(bg=_SB_BG)

    # Convenience alias — all children go into p
    sidebar = p

    # ── Logo / branding block ─────────────────────────────────────────
    logo_block = tk.Frame(sidebar, bg=_SB_BG, height=72)
    logo_block.pack(fill="x")
    logo_block.pack_propagate(False)

    logo_loaded = False
    if LOGO_PATH.exists():
        try:
            from PIL import Image
            img    = Image.open(LOGO_PATH).convert("RGBA")
            bg_pil = Image.new(
                "RGBA", img.size,
                tuple(int(_SB_BG[i:i+2], 16) for i in (1, 3, 5)) + (255,)
            )
            bg_pil.paste(img, mask=img.split()[3])
            bg_pil = bg_pil.convert("RGB")
            bg_pil.thumbnail((160, 38), Image.LANCZOS)
            self._logo_img = ctk.CTkImage(
                light_image=bg_pil, dark_image=bg_pil,
                size=(bg_pil.width, bg_pil.height)
            )
            lf = ctk.CTkFrame(logo_block, fg_color="transparent")
            lf.place(relx=0.5, rely=0.5, anchor="center")
            ctk.CTkLabel(lf, image=self._logo_img, text="",
                         fg_color="transparent").pack()
            logo_loaded = True
        except Exception:
            pass

    if not logo_loaded:
        inner = tk.Frame(logo_block, bg=_SB_BG)
        inner.place(relx=0.5, rely=0.5, anchor="center")
        mark = tk.Canvas(inner, width=24, height=26,
                         bg=_SB_BG, highlightthickness=0)
        mark.pack(side="left", padx=(0, 8))
        mark.create_rectangle(0, 8, 14, 26, fill=NAVY_LIGHT,  outline="")
        mark.create_rectangle(10, 0, 24, 18, fill=_SB_ACCENT, outline="")
        tk.Label(inner, text="BSV AI-OCR",
                 font=F(11, "bold"), fg=WHITE, bg=_SB_BG).pack(side="left")

    # Thin border under logo
    tk.Frame(sidebar, bg=_SB_BORDER, height=1).pack(fill="x")

    # ── NAVIGATION section label ──────────────────────────────────────
    tk.Frame(sidebar, bg=_SB_BG, height=12).pack()
    tk.Label(sidebar, text="NAVIGATION",
             font=F(7, "bold"), fg=_SB_TXT, bg=_SB_BG,
             anchor="w").pack(fill="x", padx=18, pady=(0, 4))

    # ── Nav items ─────────────────────────────────────────────────────
    self._nav_btns       = {}
    self._active_tab_key = tk.StringVar(value="cibi")

    nav_container = tk.Frame(sidebar, bg=_SB_BG)
    nav_container.pack(fill="x", padx=10)

    for tab_key, icon, label in _NAV_ITEMS:
        # Each pill is a fixed-height frame; pack_propagate kept True so
        # children are visible — we control height via minsize in the row.
        pill = tk.Frame(nav_container, bg=_SB_BG, cursor="hand2", height=36)
        pill.pack(fill="x", pady=1)
        pill.pack_propagate(False)

        # Left accent stripe (3 px wide, shown only when active)
        accent_bar = tk.Frame(pill, bg=_SB_ACCENT, width=3)
        # Not packed yet — shown by _apply_nav_active

        # Icon label
        icon_lbl = tk.Label(pill, text=icon,
                            font=("Segoe UI Emoji", 12),
                            fg=_SB_TXT, bg=_SB_BG,
                            width=2, anchor="center")
        icon_lbl.pack(side="left", padx=(8, 4))

        # Text label
        text_lbl = tk.Label(pill, text=label,
                            font=F(9), fg=_SB_TXT, bg=_SB_BG,
                            anchor="w")
        text_lbl.pack(side="left", fill="x", expand=True)

        self._nav_btns[tab_key] = (pill, accent_bar, icon_lbl, text_lbl)

        def _make_click(k=tab_key):
            return lambda e: self._switch_tab(k)

        def _make_enter(pill=pill, icon=icon_lbl, txt=text_lbl, key=tab_key):
            def _on(e):
                if self._active_tab_key.get() != key:
                    for w in (pill, icon, txt):
                        w.config(bg=_SB_HOVER_BG)
            return _on

        def _make_leave(pill=pill, icon=icon_lbl, txt=text_lbl, key=tab_key):
            def _off(e):
                if self._active_tab_key.get() != key:
                    for w in (pill, icon, txt):
                        w.config(bg=_SB_BG)
            return _off

        for w in (pill, icon_lbl, text_lbl):
            w.bind("<Button-1>", _make_click())
            w.bind("<Enter>",    _make_enter())
            w.bind("<Leave>",    _make_leave())

    # ── Divider ───────────────────────────────────────────────────────
    tk.Frame(sidebar, bg=_SB_BORDER, height=1).pack(fill="x", pady=(14, 0), padx=10)

    # ── Utility links ─────────────────────────────────────────────────
    util_frame = tk.Frame(sidebar, bg=_SB_BG)
    util_frame.pack(fill="x", padx=10, pady=(6, 0))

    for u_icon, u_label in [("❓", "Help"), ("⚙", "Settings")]:
        row = tk.Frame(util_frame, bg=_SB_BG, cursor="hand2", height=32)
        row.pack(fill="x", pady=1)
        row.pack_propagate(False)
        il = tk.Label(row, text=u_icon, font=("Segoe UI Emoji", 11),
                      fg=_SB_TXT, bg=_SB_BG, width=2, anchor="center")
        il.pack(side="left", padx=(8, 4))
        tl = tk.Label(row, text=u_label, font=F(9), fg=_SB_TXT,
                      bg=_SB_BG, anchor="w")
        tl.pack(side="left")
        for w in (row, il, tl):
            w.bind("<Enter>", lambda e, r=row, i=il, t=tl: [
                x.config(bg=_SB_HOVER_BG) for x in (r, i, t)])
            w.bind("<Leave>", lambda e, r=row, i=il, t=tl: [
                x.config(bg=_SB_BG) for x in (r, i, t)])

    # ── Spacer pushes footer to bottom ────────────────────────────────
    tk.Frame(sidebar, bg=_SB_BG).pack(fill="both", expand=True)

    # ── Version footer ────────────────────────────────────────────────
    ver = tk.Frame(sidebar, bg=_SB_BG)
    ver.pack(fill="x", pady=(0, 12), padx=16)
    tk.Frame(ver, bg=_SB_BORDER, height=1).pack(fill="x", pady=(0, 8))
    tk.Label(ver, text="Gemini 2.5 Flash · PaddleOCR",
             font=F(7), fg="#2D4F7A", bg=_SB_BG, anchor="w").pack(anchor="w")
    tk.Label(ver, text="BSV AI-OCR v2.0",
             font=F(7), fg="#2D4F7A", bg=_SB_BG, anchor="w").pack(anchor="w")


def _apply_nav_active(self, active_key):
    """Update sidebar nav pill styles to reflect the active tab.
    The accent stripe uses place() so it never disrupts the pack chain
    of icon_lbl / text_lbl.
    """
    self._active_tab_key.set(active_key)
    for key, (pill, accent_bar, icon_lbl, text_lbl) in self._nav_btns.items():
        if key == active_key:
            pill.config(bg=_SB_ACTIVE_BG)
            icon_lbl.config(fg=WHITE,   bg=_SB_ACTIVE_BG)
            text_lbl.config(fg=WHITE,   bg=_SB_ACTIVE_BG, font=F(9, "bold"))
            # Place the 3-px stripe flush-left; it floats above pack children
            accent_bar.place(x=0, y=0, width=3, relheight=1.0)
            accent_bar.lift()
        else:
            pill.config(bg=_SB_BG)
            icon_lbl.config(fg=_SB_TXT, bg=_SB_BG)
            text_lbl.config(fg=_SB_TXT, bg=_SB_BG, font=F(9))
            accent_bar.place_forget()


# ─────────────────────────────────────────────────────────────────────
#  RIGHT PANEL  — clean content card, no tab row
# ─────────────────────────────────────────────────────────────────────
def _build_right(self, p):
    # p is the right pane — configure it directly as the page background
    p.configure(bg=_PAGE_BG)
    page = p   # alias for readability; all children pack into p

    # ── Top header bar (title + status + copy) ────────────────────────
    hdr = tk.Frame(page, bg=_HDR_BG, height=64)
    hdr.pack(fill="x")
    hdr.pack_propagate(False)
    tk.Frame(page, bg=_HDR_BORDER, height=1).pack(fill="x")

    hdr_inner = tk.Frame(hdr, bg=_HDR_BG)
    hdr_inner.pack(side="left", fill="y", padx=(24, 0))

    self._page_title_lbl = tk.Label(
        hdr_inner, text="CIBI Mode",
        font=F(17, "bold"), fg=NAVY_DEEP, bg=_HDR_BG
    )
    self._page_title_lbl.pack(side="left", anchor="center")

    status_pill = tk.Frame(hdr, bg=LIME_MIST,
                           highlightbackground=LIME_MID, highlightthickness=1)
    status_pill.pack(side="left", padx=(14, 0), pady=18)
    self._status_lbl = tk.Label(
        status_pill, text="●  Ready",
        font=F(8, "bold"), fg=LIME_DARK,
        bg=LIME_MIST, padx=12, pady=4
    )
    self._status_lbl.pack()

    self._copy_btn = ctk.CTkButton(
        hdr, text="⎘  Copy", command=self._copy,
        width=100, height=32, corner_radius=6,
        fg_color=NAVY_MIST, hover_color=NAVY_GHOST, text_color=NAVY_MID,
        font=FF(9, "bold"),
        border_width=1, border_color=BORDER_MID
    )
    self._copy_btn.pack(side="right", padx=(0, 20))

    # ── Pipeline caption ──────────────────────────────────────────────
    pipe_row = tk.Frame(page, bg=_PAGE_BG)
    pipe_row.pack(fill="x", padx=24, pady=(8, 0))
    tk.Label(
        pipe_row,
        text="Pipeline:  PaddleOCR  →  Gemini 2.5 Flash VLM  →  Confidence Scoring  →  CIBI Credit Analysis",
        font=F(8), fg=TXT_MUTED, bg=_PAGE_BG
    ).pack(anchor="w")

    # ── Search bar ────────────────────────────────────────────────────
    search_row = tk.Frame(page, bg=_PAGE_BG)
    search_row.pack(fill="x", padx=24, pady=(6, 8))

    search_wrap = tk.Frame(search_row, bg=WHITE,
                           highlightbackground=BORDER_MID,
                           highlightthickness=1)
    search_wrap.pack(side="left", fill="x", expand=True)

    tk.Label(search_wrap, text="🔍", font=("Segoe UI Emoji", 10),
             bg=WHITE, fg=NAVY_PALE).pack(side="left", padx=(10, 2))

    self._search_var   = tk.StringVar()
    self._search_entry = tk.Entry(
        search_wrap, textvariable=self._search_var,
        font=F(10), fg=TXT_NAVY, bg=WHITE,
        relief="flat", bd=0, insertbackground=NAVY_MID, width=28
    )
    self._search_entry.pack(side="left", fill="x", expand=True, pady=6)
    self._search_entry.bind("<Return>",   lambda e: self._do_search())
    self._search_entry.bind("<KP_Enter>", lambda e: self._do_search())
    self._search_entry.bind("<Escape>",   lambda e: self._clear_search())
    self._search_var.trace_add("write",   lambda *a: self._do_search())

    self._match_lbl = tk.Label(search_wrap, text="",
                               font=F(8), fg=TXT_SOFT, bg=WHITE, padx=8)
    self._match_lbl.pack(side="left")

    nav_frame = tk.Frame(search_row, bg=_PAGE_BG)
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

    # ── Content card ──────────────────────────────────────────────────
    card_outer = tk.Frame(page, bg=_HDR_BORDER, padx=1, pady=1)
    card_outer.pack(fill="both", expand=True, padx=24, pady=(0, 16))
    card = tk.Frame(card_outer, bg=_CARD_BG)
    card.pack(fill="both", expand=True)

    # Thin top accent bar
    top_acc = tk.Canvas(card, height=4, bg=_CARD_BG, highlightthickness=0)
    top_acc.pack(fill="x")
    top_acc.bind("<Configure>",
                 lambda e, c=top_acc: self._hbar(c, NAVY_DEEP, _SB_ACCENT, 80))

    # ── Loader frame ──────────────────────────────────────────────────
    self._loader_frame = tk.Frame(card, bg=_CARD_BG)
    tk.Frame(self._loader_frame, bg=_CARD_BG).pack(expand=True, fill="both")
    center = tk.Frame(self._loader_frame, bg=_CARD_BG)
    center.pack()
    self._spinner = Spinner(center, size=96, bg=_CARD_BG)
    self._spinner.pack(pady=(0, 20))
    tk.Label(center, text="Processing…",
             font=F(15, "bold"), fg=NAVY_DEEP, bg=_CARD_BG).pack()
    self._stage_lbl = tk.Label(center, text="Initialising…",
                               font=F(10), fg=TXT_SOFT, bg=_CARD_BG)
    self._stage_lbl.pack(pady=(6, 2))
    self._pct_lbl = tk.Label(center, text="0%",
                             font=F(13, "bold"), fg=LIME_DARK, bg=_CARD_BG)
    self._pct_lbl.pack(pady=(0, 16))
    self._prog_bar = ctk.CTkProgressBar(
        center, width=300, height=8, corner_radius=4,
        fg_color=NAVY_MIST, progress_color=LIME_MID, border_width=0
    )
    self._prog_bar.set(0)
    self._prog_bar.pack()
    tk.Frame(self._loader_frame, bg=_CARD_BG).pack(expand=True, fill="both")

    # Content panels (same as before — logic unchanged)
    self._build_classifier_panel(card)

    self._analysis_frame = tk.Frame(card, bg=_CARD_BG)
    sb2 = tk.Scrollbar(self._analysis_frame, relief="flat",
                       troughcolor=OFF_WHITE,
                       bg=BORDER_LIGHT, width=8, bd=0)
    sb2.pack(side="right", fill="y")
    self._analysis_box = tk.Text(
        self._analysis_frame, wrap="word", font=FMONO(11),
        fg=TXT_NAVY, bg=_CARD_BG, relief="flat", bd=0,
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
    self._build_lu_analysis_panel(card)
    self._build_general_lookup_panel(card)
    self._build_general_summary_panel(card)

    self._put_placeholder()
    self._current_tab = "cibi"
    self._switch_tab("cibi")


# ─────────────────────────────────────────────────────────────────────
#  EXTRACT TOOLBAR  (injected at top of _txt_frame — unchanged logic)
# ─────────────────────────────────────────────────────────────────────
def _build_extract_toolbar(self, parent):
    BG = _CARD_BG

    # ① Header row
    hdr_row = tk.Frame(parent, bg=BG)
    hdr_row.pack(fill="x", padx=20, pady=(14, 6))

    lbl_block = tk.Frame(hdr_row, bg=BG)
    lbl_block.pack(side="left", fill="y")
    tk.Frame(lbl_block, bg=_SB_ACCENT, width=3, height=14).pack(
        side="left", padx=(0, 8), pady=3)
    tk.Label(lbl_block, text="FILES & ACTIONS",
             font=F(7, "bold"), fg=TXT_SOFT, bg=BG).pack(side="left")

    self._toolbar_count_lbl = tk.Label(
        hdr_row, text="No files selected",
        font=F(8), fg=TXT_MUTED, bg=BG
    )
    self._toolbar_count_lbl.pack(side="right")

    # ② Drop-zone + file list
    middle_row = tk.Frame(parent, bg=BG)
    middle_row.pack(fill="x", padx=20, pady=(0, 6))

    drop_chip = tk.Frame(
        middle_row, bg=SIDEBAR_ITEM,
        highlightbackground=_SB_ACCENT, highlightthickness=1,
        width=110, height=52
    )
    drop_chip.pack(side="left", padx=(0, 10))
    drop_chip.pack_propagate(False)
    inner_drop = tk.Frame(drop_chip, bg=SIDEBAR_ITEM)
    inner_drop.place(relx=0.5, rely=0.5, anchor="center")
    self._icon_lbl = tk.Label(
        inner_drop, text="📁",
        font=("Segoe UI Emoji", 15), fg=_SB_ACCENT, bg=SIDEBAR_ITEM
    )
    self._icon_lbl.pack()
    self._filename_lbl = tk.Label(
        inner_drop, text="Drop / Browse",
        font=F(7), fg=TXT_SOFT, bg=SIDEBAR_ITEM,
        wraplength=96, justify="center"
    )
    self._filename_lbl.pack(pady=(1, 0))

    list_outer = tk.Frame(
        middle_row, bg=SIDEBAR_ITEM,
        highlightbackground="#1E3A5F", highlightthickness=1
    )
    list_outer.pack(side="left", fill="x", expand=True, ipady=2)

    list_hdr = tk.Frame(list_outer, bg=SIDEBAR_ITEM)
    list_hdr.pack(fill="x", padx=8, pady=(3, 1))
    tk.Label(list_hdr, text="SELECTED FILES",
             font=F(7, "bold"), fg=TXT_SOFT, bg=SIDEBAR_ITEM).pack(side="left")
    self._clear_btn = tk.Label(
        list_hdr, text="✕ Clear",
        font=F(7, "bold"), fg=ACCENT_RED, bg=SIDEBAR_ITEM, cursor="hand2"
    )
    self._clear_btn.pack(side="right")
    self._clear_btn.bind("<Button-1>", lambda e: self._clear_files())

    list_body = tk.Frame(list_outer, bg=SIDEBAR_ITEM, height=36)
    list_body.pack(fill="x", padx=4, pady=(0, 3))
    list_body.pack_propagate(False)

    list_sb = tk.Scrollbar(list_body, relief="flat",
                           troughcolor=SIDEBAR_ITEM, bg=SIDEBAR_HVR,
                           width=5, bd=0)
    list_sb.pack(side="right", fill="y")

    self._file_listbox = tk.Listbox(
        list_body,
        font=F(8), fg="#C5D8F5", bg=SIDEBAR_ITEM,
        relief="flat", bd=0,
        selectbackground=SIDEBAR_HVR, selectforeground=WHITE,
        activestyle="none",
        yscrollcommand=list_sb.set,
        height=2,
    )
    self._file_listbox.pack(side="left", fill="both", expand=True)
    list_sb.config(command=self._file_listbox.yview)
    self._file_listbox.bind("<Button-3>", self._remove_selected_file)

    # ③ Action buttons
    btn_row = tk.Frame(parent, bg=BG)
    btn_row.pack(fill="x", padx=20, pady=(0, 10))

    self._browse_btn = ctk.CTkButton(
        btn_row, text="📁  Browse",
        command=self._browse,
        height=34, corner_radius=7,
        fg_color=_SB_ACCENT, hover_color=LIME_BRIGHT,
        text_color="#0F1B2D",
        font=FF(9, "bold"),
        border_width=0, width=110
    )
    self._browse_btn.pack(side="left", padx=(0, 5))

    self._add_btn = ctk.CTkButton(
        btn_row, text="➕  Add More",
        command=self._browse_add,
        height=34, corner_radius=7,
        fg_color=SIDEBAR_ITEM, hover_color=SIDEBAR_HVR,
        text_color="#C5D8F5",
        font=FF(9),
        border_width=1, border_color="#1E3A5F",
        state="disabled", width=110
    )
    self._add_btn.pack(side="left", padx=(0, 5))

    tk.Frame(btn_row, bg=BORDER_LIGHT, width=1, height=28).pack(
        side="left", padx=(4, 9), pady=3)

    self._ext_btn = ctk.CTkButton(
        btn_row, text="⚡  Extract Text",
        command=self._start_extraction,
        height=34, corner_radius=7,
        fg_color=NAVY_LIGHT, hover_color=NAVY_PALE,
        text_color=WHITE,
        font=FF(9, "bold"),
        state="disabled", border_width=0, width=130
    )
    self._ext_btn.pack(side="left", padx=(0, 5))

    self._analyze_btn = ctk.CTkButton(
        btn_row, text="🏦  Analyze CIBI",
        command=self._start_analysis,
        height=34, corner_radius=7,
        fg_color=SIDEBAR_ITEM, hover_color=SIDEBAR_HVR,
        text_color=_SB_ACCENT,
        font=FF(9, "bold"),
        state="disabled",
        border_width=1, border_color=LIME_DARK, width=140
    )
    self._analyze_btn.pack(side="left", padx=(0, 5))

    self._analyze_excel_btn = ctk.CTkButton(
        btn_row, text="📊  Analyze from Excel",
        command=self._start_cibi_analysis_from_excel,
        height=34, corner_radius=7,
        fg_color=SIDEBAR_ITEM, hover_color=SIDEBAR_HVR,
        text_color=LIME_PALE,
        font=FF(9, "bold"),
        border_width=1, border_color=LIME_DARK, width=175
    )
    self._analyze_excel_btn.pack(side="left")

    tk.Label(
        btn_row,
        text="Right-click a file in the list to remove it",
        font=F(7), fg=TXT_MUTED, bg=BG
    ).pack(side="right")

    # ④ Divider
    tk.Frame(parent, bg=BORDER_LIGHT, height=1).pack(fill="x", padx=20, pady=(0, 4))

    # Keep count label in sync
    def _update_toolbar_count(event=None):
        n = self._file_listbox.size()
        if n == 0:
            self._toolbar_count_lbl.config(text="No files selected", fg=TXT_MUTED)
            self._filename_lbl.config(text="Drop / Browse")
        elif n == 1:
            name  = self._file_listbox.get(0)
            short = (name[:22] + "…") if len(name) > 24 else name
            self._toolbar_count_lbl.config(text=short, fg=TXT_NAVY)
            self._filename_lbl.config(text="1 file")
        else:
            self._toolbar_count_lbl.config(text=f"{n} files selected", fg=_SB_ACCENT)
            self._filename_lbl.config(text=f"{n} files")

    self._file_listbox.bind("<<ListboxSelect>>", _update_toolbar_count)
    self._update_toolbar_count = _update_toolbar_count


# ─────────────────────────────────────────────────────────────────────
#  TAG CONFIGURATION  (unchanged)
# ─────────────────────────────────────────────────────────────────────
def _configure_analysis_tags(self, box):
    sz = 11
    ff = FMONO(11)[0]
    box.tag_configure("search_match",    background=LIME_PALE,    foreground=TXT_NAVY)
    box.tag_configure("search_current",  background=LIME_BRIGHT,  foreground=NAVY_DEEP)
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


# ─────────────────────────────────────────────────────────────────────
#  WRITE HELPERS  (unchanged)
# ─────────────────────────────────────────────────────────────────────
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
        "↑  Use the toolbar above to browse a file, then click  Extract Text  to begin.\n\n"
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


# ─────────────────────────────────────────────────────────────────────
#  TAB SWITCHING  — updated to drive sidebar nav active state
# ─────────────────────────────────────────────────────────────────────

# Map tab key → human-readable page title shown in the header
_TAB_TITLES = {
    "cibi":            "CIBI Mode",
    "extract":         "Extracted Content",
    "analysis":        "CIBI Analysis",
    "summary":         "Summary",
    "aiprompt":        "AI Chat",
    "samples":         "Samples",
    "lookup":          "Look-Up",
    "lookup_summary":  "LU Summary",
    "lu_analysis":     "LU Analysis",
    "general_lookup":  "General Look-Up",
    "general_summary": "General Summary",
}

def _switch_tab(self, tab):
    self._current_tab = tab

    # Update sidebar active highlight
    if hasattr(self, '_nav_btns'):
        self._apply_nav_active(tab)

    # Update page title
    if hasattr(self, '_page_title_lbl'):
        self._page_title_lbl.config(text=_TAB_TITLES.get(tab, tab.title()))

    # Hide all content frames
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

    # Show the requested frame
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
    else:  # aiprompt
        self._aiprompt_frame.pack(fill="both", expand=True)
        self.after(50, self._chat_input.focus_set)

    if self._search_var.get().strip() and tab in ("extract", "analysis"):
        self._do_search()


# ─────────────────────────────────────────────────────────────────────
#  LOADER  (unchanged logic, updated bg colour)
# ─────────────────────────────────────────────────────────────────────
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
        self._topbar_status.config(text="● Processing…", fg=ACCENT_GOLD, bg="#1A2F47")
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
        self._topbar_status.config(text="● Ready", fg=_SB_ACCENT, bg="#1A2F47")

def _set_progress(self, pct, stage=""):
    self._pct_lbl.config(text=f"{pct}%")
    if stage:
        self._stage_lbl.config(text=stage)
    self._prog_bar.set(pct / 100)


# ─────────────────────────────────────────────────────────────────────
#  GRADIENT / LAYOUT HELPERS  (unchanged)
# ─────────────────────────────────────────────────────────────────────
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
    row = tk.Frame(parent, bg=_SB_BG)
    row.pack(fill="x", pady=pady)
    tk.Frame(row, bg=_SB_ACCENT, width=3, height=14).pack(side="left", padx=(0, 8))
    tk.Label(row, text=text, font=F(7, "bold"),
             fg=_SB_TXT, bg=_SB_BG).pack(side="left")

def _file_icon_for(self, name):
    ext = Path(name).suffix.lower()
    if ext == ".pdf":              return "📄"
    if ext in (".docx", ".doc"):   return "📝"
    if ext in (".xlsx", ".xls"):   return "📊"
    if ext in IMAGE_EXTS:          return "🖼"
    return "📃"


# ─────────────────────────────────────────────────────────────────────
#  SEARCH  (unchanged)
# ─────────────────────────────────────────────────────────────────────
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


# ─────────────────────────────────────────────────────────────────────
#  COPY  (unchanged)
# ─────────────────────────────────────────────────────────────────────
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


# ═════════════════════════════════════════════════════════════════════
#  ATTACH ALL METHODS TO DocExtractorApp
# ═════════════════════════════════════════════════════════════════════
def attach(cls):
    # Core UI panels
    cls._build_topbar            = _build_topbar
    cls._build_left              = _build_left
    cls._build_right             = _build_right

    # Sidebar nav helper (NEW)
    cls._apply_nav_active        = _apply_nav_active

    # Extract-tab toolbar
    cls._build_extract_toolbar   = _build_extract_toolbar

    # Text / analysis tag config
    cls._configure_analysis_tags = _configure_analysis_tags

    # Write helpers
    cls._write                   = _write
    cls._write_analysis          = _write_analysis
    cls._insert_with_peso        = _insert_with_peso
    cls._insert_sym_line         = _insert_sym_line
    cls._put_placeholder         = _put_placeholder

    # Tab switching & loader
    cls._switch_tab              = _switch_tab
    cls._show_loader             = _show_loader
    cls._set_progress            = _set_progress

    # Gradient / layout helpers
    cls._hbar                    = _hbar
    cls._vbar                    = _vbar
    cls._vbar_full               = _vbar_full
    cls._div                     = _div
    cls._sec                     = _sec
    cls._sidebar_sec             = _sidebar_sec
    cls._file_icon_for           = _file_icon_for

    # Search
    cls._active_textbox          = _active_textbox
    cls._do_search               = _do_search
    cls._highlight_current       = _highlight_current
    cls._search_next             = _search_next
    cls._search_prev             = _search_prev
    cls._clear_search            = _clear_search

    # Copy
    cls._copy                    = _copy

    # External tab modules
    _attach_lookup(cls)
    _attach_summary(cls)
    _attach_lu_analysis(cls)
    _attach_general_lookup(cls)
    _attach_general_summary(cls)
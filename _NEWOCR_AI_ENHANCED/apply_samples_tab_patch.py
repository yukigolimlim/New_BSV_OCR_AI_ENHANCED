"""
apply_samples_tab_patch.py
==========================
Automatically wires the SamplesTabMixin into app.py.

Run:  python apply_samples_tab_patch.py

Makes a backup at app.py.bak before touching anything.
"""

from pathlib import Path
import shutil

APP_PATH = Path(__file__).parent / "app.py"

# ── Each patch: (name, find_str, replace_str) ────────────────────────────────

# PATCH 1 — Add mixin import after existing imports block
PATCH_1_FIND = "from Cibi_populator import populate_cibi_form"
PATCH_1_REPLACE = """\
from Cibi_populator import populate_cibi_form
from samples_tab import SamplesTabMixin"""

# PATCH 2 — Add mixin to class definition
PATCH_2_FIND = "class DocExtractorApp(ctk.CTk):"
PATCH_2_REPLACE = "class DocExtractorApp(SamplesTabMixin, ctk.CTk):"

# PATCH 3 — Init mixin state inside __init__, after _cibi_slots setup
PATCH_3_FIND = "        self._cibi_stage          = \"idle\""
PATCH_3_REPLACE = """\
        self._cibi_stage          = "idle"
        self._samples_init_state()   # ← SamplesTabMixin"""

# PATCH 4 — Add Samples tab button in _build_right, after AI Chat button
PATCH_4_FIND = """\
        self._tab_aiprompt_btn = ctk.CTkButton(
            tab_row, text="🤖  AI Chat", width=110, height=30,
            corner_radius=6, font=ctk.CTkFont(_FONT_FAMILY, 9, weight="bold"),
            command=lambda: self._switch_tab("aiprompt")
        )
        self._tab_aiprompt_btn.pack(side="left", padx=(0, 4))"""

PATCH_4_REPLACE = """\
        self._tab_aiprompt_btn = ctk.CTkButton(
            tab_row, text="🤖  AI Chat", width=110, height=30,
            corner_radius=6, font=ctk.CTkFont(_FONT_FAMILY, 9, weight="bold"),
            command=lambda: self._switch_tab("aiprompt")
        )
        self._tab_aiprompt_btn.pack(side="left", padx=(0, 4))

        self._tab_samples_btn = ctk.CTkButton(
            tab_row, text="🗂  Samples", width=110, height=30,
            corner_radius=6, font=ctk.CTkFont(_FONT_FAMILY, 9, weight="bold"),
            command=lambda: self._switch_tab("samples")
        )
        self._tab_samples_btn.pack(side="left", padx=(0, 4))"""

# PATCH 5 — Add tab style initialisation for the new button
PATCH_5_FIND = """\
        _tab_style(self._tab_extract_btn,  True)
        _tab_style(self._tab_cibi_btn,     False)
        _tab_style(self._tab_analysis_btn, False)
        _tab_style(self._tab_summary_btn,  False)
        _tab_style(self._tab_aiprompt_btn, False)"""

PATCH_5_REPLACE = """\
        _tab_style(self._tab_extract_btn,  True)
        _tab_style(self._tab_cibi_btn,     False)
        _tab_style(self._tab_analysis_btn, False)
        _tab_style(self._tab_summary_btn,  False)
        _tab_style(self._tab_aiprompt_btn, False)
        _tab_style(self._tab_samples_btn,  False)"""

# PATCH 6 — Register samples frame build inside _build_right's card section
PATCH_6_FIND = "        self._build_cibi_output_panel(card)"
PATCH_6_REPLACE = """\
        self._build_cibi_output_panel(card)
        self._build_samples_panel(card)"""

# PATCH 7 — Add samples to the _switch_tab hide-all block
PATCH_7_FIND = """\
        self._loader_frame.pack_forget()
        self._txt_frame.pack_forget()
        self._analysis_frame.pack_forget()
        self._summary_frame.pack_forget()
        self._aiprompt_frame.pack_forget()
        self._cibi_output_frame.pack_forget()"""

PATCH_7_REPLACE = """\
        self._loader_frame.pack_forget()
        self._txt_frame.pack_forget()
        self._analysis_frame.pack_forget()
        self._summary_frame.pack_forget()
        self._aiprompt_frame.pack_forget()
        self._cibi_output_frame.pack_forget()
        self._samples_frame.pack_forget()"""

# PATCH 8 — Add samples to the _switch_tab show block
PATCH_8_FIND = """\
        if tab == "extract":
            self._txt_frame.pack(fill="both", expand=True)
        elif tab == "analysis":
            self._analysis_frame.pack(fill="both", expand=True)
        elif tab == "summary":
            self._summary_frame.pack(fill="both", expand=True)
        elif tab == "cibi":
            self._cibi_output_frame.pack(fill="both", expand=True)
        else:
            self._aiprompt_frame.pack(fill="both", expand=True)
            self.after(50, self._chat_input.focus_set)"""

PATCH_8_REPLACE = """\
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
        else:
            self._aiprompt_frame.pack(fill="both", expand=True)
            self.after(50, self._chat_input.focus_set)"""

# PATCH 9 — Update _switch_tab's tab-style calls to include samples btn
PATCH_9_FIND = """\
        self._tab_style_fn(self._tab_extract_btn,  tab == "extract")
        self._tab_style_fn(self._tab_cibi_btn,     tab == "cibi")
        self._tab_style_fn(self._tab_analysis_btn, tab == "analysis")
        self._tab_style_fn(self._tab_summary_btn,  tab == "summary")
        self._tab_style_fn(self._tab_aiprompt_btn, tab == "aiprompt")"""

PATCH_9_REPLACE = """\
        self._tab_style_fn(self._tab_extract_btn,  tab == "extract")
        self._tab_style_fn(self._tab_cibi_btn,     tab == "cibi")
        self._tab_style_fn(self._tab_analysis_btn, tab == "analysis")
        self._tab_style_fn(self._tab_summary_btn,  tab == "summary")
        self._tab_style_fn(self._tab_aiprompt_btn, tab == "aiprompt")
        self._tab_style_fn(self._tab_samples_btn,  tab == "samples")"""

# ── F() / FMONO() helper forwarding so the mixin can call self.F() ────────────
PATCH_10_FIND = "    def _fix_windows_taskbar(self):"
PATCH_10_REPLACE = """\
    # ── Font helper forwarding for SamplesTabMixin ───────────────────────
    def F(self, size: int, weight: str = "normal") -> tuple:
        global _FONT_FAMILY
        if _FONT_FAMILY is None:
            _FONT_FAMILY = best_font()
        return (_FONT_FAMILY, size, weight)

    def FMONO(self, size: int, weight: str = "normal") -> tuple:
        import tkinter.font as tkfont
        available = set(tkfont.families())
        for f in ("JetBrains Mono", "Cascadia Code", "Consolas", "Courier New"):
            if f in available:
                return (f, size, weight)
        return ("Courier New", size, weight)

    def _fix_windows_taskbar(self):"""

# ─────────────────────────────────────────────────────────────────────────────

ALL_PATCHES = [
    ("1 — import SamplesTabMixin",          PATCH_1_FIND,  PATCH_1_REPLACE),
    ("2 — mixin in class definition",       PATCH_2_FIND,  PATCH_2_REPLACE),
    ("3 — _samples_init_state() in __init__",PATCH_3_FIND, PATCH_3_REPLACE),
    ("4 — Samples tab button",              PATCH_4_FIND,  PATCH_4_REPLACE),
    ("5 — tab style init",                  PATCH_5_FIND,  PATCH_5_REPLACE),
    ("6 — build_samples_panel call",        PATCH_6_FIND,  PATCH_6_REPLACE),
    ("7 — hide samples frame",              PATCH_7_FIND,  PATCH_7_REPLACE),
    ("8 — show samples frame in switch_tab",PATCH_8_FIND,  PATCH_8_REPLACE),
    ("9 — tab_style_fn for samples btn",    PATCH_9_FIND,  PATCH_9_REPLACE),
    ("10 — F/FMONO helper methods",         PATCH_10_FIND, PATCH_10_REPLACE),
]


def apply_patches(app_path: Path = APP_PATH) -> None:
    if not app_path.exists():
        print(f"ERROR: {app_path} not found.")
        return

    bak = app_path.with_suffix(".py.bak")
    shutil.copy2(app_path, bak)
    print(f"✓ Backup saved: {bak.name}\n")

    content  = app_path.read_text(encoding="utf-8")
    applied  = 0
    skipped  = 0
    failed   = 0

    for name, find_str, replace_str in ALL_PATCHES:
        if find_str in content:
            content = content.replace(find_str, replace_str, 1)
            print(f"  ✓ Patch {name}")
            applied += 1
        elif replace_str.strip() in content or find_str.replace(find_str, replace_str) in content:
            print(f"  ℹ Already applied — Patch {name}")
            skipped += 1
        else:
            print(f"  ✗ FAILED — Patch {name}")
            print(f"      Search string not found. Apply manually.")
            failed += 1

    app_path.write_text(content, encoding="utf-8")

    print(f"\n{'='*55}")
    print(f"  {applied} patches applied   {skipped} already present   {failed} failed")
    print(f"  Output: {app_path}")
    print(f"  Backup: {bak}")
    if failed == 0:
        print(f"\n  ✅  All patches applied successfully!")
        print(f"  Place samples_tab.py in the same folder as app.py")
        print(f"  and restart the app.")
    else:
        print(f"\n  ⚠ {failed} patch(es) failed — check the output above.")
        print(f"  The failed patches are documented in apply_samples_tab_patch.py")
        print(f"  with clear FIND/REPLACE strings for manual application.")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Wire the Samples tab into app.py")
    parser.add_argument("--file", default=str(APP_PATH), help="Path to app.py")
    args = parser.parse_args()
    apply_patches(Path(args.file))
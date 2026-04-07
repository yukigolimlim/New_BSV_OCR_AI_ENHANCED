"""
ui_summary.py — DocExtract Pro
=================================
Summary tab builder and all chart/parse helpers
attached to DocExtractorApp.
"""
import re
import tkinter as tk
import customtkinter as ctk
from app_constants import *

def _build_summary_panel(self, parent):
    self._summary_frame = tk.Frame(parent, bg=CARD_WHITE)
    sb = tk.Scrollbar(self._summary_frame, relief="flat",
                      troughcolor=OFF_WHITE, bg=BORDER_LIGHT, width=8, bd=0)
    sb.pack(side="right", fill="y")
    self._summary_canvas = tk.Canvas(
        self._summary_frame, bg=CARD_WHITE,
        highlightthickness=0, yscrollcommand=sb.set
    )
    self._summary_canvas.pack(side="left", fill="both", expand=True)
    sb.config(command=self._summary_canvas.yview)
    self._summary_inner = tk.Frame(self._summary_canvas, bg=CARD_WHITE)
    self._summary_canvas_win = self._summary_canvas.create_window(
        (0, 0), window=self._summary_inner, anchor="nw"
    )
    self._summary_inner.bind(
        "<Configure>",
        lambda e: self._summary_canvas.configure(
            scrollregion=self._summary_canvas.bbox("all")
        )
    )
    self._summary_canvas.bind(
        "<Configure>",
        lambda e: self._summary_canvas.itemconfig(
            self._summary_canvas_win, width=e.width
        )
    )
    self._summary_canvas.bind(
        "<MouseWheel>",
        lambda e: self._summary_canvas.yview_scroll(
            int(-1 * (e.delta / 120)), "units"
        )
    )
    self._summary_placeholder()

def _summary_placeholder(self):
    for w in self._summary_inner.winfo_children():
        w.destroy()
    outer = tk.Frame(self._summary_inner, bg=CARD_WHITE)
    outer.pack(expand=True, fill="both", pady=80)
    tk.Label(outer, text="📊", font=("Segoe UI Emoji", 44),
             fg=BORDER_MID, bg=CARD_WHITE).pack()
    tk.Label(outer, text="Summary & Charts",
             font=F(15, "bold"), fg=TXT_MUTED, bg=CARD_WHITE).pack(pady=(10, 4))
    tk.Label(outer,
             text="Extract a document and run  Analyze CIBI  to see\n"
                  "credit score gauges, ratio charts, and risk breakdown.",
             font=F(10), fg=TXT_MUTED, bg=CARD_WHITE, justify="center").pack()

# ══════════════════════════════════════════════════════════════════════
#  UPDATED _populate_summary (Change 5)
# ══════════════════════════════════════════════════════════════════════
def _populate_summary(self, analysis_text: str):
    for w in self._summary_inner.winfo_children():
        w.destroy()

    pad = dict(padx=24)

    # ══════════════════════════════════════════════════════════════════
    # SECTION A — VERDICT + APPLICANT (top row, always shown)
    # ══════════════════════════════════════════════════════════════════
    row1 = tk.Frame(self._summary_inner, bg=CARD_WHITE)
    row1.pack(fill="x", pady=(18, 0), **pad)
    verdict, verdict_color = self._parse_verdict(analysis_text)
    self._summary_verdict_card(row1, verdict, verdict_color)
    applicant, loan_amt = self._parse_applicant(analysis_text)
    self._summary_info_card(row1, "👤  Applicant", applicant, loan_amt)

    # ══════════════════════════════════════════════════════════════════
    # SECTION B — CIBI MODE FINDINGS
    # ══════════════════════════════════════════════════════════════════
    tk.Frame(self._summary_inner, bg=BORDER_LIGHT, height=1).pack(
        fill="x", padx=24, pady=(16, 0))
    self._summary_section_label("📋  CIBI MODE FINDINGS", NAVY_DEEP)

    cibi_row = tk.Frame(self._summary_inner, bg=CARD_WHITE)
    cibi_row.pack(fill="x", pady=(8, 0), **pad)

    # Bank CI card
    self._summary_cibi_ci_card(cibi_row)

    # Loan tier card
    self._summary_cibi_tier_card(cibi_row)

    # Documents uploaded card
    self._summary_cibi_docs_card(cibi_row)

    # ══════════════════════════════════════════════════════════════════
    # SECTION C — CIBI ANALYSIS RESULTS
    # ══════════════════════════════════════════════════════════════════
    tk.Frame(self._summary_inner, bg=BORDER_LIGHT, height=1).pack(
        fill="x", padx=24, pady=(16, 0))
    self._summary_section_label("🏦  CIBI ANALYSIS RESULTS", NAVY_DEEP)

    # Cash flow row
    cf_row = tk.Frame(self._summary_inner, bg=CARD_WHITE)
    cf_row.pack(fill="x", pady=(8, 0), **pad)
    income_data = self._parse_income_expense(analysis_text)
    self._summary_income_chart(cf_row, income_data)
    ratios = self._parse_ratios(analysis_text)
    self._summary_ratios_chart(cf_row, ratios)

    # Score + risk row
    tk.Frame(self._summary_inner, bg=CARD_WHITE, height=8).pack()
    score_row = tk.Frame(self._summary_inner, bg=CARD_WHITE)
    score_row.pack(fill="x", pady=(0, 0), **pad)
    score, stage, ecl = self._parse_score(analysis_text)
    self._summary_score_gauge(score_row, score, stage, ecl)
    flags = self._parse_flags(analysis_text)
    self._summary_risk_donut(score_row, flags)

    # DSR + NDI indicators
    tk.Frame(self._summary_inner, bg=CARD_WHITE, height=8).pack()
    dsr_row = tk.Frame(self._summary_inner, bg=CARD_WHITE)
    dsr_row.pack(fill="x", pady=(0, 0), **pad)
    self._summary_dsr_ndi_card(dsr_row, analysis_text)

    # Risk flags list
    if flags:
        tk.Frame(self._summary_inner, bg=CARD_WHITE, height=6).pack()
        self._summary_flags_card(analysis_text)

    # ══════════════════════════════════════════════════════════════════
    # SECTION D — RECOMMENDATION
    # ══════════════════════════════════════════════════════════════════
    tk.Frame(self._summary_inner, bg=BORDER_LIGHT, height=1).pack(
        fill="x", padx=24, pady=(16, 0))
    row4 = tk.Frame(self._summary_inner, bg=CARD_WHITE)
    row4.pack(fill="x", pady=(10, 20), **pad)
    product = self._parse_product(analysis_text)
    self._summary_product_card(row4, product)

def _summary_section_label(self, text: str, color=None):
    """Render a bold section divider label in the summary panel."""
    color = color or NAVY_PALE
    row = tk.Frame(self._summary_inner, bg=CARD_WHITE)
    row.pack(fill="x", padx=24, pady=(4, 2))
    tk.Frame(row, bg=LIME_MID, width=4, height=16).pack(side="left", padx=(0, 8))
    tk.Label(row, text=text, font=F(8, "bold"),
             fg=color, bg=CARD_WHITE).pack(side="left", anchor="w")

# ── CIBI Mode: Bank CI card ───────────────────────────────────────────
def _summary_cibi_ci_card(self, parent):
    result  = self._cibi_bank_ci_result
    verdict = result.get("verdict", "")
    summary = result.get("summary", "")
    proceed = result.get("proceed", None)

    if not result:
        bg = NAVY_MIST; icon = "—"; label = "Bank CI"; sub = "Not yet evaluated"
        v_color = TXT_MUTED
    elif verdict == "GOOD":
        bg = LIME_MIST; icon = "✅"; label = "Bank CI: GOOD"; sub = summary or "No adverse findings"
        v_color = LIME_DARK
    elif verdict == "BAD":
        bg = "#FEE2E2"; icon = "❌"; label = "Bank CI: ADVERSE"; sub = summary or "Negative findings detected"
        v_color = ACCENT_RED
    else:
        bg = BUBBLE_SYS; icon = "⚠️"; label = "Bank CI: UNCERTAIN"; sub = summary or "Manual review needed"
        v_color = ACCENT_GOLD

    card = self._card(parent, bg=bg, side="left", fill="both",
                      expand=True, padx=(0, 8))
    tk.Frame(card, bg=bg, height=6).pack()
    tk.Label(card, text=icon, font=("Segoe UI Emoji", 22),
             fg=v_color, bg=bg).pack(pady=(4, 0))
    tk.Label(card, text=label, font=F(9, "bold"),
             fg=v_color, bg=bg).pack(pady=(2, 0))
    tk.Label(card, text=sub, font=F(8), fg=TXT_SOFT, bg=bg,
             wraplength=200, justify="center").pack(padx=10, pady=(2, 8))

    if proceed is not None:
        pill_bg = LIME_MIST if proceed else "#FEE2E2"
        pill_fg = LIME_DARK if proceed else ACCENT_RED
        pill_txt = "PROCEED ✓" if proceed else "STOP ✗"
        tk.Label(card, text=pill_txt, font=F(8, "bold"),
                 fg=pill_fg, bg=pill_bg,
                 padx=10, pady=3).pack(pady=(0, 10))

# ── CIBI Mode: Loan Tier card ─────────────────────────────────────────
def _summary_cibi_tier_card(self, parent):
    tier = self._cibi_loan_tier
    has_cic = self._cibi_has_cic

    if tier == "above_100k":
        bg = LIME_MIST; icon = "⬆"; label = "Above ₱100,000"
        sub = "CIC required — uploaded" if has_cic else "CIC required — not uploaded"
        t_color = NAVY_DEEP
    elif tier == "below_100k":
        bg = LIME_MIST; icon = "⬇"; label = "Below ₱100,000"
        sub = "Bank CI only required"
        t_color = LIME_DARK
    else:
        bg = NAVY_MIST; icon = "?"; label = "Tier Unknown"
        sub = "Run Stage 1 to determine"
        t_color = TXT_MUTED

    card = self._card(parent, bg=bg, side="left", fill="both",
                      expand=True, padx=(0, 8))
    tk.Frame(card, bg=bg, height=6).pack()
    tk.Label(card, text=icon, font=F(22, "bold"),
             fg=t_color, bg=bg).pack(pady=(4, 0))
    tk.Label(card, text="Loan Tier", font=F(7, "bold"),
             fg=TXT_MUTED, bg=bg).pack()
    tk.Label(card, text=label, font=F(10, "bold"),
             fg=t_color, bg=bg).pack(pady=(2, 0))
    tk.Label(card, text=sub, font=F(8), fg=TXT_SOFT, bg=bg,
             wraplength=180, justify="center").pack(padx=10, pady=(2, 10))

    cic_icon = "✅" if has_cic else "—"
    cic_color = LIME_DARK if has_cic else TXT_MUTED
    tk.Label(card, text=f"CIC: {cic_icon}", font=F(8, "bold"),
             fg=cic_color, bg=bg).pack(pady=(0, 8))

# ── CIBI Mode: Documents uploaded card ───────────────────────────────
def _summary_cibi_docs_card(self, parent):
    slots = self._cibi_slots
    populated = self._summary_cibi_populated

    card = self._card(parent, bg=CARD_WHITE, side="left", fill="both",
                      expand=True, padx=(0, 0))
    tk.Label(card, text="📁  Documents", font=F(8, "bold"),
             fg=NAVY_PALE, bg=CARD_WHITE, anchor="w").pack(
        anchor="w", padx=14, pady=(10, 4))

    doc_defs = [
        ("CIC",     "📋 CIC Report"),
        ("BANK_CI", "🏦 Bank CI"),
        ("PAYSLIP", "💵 Payslip"),
        ("ITR",     "📊 ITR"),
        ("SALN",    "📄 SALN"),
    ]
    for key, label in doc_defs:
        slot = slots.get(key, {})
        has_file = bool(slot.get("path"))
        has_text = bool(slot.get("text"))
        if has_text:
            icon = "✅"; color = LIME_DARK
        elif has_file:
            icon = "⏳"; color = ACCENT_GOLD
        else:
            icon = "○";  color = TXT_MUTED

        row = tk.Frame(card, bg=CARD_WHITE)
        row.pack(fill="x", padx=14, pady=1)
        tk.Label(row, text=icon, font=F(9),
                 fg=color, bg=CARD_WHITE, width=2).pack(side="left")
        name_text = label
        if has_file and slot.get("path"):
            from pathlib import Path as _P
            short = _P(slot["path"]).name
            name_text = f"{label}  ({short[:20]}…)" if len(short) > 20 else f"{label}  ({short})"
        tk.Label(row, text=name_text, font=F(8),
                 fg=TXT_NAVY if has_file else TXT_MUTED,
                 bg=CARD_WHITE).pack(side="left", padx=(4, 0))

    # Populated status
    pop_bg = LIME_MIST if populated else NAVY_MIST
    pop_fg = LIME_DARK if populated else TXT_MUTED
    pop_txt = "✅  CIBI Excel Populated" if populated else "⏳  Not yet populated"
    tk.Frame(card, bg=BORDER_LIGHT, height=1).pack(
        fill="x", padx=10, pady=(6, 4))
    tk.Label(card, text=pop_txt, font=F(8, "bold"),
             fg=pop_fg, bg=pop_bg,
             padx=10, pady=4).pack(fill="x", padx=10, pady=(0, 10))

# ── CIBI Analysis: DSR + NDI indicator card ───────────────────────────
def _summary_dsr_ndi_card(self, parent, text: str):
    """Render DSR and NDI as coloured pill indicators side by side."""

    def _extract(patterns, t):
        for pat in patterns:
            m = re.search(pat, t, re.I)
            if m:
                try:    return float(m.group(1).replace(",", "").replace("%", "").strip())
                except: pass
        return None

    dsr = _extract([
        r'DSR\s*[=:\-]\s*([\d.]+)',
        r'[Dd]ebt.?[Ss]ervice\s*[Rr]atio\s*[=:\-]\s*([\d.]+)',
    ], text)

    ndi = _extract([
        r'[Nn]et\s*[Dd]isposable\s*[Ii]ncome\s*[=:\-]\s*₱?\s*([\d,]+)',
        r'[Nn]et\s*[Dd]isposable\s*[=:\-]\s*₱?\s*([\d,]+)',
        r'[Ff]ree\s*[Cc]ash\s*[=:\-]\s*₱?\s*([\d,]+)',
    ], text)

    # DSR thresholds
    if dsr is None:
        dsr_bg = NAVY_MIST; dsr_fg = TXT_MUTED; dsr_lbl = "DSR\nN/A"
    else:
        dsr_v = dsr if dsr <= 1 else dsr / 100
        if dsr_v <= 0.35:
            dsr_bg = LIME_MIST; dsr_fg = LIME_DARK
            dsr_lbl = f"DSR\n{dsr_v*100:.1f}%\n✅ Low Risk"
        elif dsr_v <= 0.50:
            dsr_bg = BUBBLE_SYS; dsr_fg = ACCENT_GOLD
            dsr_lbl = f"DSR\n{dsr_v*100:.1f}%\n⚠ Moderate"
        else:
            dsr_bg = "#FEE2E2"; dsr_fg = ACCENT_RED
            dsr_lbl = f"DSR\n{dsr_v*100:.1f}%\n❌ High Risk"

    # NDI display
    if ndi is None:
        ndi_bg = NAVY_MIST; ndi_fg = TXT_MUTED; ndi_lbl = "NDI\nN/A"
    elif ndi >= 0:
        ndi_bg = LIME_MIST; ndi_fg = LIME_DARK
        ndi_lbl = f"NDI\n₱{ndi:,.0f}\n✅ Positive"
    else:
        ndi_bg = "#FEE2E2"; ndi_fg = ACCENT_RED
        ndi_lbl = f"NDI\n₱{ndi:,.0f}\n❌ Negative"

    outer = tk.Frame(self._summary_inner, bg=CARD_WHITE)
    outer.pack(fill="x", padx=24, pady=(0, 4))

    for lbl_txt, bg, fg in [(dsr_lbl, dsr_bg, dsr_fg),
                              (ndi_lbl, ndi_bg, ndi_fg)]:
        pill = tk.Frame(outer, bg=bg,
                        highlightbackground=BORDER_MID,
                        highlightthickness=1)
        pill.pack(side="left", padx=(0, 8), pady=4, ipadx=20, ipady=8)
        tk.Label(pill, text=lbl_txt, font=F(9, "bold"),
                 fg=fg, bg=bg, justify="center").pack(padx=16, pady=6)

# ── CIBI Analysis: Risk flags list card ──────────────────────────────
def _summary_flags_card(self, analysis_text: str):
    """Render a compact risk flags card below the charts."""
    flags = self._parse_flags(analysis_text)
    if not flags:
        return

    outer = tk.Frame(self._summary_inner, bg=BORDER_LIGHT, padx=1, pady=1)
    outer.pack(fill="x", padx=24, pady=(0, 4))
    card = tk.Frame(outer, bg="#FFF8F0")
    card.pack(fill="both", expand=True)

    tk.Label(card, text="⚑  Risk Flags Identified",
             font=F(8, "bold"), fg=ACCENT_GOLD,
             bg="#FFF8F0", anchor="w", padx=14).pack(
        anchor="w", pady=(8, 4))

    for flag in flags:
        row = tk.Frame(card, bg="#FFF8F0")
        row.pack(fill="x", padx=14, pady=1)

        flag_lower = flag.lower()
        if any(w in flag_lower for w in ("negative", "overdue", "delinquent",
                                          "default", "fraud", "bad")):
            dot_color = ACCENT_RED
        elif any(w in flag_lower for w in ("low", "insufficient", "unstable",
                                            "irregular", "high dsr")):
            dot_color = ACCENT_GOLD
        else:
            dot_color = NAVY_PALE

        tk.Label(row, text="•", font=F(10, "bold"),
                 fg=dot_color, bg="#FFF8F0").pack(side="left", padx=(0, 6))
        tk.Label(row, text=flag, font=F(8),
                 fg=TXT_NAVY, bg="#FFF8F0",
                 wraplength=700, anchor="w",
                 justify="left").pack(side="left", fill="x")

    tk.Frame(card, bg="#FFF8F0", height=8).pack()

def _summary_score_gauge(self, parent, score, stage, ecl):
    try:
        import matplotlib; matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        import numpy as np
        fig, ax = plt.subplots(figsize=(3.6, 2.6), facecolor=CARD_WHITE)
        ax.set_facecolor(CARD_WHITE); ax.set_aspect("equal")
        bands = [(300,570,"#E74C3C","Poor"),(570,691,"#E67E22","Fair"),
                 (691,781,"#F0A800","Good"),(781,900,"#22C870","Excellent")]
        total = 600
        for lo, hi, color, _ in bands:
            t1 = 180-((lo-300)/total*180); t2 = 180-((hi-300)/total*180)
            ax.add_patch(mpatches.Wedge((0,0),1.0,t2,t1,width=0.28,color=color,alpha=0.88))
        sv = max(300, min(900, int(score) if score else 600))
        a  = np.radians(180-((sv-300)/total*180))
        ax.annotate("",xy=(0.72*np.cos(a),0.72*np.sin(a)),xytext=(0,0),
                    arrowprops=dict(arrowstyle="-|>",color=NAVY_DEEP,lw=2.2,mutation_scale=14))
        ax.add_patch(plt.Circle((0,0),0.07,color=NAVY_DEEP,zorder=5))
        ax.text(0,-0.22,score or "—",ha="center",va="center",fontsize=18,
                fontweight="bold",color=self._score_color(score))
        ax.text(0,-0.42,"Credit Score",ha="center",va="center",fontsize=7,color="#9AAACE")
        for lo,hi,color,label in bands:
            mid=(lo+hi)/2; aa=np.radians(180-((mid-300)/total*180))
            ax.text(1.12*np.cos(aa),1.12*np.sin(aa),label,ha="center",va="center",
                    fontsize=5.5,color=color,fontweight="bold")
        meta=[]
        if stage: meta.append(f"Stage {stage}")
        if ecl:   meta.append(f"ECL {ecl}")
        if meta:  ax.text(0,-0.60,"  ·  ".join(meta),ha="center",va="center",fontsize=6.5,color=NAVY_MID)
        ax.set_xlim(-1.35,1.35); ax.set_ylim(-0.75,1.15); ax.axis("off")
        fig.tight_layout(pad=0.2)
    except Exception:
        fig, ax = plt.subplots(figsize=(3.6,2.6),facecolor=CARD_WHITE)
        ax.text(0.5,0.5,f"Score:\n{score or '—'}",ha="center",va="center",fontsize=14,
                color=self._score_color(score),fontweight="bold",transform=ax.transAxes)
        ax.axis("off")
    self._embed_chart(parent, fig, "🎯  Credit Score Gauge", NAVY_MIST, side="left")

def _summary_ratios_chart(self, parent, ratios):
    try:
        import matplotlib; matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots(figsize=(3.8,2.6),facecolor=CARD_WHITE)
        ax.set_facecolor(CARD_WHITE)
        if ratios:
            labels=list(ratios.keys())
            raw_vals=[]
            for v in ratios.values():
                try: raw_vals.append(float(str(v).replace("x","").replace("%","").strip()))
                except: raw_vals.append(0.0)
            max_v=max(raw_vals) if max(raw_vals)>0 else 1
            norm=[v/max_v for v in raw_vals]
            colors=[]
            for label,val in zip(labels,raw_vals):
                if label=="DSR": colors.append("#22C870" if val<=0.35 else "#F0A800" if val<=0.50 else "#E74C3C")
                else: colors.append(NAVY_MID)
            y_pos=range(len(labels))
            bars=ax.barh(list(y_pos),norm,color=colors,height=0.55,alpha=0.85)
            for bar,real_v,raw_str in zip(bars,raw_vals,ratios.values()):
                ax.text(bar.get_width()+0.02,bar.get_y()+bar.get_height()/2,str(raw_str),
                        va="center",ha="left",fontsize=7.5,color=NAVY_DEEP,fontweight="bold")
            ax.set_yticks(list(y_pos)); ax.set_yticklabels(labels,fontsize=8,color=TXT_NAVY)
            ax.set_xlim(0,1.35); ax.set_xticks([])
            ax.spines[["top","right","bottom"]].set_visible(False)
            ax.spines["left"].set_color(BORDER_MID); ax.tick_params(axis="y",length=0)
            ax.set_title("Key Financial Ratios",fontsize=8,color=NAVY_PALE,pad=6,loc="left")
        else:
            ax.text(0.5,0.5,"No ratios detected",ha="center",va="center",fontsize=9,color=TXT_MUTED,transform=ax.transAxes); ax.axis("off")
        fig.tight_layout(pad=0.4)
    except Exception:
        fig, ax = plt.subplots(figsize=(3.8,2.6),facecolor=CARD_WHITE)
        ax.text(0.5,0.5,"Ratios\nChart",ha="center",va="center",fontsize=10,color=TXT_MUTED,transform=ax.transAxes); ax.axis("off")
    self._embed_chart(parent, fig, "📐  Key Ratios", CARD_WHITE, side="left", expand=True)

def _summary_risk_donut(self, parent, flags):
    try:
        import matplotlib; matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots(figsize=(3.6,2.6),facecolor=CARD_WHITE)
        ax.set_facecolor(CARD_WHITE)
        n=len(flags)
        high=sum(1 for f in flags if any(w in f.lower() for w in ("negative","overdue","delinquent","default","fraud","criminal","bad")))
        medium=sum(1 for f in flags if any(w in f.lower() for w in ("low","insufficient","unstable","irregular","seasonal","high dsr")))
        low=max(0,n-high-medium)
        if n==0:
            sizes=[1]; colors=["#22C870"]; labels=["No Flags"]; center_txt="✓ Clean"; center_col="#22C870"
        else:
            sizes,colors,labels=[],[],[]
            if high:   sizes.append(high);  colors.append("#E74C3C"); labels.append(f"High ({high})")
            if medium: sizes.append(medium);colors.append("#F0A800"); labels.append(f"Medium ({medium})")
            if low:    sizes.append(low);   colors.append(NAVY_MID);  labels.append(f"Low ({low})")
            center_txt=str(n); center_col="#E74C3C" if high else "#F0A800" if medium else NAVY_MID
        wedges,_=ax.pie(sizes,colors=colors,startangle=90,wedgeprops=dict(width=0.45,edgecolor=CARD_WHITE,linewidth=2))
        ax.text(0,0.06,center_txt,ha="center",va="center",fontsize=20,fontweight="bold",color=center_col)
        ax.text(0,-0.22,"flag" if n==1 else "flags",ha="center",va="center",fontsize=7.5,color=TXT_SOFT)
        ax.legend(wedges,labels,loc="lower center",fontsize=6.5,frameon=False,ncol=len(labels),bbox_to_anchor=(0.5,-0.12))
        ax.set_title("Risk Flag Breakdown",fontsize=8,color=NAVY_PALE,pad=4,loc="left")
        fig.tight_layout(pad=0.3)
    except Exception:
        fig, ax = plt.subplots(figsize=(3.6,2.6),facecolor=CARD_WHITE)
        ax.text(0.5,0.5,f"{len(flags)} Flags",ha="center",va="center",fontsize=12,color=TXT_SOFT,transform=ax.transAxes); ax.axis("off")
    self._embed_chart(parent, fig, "⚑  Risk Flags", "#FFF8F0", side="left")

def _summary_income_chart(self, parent, income_data):
    try:
        import matplotlib; matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import numpy as np
        fig, ax = plt.subplots(figsize=(3.8,2.6),facecolor=CARD_WHITE)
        ax.set_facecolor(CARD_WHITE)
        keys=list(income_data.keys()); values=list(income_data.values())
        if keys and any(v>0 for v in values):
            colors=[]
            for k in keys:
                kl=k.lower()
                if any(w in kl for w in ("income","salary","revenue","gross")): colors.append("#22C870")
                elif any(w in kl for w in ("expense","loan","obligation","payment")): colors.append("#E74C3C")
                else: colors.append(NAVY_MID)
            x=np.arange(len(keys))
            bars=ax.bar(x,values,color=colors,width=0.55,alpha=0.88,edgecolor=CARD_WHITE,linewidth=1.5)
            for bar,val in zip(bars,values):
                label=f"₱{val/1000:.0f}k" if val>=1000 else f"₱{val:.0f}"
                ax.text(bar.get_x()+bar.get_width()/2,bar.get_height()+max(values)*0.03,
                        label,ha="center",va="bottom",fontsize=6.5,color=NAVY_DEEP,fontweight="bold")
            short_keys=[k[:10]+"…" if len(k)>10 else k for k in keys]
            ax.set_xticks(x); ax.set_xticklabels(short_keys,fontsize=7,rotation=15,ha="right",color=TXT_NAVY)
            ax.set_yticks([]); ax.spines[["top","right","left"]].set_visible(False)
            ax.spines["bottom"].set_color(BORDER_MID); ax.tick_params(axis="x",length=0)
            ax.set_title("Income & Expense Overview",fontsize=8,color=NAVY_PALE,pad=6,loc="left")
        else:
            ax.text(0.5,0.5,"No financial\nfigures detected",ha="center",va="center",fontsize=9,color=TXT_MUTED,transform=ax.transAxes); ax.axis("off")
        fig.tight_layout(pad=0.4)
    except Exception:
        fig, ax = plt.subplots(figsize=(3.8,2.6),facecolor=CARD_WHITE)
        ax.text(0.5,0.5,"Income Chart",ha="center",va="center",fontsize=10,color=TXT_MUTED,transform=ax.transAxes); ax.axis("off")
    self._embed_chart(parent, fig, "💰  Income & Expenses", CARD_WHITE, side="left", expand=True)

def _embed_chart(self, parent, fig, title, bg, side="left", expand=False):
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.pyplot as plt
    outer = tk.Frame(parent, bg=BORDER_LIGHT, padx=1, pady=1)
    outer.pack(side=side, fill="both", expand=expand, padx=(0,10) if side=="left" and not expand else 0)
    card = tk.Frame(outer, bg=bg); card.pack(fill="both", expand=True)
    tk.Label(card, text=title, font=F(8,"bold"), fg=NAVY_PALE, bg=bg, anchor="w", padx=12).pack(fill="x", pady=(8,2))
    canvas = FigureCanvasTkAgg(fig, master=card); canvas.draw()
    widget = canvas.get_tk_widget(); widget.configure(bg=bg, highlightthickness=0)
    widget.pack(fill="both", expand=True, padx=6, pady=(0,8))
    plt.close(fig)

def _card(self, parent, bg=None, side="left", fill="both", expand=False, padx=(0,10), pady=0, minwidth=0):
    bg = bg or CARD_WHITE
    outer = tk.Frame(parent, bg=BORDER_LIGHT, padx=1, pady=1)
    outer.pack(side=side, fill=fill, expand=expand, padx=padx, pady=pady)
    if minwidth: outer.config(width=minwidth)
    inner = tk.Frame(outer, bg=bg); inner.pack(fill="both", expand=True)
    return inner

def _summary_verdict_card(self, parent, verdict, color):
    card = self._card(parent, bg=color, side="left", fill="y", expand=False, padx=(0,12), minwidth=190)
    tk.Frame(card, bg=color).pack(expand=True, fill="both")
    tk.Label(card, text="VERDICT", font=F(8,"bold"), fg=WHITE, bg=color).pack()
    display = verdict.replace("CONDITIONALLY ", "COND.\n")
    tk.Label(card, text=display, font=F(16,"bold"), fg=WHITE, bg=color,
             wraplength=170, justify="center", padx=18, pady=6).pack()
    tk.Frame(card, bg=color).pack(expand=True, fill="both")

def _summary_info_card(self, parent, title, applicant, loan_amt):
    card = self._card(parent, side="left", fill="both", expand=True, padx=(0,0))
    tk.Label(card, text=title, font=F(8,"bold"), fg=NAVY_PALE, bg=CARD_WHITE, anchor="w").pack(anchor="w", padx=16, pady=(12,2))
    tk.Label(card, text=applicant, font=FMONO(13,"bold"), fg=NAVY_DEEP, bg=CARD_WHITE, anchor="w", wraplength=340).pack(anchor="w", padx=16)
    if loan_amt:
        tk.Label(card, text=f"Loan Amount:  {loan_amt}", font=F(10), fg=TXT_SOFT, bg=CARD_WHITE, anchor="w").pack(anchor="w", padx=16, pady=(6,12))

def _summary_product_card(self, parent, product):
    card = self._card(parent, bg=LIME_MIST, side="left", fill="x", expand=True, padx=(0,0))
    tk.Label(card, text="🏦  TOP RECOMMENDED PRODUCT", font=F(8,"bold"), fg=LIME_DARK, bg=LIME_MIST, anchor="w").pack(anchor="w", padx=16, pady=(12,4))
    tk.Label(card, text=product if product else "See the Loan Analysis tab for full product recommendations.",
             font=FMONO(11,"bold") if product else F(10),
             fg=NAVY_DEEP if product else TXT_MUTED,
             bg=LIME_MIST, anchor="w", wraplength=880, justify="left").pack(anchor="w", padx=16, pady=(0,12))

def _pill(self, parent, text, bg, fg):
    tk.Label(parent, text=text, font=F(7,"bold"), fg=fg, bg=bg, padx=8, pady=2).pack(side="left", padx=(0,6))

def _parse_income_expense(self, text):
    data = {}
    patterns = {
        "Monthly Income":   r'(?:monthly\s+income|gross\s+income|net\s+income)\s*[=:\-]\s*₱?\s*([\d,]+(?:\.\d+)?)',
        "Monthly Expenses": r'(?:monthly\s+expenses?|total\s+expenses?|household\s+expenses?)\s*[=:\-]\s*₱?\s*([\d,]+(?:\.\d+)?)',
        "Loan Payment":     r'(?:monthly\s+(?:amortization|payment|installment|loan\s+payment))\s*[=:\-]\s*₱?\s*([\d,]+(?:\.\d+)?)',
        "Net Disposable":   r'(?:net\s+disposable|disposable\s+income|free\s+cash)\s*[=:\-]\s*₱?\s*([\d,]+(?:\.\d+)?)',
    }
    for label, pattern in patterns.items():
        m = re.search(pattern, text, re.I)
        if m:
            try: data[label] = float(m.group(1).replace(",",""))
            except: pass
    return data

def _parse_verdict(self, text):
    if re.search(r'CONDITIONALLY\s+APPROVE', text, re.I): return "CONDITIONALLY APPROVE", ACCENT_GOLD
    if re.search(r'\bDECLINE\b', text, re.I):             return "DECLINE", ACCENT_RED
    if re.search(r'\bAPPROVE\b', text, re.I):             return "APPROVE", ACCENT_SUCCESS
    return "PENDING", TXT_SOFT

def _parse_applicant(self, text):
    name=""; amt=""
    m=re.search(r'(?:full\s+name|name)\s*[:\-]\s*([^\n,]+)',text,re.I)
    if m: name=m.group(1).strip()
    m2=re.search(r'(?:loan\s+amount|amount\s+(?:applied|requested|of\s+loan))[^\n₱]*₱\s*([\d,]+(?:\.\d+)?)',text,re.I)
    if m2: amt=f"₱{m2.group(1)}"
    else:
        m3=re.search(r'₱\s*([\d,]{6,})',text)
        if m3: amt=f"₱{m3.group(1)}"
    return name or "See full report", amt

def _parse_score(self, text):
    score=stage=ecl=""
    m=re.search(r'total\s+score\s*[=:\-]\s*(\d{3})',text,re.I)
    if m: score=m.group(1)
    m2=re.search(r'stage\s*[:\-]?\s*([123])\b',text,re.I)
    if m2: stage=m2.group(1)
    m3=re.search(r'ECL\s*(?:%|percentage)?\s*[=:\-]\s*([\d.]+\s*%?)',text,re.I)
    if m3:
        ecl=m3.group(1).strip()
        if "%" not in ecl: ecl+= "%"
    return score, stage, ecl

def _parse_ratios(self, text):
    ratios={}
    patterns={"DSR":r'DSR\s*[=:\-]\s*([\d.]+)',"Liquidity":r'[Ll]iquidity\s*[=:\-]\s*([\d.]+)',
              "D/E":r'D[/\\]?E\s*[=:\-]\s*([\d.]+)',"ROE":r'ROE\s*[=:\-]\s*([\d.]+\s*%?)',
              "Asset Turn":r'[Aa]sset\s+[Tt]urn(?:over)?\s*[=:\-]\s*([\d.]+)'}
    for label,pattern in patterns.items():
        m=re.search(pattern,text)
        if m:
            val=m.group(1).strip()
            if label in ("DSR","Liquidity","D/E","Asset Turn") and "%" not in val: val+="x"
            ratios[label]=val
    return ratios

def _parse_flags(self, text):
    flags=[]
    m=re.search(r'8\.\s+RISK\s+FLAGS?\s*\n(.*?)(?=\n\s*9\.|\Z)',text,re.I|re.DOTALL)
    if not m: return flags
    for line in m.group(1).splitlines():
        line=line.strip()
        if re.match(r'^[•\-\*–]',line):
            clean=re.sub(r'^[•\-\*–]\s*','',line).strip()
            if clean: flags.append(clean)
        elif len(line)>20 and not set(line).issubset({"─","—","-","=","_"}) and not line[0].isdigit():
            flags.append(line)
    return flags[:8]

def _parse_product(self, text):
    m=re.search(r'(?:RECOMMENDED\s+PRODUCTS?.*?\n\s*[•\-\*]?\s*(?:Product\s+name\s*[:\-]\s*)?|Product\s*:\s*)([^\n]+)',text,re.I|re.DOTALL)
    if m:
        name=m.group(1).strip().lstrip("•-* ").split("\n")[0].strip()
        if 3<len(name)<120: return name
    m2=re.search(r'((?:Micro|Small|Medium|Salary|Housing|Agricultural|LIPPUP|Center)[^\n]{5,80}(?:Loan|Credit)[^\n]*)',text,re.I)
    if m2: return m2.group(1).strip()
    return ""

def _score_color(self, score_str):
    try:
        s=int(score_str)
        if s>=781: return ACCENT_SUCCESS
        if s>=691: return ACCENT_GOLD
        if s>=571: return "#E67E22"
        return ACCENT_RED
    except: return TXT_SOFT

# ══════════════════════════════════════════════════════════════════════
#  AI-PROMPT CHATBOT TAB
# ══════════════════════════════════════════════════════════════════════

# ── attach ────────────────────────────────────────────────────────────────────
def attach(cls):
    cls._build_summary_panel      = _build_summary_panel
    cls._summary_placeholder      = _summary_placeholder
    cls._populate_summary         = _populate_summary
    cls._summary_section_label    = _summary_section_label
    cls._summary_cibi_ci_card     = _summary_cibi_ci_card
    cls._summary_cibi_tier_card   = _summary_cibi_tier_card
    cls._summary_cibi_docs_card   = _summary_cibi_docs_card
    cls._summary_dsr_ndi_card     = _summary_dsr_ndi_card
    cls._summary_flags_card       = _summary_flags_card
    cls._summary_score_gauge      = _summary_score_gauge
    cls._summary_ratios_chart     = _summary_ratios_chart
    cls._summary_risk_donut       = _summary_risk_donut
    cls._summary_income_chart     = _summary_income_chart
    cls._embed_chart              = _embed_chart
    cls._card                     = _card
    cls._summary_verdict_card     = _summary_verdict_card
    cls._summary_info_card        = _summary_info_card
    cls._summary_product_card     = _summary_product_card
    cls._pill                     = _pill
    cls._parse_income_expense     = _parse_income_expense
    cls._parse_verdict            = _parse_verdict
    cls._parse_applicant          = _parse_applicant
    cls._parse_score              = _parse_score
    cls._parse_ratios             = _parse_ratios
    cls._parse_flags              = _parse_flags
    cls._parse_product            = _parse_product
    cls._score_color              = _score_color
"""
ui_chat.py — DocExtract Pro
==============================
AI Chat tab builder and all chat/RAG/KB methods
attached to DocExtractorApp.
"""
import os
import sys
import threading
from pathlib import Path
import tkinter as tk
import customtkinter as ctk
from app_constants import *
from RAG import get_rag_engine
from file_generator import detect_file_intent, generate_file

def _build_ai_prompt_panel(self, parent):
    self._aiprompt_frame = tk.Frame(parent, bg=CARD_WHITE)

    hdr = tk.Frame(self._aiprompt_frame, bg=NAVY_DEEP)
    hdr.pack(fill="x")
    tk.Label(hdr, text="🤖  BSV.AI",
             font=F(12, "bold"), fg=WHITE, bg=NAVY_DEEP,
             padx=20, pady=10).pack(side="left")

    model_badge = tk.Label(hdr, text=f"  {GEMINI_MODEL}  ",
                            font=F(7, "bold"), fg=LIME_BRIGHT,
                            bg=NAVY_MID, padx=8, pady=4)
    model_badge.pack(side="left", padx=(0, 12), pady=8)

    ctrl_row = tk.Frame(hdr, bg=NAVY_DEEP)
    ctrl_row.pack(side="right", padx=16, pady=8)

    self._kb_status_lbl = tk.Label(ctrl_row, text="KB: 0 docs",
                                    font=F(8), fg=LIME_PALE,
                                    bg=NAVY_DEEP, padx=8)
    self._kb_status_lbl.pack(side="left")

    ctk.CTkButton(ctrl_row, text="📚  Add to KB",
                  command=self._add_to_knowledge_base,
                  width=100, height=28, corner_radius=6,
                  fg_color=LIME_DARK, hover_color=LIME_MID,
                  text_color=WHITE,
                  font=FF(8, "bold"),
                  border_width=0).pack(side="left", padx=(0, 6))

    ctk.CTkButton(ctrl_row, text="🗑  Clear",
                  command=self._clear_chat,
                  width=80, height=28, corner_radius=6,
                  fg_color=SIDEBAR_ITEM, hover_color=SIDEBAR_HVR,
                  text_color=WHITE,
                  font=FF(8, "bold"),
                  border_width=0).pack(side="left")
    self._refresh_kb_status()

    ctx_bar = tk.Frame(self._aiprompt_frame, bg=NAVY_MIST,
                        highlightbackground=BORDER_LIGHT, highlightthickness=1)
    ctx_bar.pack(fill="x")
    self._ctx_lbl = tk.Label(ctx_bar,
        text="ℹ  No document extracted yet — extract a file first, then ask questions here.",
        font=F(9), fg=TXT_SOFT, bg=NAVY_MIST,
        padx=16, pady=7, anchor="w")
    self._ctx_lbl.pack(fill="x")

    chat_area = tk.Frame(self._aiprompt_frame, bg=OFF_WHITE)
    chat_area.pack(fill="both", expand=True)
    chat_sb = tk.Scrollbar(chat_area, relief="flat", troughcolor=OFF_WHITE,
                           bg=BORDER_LIGHT, width=8, bd=0)
    chat_sb.pack(side="right", fill="y")
    self._chat_canvas = tk.Canvas(chat_area, bg=OFF_WHITE,
                                   highlightthickness=0, yscrollcommand=chat_sb.set)
    self._chat_canvas.pack(side="left", fill="both", expand=True)
    chat_sb.config(command=self._chat_canvas.yview)
    self._chat_inner = tk.Frame(self._chat_canvas, bg=OFF_WHITE)
    self._chat_canvas_win = self._chat_canvas.create_window(
        (0, 0), window=self._chat_inner, anchor="nw")
    self._chat_inner.bind("<Configure>",
        lambda e: self._chat_canvas.configure(scrollregion=self._chat_canvas.bbox("all")))
    self._chat_canvas.bind("<Configure>",
        lambda e: self._chat_canvas.itemconfig(self._chat_canvas_win, width=e.width))

    def _scroll_chat(event):
        self._chat_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    self._chat_canvas.bind("<MouseWheel>", _scroll_chat)
    self._chat_inner.bind("<MouseWheel>", _scroll_chat)

    self._append_chat_bubble(
        "👋  Hello! I'm BSV-AI!\n\n"
        "Extract a document first, then ask me anything about it.\n\n"
        "You can also ask me to generate a file — just say:\n"
        "  • 'Create a Word document with the loan summary'\n"
        "  • 'Export this analysis as a PDF'\n"
        "  • 'Save the key ratios to an Excel spreadsheet'",
        role="system"
    )

    chips_frame = tk.Frame(self._aiprompt_frame, bg=WHITE,
                            highlightbackground=BORDER_LIGHT, highlightthickness=1)
    chips_frame.pack(fill="x", padx=0)

    chips_inner = tk.Frame(chips_frame, bg=WHITE)
    chips_inner.pack(fill="x", padx=14, pady=6)

    tk.Label(chips_inner, text="Quick:", font=F(7, "bold"),
             fg=TXT_MUTED, bg=WHITE).pack(side="left", padx=(0, 8))

    quick_prompts = [
        ("📋 Summarize",        "Summarize the key points of the extracted document."),
        ("💰 Loan Eligibility", "Based on the document, is the applicant eligible for a loan? Give reasons."),
        ("⚠ Risk Flags",       "List all risk flags or red flags found in the document."),
        ("📝 Worksheet",       "Create a worksheet based from the extracted sheet."),
        ("✅ Recommend",        "What is your overall recommendation for this loan application?"),
        ("📄 DOCX",             "Create a Word document with the full loan analysis and recommendation."),
        ("📊 XLSX",             "Save the financial data and key ratios to an Excel spreadsheet."),
        ("🖨 PDF",              "Generate a PDF report of the complete loan analysis."),
    ]
    for chip_label, chip_prompt in quick_prompts:
        btn = tk.Label(chips_inner, text=chip_label, font=F(7, "bold"),
                       fg=NAVY_MID, bg=NAVY_MIST,
                       padx=8, pady=3, cursor="hand2", relief="flat")
        btn.pack(side="left", padx=2)
        btn.bind("<Enter>", lambda e, b=btn: b.config(bg=NAVY_GHOST, fg=NAVY_DEEP))
        btn.bind("<Leave>", lambda e, b=btn: b.config(bg=NAVY_MIST, fg=NAVY_MID))
        btn.bind("<Button-1>", lambda e, pr=chip_prompt: self._inject_quick_prompt(pr))

    input_row = tk.Frame(self._aiprompt_frame, bg=WHITE,
                          highlightbackground=BORDER_MID, highlightthickness=1)
    input_row.pack(fill="x")

    input_wrap = tk.Frame(input_row, bg=WHITE)
    input_wrap.pack(side="left", fill="both", expand=True, padx=(14, 0), pady=10)

    self._chat_input = tk.Text(input_wrap, font=FMONO(10), fg=TXT_NAVY, bg=OFF_WHITE,
                               relief="flat", bd=0, insertbackground=NAVY_MID,
                               height=3, wrap="word",
                               padx=12, pady=8,
                               highlightbackground=BORDER_MID, highlightthickness=1)
    self._chat_input.pack(fill="x")
    self._chat_input.bind("<Return>",       self._on_chat_enter)
    self._chat_input.bind("<Shift-Return>", lambda e: None)
    self._chat_input.insert("1.0", "Ask anything about the document…")
    self._chat_input.config(fg=TXT_MUTED)
    self._chat_input.bind("<FocusIn>",  self._chat_input_focus_in)
    self._chat_input.bind("<FocusOut>", self._chat_input_focus_out)

    send_col = tk.Frame(input_row, bg=WHITE)
    send_col.pack(side="right", padx=12, pady=10)
    self._send_btn = ctk.CTkButton(
        send_col, text="Send  ➤",
        command=self._send_ai_prompt,
        width=100, height=60, corner_radius=8,
        fg_color=NAVY_DEEP, hover_color=NAVY_MID,
        text_color=WHITE,
        font=FF(10, "bold"),
        border_width=0
    )
    self._send_btn.pack()
    tk.Label(send_col, text="Enter to send\nShift+Enter for newline",
             font=F(7), fg=TXT_MUTED, bg=WHITE, justify="center").pack(pady=(4, 0))

    self._typing_frame = tk.Frame(self._chat_inner, bg=OFF_WHITE)
    self._typing_lbl   = tk.Label(self._typing_frame,
                                   text="  🤖  AI is thinking…",
                                   font=F(10), fg=TXT_SOFT, bg=OFF_WHITE,
                                   padx=10, pady=8)
    self._typing_lbl.pack(anchor="w")

# NOTE: _CHAT_PLACEHOLDER removed from module level — it is now set as a
# class attribute inside attach() so self._CHAT_PLACEHOLDER resolves correctly.

def _chat_input_focus_in(self, _event=None):
    if self._chat_input.get("1.0", "end").strip() == self._CHAT_PLACEHOLDER:
        self._chat_input.delete("1.0", "end")
        self._chat_input.config(fg=TXT_NAVY)

def _chat_input_focus_out(self, _event=None):
    if not self._chat_input.get("1.0", "end").strip():
        self._chat_input.insert("1.0", self._CHAT_PLACEHOLDER)
        self._chat_input.config(fg=TXT_MUTED)

def _inject_quick_prompt(self, prompt):
    self._chat_input.delete("1.0", "end")
    self._chat_input.insert("1.0", prompt)
    self._chat_input.config(fg=TXT_NAVY)
    self._chat_input.focus_set()

def _on_chat_enter(self, event):
    if event.state & 0x1: return
    self._send_ai_prompt()
    return "break"

def _append_chat_bubble(self, text, role="assistant"):
    outer = tk.Frame(self._chat_inner, bg=OFF_WHITE)
    outer.pack(fill="x", padx=14, pady=5)
    if role == "user":
        bubble_bg=BUBBLE_USER; bubble_fg=BUBBLE_USER_TXT; anchor="e"; side="right"; lmargin=80; rmargin=0
    elif role == "system":
        bubble_bg=BUBBLE_SYS; bubble_fg=BUBBLE_SYS_TXT; anchor="w"; side="left"; lmargin=0; rmargin=80
    else:
        bubble_bg=BUBBLE_AI; bubble_fg=BUBBLE_AI_TXT; anchor="w"; side="left"; lmargin=0; rmargin=80
    if role == "user":
        tk.Frame(outer, bg=OFF_WHITE, width=lmargin).pack(side="left")
    else:
        tk.Frame(outer, bg=OFF_WHITE, width=rmargin).pack(side="right")
    bubble = tk.Frame(outer, bg=bubble_bg,
                       highlightbackground=BORDER_MID, highlightthickness=1)
    bubble.pack(side=side, anchor=anchor, fill="x", expand=(role != "user"))
    role_label = "You" if role == "user" else ("System" if role == "system" else "AI")
    role_color = (LIME_BRIGHT if role == "user"
                  else LIME_DARK if role == "system"
                  else NAVY_PALE)
    tk.Label(bubble, text=role_label, font=F(7, "bold"),
             fg=role_color, bg=bubble_bg, padx=10, anchor="w").pack(fill="x", anchor="w", pady=(6, 0))
    msg_lbl = tk.Label(bubble, text=text, font=FMONO(10),
                        fg=bubble_fg, bg=bubble_bg,
                        justify="left", anchor="w", wraplength=560, padx=10)
    msg_lbl.pack(fill="x", anchor="w", pady=(2, 10))

    def _bind_scroll(widget):
        widget.bind("<MouseWheel>",
            lambda e: self._chat_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        for child in widget.winfo_children(): _bind_scroll(child)
    _bind_scroll(outer)

    def _update_wrap(e, lbl=msg_lbl, lm=lmargin, rm=rmargin):
        avail = e.width - lm - rm - 28
        if avail > 80: lbl.config(wraplength=avail)
    self._chat_canvas.bind("<Configure>", lambda e, f=_update_wrap: f(e), add="+")
    self._chat_canvas.update_idletasks()
    self._chat_canvas.configure(scrollregion=self._chat_canvas.bbox("all"))
    self._chat_canvas.yview_moveto(1.0)

def _show_typing(self, show):
    if show:
        self._typing_frame.pack(fill="x", padx=14, pady=(0, 8))
        self._chat_canvas.update_idletasks()
        self._chat_canvas.configure(scrollregion=self._chat_canvas.bbox("all"))
        self._chat_canvas.yview_moveto(1.0)
    else:
        self._typing_frame.pack_forget()

def _update_ctx_badge(self):
    if self._extracted_text.strip():
        chars = len(self._extracted_text)
        names = ", ".join(Path(f).name for f in (self._selected_files or [])) or "document"
        self._ctx_lbl.config(
            text=f"✅  Context loaded: {names}  ({chars:,} chars) — AI will answer based on this document.",
            fg=LIME_DARK)
    else:
        self._ctx_lbl.config(
            text="ℹ  No document extracted yet — extract a file first, then ask questions here.",
            fg=TXT_SOFT)

def _clear_chat(self):
    for w in self._chat_inner.winfo_children(): w.destroy()
    self._chat_history = []
    self._pending_file_type = None
    self._typing_frame = tk.Frame(self._chat_inner, bg=OFF_WHITE)
    self._typing_lbl = tk.Label(self._typing_frame, text="  🤖  AI is thinking…",
                                 font=F(10), fg=TXT_SOFT, bg=OFF_WHITE, padx=10, pady=8)
    self._typing_lbl.pack(anchor="w")
    self._append_chat_bubble("Chat cleared. How can I help you?", role="system")
    if self._extracted_text.strip():
        self._post_doc_loaded_message(self._selected_files or [], self._extracted_text)

def _send_ai_prompt(self):
    raw = self._chat_input.get("1.0", "end").strip()
    if not raw or raw == self._CHAT_PLACEHOLDER: return
    self._update_ctx_badge()
    self._chat_input.delete("1.0", "end")
    self._chat_input.config(fg=TXT_NAVY)
    self._append_chat_bubble(raw, role="user")
    self._pending_file_type = detect_file_intent(raw)
    if self._pending_file_type:
        self._append_chat_bubble(
            f"📎 Got it — I'll generate a {self._pending_file_type.upper()} file from my response "
            f"and save it to your Desktop (DocExtract_Files folder).", role="system")
    self._send_btn.configure(state="disabled", text="Thinking…")
    self._show_typing(True)
    system_prompt = self._build_system_prompt(raw)
    messages = [{"role": "system", "content": system_prompt}]
    for msg in self._chat_history[-8:]: messages.append(msg)
    messages.append({"role": "user", "content": raw})
    self._chat_history.append({"role": "user", "content": raw})

    def worker():
        reply, model_used = gemini_chat(messages, GEMINI_API_KEY)
        self.after(0, self._finish_ai_reply, reply, model_used)
    threading.Thread(target=worker, daemon=True).start()

def _build_system_prompt(self, user_question=""):
    ROLE = (
        "You are a highly skilled financial analyst and credit officer "
        "working for Banco San Vicente (BSV), a rural bank in the Philippines. "
        "You specialise in microfinance, agricultural loans, MSME lending, and salary loans. "
        "Always answer in clear, professional English. Be concise but thorough.\n\n"
        "Your task: answer the user's question based on the extracted document "
        "and knowledge base context provided below.\n\n"
    )
    from cic_parser import is_cic_report, CIC_ANALYSIS_PROMPT_BLOCK
    if self._extracted_text and is_cic_report(self._extracted_text):
        ROLE = CIC_ANALYSIS_PROMPT_BLOCK + "\n\n" + ROLE

    WORKSHEET_KEYWORDS = {"worksheet","work sheet","create a worksheet","sources of income",
                           "household expenses","salary worksheet","income worksheet","bsv worksheet",
                           "fill out the form","fill in the form","complete the worksheet",
                           "salary computation","income computation","average monthly income"}
    q_lower = user_question.lower()
    needs_worksheet = any(kw in q_lower for kw in WORKSHEET_KEYWORDS)
    worksheet_hint = ""
    if needs_worksheet:
        worksheet_hint = (
            "WORKSHEET GENERATION INSTRUCTIONS:\n"
            "Base format ENTIRELY on worksheet examples from the knowledge base.\n"
            "Do NOT invent a format. Fill values from the extracted document.\n"
            "Use [MISSING] for any field not found.\n\n"
        )
    rag_context = ""
    query = user_question.strip()
    if needs_worksheet and self._extracted_text.strip():
        doc_preview = self._extracted_text[:300].replace("\n", " ")
        query = f"{user_question} {doc_preview}"
    if query:
        try:
            engine = get_rag_engine()
            if engine.is_ready and engine.document_count > 0:
                rag_context = engine.query(query, n_results=3)
        except Exception: pass
    doc_section = ""
    if self._extracted_text.strip():
        snippet = self._extracted_text[:2_500]
        if len(self._extracted_text) > 2_500:
            snippet += "\n\n[… document truncated …]"
        doc_section = (
            "\n\n═══════════════════════════════════\n"
            "  EXTRACTED DOCUMENT CONTENT\n"
            "═══════════════════════════════════\n"
            f"{snippet}\n"
            "═══════════════════════════════════\n"
        )
    return ROLE + worksheet_hint + rag_context + doc_section

def _finish_ai_reply(self, reply, model_used="primary"):
    from utils import strip_thinking
    reply = strip_thinking(reply)
    self._show_typing(False)
    if model_used == "fallback":
        self._append_chat_bubble(f"⚡ Switched to {FALLBACK_MODEL}.", role="system")
    elif model_used == "trimmed":
        self._append_chat_bubble("⚠ Context was trimmed — document too large.", role="system")
    self._append_chat_bubble(reply, role="assistant")
    self._chat_history.append({"role": "assistant", "content": reply})
    self._send_btn.configure(state="normal", text="Send  ➤")
    self._status_lbl.config(text="●  Ready", fg=LIME_DARK)
    if self._pending_file_type and not reply.startswith("⚠"):
        ftype = self._pending_file_type; self._pending_file_type = None
        def _gen(ft=ftype, r=reply, dt=self._extracted_text):
            try:
                stem = "bsv_analysis"
                if self._selected_files: stem = Path(self._selected_files[0]).stem
                path = generate_file(r, ft, dt, stem)
                self.after(0, self._notify_file_saved, path)
            except Exception as e:
                self.after(0, self._append_chat_bubble, f"⚠ File generation failed:\n{e}", "system")
        threading.Thread(target=_gen, daemon=True).start()

def _notify_file_saved(self, path):
    self._append_chat_bubble(
        f"✅ File saved!\n\n📁  {path}\n\nSaved to Desktop → DocExtract_Files folder.",
        role="system")
    try:
        folder = str(path.parent)
        if sys.platform == "win32":    os.startfile(folder)
        elif sys.platform == "darwin": os.system(f'open "{folder}"')
        else:                          os.system(f'xdg-open "{folder}"')
    except Exception: pass

def _refresh_kb_status(self):
    self._kb_status_lbl.config(text="KB: loading…")
    def _worker():
        try:
            engine = get_rag_engine()
            count  = len(engine.list_documents()) if engine.is_ready else 0
            label  = f"KB: {count} doc(s)" if engine.is_ready else "KB: offline"
            self.after(0, self._kb_status_lbl.config, {"text": label})
        except Exception:
            self.after(0, self._kb_status_lbl.config, {"text": "KB: offline"})
    threading.Thread(target=_worker, daemon=True).start()

def _add_to_knowledge_base(self):
    if not self._extracted_text.strip():
        self._append_chat_bubble("⚠ No extracted text found. Please extract a document first.", role="system")
        return
    try:
        engine = get_rag_engine()
        if not engine.is_ready:
            self._append_chat_bubble(
                "⚠ Knowledge base is offline.\n\npip install chromadb sentence-transformers",
                role="system")
            return
        if self._selected_files:
            names  = [Path(f).stem for f in self._selected_files]
            doc_id = "__".join(names[:3]).lower().replace(" ", "_")
            source = ", ".join(Path(f).name for f in self._selected_files[:3])
        else:
            doc_id = "document"; source = "unknown"
        doc_type = "loan_application"
        if any(kw in self._extracted_text.upper() for kw in ("WORKSHEET","SOURCES OF INCOME","HOUSEHOLD EXPENSES")):
            doc_type = "worksheet"
        elif any(kw in self._extracted_text.upper() for kw in ("TEMPLATE","CRITERIA","SCORING RUBRIC")):
            doc_type = "scoring_rubric"
        self._append_chat_bubble(f"⏳ Adding '{source}' to the knowledge base…", role="system")
        def _worker():
            try:
                n = engine.add_document(text=self._extracted_text, doc_id=doc_id,
                                         metadata={"source": source, "type": doc_type})
                self.after(0, self._finish_add_to_kb, source, n)
            except Exception as e:
                self.after(0, self._append_chat_bubble, f"⚠ Failed:\n{e}", "system")
        threading.Thread(target=_worker, daemon=True).start()
    except Exception as e:
        self._append_chat_bubble(f"⚠ Knowledge base error:\n{e}", role="system")

def _finish_add_to_kb(self, source, chunk_count):
    self._append_chat_bubble(
        f"✅ '{source}' added to the knowledge base ({chunk_count} chunk(s)).", role="system")
    self._refresh_kb_status()

# ══════════════════════════════════════════════════════════════════════
#  CIBI MODE — INTELLIGENT STEP-BY-STEP WORKFLOW
# ══════════════════════════════════════════════════════════════════════


# ── attach ────────────────────────────────────────────────────────────────────
def attach(cls):
    # FIX: _CHAT_PLACEHOLDER assigned as a class attribute here so that
    # self._CHAT_PLACEHOLDER resolves correctly in all bound methods.
    cls._CHAT_PLACEHOLDER         = "Ask anything about the document…"
    cls._build_ai_prompt_panel    = _build_ai_prompt_panel
    cls._chat_input_focus_in      = _chat_input_focus_in
    cls._chat_input_focus_out     = _chat_input_focus_out
    cls._inject_quick_prompt      = _inject_quick_prompt
    cls._on_chat_enter            = _on_chat_enter
    cls._append_chat_bubble       = _append_chat_bubble
    cls._show_typing              = _show_typing
    cls._update_ctx_badge         = _update_ctx_badge
    cls._clear_chat               = _clear_chat
    cls._send_ai_prompt           = _send_ai_prompt
    cls._build_system_prompt      = _build_system_prompt
    cls._finish_ai_reply          = _finish_ai_reply
    cls._notify_file_saved        = _notify_file_saved
    cls._refresh_kb_status        = _refresh_kb_status
    cls._add_to_knowledge_base    = _add_to_knowledge_base
    cls._finish_add_to_kb         = _finish_add_to_kb
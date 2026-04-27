"""
admin_account.py — DocExtract Pro
===================================
Accounts management tab (super admin only).

Features
--------
• Live search / filter (resets to page 1)
• Clickable column headers → ascending / descending sort with ▲ / ▼ arrows
• Rows fill the full table height (no blank gap at bottom)
• Role / Status pill badges
• Working Edit dialog  → updates role + status in DB
• Working Delete dialog → removes user from DB after confirmation
• Add User dialog → inserts new user (SHA-256 password)
• Pagination 20 rows / page  (Prev · Next · Loaded N rows)
"""

import os
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime

from dotenv import load_dotenv
load_dotenv()

# ── Design tokens ──────────────────────────────────────────────────────────────
_PAGE_BG        = "#F4F6FA"
_SB_BG          = "#0B1622"
_SB_ACCENT      = "#5BBF3E"
_NAVY_MID       = "#1E3A5F"
_NAVY_LIGHT     = "#2D5F8A"
_NAVY_MIST      = "#EBF0F7"
_NAVY_GHOST     = "#D6E4F0"
_BORDER_MID     = "#C8D8E8"
_BORDER_LIGHT   = "#E8ECF2"
_TXT_NAVY       = "#1A2F47"
_TXT_SOFT       = "#4A6480"
_TXT_MUTED      = "#8A9BB0"
_ACCENT_RED     = "#E05555"
_ACCENT_GOLD    = "#D4A017"
_ACCENT_SUCCESS = "#2E8B4A"
_WHITE          = "#FFFFFF"

# ── Table visuals ──────────────────────────────────────────────────────────────
_HDR_BG  = "#1B2B3A"
_HDR_FG  = "#C8D8E8"
_HDR_H   = 42
_ROW_H   = 38
_ROW_BG  = _WHITE
_ROW_HOV = "#F0F6FF"
_ROW_ALT = "#FAFBFD"
_SEP_CLR = "#E8ECF2"

_PAGE_SIZE = 20

# ── Column definitions: (label, db_key, width, anchor, sortable?) ──────────────
_COLUMNS = [
    ("ID",          "id",            52,   "center", True),
    ("Username",    "username",      130,  "w",      True),
    ("Email",       "email",         200,  "w",      True),
    ("Position",    "position",      145,  "w",      True),
    ("Role",        "role",          108,  "center", True),
    ("Status",      "status",        88,   "center", True),
    ("Created",     "created_at",    138,  "center", True),
    ("Last Login",  "last_login_at", 138,  "center", True),
    ("Delete",      "_delete",       52,   "center", False),
    ("Edit",        "_edit",         52,   "center", False),
]

_ROLE_BADGE = {
    "super admin": ("#FFFFFF",      "#1E5C1E", "#1E5C1E"),
    "Account Officer":       (_NAVY_MID,      _NAVY_MIST, _NAVY_LIGHT),
    "Credit Risk Officer":        (_TXT_SOFT,      "#F0F4F8",  _BORDER_MID),
}
_STATUS_BADGE = {
    "active":   (_ACCENT_SUCCESS, "#E8F7EE", "#2E8B4A"),
    "inactive": (_ACCENT_RED,     "#FEF0F0", _ACCENT_RED),
    "pending":  (_ACCENT_GOLD,    "#FEF8E8", _ACCENT_GOLD),
}


# ═══════════════════════════════════════════════════════════════════════
#  DATABASE CONNECTION
# ═══════════════════════════════════════════════════════════════════════

def _db_connect():
    import psycopg2
    conn = psycopg2.connect(
        host=os.getenv("DB_HOST"),
        port=int(os.getenv("DB_PORT", 5432)),
        dbname=os.getenv("DB_NAME"),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
    )
    return conn


# ═══════════════════════════════════════════════════════════════════════
#  AUDIT LOG HELPER
# ═══════════════════════════════════════════════════════════════════════

def _log_action(self, action: str, description: str):
    """Write an audit log entry to the logs table."""
    try:
        user_id = getattr(self, "_current_user_id", None)
        email   = getattr(self, "_current_username", None) or ""
        with _db_connect() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO logs (user_id, email, action, description, time) "
                "VALUES (%s, %s, %s, %s, NOW())",
                (user_id, email, action, description)
            )
            conn.commit()
            cur.close()
    except Exception as e:
        print(f"[log_action] failed: {e}")


# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _fmt_dt(val):
    if val is None:
        return "—"
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d %H:%M")
    s = str(val)
    return s[:16] if len(s) >= 16 else s


def _center_dialog(dlg, w, h):
    dlg.update_idletasks()
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    dlg.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")


def _pill(parent, text, fg, bg, border):
    wrap = tk.Frame(parent, bg=border, padx=1, pady=1)
    tk.Label(wrap, text=text, font=("Segoe UI", 8, "bold"),
             fg=fg, bg=bg, padx=7, pady=2).pack()
    return wrap


def _all_children(w):
    res = list(w.winfo_children())
    for c in w.winfo_children():
        res.extend(_all_children(c))
    return res


def _safe_bg(w, color):
    try:
        if getattr(w, "_is_badge", False):
            return
        if w.winfo_class() in ("Frame", "Label"):
            w.config(bg=color)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
#  SORTING
# ─────────────────────────────────────────────────────────────────────────────

def _sort_by(self, col_key):
    if getattr(self, "_acct_sort_col", None) == col_key:
        self._acct_sort_asc = not getattr(self, "_acct_sort_asc", False)
    else:
        self._acct_sort_col = col_key
        self._acct_sort_asc = (col_key != "id")

    self._acct_page = 1
    _render_table(self)
    _refresh_header_labels(self)


def _sorted_data(self):
    data    = _filtered(self)
    col     = getattr(self, "_acct_sort_col", "id")
    asc     = getattr(self, "_acct_sort_asc", False)

    def _key(row):
        v = row.get(col)
        if v is None:
            return ("", )
        if isinstance(v, (int, float)):
            return (0, v)
        if isinstance(v, datetime):
            return (0, v.timestamp())
        return (1, str(v).lower())

    return sorted(data, key=_key, reverse=not asc)


def _col_header_text(self, label, col_key, sortable):
    if not sortable:
        return label
    if getattr(self, "_acct_sort_col", "id") == col_key:
        arrow = " ▲" if getattr(self, "_acct_sort_asc", False) else " ▼"
        return label + arrow
    return label


def _refresh_header_labels(self):
    for col_label, col_key, col_w, col_anchor, sortable in _COLUMNS:
        lbl = self._acct_hdr_labels.get(col_key)
        if lbl:
            lbl.config(text=_col_header_text(self, col_label, col_key, sortable))


# ─────────────────────────────────────────────────────────────────────────────
#  DATA
# ─────────────────────────────────────────────────────────────────────────────

def _load_users(self):
    try:
        conn = self.get_conn()
        if conn is None:
            raise Exception("No database connection.")
        cur = conn.cursor()
        cur.execute("""
            SELECT id, username, email, position, role, status,
                   created_at, last_login_at
            FROM users
        """)
        rows = cur.fetchall()
        cur.close()
        keys = [c[1] for c in _COLUMNS if not c[1].startswith("_")]
        self._acct_data = [dict(zip(keys, r)) for r in rows]
    except Exception as e:
        self._acct_data = []
        self._acct_err_var.set(f"⚠  Could not load users: {e}")
        self._acct_err_lbl.pack(anchor="w", padx=24, pady=(4, 0))
        return

    self._acct_err_lbl.pack_forget()
    self._acct_err_var.set("")
    self._acct_page = 1
    _render_table(self)


def _filtered(self):
    q = self._acct_search.get().strip().lower()
    if not q:
        return self._acct_data
    return [r for r in self._acct_data
            if any(q in str(v).lower() for v in r.values() if v is not None)]


# ─────────────────────────────────────────────────────────────────────────────
#  DIALOGS
# ─────────────────────────────────────────────────────────────────────────────

def _open_edit_dialog(self, user_row):
    """Edit username, email, position, role and status for an existing user."""
    dlg = tk.Toplevel(self)
    dlg.title(f"Edit User — {user_row.get('username','')}")
    dlg.configure(bg=_PAGE_BG)
    dlg.resizable(False, False)
    dlg.grab_set()
    _center_dialog(dlg, 460, 530)
    dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

    # ── Header ────────────────────────────────────────────────────────────
    hdr = tk.Frame(dlg, bg=_HDR_BG, height=54)
    hdr.pack(fill="x"); hdr.pack_propagate(False)
    tk.Label(hdr, text=f"✏   Edit User  ·  {user_row.get('username','')}",
             font=("Segoe UI", 12, "bold"), fg=_WHITE, bg=_HDR_BG,
             anchor="w").pack(side="left", padx=20, fill="y")

    body = tk.Frame(dlg, bg=_PAGE_BG)
    body.pack(fill="both", expand=True, padx=26, pady=18)

    # ── Read-only ID ──────────────────────────────────────────────────────
    r = tk.Frame(body, bg=_PAGE_BG); r.pack(fill="x", pady=3)
    tk.Label(r, text="ID:", font=("Segoe UI", 9, "bold"),
             fg=_TXT_SOFT, bg=_PAGE_BG, width=12, anchor="w").pack(side="left")
    tk.Label(r, text=str(user_row.get("id") or "—"),
             font=("Segoe UI", 9), fg=_TXT_MUTED, bg=_PAGE_BG,
             anchor="w").pack(side="left")

    tk.Frame(body, bg=_BORDER_LIGHT, height=1).pack(fill="x", pady=(6, 4))

    # ── Editable text fields ──────────────────────────────────────────────
    edit_vars = {}

    def _entry_field(label, key):
        r = tk.Frame(body, bg=_PAGE_BG); r.pack(fill="x", pady=4)
        tk.Label(r, text=label + ":", font=("Segoe UI", 9, "bold"),
                 fg=_TXT_SOFT, bg=_PAGE_BG, width=12, anchor="w").pack(side="left")
        wrap = tk.Frame(r, bg=_WHITE,
                        highlightbackground=_BORDER_MID, highlightthickness=1)
        wrap.pack(side="left", fill="x", expand=True)
        var = tk.StringVar(value=str(user_row.get(key) or ""))
        tk.Entry(wrap, textvariable=var, font=("Segoe UI", 10),
                 fg=_TXT_NAVY, bg=_WHITE, relief="flat", bd=4,
                 insertbackground=_NAVY_MID).pack(fill="x")
        edit_vars[key] = var

    _entry_field("Username", "username")
    _entry_field("Email",    "email")
    _entry_field("Position", "position")

    tk.Frame(body, bg=_BORDER_MID, height=1).pack(fill="x", pady=(10, 6))

    # ── Editable dropdowns ────────────────────────────────────────────────
    role_var   = tk.StringVar(value=str(user_row.get("role")   or "user"))
    status_var = tk.StringVar(value=str(user_row.get("status") or "active"))

    for label, var, vals in [
        ("Role:",   role_var,   ["super admin", "Account Officer", "Credit Risk Officer"]),
        ("Status:", status_var, ["active", "inactive", "pending"]),
    ]:
        r = tk.Frame(body, bg=_PAGE_BG); r.pack(fill="x", pady=6)
        tk.Label(r, text=label, font=("Segoe UI", 9, "bold"),
                 fg=_TXT_SOFT, bg=_PAGE_BG, width=12, anchor="w").pack(side="left")
        ctk.CTkOptionMenu(
            r, variable=var, values=vals,
            width=220, height=32, corner_radius=6,
            fg_color=_NAVY_MIST, button_color=_NAVY_MID,
            button_hover_color=_NAVY_LIGHT,
            text_color=_TXT_NAVY, dropdown_fg_color=_WHITE,
        ).pack(side="left")

    # ── Divider above buttons ─────────────────────────────────────────────
    tk.Frame(dlg, bg=_BORDER_LIGHT, height=1).pack(fill="x", pady=(4, 0))

    # ── Buttons ───────────────────────────────────────────────────────────
    btn_row = tk.Frame(dlg, bg=_PAGE_BG)
    btn_row.pack(fill="x", padx=26, pady=14)

    def _save():
        uid          = user_row.get("id")
        new_username = edit_vars["username"].get().strip()
        new_email    = edit_vars["email"].get().strip()
        new_position = edit_vars["position"].get().strip()
        new_role     = role_var.get().strip()
        new_status   = status_var.get().strip()

        if not new_username or not new_email:
            messagebox.showwarning("Missing Fields",
                                   "Username and Email are required.", parent=dlg)
            return
        try:
            conn = self.get_conn()
            cur  = conn.cursor()
            cur.execute(
                """UPDATE users
                   SET username=%s, email=%s, position=%s,
                       role=%s, status=%s, updated_at=NOW()
                   WHERE id=%s""",
                (new_username, new_email, new_position or None,
                 new_role, new_status, uid),
            )
            conn.commit()
            cur.close()

            # ── Audit log ──────────────────────────────────────────────
            old_username = user_row.get("username", "")
            old_role     = user_row.get("role", "")
            old_status   = user_row.get("status", "")
            changes = []
            if new_username != old_username:
                changes.append(f"username: '{old_username}' → '{new_username}'")
            if new_role != old_role:
                changes.append(f"role: '{old_role}' → '{new_role}'")
            if new_status != old_status:
                changes.append(f"status: '{old_status}' → '{new_status}'")
            change_str = ", ".join(changes) if changes else "no field changes"
            _log_action(self, "edit_user",
                        f"Edited user id={uid} ({new_username}): {change_str}")

            messagebox.showinfo("Saved",
                                f"User #{uid} updated successfully.", parent=dlg)
            dlg.destroy()
            _load_users(self)
        except Exception as e:
            messagebox.showerror("DB Error",
                                 f"Could not update user:\n{e}", parent=dlg)

    ctk.CTkButton(btn_row, text="💾  Save Changes", command=_save,
                  width=140, height=36, corner_radius=7,
                  fg_color=_SB_ACCENT, hover_color="#4CAF35",
                  text_color="#0A1628", font=("Segoe UI", 10, "bold")).pack(side="left")
    ctk.CTkButton(btn_row, text="Cancel", command=dlg.destroy,
                  width=90, height=36, corner_radius=7,
                  fg_color=_NAVY_MIST, hover_color=_NAVY_GHOST,
                  text_color=_TXT_SOFT, font=("Segoe UI", 10)).pack(side="left", padx=(10, 0))


def _open_delete_dialog(self, user_row):
    """Confirm then delete a user from the DB."""
    uid   = user_row.get("id", "")
    uname = user_row.get("username", "")

    dlg = tk.Toplevel(self)
    dlg.title("Confirm Delete")
    dlg.configure(bg=_PAGE_BG)
    dlg.resizable(False, False)
    dlg.grab_set()
    _center_dialog(dlg, 420, 290)
    dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

    # ── Header (danger red) ───────────────────────────────────────────────
    hdr = tk.Frame(dlg, bg="#7A1818", height=54)
    hdr.pack(fill="x"); hdr.pack_propagate(False)
    tk.Label(hdr, text="🗑   Delete User",
             font=("Segoe UI", 12, "bold"), fg=_WHITE, bg="#7A1818",
             anchor="w").pack(side="left", padx=20, fill="y")

    body = tk.Frame(dlg, bg=_PAGE_BG)
    body.pack(fill="both", expand=True, padx=26, pady=22)

    tk.Label(body, text="Are you sure you want to permanently delete:",
             font=("Segoe UI", 10), fg=_TXT_NAVY, bg=_PAGE_BG,
             anchor="w").pack(anchor="w")

    info_box = tk.Frame(body, bg="#FEF0F0",
                        highlightbackground="#F5CCCC", highlightthickness=1)
    info_box.pack(fill="x", pady=(8, 0))
    tk.Label(info_box,
             text=f"  ID {uid}  ·  {uname}  ·  {user_row.get('email','—')}",
             font=("Segoe UI", 10, "bold"), fg="#7A1818", bg="#FEF0F0",
             anchor="w", pady=8).pack(anchor="w", padx=8)

    tk.Label(body, text="⚠  This action cannot be undone.",
             font=("Segoe UI", 9, "italic"), fg=_ACCENT_RED, bg=_PAGE_BG,
             anchor="w").pack(anchor="w", pady=(10, 0))

    # ── Divider above buttons ─────────────────────────────────────────────
    tk.Frame(dlg, bg=_BORDER_LIGHT, height=1).pack(fill="x", pady=(0, 0))

    btn_row = tk.Frame(dlg, bg=_PAGE_BG)
    btn_row.pack(fill="x", padx=26, pady=14)

    def _confirm():
        try:
            conn = self.get_conn()
            cur  = conn.cursor()
            cur.execute("DELETE FROM users WHERE id=%s", (uid,))
            conn.commit()
            cur.close()

            # ── Audit log ──────────────────────────────────────────────
            _log_action(self, "delete_user",
                        f"Deleted user id={uid} "
                        f"(username='{uname}', "
                        f"email='{user_row.get('email', '')}', "
                        f"role='{user_row.get('role', '')}')")

            messagebox.showinfo("Deleted",
                                f"User '{uname}' has been deleted.", parent=dlg)
            dlg.destroy()
            _load_users(self)
        except Exception as e:
            messagebox.showerror("DB Error",
                                 f"Could not delete user:\n{e}", parent=dlg)

    ctk.CTkButton(btn_row, text="🗑  Yes, Delete", command=_confirm,
                  width=130, height=36, corner_radius=7,
                  fg_color=_ACCENT_RED, hover_color="#BF2222",
                  text_color=_WHITE, font=("Segoe UI", 10, "bold")).pack(side="left")
    ctk.CTkButton(btn_row, text="Cancel", command=dlg.destroy,
                  width=90, height=36, corner_radius=7,
                  fg_color=_NAVY_MIST, hover_color=_NAVY_GHOST,
                  text_color=_TXT_SOFT, font=("Segoe UI", 10)).pack(side="left", padx=(10, 0))


def _open_add_dialog(self):
    """Create a new user."""
    dlg = tk.Toplevel(self)
    dlg.title("Add New User")
    dlg.configure(bg=_PAGE_BG)
    dlg.resizable(False, False)
    dlg.grab_set()
    _center_dialog(dlg, 450, 430)
    dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)

    hdr = tk.Frame(dlg, bg=_HDR_BG, height=54)
    hdr.pack(fill="x"); hdr.pack_propagate(False)
    tk.Label(hdr, text="➕   Add New User",
             font=("Segoe UI", 12, "bold"), fg=_WHITE, bg=_HDR_BG,
             anchor="w").pack(side="left", padx=20, fill="y")

    body = tk.Frame(dlg, bg=_PAGE_BG)
    body.pack(fill="both", expand=True, padx=26, pady=16)
    fields = {}

    def _entry_field(label, key, show=""):
        r = tk.Frame(body, bg=_PAGE_BG); r.pack(fill="x", pady=5)
        tk.Label(r, text=label + ":", font=("Segoe UI", 9, "bold"),
                 fg=_TXT_SOFT, bg=_PAGE_BG, width=12, anchor="w").pack(side="left")
        wrap = tk.Frame(r, bg=_WHITE,
                        highlightbackground=_BORDER_MID, highlightthickness=1)
        wrap.pack(side="left", fill="x", expand=True)
        var = tk.StringVar()
        tk.Entry(wrap, textvariable=var, font=("Segoe UI", 10),
                 fg=_TXT_NAVY, bg=_WHITE, relief="flat", bd=4,
                 show=show, insertbackground=_NAVY_MID).pack(fill="x")
        fields[key] = var

    _entry_field("Username", "username")
    _entry_field("Email",    "email")
    _entry_field("Password", "password", show="•")
    _entry_field("Position", "position")

    role_var   = tk.StringVar(value="user")
    status_var = tk.StringVar(value="active")

    for label, var, opts in [
        ("Role",   role_var,   ["super admin", "Account Officer", "Credit Risk Officer"]),
        ("Status", status_var, ["active", "inactive", "pending"]),
    ]:
        r = tk.Frame(body, bg=_PAGE_BG); r.pack(fill="x", pady=5)
        tk.Label(r, text=label + ":", font=("Segoe UI", 9, "bold"),
                 fg=_TXT_SOFT, bg=_PAGE_BG, width=12, anchor="w").pack(side="left")
        ctk.CTkOptionMenu(r, variable=var, values=opts,
                          width=200, height=32, corner_radius=6,
                          fg_color=_NAVY_MIST, button_color=_NAVY_MID,
                          button_hover_color=_NAVY_LIGHT,
                          text_color=_TXT_NAVY,
                          dropdown_fg_color=_WHITE).pack(side="left")

    tk.Frame(dlg, bg=_BORDER_LIGHT, height=1).pack(fill="x", pady=(4, 0))

    btn_row = tk.Frame(dlg, bg=_PAGE_BG)
    btn_row.pack(fill="x", padx=26, pady=14)

    def _create():
        username = fields["username"].get().strip()
        email    = fields["email"].get().strip()
        password = fields["password"].get().strip()
        position = fields["position"].get().strip()

        if not username or not email or not password:
            messagebox.showwarning("Missing Fields",
                                "Username, Email and Password are required.",
                                parent=dlg)
            return
        try:
            conn = self.get_conn()
            cur  = conn.cursor()
            cur.execute(
                """INSERT INTO users
                (username, email, password, position, role, status,
                    created_at, updated_at)
                VALUES (%s, %s, crypt(%s, gen_salt('bf')), %s, %s, %s, NOW(), NOW())""",
                (username, email, password,
                position or None,
                role_var.get().strip(), status_var.get().strip()),
            )
            conn.commit()
            cur.close()

            # ── Audit log ──────────────────────────────────────────────
            _log_action(self, "add_user",
                        f"Created new user: username='{username}', "
                        f"email='{email}', "
                        f"role='{role_var.get().strip()}', "
                        f"status='{status_var.get().strip()}'")

            messagebox.showinfo("Created",
                                f"User '{username}' created successfully.", parent=dlg)
            dlg.destroy()
            _load_users(self)
        except Exception as e:
            messagebox.showerror("DB Error",
                                f"Could not create user:\n{e}", parent=dlg)

    ctk.CTkButton(btn_row, text="✔  Create User", command=_create,
                  width=140, height=36, corner_radius=7,
                  fg_color=_SB_ACCENT, hover_color="#4CAF35",
                  text_color="#0A1628", font=("Segoe UI", 10, "bold")).pack(side="left")
    ctk.CTkButton(btn_row, text="Cancel", command=dlg.destroy,
                  width=90, height=36, corner_radius=7,
                  fg_color=_NAVY_MIST, hover_color=_NAVY_GHOST,
                  text_color=_TXT_SOFT, font=("Segoe UI", 10)).pack(side="left", padx=(10, 0))


# ─────────────────────────────────────────────────────────────────────────────
#  TABLE RENDERING
# ─────────────────────────────────────────────────────────────────────────────

def _render_table(self, *_):
    for w in self._acct_body.winfo_children():
        w.destroy()

    data        = _sorted_data(self)
    page        = max(1, getattr(self, "_acct_page", 1))
    total_pages = max(1, -(-len(data) // _PAGE_SIZE))
    page        = min(page, total_pages)
    self._acct_page = page

    start     = (page - 1) * _PAGE_SIZE
    page_data = data[start: start + _PAGE_SIZE]
    loaded    = min(page * _PAGE_SIZE, len(data))

    self._loaded_lbl.config(text=f"✔  Loaded {loaded} rows")
    total  = len(self._acct_data)
    filt   = len(_filtered(self))
    end    = min(start + _PAGE_SIZE, filt)
    suffix = f" (filtered from {total})" if self._acct_search.get().strip() else ""
    self._count_lbl.config(
        text=f"Page {page}/{total_pages}  ·  {start+1}–{end} of {filt}{suffix}")

    _set_nav(self, page, total_pages)

    if not page_data:
        tk.Label(self._acct_body, text="No users found.",
                 font=("Segoe UI", 11), fg=_TXT_MUTED,
                 bg=_ROW_BG).pack(pady=50)
        return

    for idx, row in enumerate(page_data):
        bg = _ROW_BG if idx % 2 == 0 else _ROW_ALT
        _build_row(self, row, idx, bg)

    filler = tk.Frame(self._acct_body, bg=_ROW_BG)
    filler.pack(fill="both", expand=True)

    self._acct_body.update_idletasks()
    self._acct_canvas.configure(
        scrollregion=self._acct_canvas.bbox("all"))
    self._acct_canvas.yview_moveto(0)


def _build_row(self, row, idx, bg):
    uid = row.get("id")

    row_frame = tk.Frame(self._acct_body, bg=bg, height=_ROW_H)
    row_frame.pack(fill="x")
    row_frame.pack_propagate(False)

    for col_label, col_key, col_w, col_anchor, sortable in _COLUMNS:
        cell = tk.Frame(row_frame, bg=bg, width=col_w)
        if col_key == "email":
            cell.pack(side="left", fill="both", expand=True)
        elif col_key in ("_edit", "_delete"):
            cell.pack(side="right", fill="y")
        else:
            cell.pack(side="left", fill="y")
        cell.pack_propagate(False)

        val = row.get(col_key)

        if col_key == "_edit":
            ico = tk.Label(cell, text="✏", font=("Segoe UI Emoji", 13),
                           fg=_NAVY_LIGHT, bg=bg, cursor="hand2")
            ico.place(relx=0.5, rely=0.5, anchor="center")
            for w in (ico, cell):
                w.bind("<Button-1>", lambda e, r=row: _open_edit_dialog(self, r))
                w.bind("<Enter>",    lambda e, i=ico: i.config(fg=_SB_ACCENT))
                w.bind("<Leave>",    lambda e, i=ico, c=bg: i.config(fg=_NAVY_LIGHT))

        elif col_key == "_delete":
            ico = tk.Label(cell, text="🗑", font=("Segoe UI Emoji", 13),
                           fg="#CC3333", bg=bg, cursor="hand2")
            ico.place(relx=0.5, rely=0.5, anchor="center")
            for w in (ico, cell):
                w.bind("<Button-1>", lambda e, r=row: _open_delete_dialog(self, r))
                w.bind("<Enter>",    lambda e, i=ico: i.config(fg="#881111"))
                w.bind("<Leave>",    lambda e, i=ico, c=bg: i.config(fg="#CC3333"))

        elif col_key == "role":
            v = str(val).lower() if val else "user"
            fg, pbg, bd = _ROLE_BADGE.get(v, (_TXT_SOFT, "#F0F4F8", _BORDER_MID))
            pill = _pill(cell, val or "user", fg, pbg, bd)
            pill.place(relx=0.5, rely=0.5, anchor="center")
            pill._is_badge = True
            for child in _all_children(pill):
                child._is_badge = True

        elif col_key == "status":
            v = str(val).lower() if val else "unknown"
            fg, pbg, bd = _STATUS_BADGE.get(v, (_TXT_SOFT, _NAVY_MIST, _BORDER_MID))
            pill = _pill(cell, v.capitalize(), fg, pbg, bd)
            pill.place(relx=0.5, rely=0.5, anchor="center")
            pill._is_badge = True
            for child in _all_children(pill):
                child._is_badge = True

        elif col_key in ("created_at", "last_login_at"):
            tk.Label(cell, text=_fmt_dt(val), font=("Segoe UI", 9),
                     fg=_TXT_SOFT, bg=bg).place(relx=0.5, rely=0.5, anchor="center")

        elif col_key == "id":
            tk.Label(cell, text=str(val) if val is not None else "—",
                     font=("Segoe UI", 9), fg=_TXT_SOFT,
                     bg=bg).place(relx=0.5, rely=0.5, anchor="center")

        else:
            display  = str(val) if val is not None else "—"
            anchor_x = 0.5 if col_anchor == "center" else 0.05
            lbl = tk.Label(cell, text=display, font=("Segoe UI", 9),
                           fg=_TXT_NAVY, bg=bg, anchor=col_anchor)
            lbl.place(relx=anchor_x, rely=0.5,
                      anchor="center" if col_anchor == "center" else "w")

    tk.Frame(self._acct_body, bg=_SEP_CLR, height=1).pack(fill="x")

    all_w = _all_children(row_frame)

    def _enter(*_):
        row_frame.config(bg=_ROW_HOV)
        for c in all_w: _safe_bg(c, _ROW_HOV)

    def _leave(*_):
        row_frame.config(bg=bg)
        for c in all_w: _safe_bg(c, bg)

    for w in [row_frame] + all_w:
        w.bind("<Enter>",           _enter)
        w.bind("<Leave>",           _leave)
        w.bind("<Double-Button-1>", lambda e, r=row: _open_edit_dialog(self, r))


# ─────────────────────────────────────────────────────────────────────────────
#  PAGINATION HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _set_nav(self, page, total_pages):
    for btn, enabled in ((self._prev_btn, page > 1),
                          (self._next_btn, page < total_pages)):
        if enabled:
            btn.configure(state="normal",  fg_color=_NAVY_MIST,
                          text_color=_TXT_NAVY, border_color=_BORDER_MID)
        else:
            btn.configure(state="disabled", fg_color="#F0F0F0",
                          text_color=_TXT_MUTED, border_color=_BORDER_LIGHT)


def _go_page(self, delta):
    data        = _sorted_data(self)
    total_pages = max(1, -(-len(data) // _PAGE_SIZE))
    new_page    = getattr(self, "_acct_page", 1) + delta
    if 1 <= new_page <= total_pages:
        self._acct_page = new_page
        _render_table(self)


# ─────────────────────────────────────────────────────────────────────────────
#  PANEL BUILD
# ─────────────────────────────────────────────────────────────────────────────

def _build_accounts_panel(self, parent):
    self._acct_data      = []
    self._acct_page      = 1
    self._acct_search    = tk.StringVar()
    self._acct_sort_col  = "id"
    self._acct_sort_asc  = False
    self._acct_hdr_labels = {}

    self._accounts_frame = tk.Frame(parent, bg=_PAGE_BG)
    frame = self._accounts_frame

    # ── Toolbar ───────────────────────────────────────────────────────────
    toolbar = tk.Frame(frame, bg=_PAGE_BG)
    toolbar.pack(fill="x", padx=20, pady=(14, 6))

    lbl_blk = tk.Frame(toolbar, bg=_PAGE_BG)
    lbl_blk.pack(side="left", fill="y")
    tk.Frame(lbl_blk, bg=_SB_ACCENT, width=3, height=16).pack(
        side="left", padx=(0, 8), pady=2)
    tk.Label(lbl_blk, text="USER ACCOUNTS",
             font=("Segoe UI", 8, "bold"), fg=_TXT_SOFT,
             bg=_PAGE_BG).pack(side="left")

    self._count_lbl = tk.Label(toolbar, text="Loading…",
                                font=("Segoe UI", 8), fg=_TXT_MUTED, bg=_PAGE_BG)
    self._count_lbl.pack(side="left", padx=(14, 0))

    ctk.CTkButton(toolbar, text="➕  Add User",
                  command=lambda: _open_add_dialog(self),
                  width=110, height=30, corner_radius=6,
                  fg_color=_SB_ACCENT, hover_color="#4CAF35",
                  text_color="#0A1628", font=("Segoe UI", 9, "bold"),
                  ).pack(side="right")
    ctk.CTkButton(toolbar, text="⟳  Refresh",
                  command=lambda: _load_users(self),
                  width=90, height=30, corner_radius=6,
                  fg_color=_NAVY_MIST, hover_color=_NAVY_GHOST,
                  text_color=_TXT_SOFT, font=("Segoe UI", 9),
                  border_width=1, border_color=_BORDER_MID,
                  ).pack(side="right", padx=(0, 6))

    # ── Search bar ────────────────────────────────────────────────────────
    sb_row = tk.Frame(frame, bg=_PAGE_BG)
    sb_row.pack(fill="x", padx=20, pady=(0, 8))

    sb_wrap = tk.Frame(sb_row, bg=_WHITE,
                       highlightbackground=_BORDER_MID, highlightthickness=1)
    sb_wrap.pack(side="left", fill="x", expand=True)
    tk.Label(sb_wrap, text="🔍", font=("Segoe UI Emoji", 10),
             bg=_WHITE, fg="#AABDD0").pack(side="left", padx=(10, 2))
    tk.Entry(sb_wrap, textvariable=self._acct_search,
             font=("Segoe UI", 10), fg=_TXT_NAVY, bg=_WHITE,
             relief="flat", bd=0, insertbackground=_NAVY_MID,
             width=36).pack(side="left", fill="x", expand=True, pady=6)
    clr = tk.Label(sb_wrap, text="✕", font=("Segoe UI", 9),
                   fg=_ACCENT_RED, bg=_WHITE, padx=8, cursor="hand2")
    clr.pack(side="right")
    clr.bind("<Button-1>", lambda e: self._acct_search.set(""))

    def _on_search(*_):
        self._acct_page = 1
        _render_table(self)

    self._acct_search.trace_add("write", _on_search)

    # ── Error label ───────────────────────────────────────────────────────
    self._acct_err_var = tk.StringVar()
    self._acct_err_lbl = tk.Label(
        frame, textvariable=self._acct_err_var,
        font=("Segoe UI", 9), fg=_ACCENT_RED, bg=_PAGE_BG, anchor="w")

    # ══ TABLE CARD ════════════════════════════════════════════════════════
    outer = tk.Frame(frame, bg=_BORDER_MID, padx=1, pady=1)
    outer.pack(fill="both", expand=True, padx=20)

    card = tk.Frame(outer, bg=_WHITE)
    card.pack(fill="both", expand=True)

    # ── Column headers (clickable) ────────────────────────────────────────
    hdr_frame = tk.Frame(card, bg=_HDR_BG, height=_HDR_H)
    hdr_frame.pack(fill="x")
    hdr_frame.pack_propagate(False)

    for col_label, col_key, col_w, col_anchor, sortable in _COLUMNS:
        hcell = tk.Frame(hdr_frame, bg=_HDR_BG, width=col_w,
                 cursor="hand2" if sortable else "arrow")
        if col_key == "email":
            hcell.pack(side="left", fill="both", expand=True)
        elif col_key in ("_edit", "_delete"):
            hcell.pack(side="right", fill="y")
        else:
            hcell.pack(side="left", fill="y")
        hcell.pack_propagate(False)

        init_text = _col_header_text(self, col_label, col_key, sortable)
        anchor_x  = 0.5 if col_anchor == "center" else 0.07

        lbl = tk.Label(hcell, text=init_text,
                       font=("Segoe UI", 9, "bold"),
                       fg=_SB_ACCENT if (col_key == getattr(self, "_acct_sort_col", "id"))
                          else _HDR_FG,
                       bg=_HDR_BG, anchor=col_anchor, cursor="hand2" if sortable else "arrow")
        lbl.place(relx=anchor_x, rely=0.5,
                  anchor="center" if col_anchor == "center" else "w")

        self._acct_hdr_labels[col_key] = lbl

        if sortable:
            def _make_sort_cb(k):
                def _cb(e): _sort_by(self, k)
                return _cb

            cb = _make_sort_cb(col_key)
            lbl.bind("<Button-1>",   cb)
            hcell.bind("<Button-1>", cb)

            def _hdr_enter(e, c=hcell, l=lbl):
                c.config(bg="#243F56")
                l.config(bg="#243F56")

            def _hdr_leave(e, c=hcell, l=lbl):
                c.config(bg=_HDR_BG)
                l.config(bg=_HDR_BG)

            hcell.bind("<Enter>", _hdr_enter)
            hcell.bind("<Leave>", _hdr_leave)
            lbl.bind("<Enter>",   _hdr_enter)
            lbl.bind("<Leave>",   _hdr_leave)

    tk.Frame(card, bg="#263D52", height=1).pack(fill="x")

    # ── Scrollable body ───────────────────────────────────────────────────
    body_wrap = tk.Frame(card, bg=_WHITE)
    body_wrap.pack(fill="both", expand=True)

    vsb = tk.Scrollbar(body_wrap, orient="vertical", relief="flat",
                   troughcolor="#F0F4F8", bg=_BORDER_MID, width=8, bd=0)
    vsb.pack(side="right", fill="y", padx=(0, 0))

    self._acct_canvas = tk.Canvas(body_wrap, bg=_WHITE,
                                   highlightthickness=0,
                                   yscrollcommand=vsb.set)
    self._acct_canvas.pack(side="left", fill="both", expand=True)
    vsb.config(command=self._acct_canvas.yview)

    self._acct_body = tk.Frame(self._acct_canvas, bg=_WHITE)
    _win = self._acct_canvas.create_window(
        (0, 0), window=self._acct_body, anchor="nw")

    self._acct_body.bind(
        "<Configure>",
        lambda e: self._acct_canvas.configure(
            scrollregion=self._acct_canvas.bbox("all")))
    self._acct_canvas.bind(
        "<Configure>",
        lambda e: self._acct_canvas.itemconfig(_win, width=e.width))

    def _wheel(e):
        self._acct_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    self._acct_canvas.bind("<MouseWheel>", _wheel)
    self._acct_body.bind("<MouseWheel>",   _wheel)

    # ══ FOOTER ════════════════════════════════════════════════════════════
    footer = tk.Frame(frame, bg=_WHITE,
                      highlightbackground=_BORDER_MID, highlightthickness=1)
    footer.pack(fill="x", padx=20, pady=(0, 14))

    inner = tk.Frame(footer, bg=_WHITE)
    inner.pack(fill="x", padx=10, pady=6)

    self._prev_btn = ctk.CTkButton(
        inner, text="◀  Prev",
        command=lambda: _go_page(self, -1),
        width=76, height=26, corner_radius=4,
        fg_color=_NAVY_MIST, hover_color=_NAVY_GHOST,
        text_color=_TXT_NAVY, font=("Segoe UI", 8, "bold"),
        border_width=1, border_color=_BORDER_MID)
    self._prev_btn.pack(side="left")

    self._next_btn = ctk.CTkButton(
        inner, text="Next  ▶",
        command=lambda: _go_page(self, +1),
        width=76, height=26, corner_radius=4,
        fg_color=_NAVY_MIST, hover_color=_NAVY_GHOST,
        text_color=_TXT_NAVY, font=("Segoe UI", 8, "bold"),
        border_width=1, border_color=_BORDER_MID)
    self._next_btn.pack(side="left", padx=(6, 0))

    self._loaded_lbl = tk.Label(inner, text="✔  Loaded 0 rows",
                                 font=("Segoe UI", 9), fg=_ACCENT_SUCCESS, bg=_WHITE)
    self._loaded_lbl.pack(side="left", padx=(16, 0))

    tk.Label(inner,
             text="Double-click a row to edit  ·  Passwords stored as bcrypt",
             font=("Segoe UI", 7), fg=_TXT_MUTED, bg=_WHITE).pack(side="right")

    frame.after(120, lambda: _load_users(self))


# ─────────────────────────────────────────────────────────────────────────────
#  ATTACH
# ─────────────────────────────────────────────────────────────────────────────

def attach(cls):
    cls._build_accounts_panel = _build_accounts_panel
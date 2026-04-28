import customtkinter as ctk
import sqlite3
import csv
import os
from datetime import datetime
from tkinter import messagebox, filedialog
import tkinter as tk

# ── App Config ───────────────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "doc_tracker.db")

ACCENT      = "#1B4F72"
ACCENT_DARK = "#154360"
SURFACE     = "#F4F6F8"
CARD        = "#FFFFFF"
BORDER      = "#D5D8DC"
TEXT_MAIN   = "#1C2833"
TEXT_SUB    = "#5D6D7E"
DANGER      = "#922B21"
ROW_ALT     = "#EAF0F6"
WARN_BG     = "#FEF9E7"
WARN_BORDER = "#F9A825"

TAB_IDLE_BG   = "#D6E4F0"
TAB_IDLE_TEXT = ACCENT

TAB_COLORS = {
    "Incoming Memo": {"accent": "#1B4F72", "dark": "#154360", "icon": "📥"},
    "Outgoing Memo": {"accent": "#1A5276", "dark": "#154360", "icon": "📤"},
}

DEPARTMENTS = [
    "Select Department…",
    "Administration", "Finance", "Human Resources",
    "Information Technology", "Legal", "Logistics",
    "Marketing", "Operations", "Procurement",
    "Public Relations", "Research & Development",
    "Sales", "Security", "Other",
]


# ── Database ──────────────────────────────────────────────────────────────────
def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS memos (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            date       TEXT NOT NULL,
            department TEXT NOT NULL,
            subject    TEXT NOT NULL,
            receiver   TEXT NOT NULL
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS outgoing (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            date        TEXT NOT NULL,
            subject     TEXT NOT NULL,
            recipient   TEXT NOT NULL,
            remarks     TEXT
        )
    """)
    con.commit()
    con.close()

# ── Generic DB helpers ────────────────────────────────────────────────────────
def _fetch(table, search_cols, query=""):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    if query:
        q = f"%{query}%"
        where = " OR ".join(f"{c} LIKE ?" for c in search_cols)
        cur.execute(f"SELECT * FROM {table} WHERE {where} ORDER BY id DESC",
                    tuple(q for _ in search_cols))
    else:
        cur.execute(f"SELECT * FROM {table} ORDER BY id DESC")
    rows = cur.fetchall()
    con.close()
    return rows

def _delete(table, row_id):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute(f"DELETE FROM {table} WHERE id=?", (row_id,))
    con.commit()
    con.close()

def _insert(table, cols, vals):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    ph = ",".join("?" for _ in vals)
    cur.execute(f"INSERT INTO {table} ({','.join(cols)}) VALUES ({ph})", vals)
    con.commit()
    con.close()

def now_str():
    return datetime.now().strftime("%Y-%m-%d  %H:%M")

# ── Reusable widget helpers ───────────────────────────────────────────────────
def make_label(parent, text, size=11, weight="bold", color=TEXT_SUB, **kw):
    return ctk.CTkLabel(parent, text=text,
                        font=ctk.CTkFont(size=size, weight=weight),
                        text_color=color, **kw)

def make_entry(parent, var, placeholder="", width=254):
    return ctk.CTkEntry(parent, textvariable=var, placeholder_text=placeholder,
                        fg_color=CARD, border_color=BORDER, text_color=TEXT_MAIN,
                        font=ctk.CTkFont(size=13), width=width, height=36)

def make_option(parent, var, values, accent, dark, width=254):
    return ctk.CTkOptionMenu(parent, variable=var, values=values,
                             fg_color=CARD, button_color=accent,
                             button_hover_color=dark, text_color=TEXT_MAIN,
                             dropdown_fg_color=CARD, dropdown_text_color=TEXT_MAIN,
                             dropdown_hover_color=ROW_ALT,
                             font=ctk.CTkFont(size=13), width=width)

def make_button(parent, text, cmd, fg, hover, text_color="#FFFFFF",
                border=None, height=42, width=254, radius=8):
    kw = dict(text=text, command=cmd, fg_color=fg, hover_color=hover,
              text_color=text_color, font=ctk.CTkFont(size=13, weight="bold"),
              height=height, corner_radius=radius, width=width)
    if border:
        kw["border_color"] = border
        kw["border_width"] = 1
    return ctk.CTkButton(parent, **kw)


# ── Past Entry Popup ──────────────────────────────────────────────────────────
class PastEntryPopup(ctk.CTkToplevel):
    """
    Modal popup for logging a record with a manually typed date.
    Subclasses implement _build_fields() and _collect_values().
    """
    TITLE  = "Log Past Entry"
    ACCENT = ACCENT

    def __init__(self, parent, on_save):
        super().__init__(parent)
        self._on_save = on_save
        self.title(self.TITLE)
        self.geometry("380x520")
        self.resizable(False, False)
        self.grab_set()
        self.configure(fg_color=SURFACE)
        self._build_ui()

    def _build_ui(self):
        ctk.CTkFrame(self, fg_color=self.ACCENT, corner_radius=0, height=6
                     ).pack(fill="x")

        make_label(self, self.TITLE, size=15, weight="bold",
                   color=self.ACCENT).pack(anchor="w", padx=28, pady=(18, 2))
        make_label(self, "Enter the original date of this memo.",
                   size=11, weight="normal", color=TEXT_SUB
                   ).pack(anchor="w", padx=28, pady=(0, 14))

        # Date field with yellow highlight
        date_card = ctk.CTkFrame(self, fg_color=WARN_BG, corner_radius=8,
                                 border_width=1, border_color=WARN_BORDER)
        date_card.pack(fill="x", padx=28, pady=(0, 16))
        make_label(date_card, "📅  Date  (YYYY-MM-DD)", size=11,
                   weight="bold", color="#7D6608"
                   ).pack(anchor="w", padx=14, pady=(10, 2))
        self._date_var = ctk.StringVar()
        ctk.CTkEntry(date_card, textvariable=self._date_var,
                     placeholder_text="e.g. 2025-03-15",
                     fg_color=CARD, border_color=WARN_BORDER,
                     text_color=TEXT_MAIN,
                     font=ctk.CTkFont(size=13), height=34,
                     ).pack(fill="x", padx=14, pady=(0, 10))

        self._build_fields()

        btn_row = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        btn_row.pack(fill="x", padx=28, pady=(16, 24))
        make_button(btn_row, "  ✔  Save Past Entry", self._save,
                    self.ACCENT, "#154360", height=40, width=190
                    ).pack(side="left")
        make_button(btn_row, "Cancel", self.destroy,
                    BORDER, "#BFC9CA", text_color=TEXT_MAIN,
                    height=40, width=100, radius=8
                    ).pack(side="left", padx=(10, 0))

    def _build_fields(self):
        pass

    def _collect_values(self):
        return ()

    def _validate_date(self, s):
        try:
            datetime.strptime(s.strip(), "%Y-%m-%d")
            return True
        except ValueError:
            return False

    def _save(self):
        date_str = self._date_var.get().strip()
        if not self._validate_date(date_str):
            messagebox.showwarning("Invalid Date",
                                   "Please enter a valid date in YYYY-MM-DD format.\n"
                                   "Example: 2025-03-15",
                                   parent=self)
            return
        vals = self._collect_values()
        if vals is None:
            return
        self._on_save(date_str, *vals)
        self.destroy()

    def _field_label(self, text):
        make_label(self, text, size=11, weight="bold",
                   color=TEXT_SUB).pack(anchor="w", padx=28, pady=(6, 0))


# ── Past Memo Popup ───────────────────────────────────────────────────────────
class PastMemoPopup(PastEntryPopup):
    TITLE  = "Log Past Incoming Memo"
    ACCENT = TAB_COLORS["Incoming Memo"]["accent"]

    def _build_fields(self):
        self._dept_var     = ctk.StringVar(value=DEPARTMENTS[0])
        self._subject_var  = ctk.StringVar()
        self._receiver_var = ctk.StringVar()

        self._field_label("From Department")
        make_option(self, self._dept_var, DEPARTMENTS,
                    self.ACCENT, "#154360").pack(anchor="w", padx=28, pady=(4, 0))

        self._field_label("Subject / Memo Topic")
        make_entry(self, self._subject_var,
                   "e.g. Monthly Budget Report").pack(anchor="w", padx=28, pady=(4, 0))

        self._field_label("Received By")
        make_entry(self, self._receiver_var,
                   "e.g. Maria Santos").pack(anchor="w", padx=28, pady=(4, 0))

    def _collect_values(self):
        dept    = self._dept_var.get()
        subject = self._subject_var.get().strip()
        rcvr    = self._receiver_var.get().strip()
        if dept == DEPARTMENTS[0]:
            messagebox.showwarning("Missing", "Please select a department.", parent=self)
            return None
        if not subject:
            messagebox.showwarning("Missing", "Please enter the subject.", parent=self)
            return None
        if not rcvr:
            messagebox.showwarning("Missing", "Please enter the receiver's name.", parent=self)
            return None
        return (dept, subject, rcvr)


# ── Past Outgoing Popup ───────────────────────────────────────────────────────
class PastOutgoingPopup(PastEntryPopup):
    TITLE  = "Log Past Outgoing Memo"
    ACCENT = TAB_COLORS["Outgoing Memo"]["accent"]

    def _build_fields(self):
        self._subject_var   = ctk.StringVar()
        self._recipient_var = ctk.StringVar(value=DEPARTMENTS[0])
        self._remarks_var   = ctk.StringVar()

        self._field_label("Subject / Title")
        make_entry(self, self._subject_var,
                   "e.g. Q2 Report").pack(anchor="w", padx=28, pady=(4, 0))

        self._field_label("Department Sent To")
        make_option(self, self._recipient_var, DEPARTMENTS,
                    self.ACCENT, "#154360").pack(anchor="w", padx=28, pady=(4, 0))

        self._field_label("Remarks (optional)")
        make_entry(self, self._remarks_var,
                   "e.g. Via courier").pack(anchor="w", padx=28, pady=(4, 0))

    def _collect_values(self):
        subject   = self._subject_var.get().strip()
        recipient = self._recipient_var.get()
        remarks   = self._remarks_var.get().strip()
        if not subject:
            messagebox.showwarning("Missing", "Please enter the subject.", parent=self)
            return None
        if recipient == DEPARTMENTS[0]:
            messagebox.showwarning("Missing", "Please select the destination department.", parent=self)
            return None
        return (subject, recipient, remarks)


# ── Base Tab Panel ────────────────────────────────────────────────────────────
class BaseTab(ctk.CTkFrame):
    TABLE_NAME      = ""
    SEARCH_COLS     = []
    FORM_TITLE      = ""
    FORM_HINT       = ""
    COL_SPECS       = []
    ACCENT          = ACCENT
    DARK            = ACCENT_DARK
    ICON            = "📋"
    PAST_POPUP_CLS  = None

    def __init__(self, parent):
        super().__init__(parent, fg_color=SURFACE, corner_radius=0)
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._build_form_panel()
        self._build_records_panel()
        self._refresh()

    def _build_form_panel(self):
        self._form_panel = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0,
                                        width=310, border_width=0)
        self._form_panel.grid(row=0, column=0, sticky="nsew")
        self._form_panel.pack_propagate(False)

        ctk.CTkFrame(self._form_panel, fg_color=self.ACCENT,
                     corner_radius=0, height=6).pack(fill="x")

        make_label(self._form_panel, self.FORM_TITLE,
                   size=16, weight="bold", color=self.ACCENT,
                   ).pack(anchor="w", padx=28, pady=(20, 2))
        make_label(self._form_panel, self.FORM_HINT,
                   size=11, weight="normal", color=TEXT_SUB,
                   ).pack(anchor="w", padx=28, pady=(0, 16))

        date_frame = ctk.CTkFrame(self._form_panel, fg_color=ROW_ALT, corner_radius=8)
        date_frame.pack(fill="x", padx=28, pady=(0, 14))
        make_label(date_frame, "Date & Time", size=11, weight="bold",
                   color=TEXT_SUB).pack(anchor="w", padx=14, pady=(10, 0))
        self._date_disp = make_label(date_frame, now_str(), size=13,
                                     weight="bold", color=self.ACCENT)
        self._date_disp.pack(anchor="w", padx=14, pady=(2, 10))

        self._build_fields(self._form_panel)

        # Primary log button
        make_button(self._form_panel, f"  ＋  Log {self.ICON}",
                    self._log_entry, self.ACCENT, self.DARK,
                    height=42).pack(anchor="w", padx=28, pady=(4, 6))

        # Past entry button
        make_button(self._form_panel, "  🕓  Log Past Memo",
                    self._open_past_popup,
                    SURFACE, ROW_ALT,
                    text_color=self.ACCENT, border=self.ACCENT,
                    height=36).pack(anchor="w", padx=28)

        ctk.CTkFrame(self._form_panel, fg_color=BORDER, height=1
                     ).pack(fill="x", padx=28, pady=18)

        make_button(self._form_panel, "  ↓  Export to CSV",
                    self._export_csv, SURFACE, ROW_ALT,
                    text_color=self.ACCENT, border=self.ACCENT,
                    height=36).pack(anchor="w", padx=28)

    def _open_past_popup(self):
        if self.PAST_POPUP_CLS:
            self.PAST_POPUP_CLS(self, on_save=self._save_past_entry)

    def _save_past_entry(self, date_str, *vals):
        pass

    def _field_label(self, parent, text):
        make_label(parent, text, size=11, weight="bold",
                   color=TEXT_SUB).pack(anchor="w", padx=28, pady=(6, 0))

    def _build_fields(self, parent):
        pass

    def _build_records_panel(self):
        panel = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        panel.grid(row=0, column=1, sticky="nsew")

        top = ctk.CTkFrame(panel, fg_color=SURFACE, corner_radius=0)
        top.pack(fill="x", padx=24, pady=(18, 0))
        make_label(top, "Records", size=16, weight="bold",
                   color=self.ACCENT).pack(side="left")
        self._count_lbl = make_label(top, "", size=11, weight="normal", color=TEXT_SUB)
        self._count_lbl.pack(side="left", padx=10)

        sf = ctk.CTkFrame(panel, fg_color=SURFACE, corner_radius=0)
        sf.pack(fill="x", padx=24, pady=(10, 0))
        self._search_var = ctk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._refresh())
        ctk.CTkEntry(sf, textvariable=self._search_var,
                     placeholder_text="🔍  Search records…",
                     fg_color=CARD, border_color=BORDER, text_color=TEXT_MAIN,
                     font=ctk.CTkFont(size=12), height=36, corner_radius=8,
                     ).pack(side="left", fill="x", expand=True)
        make_button(sf, "Clear", lambda: self._search_var.set(""),
                    BORDER, "#BFC9CA", text_color=TEXT_MAIN,
                    height=36, width=68, radius=8
                    ).pack(side="left", padx=(8, 0))

        tcard = ctk.CTkFrame(panel, fg_color=CARD, corner_radius=10,
                             border_width=1, border_color=BORDER)
        tcard.pack(fill="both", expand=True, padx=24, pady=12)

        hdr = ctk.CTkFrame(tcard, fg_color=self.ACCENT, corner_radius=0, height=38)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        for col_name, col_w, col_anchor in self.COL_SPECS:
            ctk.CTkLabel(hdr, text=col_name,
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color="#FFFFFF", width=col_w, anchor=col_anchor,
                         ).pack(side="left", padx=(10 if col_name == "#" else 6, 0))

        self._rows_frame = ctk.CTkScrollableFrame(
            tcard, fg_color=CARD, corner_radius=0, scrollbar_button_color=BORDER)
        self._rows_frame.pack(fill="both", expand=True)

    def _refresh(self):
        for w in self._rows_frame.winfo_children():
            w.destroy()
        q = self._search_var.get().strip()
        rows = _fetch(self.TABLE_NAME, self.SEARCH_COLS, q)
        n = len(rows)
        self._count_lbl.configure(text=f"({n} record{'s' if n!=1 else ''})")
        if hasattr(self, "_date_disp"):
            self._date_disp.configure(text=now_str())
        if not rows:
            make_label(self._rows_frame, "No records found.",
                       size=13, weight="normal", color=TEXT_SUB).pack(pady=40)
            return
        for idx, row in enumerate(rows):
            bg = CARD if idx % 2 == 0 else ROW_ALT
            rf = ctk.CTkFrame(self._rows_frame, fg_color=bg,
                              corner_radius=0, height=36)
            rf.pack(fill="x")
            rf.pack_propagate(False)
            self._render_row(rf, row)

    def _render_row(self, frame, row):
        pass

    def _add_delete_btn(self, frame, row_id):
        ctk.CTkButton(frame, text="✕", width=36, height=26,
                      fg_color="#FADBD8", hover_color="#F1948A",
                      text_color=DANGER,
                      font=ctk.CTkFont(size=12, weight="bold"),
                      corner_radius=6,
                      command=lambda: self._delete(row_id),
                      ).pack(side="left", padx=6)

    def _delete(self, row_id):
        if messagebox.askyesno("Confirm Delete",
                               f"Delete record #{row_id}? This cannot be undone."):
            _delete(self.TABLE_NAME, row_id)
            self._refresh()

    def _log_entry(self):
        pass

    def _export_csv(self):
        q = self._search_var.get().strip()
        rows = _fetch(self.TABLE_NAME, self.SEARCH_COLS, q)
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"{self.TABLE_NAME}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            title="Save CSV Export",
        )
        if path:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([c[0] for c in self.COL_SPECS])
                writer.writerows(rows)
            messagebox.showinfo("Export Successful", f"Records exported to:\n{path}")


# ── Tab 1 — Incoming Memo ─────────────────────────────────────────────────────
class MemoTab(BaseTab):
    TABLE_NAME     = "memos"
    SEARCH_COLS    = ["date", "department", "subject", "receiver"]
    FORM_TITLE     = "New Incoming Memo"
    FORM_HINT      = "Log a memo received from another department."
    ACCENT         = TAB_COLORS["Incoming Memo"]["accent"]
    DARK           = TAB_COLORS["Incoming Memo"]["dark"]
    ICON           = "📥"
    PAST_POPUP_CLS = PastMemoPopup
    COL_SPECS      = [
        ("#",           42,  "center"),
        ("Date & Time", 148, "w"),
        ("Department",  140, "w"),
        ("Subject",     210, "w"),
        ("Received By", 130, "w"),
        ("Action",       72, "center"),
    ]

    def _build_fields(self, parent):
        self._dept_var     = ctk.StringVar(value=DEPARTMENTS[0])
        self._subject_var  = ctk.StringVar()
        self._receiver_var = ctk.StringVar()

        self._field_label(parent, "From Department")
        make_option(parent, self._dept_var, DEPARTMENTS,
                    self.ACCENT, self.DARK).pack(anchor="w", padx=28, pady=(4, 10))

        self._field_label(parent, "Subject / Memo Topic")
        make_entry(parent, self._subject_var,
                   "e.g. Monthly Budget Report").pack(anchor="w", padx=28, pady=(4, 10))

        self._field_label(parent, "Received By")
        make_entry(parent, self._receiver_var,
                   "e.g. Maria Santos").pack(anchor="w", padx=28, pady=(4, 14))

    def _log_entry(self):
        dept    = self._dept_var.get()
        subject = self._subject_var.get().strip()
        rcvr    = self._receiver_var.get().strip()
        if dept == DEPARTMENTS[0]:
            return messagebox.showwarning("Missing", "Please select a department.")
        if not subject:
            return messagebox.showwarning("Missing", "Please enter the subject.")
        if not rcvr:
            return messagebox.showwarning("Missing", "Please enter the receiver's name.")
        _insert("memos", ["date","department","subject","receiver"],
                (now_str(), dept, subject, rcvr))
        self._dept_var.set(DEPARTMENTS[0])
        self._subject_var.set("")
        self._receiver_var.set("")
        self._refresh()
        messagebox.showinfo("Logged", f"Memo from '{dept}' has been recorded.")

    def _save_past_entry(self, date_str, dept, subject, rcvr):
        _insert("memos", ["date","department","subject","receiver"],
                (date_str, dept, subject, rcvr))
        self._refresh()
        messagebox.showinfo("Logged",
                            f"Past memo from '{dept}' ({date_str}) has been recorded.")

    def _render_row(self, frame, row):
        memo_id, date, dept, subj, recv = row
        for text, w, anch in [
            (str(memo_id), 42, "center"),
            (date, 148, "w"),
            (dept, 140, "w"),
            (subj, 210, "w"),
            (recv, 130, "w"),
        ]:
            ctk.CTkLabel(frame, text=text, font=ctk.CTkFont(size=12),
                         text_color=TEXT_MAIN, width=w, anchor=anch,
                         ).pack(side="left", padx=(10 if w==42 else 6, 0))
        self._add_delete_btn(frame, memo_id)


# ── Tab 2 — Outgoing Memo ─────────────────────────────────────────────────────
class OutgoingTab(BaseTab):
    TABLE_NAME     = "outgoing"
    SEARCH_COLS    = ["date","subject","recipient"]
    FORM_TITLE     = "New Outgoing Memo"
    FORM_HINT      = "Record a memo sent out from your office."
    ACCENT         = TAB_COLORS["Outgoing Memo"]["accent"]
    DARK           = TAB_COLORS["Outgoing Memo"]["dark"]
    ICON           = "📤"
    PAST_POPUP_CLS = PastOutgoingPopup
    COL_SPECS      = [
        ("#",             42,  "center"),
        ("Date",         132, "w"),
        ("Subject",      250, "w"),
        ("Sent To Dept.", 148, "w"),
        ("Remarks",      140, "w"),
        ("Action",        72, "center"),
    ]

    def _build_fields(self, parent):
        self._subject_var   = ctk.StringVar()
        self._recipient_var = ctk.StringVar(value=DEPARTMENTS[0])
        self._remarks_var   = ctk.StringVar()

        self._field_label(parent, "Subject / Title")
        make_entry(parent, self._subject_var,
                   "e.g. Q2 Report").pack(anchor="w", padx=28, pady=(4, 10))

        self._field_label(parent, "Department Sent To")
        make_option(parent, self._recipient_var, DEPARTMENTS,
                    self.ACCENT, self.DARK).pack(anchor="w", padx=28, pady=(4, 10))

        self._field_label(parent, "Remarks (optional)")
        make_entry(parent, self._remarks_var,
                   "e.g. Via courier").pack(anchor="w", padx=28, pady=(4, 14))

    def _log_entry(self):
        subject   = self._subject_var.get().strip()
        recipient = self._recipient_var.get()
        remarks   = self._remarks_var.get().strip()
        if not subject:
            return messagebox.showwarning("Missing", "Please enter the subject.")
        if recipient == DEPARTMENTS[0]:
            return messagebox.showwarning("Missing", "Please select the destination department.")
        _insert("outgoing",
                ["date","subject","recipient","remarks"],
                (now_str(), subject, recipient, remarks))
        self._recipient_var.set(DEPARTMENTS[0])
        self._subject_var.set("")
        self._remarks_var.set("")
        self._refresh()
        messagebox.showinfo("Logged", "Outgoing memo recorded successfully.")

    def _save_past_entry(self, date_str, subject, recipient, remarks):
        _insert("outgoing",
                ["date","subject","recipient","remarks"],
                (date_str, subject, recipient, remarks))
        self._refresh()
        messagebox.showinfo("Logged",
                            f"Past outgoing memo ({date_str}) has been recorded.")

    def _render_row(self, frame, row):
        rid, date, subj, recip, remarks = row
        for text, w, anch in [
            (str(rid),  42,  "center"),
            (date,     132, "w"),
            (subj,     250, "w"),
            (recip,    148, "w"),
            (remarks,  140, "w"),
        ]:
            ctk.CTkLabel(frame, text=text, font=ctk.CTkFont(size=12),
                         text_color=TEXT_MAIN, width=w, anchor=anch,
                         ).pack(side="left", padx=(10 if w==42 else 6, 0))
        self._add_delete_btn(frame, rid)


# ── Main App Window ───────────────────────────────────────────────────────────
class DocTrackerApp(ctk.CTk):
    TABS = [
        ("Incoming Memo", MemoTab),
        ("Outgoing Memo", OutgoingTab),
    ]

    def __init__(self):
        super().__init__()
        self.title("Office Memo Tracker")
        self.geometry("1100x700")
        self.minsize(960, 600)
        self.configure(fg_color=SURFACE)
        init_db()
        self._active_tab = None
        self._tab_widgets = {}
        self._tab_panels  = {}
        self._build_ui()
        self._switch_tab("Incoming Memo")

    def _build_ui(self):
        hdr = ctk.CTkFrame(self, fg_color=ACCENT, corner_radius=0, height=62)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr,
                     text="  🗂  Office Memo Tracker",
                     font=ctk.CTkFont(family="Georgia", size=21, weight="bold"),
                     text_color="#FFFFFF", anchor="w",
                     ).pack(side="left", padx=24)
        self._clock = ctk.CTkLabel(hdr, text="",
                                   font=ctk.CTkFont(size=13),
                                   text_color="#A9CCE3")
        self._clock.pack(side="right", padx=24)
        self._tick()

        tab_bar = ctk.CTkFrame(self, fg_color="#D6E4F0", corner_radius=0, height=46)
        tab_bar.pack(fill="x")
        tab_bar.pack_propagate(False)

        for name, _ in self.TABS:
            info = TAB_COLORS[name]
            btn = ctk.CTkButton(
                tab_bar,
                text=f"  {info['icon']}  {name}",
                command=lambda n=name: self._switch_tab(n),
                fg_color=TAB_IDLE_BG,
                hover_color=TAB_IDLE_BG,
                text_color=TAB_IDLE_TEXT,
                font=ctk.CTkFont(size=13, weight="bold"),
                height=46, width=200, corner_radius=0,
            )
            btn.pack(side="left")
            self._tab_widgets[name] = btn

        content = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        for name, TabClass in self.TABS:
            panel = TabClass(content)
            panel.grid(row=0, column=0, sticky="nsew")
            self._tab_panels[name] = panel

    def _switch_tab(self, name):
        if self._active_tab == name:
            return
        self._active_tab = name
        info = TAB_COLORS[name]
        for tab_name, btn in self._tab_widgets.items():
            if tab_name == name:
                btn.configure(fg_color=info["accent"],
                              hover_color=info["dark"],
                              text_color="#FFFFFF")
            else:
                btn.configure(fg_color=TAB_IDLE_BG,
                              hover_color=TAB_IDLE_BG,
                              text_color=TAB_IDLE_TEXT)
        self._tab_panels[name].tkraise()

    def _tick(self):
        self._clock.configure(
            text=datetime.now().strftime("  %A, %B %d %Y   %H:%M:%S"))
        self.after(1000, self._tick)


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = DocTrackerApp()
    app.mainloop()
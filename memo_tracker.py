import customtkinter as ctk
import sqlite3
import csv
import os
from datetime import datetime
from tkinter import messagebox, filedialog
import tkinter as tk
try:
    import openpyxl
    XLSX_SUPPORTED = True
except ImportError:
    XLSX_SUPPORTED = False

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
ACTIVE_FILTER_BG = "#D5F5E3"
ACTIVE_FILTER_FG = "#1E8449"

TAB_IDLE_BG   = "#D6E4F0"
TAB_IDLE_TEXT = ACCENT

TAB_COLORS = {
    "Incoming Memo": {"accent": "#1B4F72", "dark": "#154360", "icon": "📥"},
    "Outgoing Memo": {"accent": "#1A5276", "dark": "#154360", "icon": "📤"},
    "Departments":   {"accent": "#1D6A4A", "dark": "#155236", "icon": "🏢"},
    "Endorsements":  {"accent": "#4A235A", "dark": "#3B1A47", "icon": "📋"},
}

DEPARTMENTS = []


# ── Database ──────────────────────────────────────────────────────────────────
def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS memos (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            date       TEXT NOT NULL,
            department TEXT NOT NULL,
            subject      TEXT NOT NULL,
            receiver     TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS outgoing (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            date        TEXT NOT NULL,
            subject     TEXT NOT NULL,
            recipient   TEXT NOT NULL,
            remarks      TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS departments (
            id   INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        )
    """)
    cur.execute("SELECT COUNT(*) FROM departments")
    if cur.fetchone()[0] == 0:
        defaults = [
            "Office of the Mayor", "Office of the Vice Mayor", "City Council",
            "Information and Communication Technology Office",
            "Commission on Election (COMELEC)", "General Service Department",
            "Accounting Office", "Human Resource Development",
            "Office of the Secretary to the Sangguniang Panlungsod",
            "City Administrator's Office",
        ]
        cur.executemany("INSERT INTO departments (name) VALUES (?)",
                        [(d,) for d in defaults])
    cur.execute("""
        CREATE TABLE IF NOT EXISTS endorsements (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            date        TEXT NOT NULL,
            name        TEXT NOT NULL,
            subject     TEXT NOT NULL,
            endorsed_to TEXT NOT NULL,
            office      TEXT,
            remarks     TEXT
        )
    """)
    # Fix column order if office was added via ALTER TABLE (it would be last)
    cur.execute("PRAGMA table_info(endorsements)")
    cols = [r[1] for r in cur.fetchall()]
    if cols and cols.index("office") != 5:
        # Recreate with correct order, preserving data
        cur.execute("""
            CREATE TABLE IF NOT EXISTS endorsements_new (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                date        TEXT NOT NULL,
                name        TEXT NOT NULL,
                subject     TEXT NOT NULL,
                endorsed_to TEXT NOT NULL,
                office      TEXT,
                remarks     TEXT
            )
        """)
        cur.execute("""
            INSERT INTO endorsements_new (id, date, name, subject, endorsed_to, office, remarks)
            SELECT id, date, name, subject, endorsed_to, NULL, remarks FROM endorsements
        """)
        cur.execute("DROP TABLE endorsements")
        cur.execute("ALTER TABLE endorsements_new RENAME TO endorsements")
    con.commit()
    con.close()


def load_departments():
    global DEPARTMENTS
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("SELECT name FROM departments ORDER BY name")
    rows = cur.fetchall()
    con.close()
    DEPARTMENTS = ["Select Department…"] + [r[0] for r in rows]


def fetch_department_rows():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("SELECT id, name FROM departments ORDER BY name")
    rows = cur.fetchall()
    con.close()
    return rows


def insert_department(name):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    try:
        cur.execute("INSERT INTO departments (name) VALUES (?)", (name,))
        con.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        con.close()


def update_department(dept_id, name):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    try:
        cur.execute("UPDATE departments SET name=? WHERE id=?", (name, dept_id))
        con.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        con.close()


def delete_department(dept_id):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("DELETE FROM departments WHERE id=?", (dept_id,))
    con.commit()
    con.close()


def get_monthly_summary():
    """Return (incoming_count, outgoing_count) for the current month."""
    now = datetime.now()
    prefix = f"{now.year}-{now.month:02d}%"
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("SELECT COUNT(*) FROM memos WHERE date LIKE ?", (prefix,))
    inc = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM outgoing WHERE date LIKE ?", (prefix,))
    out = cur.fetchone()[0]
    con.close()
    return inc, out


# ── Generic DB helpers ────────────────────────────────────────────────────────
MONTH_NUMS = {
    "January": "01", "February": "02", "March": "03", "April": "04",
    "May": "05", "June": "06", "July": "07", "August": "08",
    "September": "09", "October": "10", "November": "11", "December": "12",
}


def _fetch(table, search_cols, query="", sort_col=None, sort_asc=True,
           month=None, year=None, dept_col=None, dept_val=None):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    direction = "ASC" if sort_asc else "DESC"
    order = f"{sort_col} {direction}" if sort_col else "id DESC"

    conditions = []
    params = []

    if query:
        q = f"%{query}%"
        conditions.append("(" + " OR ".join(f"{c} LIKE ?" for c in search_cols) + ")")
        params.extend([q] * len(search_cols))

    has_year  = year  and year  != "All Years"
    has_month = month and month != "All Months"
    if has_year and has_month:
        conditions.append("date LIKE ?")
        params.append(f"{year}-{MONTH_NUMS[month]}%")
    elif has_year:
        conditions.append("date LIKE ?")
        params.append(f"{year}%")
    elif has_month:
        conditions.append("date LIKE ?")
        params.append(f"____-{MONTH_NUMS[month]}%")

    if dept_col and dept_val and dept_val != "All Departments":
        conditions.append(f"{dept_col} = ?")
        params.append(dept_val)

    where_clause = f"WHERE {' AND '.join(conditions)}" if conditions else ""
    cur.execute(f"SELECT * FROM {table} {where_clause} ORDER BY {order}", params)
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


def _update(table, cols, vals, row_id):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    set_clause = ", ".join(f"{c}=?" for c in cols)
    cur.execute(f"UPDATE {table} SET {set_clause} WHERE id=?", (*vals, row_id))
    con.commit()
    con.close()


def read_import_file(path):
    """
    Read a .csv or .xlsx file and return (headers, rows).
    Returns (None, error_message) on failure.
    """
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".xlsx":
            if not XLSX_SUPPORTED:
                return None, "openpyxl is not installed.\nRun: pip install openpyxl"
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            all_rows = list(ws.iter_rows(values_only=True))
            wb.close()
            if not all_rows:
                return None, "The Excel file appears to be empty."
            # Find the first non-empty row as the header
            header_row = None
            data_start = 0
            for i, row in enumerate(all_rows):
                if any(v is not None for v in row):
                    header_row = row
                    data_start = i + 1
                    break
            if header_row is None:
                return None, "Could not find a header row in the Excel file."
            headers = [str(h).strip() if h is not None else f"Col{i}"
                       for i, h in enumerate(header_row)]
            rows = []
            for r in all_rows[data_start:]:
                if not any(v is not None for v in r):
                    continue  # skip empty rows
                row_dict = {}
                for i, v in enumerate(r):
                    if i < len(headers):
                        row_dict[headers[i]] = str(v).strip() if v is not None else ""
                rows.append(row_dict)
            return headers, rows
        else:
            with open(path, newline="", encoding="utf-8-sig") as f:
                raw = f.read()
            # Strip leading blank lines before passing to DictReader
            lines = raw.splitlines()
            non_empty = [l for l in lines if l.strip().replace(",", "")]
            if not non_empty:
                return None, "The CSV file appears to be empty."
            cleaned = "\n".join(non_empty)
            import io
            reader = csv.DictReader(io.StringIO(cleaned))
            headers = [h.strip() for h in (reader.fieldnames or [])]
            rows = []
            for r in reader:
                row_dict = {k.strip(): (v.strip() if v else "")
                            for k, v in r.items() if k is not None}
                rows.append(row_dict)
            return headers, rows
    except Exception as e:
        return None, str(e)


def _exists(table, match_cols, match_vals):
    """Return True if a row matching all given col=val pairs already exists."""
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    where = " AND ".join(f"{c} = ?" for c in match_cols)
    cur.execute(f"SELECT 1 FROM {table} WHERE {where} LIMIT 1", match_vals)
    found = cur.fetchone() is not None
    con.close()
    return found


def now_str():
    return datetime.now().strftime("%Y-%m-%d  %H:%M")


def now_date_str():
    return datetime.now().strftime("%Y-%m-%d")


def parse_date_flexible(s):
    """
    Accept many date formats and return YYYY-MM-DD string.
    Returns None if no format matches.
    """
    s = s.strip()
    formats = [
        "%m/%d/%Y",   # 07/14/2025
        "%m-%d-%Y",   # 07-14-2025
        "%m/%d/%y",   # 07/14/25
        "%B %d, %Y",  # July 14, 2025
        "%b %d, %Y",  # Jul 14, 2025
        "%d/%m/%Y",   # 14/07/2025
        "%Y-%m-%d",   # 2025-07-14  (already stored format)
        "%Y/%m/%d",   # 2025/07/14
    ]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


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


# ── Print helper ──────────────────────────────────────────────────────────────
def print_record(fields, title="Memo Record"):
    """Generate a simple HTML file and open it in the default browser for printing."""
    import tempfile, webbrowser
    rows_html = ""
    for label, value in fields:
        rows_html += f"""
        <tr>
            <td style="font-weight:bold;color:#5D6D7E;padding:8px 12px;
                       border-bottom:1px solid #D5D8DC;width:160px;">{label}</td>
            <td style="padding:8px 12px;border-bottom:1px solid #D5D8DC;">{value or '—'}</td>
        </tr>"""
    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>{title}</title>
<style>
  body {{ font-family: Georgia, serif; margin: 40px; color: #1C2833; }}
  h2   {{ color: #1B4F72; border-bottom: 2px solid #1B4F72; padding-bottom: 8px; }}
  table {{ border-collapse: collapse; width: 100%; max-width: 600px; }}
  @media print {{ button {{ display:none; }} }}
</style>
</head><body>
<h2>🗂 Office Memo Tracker — {title}</h2>
<p style="color:#5D6D7E;font-size:13px;">Printed: {datetime.now().strftime('%B %d, %Y  %H:%M')}</p>
<table>{rows_html}</table>
<br>
<button onclick="window.print()" style="padding:8px 20px;font-size:14px;
  background:#1B4F72;color:white;border:none;border-radius:6px;cursor:pointer;">
  🖨 Print
</button>
</body></html>"""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".html",
                                     mode="w", encoding="utf-8")
    tmp.write(html)
    tmp.close()
    webbrowser.open(f"file://{tmp.name}")


# ── Edit Record Popups ────────────────────────────────────────────────────────
class EditMemoPopup(ctk.CTkToplevel):
    def __init__(self, parent, row, on_save, on_delete):
        super().__init__(parent)
        memo_id, date, dept, subj, recv = row
        self._id        = memo_id
        self._on_save   = on_save
        self._on_delete = on_delete
        self.title(f"Edit Memo #{memo_id}")
        self.geometry("520x540")
        self.resizable(False, False)
        self.grab_set()
        self.configure(fg_color=SURFACE)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        accent = TAB_COLORS["Incoming Memo"]["accent"]
        ctk.CTkFrame(self, fg_color=accent, corner_radius=0, height=6
                     ).grid(row=0, column=0, sticky="ew")

        scroll = ctk.CTkScrollableFrame(self, fg_color=SURFACE, corner_radius=0,
                                        scrollbar_button_color=BORDER)
        scroll.grid(row=1, column=0, sticky="nsew")

        make_label(scroll, f"Edit Incoming Memo #{memo_id}", size=15,
                   weight="bold", color=accent).pack(anchor="w", padx=28, pady=(18, 16))

        for lbl, var_name, default in [
            ("Date (YYYY-MM-DD)", "_date_var", date.split("  ")[0].strip()),
            ("Subject", "_subj_var", subj),
            ("Received By", "_recv_var", recv),
        ]:
            make_label(scroll, lbl, size=11, weight="bold",
                       color=TEXT_SUB).pack(anchor="w", padx=28, pady=(0, 2))
            setattr(self, var_name, ctk.StringVar(value=default))
            make_entry(scroll, getattr(self, var_name), width=440
                       ).pack(fill="x", padx=28, pady=(0, 10))

        make_label(scroll, "From Department", size=11, weight="bold",
                   color=TEXT_SUB).pack(anchor="w", padx=28, pady=(0, 2))
        self._dept_var = ctk.StringVar(value=dept)
        make_option(scroll, self._dept_var, DEPARTMENTS[1:],
                    accent, "#154360", width=440).pack(anchor="w", padx=28, pady=(0, 10))

        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.pack(fill="x", padx=28, pady=(10, 24))
        make_button(btn_frame, "  💾  Save Changes", self._save,
                    accent, "#154360", height=40, width=180).pack(side="left")
        make_button(btn_frame, "Cancel", self.destroy,
                    BORDER, "#BFC9CA", text_color=TEXT_MAIN,
                    height=40, width=90, radius=8).pack(side="left", padx=(10, 0))
        make_button(btn_frame, "  🗑  Delete", self._delete,
                    "#FADBD8", "#F1948A", text_color=DANGER,
                    height=40, width=110, radius=8).pack(side="right")

    def _save(self):
        date_raw = self._date_var.get().strip()
        try:
            datetime.strptime(date_raw, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("Invalid Date", "Use YYYY-MM-DD format.", parent=self)
            return
        subj = self._subj_var.get().strip()
        recv = self._recv_var.get().strip()
        if not subj:
            messagebox.showwarning("Missing", "Subject cannot be empty.", parent=self)
            return
        if not recv:
            messagebox.showwarning("Missing", "Received By cannot be empty.", parent=self)
            return
        # Preserve original time if present, otherwise use current time
        self._on_save(self._id, date_raw, self._dept_var.get(), subj, recv)
        self.destroy()

    def _delete(self):
        if messagebox.askyesno("Confirm Delete",
                               f"Delete memo #{self._id}?", parent=self):
            self._on_delete(self._id)
            self.destroy()


class EditOutgoingPopup(ctk.CTkToplevel):
    def __init__(self, parent, row, on_save, on_delete):
        super().__init__(parent)
        rid, date, subj, recip, remarks = row
        self._id        = rid
        self._on_save   = on_save
        self._on_delete = on_delete
        self.title(f"Edit Outgoing Memo #{rid}")
        self.geometry("520x520")
        self.resizable(False, False)
        self.grab_set()
        self.configure(fg_color=SURFACE)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        accent = TAB_COLORS["Outgoing Memo"]["accent"]
        ctk.CTkFrame(self, fg_color=accent, corner_radius=0, height=6
                     ).grid(row=0, column=0, sticky="ew")

        scroll = ctk.CTkScrollableFrame(self, fg_color=SURFACE, corner_radius=0,
                                        scrollbar_button_color=BORDER)
        scroll.grid(row=1, column=0, sticky="nsew")

        make_label(scroll, f"Edit Outgoing Memo #{rid}", size=15,
                   weight="bold", color=accent).pack(anchor="w", padx=28, pady=(18, 16))

        for lbl, var_name, default in [
            ("Date (YYYY-MM-DD)", "_date_var", date.split("  ")[0].strip()),
            ("Subject", "_subj_var", subj),
            ("Remarks (optional)", "_remarks_var", remarks or ""),
        ]:
            make_label(scroll, lbl, size=11, weight="bold",
                       color=TEXT_SUB).pack(anchor="w", padx=28, pady=(0, 2))
            setattr(self, var_name, ctk.StringVar(value=default))
            make_entry(scroll, getattr(self, var_name), width=440
                       ).pack(fill="x", padx=28, pady=(0, 10))

        make_label(scroll, "Department Sent To", size=11, weight="bold",
                   color=TEXT_SUB).pack(anchor="w", padx=28, pady=(0, 2))
        self._recip_var = ctk.StringVar(value=recip)
        make_option(scroll, self._recip_var, DEPARTMENTS[1:],
                    accent, "#154360", width=440).pack(anchor="w", padx=28, pady=(0, 10))

        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.pack(fill="x", padx=28, pady=(10, 24))
        make_button(btn_frame, "  💾  Save Changes", self._save,
                    accent, "#154360", height=40, width=180).pack(side="left")
        make_button(btn_frame, "Cancel", self.destroy,
                    BORDER, "#BFC9CA", text_color=TEXT_MAIN,
                    height=40, width=90, radius=8).pack(side="left", padx=(10, 0))
        make_button(btn_frame, "  🗑  Delete", self._delete,
                    "#FADBD8", "#F1948A", text_color=DANGER,
                    height=40, width=110, radius=8).pack(side="right")

    def _save(self):
        date_raw = self._date_var.get().strip()
        try:
            datetime.strptime(date_raw, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("Invalid Date", "Use YYYY-MM-DD format.", parent=self)
            return
        subj = self._subj_var.get().strip()
        if not subj:
            messagebox.showwarning("Missing", "Subject cannot be empty.", parent=self)
            return
        self._on_save(self._id, date_raw, self._subj_var.get().strip(),
                      self._recip_var.get(), self._remarks_var.get().strip())
        self.destroy()

    def _delete(self):
        if messagebox.askyesno("Confirm Delete",
                               f"Delete outgoing memo #{self._id}?", parent=self):
            self._on_delete(self._id)
            self.destroy()



# ── Edit Endorsement Popup ────────────────────────────────────────────────────
class EditEndorsementPopup(ctk.CTkToplevel):
    def __init__(self, parent, row, on_save, on_delete):
        super().__init__(parent)
        if len(row) == 7:
            rid, date, name, subj, endorsed_to, office, remarks = row
        else:
            rid, date, name, subj, endorsed_to, remarks = row
            office = ""
        self._id        = rid
        self._on_save   = on_save
        self._on_delete = on_delete
        self.title(f"Edit Endorsement #{rid}")
        self.geometry("520x520")
        self.resizable(False, False)
        self.grab_set()
        self.configure(fg_color=SURFACE)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        accent = TAB_COLORS["Endorsements"]["accent"]
        ctk.CTkFrame(self, fg_color=accent, corner_radius=0, height=6
                     ).grid(row=0, column=0, sticky="ew")

        scroll = ctk.CTkScrollableFrame(self, fg_color=SURFACE, corner_radius=0,
                                        scrollbar_button_color=BORDER)
        scroll.grid(row=1, column=0, sticky="nsew")

        make_label(scroll, f"Edit Endorsement #{rid}", size=15,
                   weight="bold", color=accent).pack(anchor="w", padx=28, pady=(18, 16))

        for lbl, var_name, default in [
            ("Date (YYYY-MM-DD)",       "_date_var",        date.split("  ")[0].strip()),
            ("Name (Who is Endorsed)",  "_name_var",        name),
            ("Subject",                 "_subj_var",        subj),
            ("Endorsed To",             "_endorsed_to_var", endorsed_to),
            ("Office",                  "_office_var",      office or ""),
            ("Remarks (optional)",      "_remarks_var",     remarks or ""),
        ]:
            make_label(scroll, lbl, size=11, weight="bold",
                       color=TEXT_SUB).pack(anchor="w", padx=28, pady=(0, 2))
            setattr(self, var_name, ctk.StringVar(value=default))
            make_entry(scroll, getattr(self, var_name), width=440
                       ).pack(fill="x", padx=28, pady=(0, 10))

        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.pack(fill="x", padx=28, pady=(10, 24))
        make_button(btn_frame, "  💾  Save Changes", self._save,
                    accent, "#3B1A47", height=40, width=180).pack(side="left")
        make_button(btn_frame, "Cancel", self.destroy,
                    BORDER, "#BFC9CA", text_color=TEXT_MAIN,
                    height=40, width=90, radius=8).pack(side="left", padx=(10, 0))
        make_button(btn_frame, "  🗑  Delete", self._delete,
                    "#FADBD8", "#F1948A", text_color=DANGER,
                    height=40, width=110, radius=8).pack(side="right")

    def _save(self):
        date_raw = self._date_var.get().strip()
        parsed = parse_date_flexible(date_raw) if date_raw else None
        if not parsed:
            messagebox.showwarning("Invalid Date",
                                   "Use YYYY-MM-DD or MM/DD/YYYY format.", parent=self)
            return
        name        = self._name_var.get().strip()
        subj        = self._subj_var.get().strip()
        endorsed_to = self._endorsed_to_var.get().strip()
        office      = self._office_var.get().strip()
        remarks     = self._remarks_var.get().strip()
        if not name:
            messagebox.showwarning("Missing", "Name cannot be empty.", parent=self)
            return
        if not subj:
            messagebox.showwarning("Missing", "Subject cannot be empty.", parent=self)
            return
        self._on_save(self._id, parsed, name, subj, endorsed_to, office, remarks)
        self.destroy()

    def _delete(self):
        if messagebox.askyesno("Confirm Delete",
                               f"Delete endorsement #{self._id}?", parent=self):
            self._on_delete(self._id)
            self.destroy()


# ── Past Entry Popups ─────────────────────────────────────────────────────────
class PastEntryPopup(ctk.CTkToplevel):
    TITLE  = "Log Past Entry"
    ACCENT = ACCENT

    def __init__(self, parent, on_save):
        super().__init__(parent)
        self._on_save = on_save
        self.title(self.TITLE)
        self.geometry("480x600")
        self.minsize(480, 600)
        self.resizable(False, False)
        self.grab_set()
        self.configure(fg_color=SURFACE)
        self.rowconfigure(2, weight=1)
        self.columnconfigure(0, weight=1)
        self._build_ui()

    def _build_ui(self):
        ctk.CTkFrame(self, fg_color=self.ACCENT, corner_radius=0, height=6
                     ).grid(row=0, column=0, sticky="ew")
        hdr = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        hdr.grid(row=1, column=0, sticky="ew", padx=28, pady=(18, 0))
        make_label(hdr, self.TITLE, size=15, weight="bold",
                   color=self.ACCENT).pack(anchor="w")
        make_label(hdr, "Enter the original date of this memo.",
                   size=11, weight="normal", color=TEXT_SUB
                   ).pack(anchor="w", pady=(2, 14))

        scroll = ctk.CTkScrollableFrame(self, fg_color=SURFACE, corner_radius=0,
                                        scrollbar_button_color=BORDER)
        scroll.grid(row=2, column=0, sticky="nsew")

        # Date
        date_card = ctk.CTkFrame(scroll, fg_color=WARN_BG, corner_radius=8,
                                 border_width=1, border_color=WARN_BORDER)
        date_card.pack(fill="x", padx=28, pady=(0, 16))
        make_label(date_card, "📅  Date  (MM/DD/YYYY)", size=11,
                   weight="bold", color="#7D6608"
                   ).pack(anchor="w", padx=14, pady=(10, 2))
        self._date_var = ctk.StringVar()
        ctk.CTkEntry(date_card, textvariable=self._date_var,
                     placeholder_text="e.g. 07/14/2025",
                     fg_color=CARD, border_color=WARN_BORDER,
                     text_color=TEXT_MAIN, font=ctk.CTkFont(size=13), height=34,
                     ).pack(fill="x", padx=14, pady=(0, 10))

        self._scroll = scroll
        self._build_fields()

        btn_row = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        btn_row.grid(row=3, column=0, sticky="ew", padx=28, pady=(12, 20))
        make_button(btn_row, "  ✔  Save Past Entry", self._save,
                    self.ACCENT, "#154360", height=40, width=190).pack(side="left")
        make_button(btn_row, "Cancel", self.destroy,
                    BORDER, "#BFC9CA", text_color=TEXT_MAIN,
                    height=40, width=100, radius=8).pack(side="left", padx=(10, 0))

    def _build_fields(self):
        pass

    def _collect_values(self):
        return ()

    def _save(self):
        raw = self._date_var.get().strip()
        date_str = parse_date_flexible(raw)
        if not date_str:
            messagebox.showwarning("Invalid Date",
                                   "Could not read that date.\n"
                                   "Try formats like:  07/14/2025  or  July 14, 2025",
                                   parent=self)
            return
        vals = self._collect_values()
        if vals is None:
            return
        self._on_save(date_str, *vals)
        self.destroy()

    def _field_label(self, text):
        make_label(self._scroll, text, size=11, weight="bold",
                   color=TEXT_SUB).pack(anchor="w", padx=28, pady=(6, 0))


class PastMemoPopup(PastEntryPopup):
    TITLE  = "Log Past Incoming Memo"
    ACCENT = TAB_COLORS["Incoming Memo"]["accent"]

    def _build_fields(self):
        self._dept_var     = ctk.StringVar(value=DEPARTMENTS[0])
        self._subject_var  = ctk.StringVar()
        self._receiver_var = ctk.StringVar()

        self._field_label("From Department")
        make_option(self._scroll, self._dept_var, DEPARTMENTS,
                    self.ACCENT, "#154360", width=400
                    ).pack(anchor="w", padx=28, pady=(4, 0))
        self._field_label("Subject / Memo Topic")
        make_entry(self._scroll, self._subject_var,
                   "e.g. Monthly Budget Report", width=400
                   ).pack(fill="x", padx=28, pady=(4, 0))
        self._field_label("Received By")
        make_entry(self._scroll, self._receiver_var,
                   "e.g. Maria Santos", width=400
                   ).pack(fill="x", padx=28, pady=(4, 0))

    def _collect_values(self):
        dept = self._dept_var.get()
        subj = self._subject_var.get().strip()
        recv = self._receiver_var.get().strip()
        if dept == DEPARTMENTS[0]:
            messagebox.showwarning("Missing", "Please select a department.", parent=self)
            return None
        if not subj:
            messagebox.showwarning("Missing", "Please enter the subject.", parent=self)
            return None
        if not recv:
            messagebox.showwarning("Missing", "Please enter the receiver's name.", parent=self)
            return None
        return (dept, subj, recv)


class PastOutgoingPopup(PastEntryPopup):
    TITLE  = "Log Past Outgoing Memo"
    ACCENT = TAB_COLORS["Outgoing Memo"]["accent"]

    def _build_fields(self):
        self._subject_var   = ctk.StringVar()
        self._recipient_var = ctk.StringVar(value=DEPARTMENTS[0])
        self._remarks_var   = ctk.StringVar()

        self._field_label("Subject / Title")
        make_entry(self._scroll, self._subject_var,
                   "e.g. Q2 Report", width=400
                   ).pack(fill="x", padx=28, pady=(4, 0))
        self._field_label("Department Sent To")
        make_option(self._scroll, self._recipient_var, DEPARTMENTS,
                    self.ACCENT, "#154360", width=400
                    ).pack(anchor="w", padx=28, pady=(4, 0))
        self._field_label("Remarks (optional)")
        make_entry(self._scroll, self._remarks_var,
                   "e.g. Via courier", width=400
                   ).pack(fill="x", padx=28, pady=(4, 0))

    def _collect_values(self):
        subj  = self._subject_var.get().strip()
        recip = self._recipient_var.get()
        remarks = self._remarks_var.get().strip()
        if not subj:
            messagebox.showwarning("Missing", "Please enter the subject.", parent=self)
            return None
        if recip == DEPARTMENTS[0]:
            messagebox.showwarning("Missing", "Please select the destination department.", parent=self)
            return None
        return (subj, recip, remarks)


class PastEndorsementPopup(PastEntryPopup):
    TITLE  = "Log Past Endorsement"
    ACCENT = TAB_COLORS["Endorsements"]["accent"]

    def _build_fields(self):
        self._name_var        = ctk.StringVar()
        self._subject_var     = ctk.StringVar()
        self._endorsed_to_var = ctk.StringVar()
        self._office_var      = ctk.StringVar()
        self._remarks_var     = ctk.StringVar()

        self._field_label("Name (Who is Endorsed)")
        make_entry(self._scroll, self._name_var,
                   "e.g. Juan dela Cruz", width=400
                   ).pack(fill="x", padx=28, pady=(4, 0))
        self._field_label("Subject")
        make_entry(self._scroll, self._subject_var,
                   "e.g. Employment Endorsement", width=400
                   ).pack(fill="x", padx=28, pady=(4, 0))
        self._field_label("Endorsed To")
        make_entry(self._scroll, self._endorsed_to_var,
                   "e.g. City Mayor", width=400
                   ).pack(fill="x", padx=28, pady=(4, 0))
        self._office_var = ctk.StringVar()
        self._field_label("Office")
        make_entry(self._scroll, self._office_var,
                   "e.g. Office of the City Mayor", width=400
                   ).pack(fill="x", padx=28, pady=(4, 0))
        self._field_label("Remarks (optional)")
        make_entry(self._scroll, self._remarks_var,
                   "e.g. For signature", width=400
                   ).pack(fill="x", padx=28, pady=(4, 0))

    def _collect_values(self):
        name        = self._name_var.get().strip()
        subj        = self._subject_var.get().strip()
        endorsed_to = self._endorsed_to_var.get().strip()
        office      = self._office_var.get().strip()
        remarks     = self._remarks_var.get().strip()
        if not name:
            messagebox.showwarning("Missing", "Please enter the name.", parent=self)
            return None
        if not subj:
            messagebox.showwarning("Missing", "Please enter the subject.", parent=self)
            return None
        if not endorsed_to:
            messagebox.showwarning("Missing", "Please enter who it is endorsed to.", parent=self)
            return None
        return (name, subj, endorsed_to, office, remarks)


# ── Record Detail Popup ───────────────────────────────────────────────────────
class RecordDetailPopup(ctk.CTkToplevel):
    def __init__(self, parent, fields, accent):
        super().__init__(parent)
        self.title("Record Details")
        self.geometry("560x480")
        self.resizable(False, False)
        self.grab_set()
        self.configure(fg_color=SURFACE)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkFrame(self, fg_color=accent, corner_radius=0, height=6
                     ).grid(row=0, column=0, sticky="ew")
        make_label(self, "Record Details", size=15, weight="bold",
                   color=accent).grid(row=1, column=0, sticky="w",
                                      padx=28, pady=(18, 8))

        card = ctk.CTkScrollableFrame(self, fg_color=CARD, corner_radius=10,
                                      border_width=1, border_color=BORDER)
        card.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 8))
        card.grid_columnconfigure(0, weight=1)

        for label, value in fields:
            make_label(card, label, size=11, weight="bold",
                       color=TEXT_SUB).pack(anchor="w", padx=16, pady=(12, 0))
            ctk.CTkLabel(card, text=value or "—",
                         font=ctk.CTkFont(size=13), text_color=TEXT_MAIN,
                         justify="left", anchor="w", wraplength=460,
                         ).pack(anchor="w", fill="x", padx=16, pady=(2, 0))
            ctk.CTkFrame(card, fg_color=BORDER, height=1
                         ).pack(fill="x", padx=16, pady=(10, 0))

        btn_row = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        btn_row.grid(row=3, column=0, pady=(0, 16))
        make_button(btn_row, "Close", self.destroy, BORDER, "#BFC9CA",
                    text_color=TEXT_MAIN, height=36, width=100, radius=8
                    ).pack(side="left")


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
    SORT_MAP        = {}
    DEPT_FILTER_COL = None
    PAST_LABEL      = "Memo"
    COMPACT         = False  # set True on tabs with many fields

    def __init__(self, parent):
        super().__init__(parent, fg_color=SURFACE, corner_radius=0)
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._sort_col     = None
        self._sort_asc     = True
        self._page         = 0
        self._page_size    = 50
        self._total_rows   = 0
        self._search_after = None
        self._filter_month = "All Months"
        self._filter_year  = "All Years"
        self._filter_dept  = "All Departments"
        self._build_form_panel()
        self._build_records_panel()
        self._refresh()

    def _build_form_panel(self):
        self._form_panel = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0,
                                        width=330, border_width=0)
        self._form_panel.grid(row=0, column=0, sticky="nsew")
        self._form_panel.pack_propagate(False)

        ctk.CTkFrame(self._form_panel, fg_color=self.ACCENT,
                     corner_radius=0, height=6).pack(fill="x")
        tp = (10, 1) if self.COMPACT else (14, 1)
        make_label(self._form_panel, self.FORM_TITLE, size=15, weight="bold",
                   color=self.ACCENT).pack(anchor="w", padx=20, pady=tp)
        make_label(self._form_panel, self.FORM_HINT, size=11, weight="normal",
                   color=TEXT_SUB).pack(anchor="w", padx=20, pady=(0, 4 if self.COMPACT else 8))

        self._build_fields(self._form_panel)

        make_button(self._form_panel, f"  ＋  Log {self.ICON}",
                    self._log_entry, self.ACCENT, self.DARK,
                    height=38, width=290).pack(anchor="w", padx=20,
                    pady=(4 if self.COMPACT else 6, 3 if self.COMPACT else 4))
        if self.PAST_POPUP_CLS:
            make_button(self._form_panel, f"  🕓  Log Past {self.PAST_LABEL}",
                        self._open_past_popup, SURFACE, ROW_ALT,
                        text_color=self.ACCENT, border=self.ACCENT,
                        height=34, width=290).pack(anchor="w", padx=20)
        ctk.CTkFrame(self._form_panel, fg_color=BORDER, height=1
                     ).pack(fill="x", padx=20, pady=6)
        make_button(self._form_panel, "  ↓  Export to CSV",
                    self._export_csv, SURFACE, ROW_ALT,
                    text_color=self.ACCENT, border=self.ACCENT,
                    height=34, width=290).pack(anchor="w", padx=20, pady=(0, 4))
        make_button(self._form_panel, "  ↑  Import from CSV",
                    self._import_csv, SURFACE, ROW_ALT,
                    text_color=self.ACCENT, border=self.ACCENT,
                    height=34, width=290).pack(anchor="w", padx=20)

    def _open_past_popup(self):
        if self.PAST_POPUP_CLS:
            self.PAST_POPUP_CLS(self, on_save=self._save_past_entry)

    def _save_past_entry(self, date_str, *vals):
        pass

    def _field_label(self, parent, text):
        make_label(parent, text, size=11, weight="bold",
                   color=TEXT_SUB).pack(anchor="w", padx=20, pady=(4, 0))

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
        # Active filter indicator
        self._filter_indicator = make_label(top, "", size=11, weight="bold",
                                            color=ACTIVE_FILTER_FG)
        self._filter_indicator.pack(side="left", padx=4)

        sf = ctk.CTkFrame(panel, fg_color=SURFACE, corner_radius=0)
        sf.pack(fill="x", padx=24, pady=(10, 0))
        self._search_var   = ctk.StringVar()
        self._search_after = None
        self._search_var.trace_add("write", lambda *_: self._debounced_refresh())
        ctk.CTkEntry(sf, textvariable=self._search_var,
                     placeholder_text="🔍  Search records…",
                     fg_color=CARD, border_color=BORDER, text_color=TEXT_MAIN,
                     font=ctk.CTkFont(size=12), height=36, corner_radius=8,
                     ).pack(side="left", fill="x", expand=True)
        make_button(sf, "Clear", lambda: self._search_var.set(""),
                    BORDER, "#BFC9CA", text_color=TEXT_MAIN,
                    height=36, width=68, radius=8
                    ).pack(side="left", padx=(8, 0))

        # Filter bar
        fbar = ctk.CTkFrame(panel, fg_color=CARD, corner_radius=8,
                            border_width=1, border_color=BORDER)
        fbar.pack(fill="x", padx=24, pady=(6, 0))

        make_label(fbar, "Month:", size=11, weight="bold",
                   color=TEXT_SUB).pack(side="left", padx=(12, 4), pady=8)
        months = ["All Months","January","February","March","April","May","June",
                  "July","August","September","October","November","December"]
        self._month_var = ctk.StringVar(value="All Months")
        ctk.CTkOptionMenu(fbar, variable=self._month_var, values=months,
                          command=lambda _: self._apply_filters(),
                          fg_color=CARD, button_color=self.ACCENT,
                          button_hover_color=self.DARK, text_color=TEXT_MAIN,
                          dropdown_fg_color=CARD, dropdown_text_color=TEXT_MAIN,
                          dropdown_hover_color=ROW_ALT,
                          font=ctk.CTkFont(size=12), width=130, height=30,
                          ).pack(side="left", padx=(0, 10), pady=6)

        make_label(fbar, "Year:", size=11, weight="bold",
                   color=TEXT_SUB).pack(side="left", padx=(0, 4))
        cur_year = datetime.now().year
        years = ["All Years"] + [str(y) for y in range(cur_year, cur_year - 10, -1)]
        self._year_var = ctk.StringVar(value="All Years")
        ctk.CTkOptionMenu(fbar, variable=self._year_var, values=years,
                          command=lambda _: self._apply_filters(),
                          fg_color=CARD, button_color=self.ACCENT,
                          button_hover_color=self.DARK, text_color=TEXT_MAIN,
                          dropdown_fg_color=CARD, dropdown_text_color=TEXT_MAIN,
                          dropdown_hover_color=ROW_ALT,
                          font=ctk.CTkFont(size=12), width=100, height=30,
                          ).pack(side="left", padx=(0, 14), pady=6)

        if self.DEPT_FILTER_COL:
            make_label(fbar, "Department:", size=11, weight="bold",
                       color=TEXT_SUB).pack(side="left", padx=(0, 4))
            self._dept_filter_var = ctk.StringVar(value="All Departments")
            self._dept_filter_menu = ctk.CTkOptionMenu(
                fbar, variable=self._dept_filter_var,
                values=["All Departments"] + DEPARTMENTS[1:],
                command=lambda _: self._apply_filters(),
                fg_color=CARD, button_color=self.ACCENT,
                button_hover_color=self.DARK, text_color=TEXT_MAIN,
                dropdown_fg_color=CARD, dropdown_text_color=TEXT_MAIN,
                dropdown_hover_color=ROW_ALT,
                font=ctk.CTkFont(size=12), width=200, height=30)
            self._dept_filter_menu.pack(side="left", padx=(0, 10), pady=6)
        else:
            self._dept_filter_var  = ctk.StringVar(value="All Departments")
            self._dept_filter_menu = None

        make_button(fbar, "Reset", self._reset_filters,
                    BORDER, "#BFC9CA", text_color=TEXT_MAIN,
                    height=30, width=70, radius=6
                    ).pack(side="left", padx=(0, 10), pady=6)

        tcard = ctk.CTkFrame(panel, fg_color=CARD, corner_radius=10,
                             border_width=1, border_color=BORDER)
        tcard.pack(fill="both", expand=True, padx=24, pady=6)

        hdr = ctk.CTkFrame(tcard, fg_color=self.ACCENT, corner_radius=0, height=38)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        self._hdr_frame = hdr
        self._build_header()

        self._rows_frame = ctk.CTkScrollableFrame(
            tcard, fg_color=CARD, corner_radius=0, scrollbar_button_color=BORDER)
        self._rows_frame.pack(fill="both", expand=True)

        # Pagination bar
        self._page_bar = ctk.CTkFrame(tcard, fg_color=SURFACE, corner_radius=0, height=36)
        self._page_bar.pack(fill="x")
        self._page_bar.pack_propagate(False)
        self._prev_btn = ctk.CTkButton(self._page_bar, text="◀  Prev",
                                       command=self._prev_page,
                                       fg_color=BORDER, hover_color="#BFC9CA",
                                       text_color=TEXT_MAIN,
                                       font=ctk.CTkFont(size=11),
                                       height=28, width=80, corner_radius=6)
        self._prev_btn.pack(side="left", padx=(8, 4), pady=4)
        self._page_lbl = ctk.CTkLabel(self._page_bar, text="",
                                      font=ctk.CTkFont(size=11), text_color=TEXT_SUB)
        self._page_lbl.pack(side="left", padx=8)
        self._next_btn = ctk.CTkButton(self._page_bar, text="Next  ▶",
                                       command=self._next_page,
                                       fg_color=BORDER, hover_color="#BFC9CA",
                                       text_color=TEXT_MAIN,
                                       font=ctk.CTkFont(size=11),
                                       height=28, width=80, corner_radius=6)
        self._next_btn.pack(side="left", padx=(4, 8), pady=4)

    def _refresh(self):
        for w in self._rows_frame.winfo_children():
            w.destroy()
        q = self._search_var.get().strip()
        all_rows = _fetch(self.TABLE_NAME, self.SEARCH_COLS, q,
                          self._sort_col, self._sort_asc,
                          month=self._filter_month, year=self._filter_year,
                          dept_col=self.DEPT_FILTER_COL, dept_val=self._filter_dept)
        self._total_rows = len(all_rows)

        total_pages = max(1, -(-self._total_rows // self._page_size))
        if self._page >= total_pages:
            self._page = total_pages - 1

        start = self._page * self._page_size
        rows  = all_rows[start: start + self._page_size]

        self._count_lbl.configure(
            text=f"({self._total_rows} record{'s' if self._total_rows!=1 else ''})")

        # Active filter indicator
        active = []
        if self._filter_month != "All Months": active.append(self._filter_month)
        if self._filter_year  != "All Years":  active.append(self._filter_year)
        if self._filter_dept  != "All Departments": active.append("dept filtered")
        self._filter_indicator.configure(
            text=f"● Filtered: {', '.join(active)}" if active else "")

        self._page_lbl.configure(
            text=(f"Page {self._page+1} of {total_pages}  "
                  f"({start+1}–{min(start+self._page_size, self._total_rows)} "
                  f"of {self._total_rows})") if self._total_rows else "No records")
        self._prev_btn.configure(state="normal" if self._page > 0 else "disabled")
        self._next_btn.configure(state="normal" if self._page < total_pages-1 else "disabled")

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

    def _prev_page(self):
        if self._page > 0:
            self._page -= 1
            self._refresh()

    def _next_page(self):
        total_pages = max(1, -(-self._total_rows // self._page_size))
        if self._page < total_pages - 1:
            self._page += 1
            self._refresh()

    def _apply_filters(self):
        self._filter_month = self._month_var.get()
        self._filter_year  = self._year_var.get()
        self._filter_dept  = self._dept_filter_var.get()
        self._page = 0
        self._debounced_refresh()

    def _reset_filters(self):
        self._month_var.set("All Months")
        self._year_var.set("All Years")
        self._dept_filter_var.set("All Departments")
        self._filter_month = "All Months"
        self._filter_year  = "All Years"
        self._filter_dept  = "All Departments"
        self._page = 0
        self._refresh()

    def _debounced_refresh(self):
        if self._search_after is not None:
            self.after_cancel(self._search_after)
        self._search_after = self.after(200, self._refresh)

    def _render_row(self, frame, row):
        pass

    def _build_header(self):
        for w in self._hdr_frame.winfo_children():
            w.destroy()
        for col_name, col_w, col_anchor in self.COL_SPECS:
            db_col = self.SORT_MAP.get(col_name)
            is_sorted = db_col and db_col == self._sort_col
            arrow = (" ▲" if self._sort_asc else " ▼") if is_sorted else ""
            cell = ctk.CTkFrame(self._hdr_frame, fg_color="transparent",
                                width=col_w, height=38)
            cell.pack(side="left", padx=0)
            cell.pack_propagate(False)
            if db_col:
                btn = ctk.CTkButton(cell, text=f"{col_name}{arrow}",
                              command=lambda c=db_col: self._sort_by(c),
                              fg_color="transparent", hover_color="#2E6DA4",
                              text_color="#FFFFFF",
                              font=ctk.CTkFont(size=12, weight="bold"),
                              anchor=col_anchor, height=38, corner_radius=0,
                              )
                btn.place(relx=0, rely=0, relwidth=1, relheight=1)
            else:
                ctk.CTkLabel(cell, text=col_name,
                             font=ctk.CTkFont(size=12, weight="bold"),
                             text_color="#FFFFFF", anchor=col_anchor,
                             ).place(relx=0, rely=0, relwidth=1, relheight=1)

    def _sort_by(self, db_col):
        if self._sort_col == db_col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = db_col
            self._sort_asc = True
        self._build_header()
        self._refresh()

    def _add_edit_btn(self, frame, row, accent):
        ctk.CTkButton(frame, text="✏️", width=36, height=26,
                      fg_color="#D6EAF8", hover_color="#AED6F1",
                      text_color=accent,
                      font=ctk.CTkFont(size=12, weight="bold"),
                      corner_radius=6,
                      command=lambda: self._open_edit(row),
                      ).pack(side="left", padx=(4, 6))

    def _open_edit(self, row):
        pass

    def _delete_record(self, row_id):
        _delete(self.TABLE_NAME, row_id)
        self._refresh()

    def _log_entry(self):
        pass

    def _export_csv(self):
        """Export respecting active filters."""
        q = self._search_var.get().strip()
        rows = _fetch(self.TABLE_NAME, self.SEARCH_COLS, q,
                      self._sort_col, self._sort_asc,
                      month=self._filter_month, year=self._filter_year,
                      dept_col=self.DEPT_FILTER_COL, dept_val=self._filter_dept)
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

    def _import_csv(self):
        """Override in subclass to handle tab-specific column mapping."""
        pass

    def _render_cells(self, frame, cells, open_detail_fn):
        for text, w, anch in cells:
            cell = ctk.CTkFrame(frame, fg_color="transparent", width=w, height=34)
            cell.pack(side="left", padx=0)
            cell.pack_propagate(False)
            cell.bind("<Button-1>", open_detail_fn)
            cell.configure(cursor="hand2")
            if anch == "center":
                lbl = ctk.CTkLabel(cell, text=text, font=ctk.CTkFont(size=12),
                                   text_color=TEXT_MAIN, anchor="center", cursor="hand2")
                lbl.place(relx=0.5, rely=0.5, anchor="center")
            else:
                lbl = ctk.CTkLabel(cell, text=text, font=ctk.CTkFont(size=12),
                                   text_color=TEXT_MAIN, anchor="w", cursor="hand2")
                lbl.place(x=8, rely=0.5, anchor="w")
            lbl.bind("<Button-1>", open_detail_fn)
        frame.bind("<Button-1>", open_detail_fn)
        frame.configure(cursor="hand2")


# ── Tab 1 — Incoming Memo ─────────────────────────────────────────────────────
class MemoTab(BaseTab):
    TABLE_NAME      = "memos"
    SEARCH_COLS     = ["date", "department", "subject", "receiver"]
    FORM_TITLE      = "New Incoming Memo"
    FORM_HINT       = "Log a memo received from another department."
    ACCENT          = TAB_COLORS["Incoming Memo"]["accent"]
    DARK            = TAB_COLORS["Incoming Memo"]["dark"]
    ICON            = "📥"
    PAST_POPUP_CLS  = PastMemoPopup
    DEPT_FILTER_COL = "department"
    COL_SPECS       = [
        ("#",           50,  "center"),
        ("Date & Time", 150, "w"),
        ("Department",  250, "w"),
        ("Subject",     300, "w"),
        ("Received By", 180, "w"),
        ("Action",       80, "center"),
    ]
    SORT_MAP = {
        "Date & Time": "date",
        "Department":  "department",
        "Subject":     "subject",
        "Received By": "receiver",
    }

    def _build_fields(self, parent):
        self._dept_var     = ctk.StringVar(value=DEPARTMENTS[0] if DEPARTMENTS else "")
        self._subject_var  = ctk.StringVar()
        self._receiver_var = ctk.StringVar()

        self._field_label(parent, "From Department")
        make_option(parent, self._dept_var, DEPARTMENTS,
                    self.ACCENT, self.DARK, width=290).pack(anchor="w", padx=20, pady=(2, 6))
        self._field_label(parent, "Subject / Memo Topic")
        make_entry(parent, self._subject_var,
                   "e.g. Monthly Budget Report", width=290).pack(anchor="w", padx=20, pady=(2, 6))
        self._field_label(parent, "Received By")
        make_entry(parent, self._receiver_var,
                   "e.g. Maria Santos", width=290).pack(anchor="w", padx=20, pady=(2, 8))

    def _log_entry(self):
        dept = self._dept_var.get()
        subj = self._subject_var.get().strip()
        recv = self._receiver_var.get().strip()
        if dept == DEPARTMENTS[0]:
            return messagebox.showwarning("Missing", "Please select a department.")
        if not subj:
            return messagebox.showwarning("Missing", "Please enter the subject.")
        if not recv:
            return messagebox.showwarning("Missing", "Please enter the receiver's name.")
        _insert("memos", ["date","department","subject","receiver"],
                (now_str(), dept, subj, recv))
        self._dept_var.set(DEPARTMENTS[0])
        self._subject_var.set("")
        self._receiver_var.set("")
        self._refresh()
        messagebox.showinfo("Logged", f"Memo from '{dept}' has been recorded.")

    def _save_past_entry(self, date_str, dept, subj, recv):
        _insert("memos", ["date","department","subject","receiver"],
                (date_str, dept, subj, recv))
        self._refresh()
        messagebox.showinfo("Logged", f"Past memo from '{dept}' ({date_str}) recorded.")

    def _import_csv(self):
        filetypes = [("Spreadsheet files", "*.csv *.xlsx"),
                     ("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
        path = filedialog.askopenfilename(filetypes=filetypes,
                                          title="Import Incoming Memo Records")
        if not path:
            return
        headers, result = read_import_file(path)
        if headers is None:
            messagebox.showerror("Import Failed", result)
            return
        headers_lower = [h.lower() for h in headers]

        def find_col(keywords):
            for kw in keywords:
                for orig, low in zip(headers, headers_lower):
                    if kw in low:
                        return orig
            return None

        col_date = find_col(["date", "time"])
        col_dept = find_col(["dept", "department", "from"])
        col_subj = find_col(["subject", "topic", "memo"])
        col_recv = find_col(["received", "receiver", "by", "name"])

        # Show what columns were detected — helps diagnose mismatches
        if not all([col_date, col_dept, col_subj, col_recv]):
            missing = []
            if not col_date: missing.append("Date")
            if not col_dept: missing.append("Department")
            if not col_subj: missing.append("Subject")
            if not col_recv: missing.append("Received By")
            ans = messagebox.askyesno("Column Detection",
                f"Could not auto-detect these columns: {', '.join(missing)}\n\n"
                f"Columns found in file:\n{chr(10).join(headers)}\n\n"
                f"Continue anyway? Rows with missing data will be skipped.")
            if not ans:
                return

        imported = skipped = 0
        for row in result:
            date = (row.get(col_date, "") if col_date else "").strip()
            dept = (row.get(col_dept, "") if col_dept else "").strip()
            subj = (row.get(col_subj, "") if col_subj else "").strip()
            recv = (row.get(col_recv, "") if col_recv else "").strip()
            parsed_date = parse_date_flexible(date) if date else None
            if not parsed_date:
                parsed_date = date
            # Skip only if the row is completely empty
            if not any([parsed_date, dept, subj, recv]):
                skipped += 1
                continue
            # Skip duplicates
            if parsed_date and dept and subj and _exists(
                    "memos", ["date","department","subject"],
                    (parsed_date, dept, subj)):
                skipped += 1
                continue
            _insert("memos", ["date","department","subject","receiver"],
                    (parsed_date or "", dept, subj, recv))
            imported += 1
        self._refresh()
        parts = [f"✔ {imported} record{'s' if imported!=1 else ''} imported."]
        if skipped:
            parts.append(f"⚠ {skipped} row(s) skipped (duplicates or missing fields).")
        messagebox.showinfo("Import Complete", "\n".join(parts))

    def _open_edit(self, row):
        EditMemoPopup(self, row, on_save=self._save_edit,
                      on_delete=self._delete_record)

    def _save_edit(self, row_id, date, dept, subj, recv):
        _update("memos", ["date","department","subject","receiver"],
                (date, dept, subj, recv), row_id)
        self._refresh()
        messagebox.showinfo("Saved", f"Memo #{row_id} updated.")

    def _render_row(self, frame, row):
        memo_id, date, dept, subj, recv = row
        fields = [("Date & Time", date), ("Department", dept),
                  ("Subject", subj), ("Received By", recv)]

        def open_detail(e=None):
            RecordDetailPopup(self, fields, self.ACCENT)

        self._render_cells(frame, [
            (str(memo_id),  50, "center"),
            (date,         150, "w"),
            (dept,         250, "w"),
            (subj,         300, "w"),
            (recv,         180, "w"),
        ], open_detail)
        self._add_edit_btn(frame, row, self.ACCENT)


# ── Tab 2 — Outgoing Memo ─────────────────────────────────────────────────────
class OutgoingTab(BaseTab):
    TABLE_NAME      = "outgoing"
    SEARCH_COLS     = ["date", "subject", "recipient"]
    FORM_TITLE      = "New Outgoing Memo"
    FORM_HINT       = "Record a memo sent out from your office."
    ACCENT          = TAB_COLORS["Outgoing Memo"]["accent"]
    DARK            = TAB_COLORS["Outgoing Memo"]["dark"]
    ICON            = "📤"
    PAST_POPUP_CLS  = PastOutgoingPopup
    DEPT_FILTER_COL = "recipient"
    COL_SPECS       = [
        ("#",             50, "center"),
        ("Date",         130, "w"),
        ("Subject",      300, "w"),
        ("Sent To Dept.", 300, "w"),
        ("Remarks",      200, "w"),
        ("Action",        80, "center"),
    ]
    SORT_MAP = {
        "Date":          "date",
        "Subject":       "subject",
        "Sent To Dept.": "recipient",
    }

    def _build_fields(self, parent):
        self._subject_var   = ctk.StringVar()
        self._recipient_var = ctk.StringVar(value=DEPARTMENTS[0] if DEPARTMENTS else "")
        self._remarks_var   = ctk.StringVar()

        self._field_label(parent, "Subject / Title")
        make_entry(parent, self._subject_var,
                   "e.g. Q2 Report", width=290).pack(anchor="w", padx=20, pady=(2, 6))
        self._field_label(parent, "Department Sent To")
        make_option(parent, self._recipient_var, DEPARTMENTS,
                    self.ACCENT, self.DARK, width=290).pack(anchor="w", padx=20, pady=(2, 6))
        self._field_label(parent, "Remarks (optional)")
        make_entry(parent, self._remarks_var,
                   "e.g. Via courier", width=290).pack(anchor="w", padx=20, pady=(2, 8))

    def _log_entry(self):
        subj    = self._subject_var.get().strip()
        recip   = self._recipient_var.get()
        remarks = self._remarks_var.get().strip()
        if not subj:
            return messagebox.showwarning("Missing", "Please enter the subject.")
        if recip == DEPARTMENTS[0]:
            return messagebox.showwarning("Missing", "Please select the destination department.")
        _insert("outgoing", ["date","subject","recipient","remarks"],
                (now_str(), subj, recip, remarks))
        self._recipient_var.set(DEPARTMENTS[0])
        self._subject_var.set("")
        self._remarks_var.set("")
        self._refresh()
        messagebox.showinfo("Logged", "Outgoing memo recorded successfully.")

    def _save_past_entry(self, date_str, subj, recip, remarks):
        _insert("outgoing", ["date","subject","recipient","remarks"],
                (date_str, subj, recip, remarks))
        self._refresh()
        messagebox.showinfo("Logged", f"Past outgoing memo ({date_str}) recorded.")

    def _import_csv(self):
        filetypes = [("Spreadsheet files", "*.csv *.xlsx"),
                     ("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
        path = filedialog.askopenfilename(filetypes=filetypes,
                                          title="Import Outgoing Memo Records")
        if not path:
            return
        headers, result = read_import_file(path)
        if headers is None:
            messagebox.showerror("Import Failed", result)
            return
        headers_lower = [h.lower() for h in headers]

        def find_col(keywords):
            for kw in keywords:
                for orig, low in zip(headers, headers_lower):
                    if kw in low:
                        return orig
            return None

        col_date    = find_col(["date", "time"])
        col_subj    = find_col(["subject", "topic", "memo"])
        col_recip   = find_col(["sent", "recipient", "dept", "department", "to"])
        col_remarks = find_col(["remark", "note", "comment"])

        if not all([col_date, col_subj, col_recip]):
            missing = []
            if not col_date:  missing.append("Date")
            if not col_subj:  missing.append("Subject")
            if not col_recip: missing.append("Department Sent To")
            ans = messagebox.askyesno("Column Detection",
                f"Could not auto-detect these columns: {', '.join(missing)}\n\n"
                f"Columns found in file:\n{chr(10).join(headers)}\n\n"
                f"Continue anyway? Rows with missing data will be skipped.")
            if not ans:
                return

        imported = skipped = 0
        for row in result:
            date    = (row.get(col_date,    "") if col_date    else "").strip()
            subj    = (row.get(col_subj,    "") if col_subj    else "").strip()
            recip   = (row.get(col_recip,   "") if col_recip   else "").strip()
            remarks = (row.get(col_remarks, "") if col_remarks else "").strip()
            parsed_date = parse_date_flexible(date) if date else None
            if not parsed_date:
                parsed_date = date
            # Skip only if the row is completely empty
            if not any([parsed_date, subj, recip]):
                skipped += 1
                continue
            # Skip duplicates
            if parsed_date and subj and recip and _exists(
                    "outgoing", ["date","subject","recipient"],
                    (parsed_date, subj, recip)):
                skipped += 1
                continue
            _insert("outgoing", ["date","subject","recipient","remarks"],
                    (parsed_date or "", subj, recip, remarks))
            imported += 1
        self._refresh()
        parts = [f"✔ {imported} record{'s' if imported!=1 else ''} imported."]
        if skipped:
            parts.append(f"⚠ {skipped} row(s) skipped (duplicates or missing fields).")
        messagebox.showinfo("Import Complete", "\n".join(parts))

    def _open_edit(self, row):
        EditOutgoingPopup(self, row, on_save=self._save_edit,
                          on_delete=self._delete_record)

    def _save_edit(self, row_id, date, subj, recip, remarks):
        _update("outgoing", ["date","subject","recipient","remarks"],
                (date, subj, recip, remarks), row_id)
        self._refresh()
        messagebox.showinfo("Saved", f"Outgoing memo #{row_id} updated.")

    def _render_row(self, frame, row):
        rid, date, subj, recip, remarks = row
        fields = [("Date", date), ("Subject", subj),
                  ("Sent To Department", recip), ("Remarks", remarks)]

        def open_detail(e=None):
            RecordDetailPopup(self, fields, self.ACCENT)

        self._render_cells(frame, [
            (str(rid),      50, "center"),
            (date,         130, "w"),
            (subj,         300, "w"),
            (recip,        300, "w"),
            (remarks or "", 200, "w"),
        ], open_detail)
        self._add_edit_btn(frame, row, self.ACCENT)



# ── Tab 3 — Endorsements ─────────────────────────────────────────────────────
class EndorsementTab(BaseTab):
    TABLE_NAME      = "endorsements"
    SEARCH_COLS     = ["date", "name", "subject", "endorsed_to"]
    FORM_TITLE      = "New Endorsement"
    FORM_HINT       = "Log an endorsement letter."
    ACCENT          = TAB_COLORS["Endorsements"]["accent"]
    DARK            = TAB_COLORS["Endorsements"]["dark"]
    ICON            = "📋"
    PAST_POPUP_CLS  = PastEndorsementPopup
    PAST_LABEL      = "Endorsement"
    DEPT_FILTER_COL = None
    COMPACT         = True
    COL_SPECS       = [
        ("#",            50, "center"),
        ("Date",        120, "w"),
        ("Name",        180, "w"),
        ("Subject",     220, "w"),
        ("Endorsed To", 160, "w"),
        ("Office",      160, "w"),
        ("Remarks",     110, "w"),
        ("Action",       80, "center"),
    ]
    SORT_MAP = {
        "Date":        "date",
        "Name":        "name",
        "Subject":     "subject",
        "Endorsed To": "endorsed_to",
        "Office":      "office",
    }

    def _build_fields(self, parent):
        self._name_var        = ctk.StringVar()
        self._subject_var     = ctk.StringVar()
        self._endorsed_to_var = ctk.StringVar()
        self._office_var      = ctk.StringVar()
        self._remarks_var     = ctk.StringVar()

        for label, var, placeholder in [
            ("Name (Who is Endorsed)", self._name_var,        "e.g. Juan dela Cruz"),
            ("Subject",                self._subject_var,     "e.g. Employment Endorsement"),
            ("Endorsed To",            self._endorsed_to_var, "e.g. City Mayor"),
            ("Office",                 self._office_var,      "e.g. Office of the City Mayor"),
            ("Remarks (optional)",     self._remarks_var,     "e.g. For signature"),
        ]:
            make_label(parent, label, size=11, weight="bold",
                       color=TEXT_SUB).pack(anchor="w", padx=20, pady=(3, 0))
            ctk.CTkEntry(parent, textvariable=var, placeholder_text=placeholder,
                         fg_color=CARD, border_color=BORDER, text_color=TEXT_MAIN,
                         font=ctk.CTkFont(size=12), width=290, height=30,
                         ).pack(anchor="w", padx=20, pady=(2, 0))

    def _log_entry(self):
        name        = self._name_var.get().strip()
        subj        = self._subject_var.get().strip()
        endorsed_to = self._endorsed_to_var.get().strip()
        office      = self._office_var.get().strip()
        remarks     = self._remarks_var.get().strip()
        if not name:
            return messagebox.showwarning("Missing", "Please enter the name.")
        if not subj:
            return messagebox.showwarning("Missing", "Please enter the subject.")
        if not endorsed_to:
            return messagebox.showwarning("Missing", "Please enter who it is endorsed to.")
        _insert("endorsements",
                ["date","name","subject","endorsed_to","office","remarks"],
                (now_str(), name, subj, endorsed_to, office, remarks))
        self._name_var.set("")
        self._subject_var.set("")
        self._endorsed_to_var.set("")
        self._office_var.set("")
        self._remarks_var.set("")
        self._refresh()
        messagebox.showinfo("Logged", f"Endorsement for '{name}' has been recorded.")

    def _save_past_entry(self, date_str, name, subj, endorsed_to, office, remarks):
        _insert("endorsements",
                ["date","name","subject","endorsed_to","office","remarks"],
                (date_str, name, subj, endorsed_to, office, remarks))
        self._refresh()
        messagebox.showinfo("Logged",
                            f"Past endorsement for '{name}' ({date_str}) recorded.")

    def _open_edit(self, row):
        EditEndorsementPopup(self, row,
                             on_save=self._save_edit,
                             on_delete=self._delete_record)

    def _save_edit(self, row_id, date, name, subj, endorsed_to, office, remarks):
        _update("endorsements",
                ["date","name","subject","endorsed_to","office","remarks"],
                (date, name, subj, endorsed_to, office, remarks), row_id)
        self._refresh()
        messagebox.showinfo("Saved", f"Endorsement #{row_id} updated.")

    def _import_csv(self):
        filetypes = [("Spreadsheet files", "*.csv *.xlsx"),
                     ("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
        path = filedialog.askopenfilename(filetypes=filetypes,
                                          title="Import Endorsement Records")
        if not path:
            return
        headers, result = read_import_file(path)
        if headers is None:
            messagebox.showerror("Import Failed", result)
            return
        headers_lower = [h.lower() for h in headers]

        def find_col(keywords):
            for kw in keywords:
                for orig, low in zip(headers, headers_lower):
                    if kw in low:
                        return orig
            return None

        col_date = find_col(["date"])
        col_name = find_col(["name", "endorsed", "who"])
        col_subj = find_col(["subject", "topic"])
        col_to   = find_col(["to", "addressee", "endorsed to"])
        col_rem  = find_col(["remark", "note", "comment"])

        imported = skipped = 0
        for row in result:
            date        = (row.get(col_date, "") if col_date else "").strip()
            name        = (row.get(col_name, "") if col_name else "").strip()
            subj        = (row.get(col_subj, "") if col_subj else "").strip()
            endorsed_to = (row.get(col_to,   "") if col_to   else "").strip()
            remarks     = (row.get(col_rem,  "") if col_rem  else "").strip()
            parsed_date = parse_date_flexible(date) if date else None
            if not parsed_date:
                parsed_date = date
            if not any([parsed_date, name, subj, endorsed_to]):
                skipped += 1
                continue
            if parsed_date and name and subj and _exists(
                    "endorsements", ["date","name","subject"],
                    (parsed_date, name, subj)):
                skipped += 1
                continue
            _insert("endorsements",
                    ["date","name","subject","endorsed_to","remarks"],
                    (parsed_date or "", name, subj, endorsed_to, remarks))
            imported += 1
        self._refresh()
        parts = [f"✔ {imported} record{'s' if imported!=1 else ''} imported."]
        if skipped:
            parts.append(f"⚠ {skipped} row(s) skipped (duplicates or empty rows).")
        messagebox.showinfo("Import Complete", "\n".join(parts))

    def _render_row(self, frame, row):
        if len(row) == 7:
            rid, date, name, subj, endorsed_to, office, remarks = row
        else:
            rid, date, name, subj, endorsed_to, remarks = row
            office = ""
        fields = [("Date", date), ("Name", name), ("Subject", subj),
                  ("Endorsed To", endorsed_to), ("Office", office),
                  ("Remarks", remarks)]

        def open_detail(e=None):
            RecordDetailPopup(self, fields, self.ACCENT)

        self._render_cells(frame, [
            (str(rid),       50, "center"),
            (date,          120, "w"),
            (name,          180, "w"),
            (subj,          220, "w"),
            (endorsed_to,   160, "w"),
            (office or "",  160, "w"),
            (remarks or "", 110, "w"),
        ], open_detail)
        self._add_edit_btn(frame, row, self.ACCENT)

# ── Tab 3 — Department Editor ─────────────────────────────────────────────────
class DepartmentsTab(ctk.CTkFrame):
    ACCENT = TAB_COLORS["Departments"]["accent"]
    DARK   = TAB_COLORS["Departments"]["dark"]

    def __init__(self, parent, on_departments_changed):
        super().__init__(parent, fg_color=SURFACE, corner_radius=0)
        self._on_changed = on_departments_changed
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._build_form()
        self._build_list()
        self._refresh()

    def _build_form(self):
        panel = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0,
                             width=310, border_width=0)
        panel.grid(row=0, column=0, sticky="nsew")
        panel.pack_propagate(False)
        ctk.CTkFrame(panel, fg_color=self.ACCENT,
                     corner_radius=0, height=6).pack(fill="x")
        make_label(panel, "Department Editor", size=16, weight="bold",
                   color=self.ACCENT).pack(anchor="w", padx=28, pady=(20, 2))
        make_label(panel, "Add or remove departments from the list.",
                   size=11, weight="normal", color=TEXT_SUB
                   ).pack(anchor="w", padx=28, pady=(0, 20))
        make_label(panel, "New Department Name", size=11, weight="bold",
                   color=TEXT_SUB).pack(anchor="w", padx=28)
        self._new_dept_var = ctk.StringVar()
        make_entry(panel, self._new_dept_var,
                   "e.g. City Health Office").pack(anchor="w", padx=28, pady=(4, 14))
        make_button(panel, "  ＋  Add Department", self._add_dept,
                    self.ACCENT, self.DARK, height=42).pack(anchor="w", padx=28)

    def _build_list(self):
        panel = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        panel.grid(row=0, column=1, sticky="nsew")
        top = ctk.CTkFrame(panel, fg_color=SURFACE, corner_radius=0)
        top.pack(fill="x", padx=24, pady=(18, 0))
        make_label(top, "Current Departments", size=16, weight="bold",
                   color=self.ACCENT).pack(side="left")
        self._count_lbl = make_label(top, "", size=11, weight="normal", color=TEXT_SUB)
        self._count_lbl.pack(side="left", padx=10)
        tcard = ctk.CTkFrame(panel, fg_color=CARD, corner_radius=10,
                             border_width=1, border_color=BORDER)
        tcard.pack(fill="both", expand=True, padx=24, pady=12)
        hdr = ctk.CTkFrame(tcard, fg_color=self.ACCENT, corner_radius=0, height=38)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        for text, w, anch in [("#", 42, "center"), ("Department Name", 500, "w"),
                               ("Action", 80, "center")]:
            cell = ctk.CTkFrame(hdr, fg_color="transparent", width=w, height=38)
            cell.pack(side="left", padx=(4 if text == "#" else 2, 0))
            cell.pack_propagate(False)
            ctk.CTkLabel(cell, text=text,
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color="#FFFFFF", anchor=anch,
                         ).place(relx=0.05 if anch == "w" else 0.5,
                                 rely=0.5, anchor="w" if anch == "w" else "center")
        self._rows_frame = ctk.CTkScrollableFrame(
            tcard, fg_color=CARD, corner_radius=0, scrollbar_button_color=BORDER)
        self._rows_frame.pack(fill="both", expand=True)

    def _refresh(self):
        for w in self._rows_frame.winfo_children():
            w.destroy()
        rows = fetch_department_rows()
        n = len(rows)
        self._count_lbl.configure(text=f"({n} department{'s' if n != 1 else ''})")
        if not rows:
            make_label(self._rows_frame, "No departments added yet.",
                       size=13, weight="normal", color=TEXT_SUB).pack(pady=40)
            return
        for idx, (dept_id, name) in enumerate(rows):
            bg = CARD if idx % 2 == 0 else ROW_ALT
            rf = ctk.CTkFrame(self._rows_frame, fg_color=bg,
                              corner_radius=0, height=38)
            rf.pack(fill="x")
            rf.pack_propagate(False)
            id_cell = ctk.CTkFrame(rf, fg_color="transparent", width=42, height=36)
            id_cell.pack(side="left", padx=(4, 2))
            id_cell.pack_propagate(False)
            ctk.CTkLabel(id_cell, text=str(dept_id), font=ctk.CTkFont(size=12),
                         text_color=TEXT_MAIN, anchor="center"
                         ).place(relx=0.5, rely=0.5, anchor="center")
            name_var = ctk.StringVar(value=name)
            name_entry = ctk.CTkEntry(rf, textvariable=name_var,
                                      fg_color="transparent", border_color=BORDER,
                                      border_width=0, text_color=TEXT_MAIN,
                                      font=ctk.CTkFont(size=12), width=480, height=30)
            name_entry.pack(side="left", padx=(2, 4))
            name_entry.bind("<FocusIn>",
                            lambda e, w=name_entry: w.configure(border_width=1))
            name_entry.bind("<FocusOut>",
                            lambda e, did=dept_id, var=name_var, orig=name:
                            self._inline_save(did, var.get().strip(), orig))
            name_entry.bind("<Return>",
                            lambda e, did=dept_id, var=name_var, orig=name:
                            self._inline_save(did, var.get().strip(), orig))
            ctk.CTkButton(rf, text="🗑", width=36, height=26,
                          fg_color="#FADBD8", hover_color="#F1948A",
                          text_color=DANGER,
                          font=ctk.CTkFont(size=12, weight="bold"),
                          corner_radius=6,
                          command=lambda did=dept_id, n=name: self._delete_dept(did, n),
                          ).pack(side="left", padx=(4, 6))

    def _inline_save(self, dept_id, new_name, original):
        if new_name == original or not new_name:
            return
        if not update_department(dept_id, new_name):
            messagebox.showwarning("Duplicate", f"'{new_name}' already exists.")
        else:
            load_departments()
            self._on_changed()
            self._refresh()

    def _add_dept(self):
        name = self._new_dept_var.get().strip()
        if not name:
            return messagebox.showwarning("Missing", "Please enter a department name.")
        if not insert_department(name):
            return messagebox.showwarning("Duplicate", f"'{name}' already exists.")
        self._new_dept_var.set("")
        load_departments()
        self._on_changed()
        self._refresh()
        messagebox.showinfo("Added", f"'{name}' has been added.")

    def _delete_dept(self, dept_id, name):
        if messagebox.askyesno("Confirm Delete",
                               f"Remove '{name}' from the department list?"):
            delete_department(dept_id)
            load_departments()
            self._on_changed()
            self._refresh()


# ── Summary Dashboard ─────────────────────────────────────────────────────────
class SummaryBar(ctk.CTkFrame):
    """Small dashboard strip shown at the top of the content area."""
    def __init__(self, parent):
        super().__init__(parent, fg_color=CARD, corner_radius=0,
                         border_width=0, height=54)
        self.pack_propagate(False)
        self._build()
        self._update()

    def _build(self):
        month_lbl = make_label(self, f"This Month ({datetime.now().strftime('%B %Y')}):",
                               size=11, weight="bold", color=TEXT_SUB)
        month_lbl.pack(side="left", padx=(20, 10))

        self._inc_lbl = ctk.CTkLabel(self, text="",
                                     font=ctk.CTkFont(size=13, weight="bold"),
                                     text_color=TAB_COLORS["Incoming Memo"]["accent"])
        self._inc_lbl.pack(side="left", padx=(0, 16))

        self._out_lbl = ctk.CTkLabel(self, text="",
                                     font=ctk.CTkFont(size=13, weight="bold"),
                                     text_color=TAB_COLORS["Outgoing Memo"]["accent"])
        self._out_lbl.pack(side="left")

    def _update(self):
        inc, out = get_monthly_summary()
        self._inc_lbl.configure(text=f"📥 Incoming: {inc}")
        self._out_lbl.configure(text=f"📤 Outgoing: {out}")
        self.after(30000, self._update)  # refresh every 30 seconds


# ── Main App Window ───────────────────────────────────────────────────────────
class DocTrackerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Office Memo Tracker")
        self.geometry("1440x780")
        self.minsize(1200, 650)
        self.configure(fg_color=SURFACE)
        init_db()
        load_departments()
        self._active_tab  = None
        self._tab_widgets = {}
        self._tab_panels  = {}
        self._build_ui()
        self._switch_tab("Incoming Memo")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        if messagebox.askyesno("Exit", "Are you sure you want to close the app?"):
            self.destroy()

    def _build_ui(self):
        # Header
        hdr = ctk.CTkFrame(self, fg_color=ACCENT, corner_radius=0, height=62)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="  🗂  Office Memo Tracker",
                     font=ctk.CTkFont(family="Georgia", size=21, weight="bold"),
                     text_color="#FFFFFF", anchor="w",
                     ).pack(side="left", padx=24)
        self._clock = ctk.CTkLabel(hdr, text="",
                                   font=ctk.CTkFont(size=13), text_color="#A9CCE3")
        self._clock.pack(side="right", padx=24)
        self._tick()

        # Tab bar
        tab_bar = ctk.CTkFrame(self, fg_color="#D6E4F0", corner_radius=0, height=46)
        tab_bar.pack(fill="x")
        tab_bar.pack_propagate(False)
        for name in ["Incoming Memo", "Outgoing Memo", "Endorsements", "Departments"]:
            info = TAB_COLORS[name]
            btn = ctk.CTkButton(
                tab_bar, text=f"  {info['icon']}  {name}",
                command=lambda n=name: self._switch_tab(n),
                fg_color=TAB_IDLE_BG, hover_color=TAB_IDLE_BG,
                text_color=TAB_IDLE_TEXT,
                font=ctk.CTkFont(size=13, weight="bold"),
                height=46, width=200, corner_radius=0,
            )
            btn.pack(side="left")
            self._tab_widgets[name] = btn

        # Summary bar
        self._summary_bar = SummaryBar(self)
        self._summary_bar.pack(fill="x")
        ctk.CTkFrame(self, fg_color=BORDER, height=1).pack(fill="x")

        # Content area
        content = ctk.CTkFrame(self, fg_color=SURFACE, corner_radius=0)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        memo_panel    = MemoTab(content)
        out_panel     = OutgoingTab(content)
        endorse_panel = EndorsementTab(content)
        dept_panel    = DepartmentsTab(content,
                                    on_departments_changed=self._on_depts_changed)

        memo_panel.grid(row=0, column=0, sticky="nsew")
        out_panel.grid(row=0, column=0, sticky="nsew")
        endorse_panel.grid(row=0, column=0, sticky="nsew")
        dept_panel.grid(row=0, column=0, sticky="nsew")

        self._tab_panels["Incoming Memo"] = memo_panel
        self._tab_panels["Outgoing Memo"] = out_panel
        self._tab_panels["Endorsements"]  = endorse_panel
        self._tab_panels["Departments"]   = dept_panel

    def _on_depts_changed(self):
        for tab_name in ("Incoming Memo", "Outgoing Memo"):
            panel = self._tab_panels[tab_name]
            self._update_option_menus(panel._form_panel)
        for tab_name in ("Incoming Memo", "Outgoing Memo"):
            panel = self._tab_panels[tab_name]
            if hasattr(panel, "_dept_filter_menu") and panel._dept_filter_menu:
                current = panel._dept_filter_var.get()
                new_vals = ["All Departments"] + DEPARTMENTS[1:]
                panel._dept_filter_menu.configure(values=new_vals)
                if current not in new_vals:
                    panel._dept_filter_var.set("All Departments")
                    panel._filter_dept = "All Departments"

    def _update_option_menus(self, parent):
        for widget in parent.winfo_children():
            if isinstance(widget, ctk.CTkOptionMenu):
                current = widget.get()
                widget.configure(values=DEPARTMENTS)
                widget.set(current if current in DEPARTMENTS else DEPARTMENTS[0])
            else:
                try:
                    self._update_option_menus(widget)
                except Exception:
                    pass

    def _switch_tab(self, name):
        if self._active_tab == name:
            return
        self._active_tab = name
        info = TAB_COLORS[name]
        for tab_name, btn in self._tab_widgets.items():
            if tab_name == name:
                btn.configure(fg_color=info["accent"], hover_color=info["dark"],
                              text_color="#FFFFFF")
            else:
                btn.configure(fg_color=TAB_IDLE_BG, hover_color=TAB_IDLE_BG,
                              text_color=TAB_IDLE_TEXT)
        self._tab_panels[name].tkraise()
        # Refresh summary whenever switching tabs
        self._summary_bar._update()

    def _tick(self):
        self._clock.configure(
            text=datetime.now().strftime("  %A, %B %d %Y   %H:%M:%S"))
        self.after(1000, self._tick)


if __name__ == "__main__":
    app = DocTrackerApp()
    app.mainloop()
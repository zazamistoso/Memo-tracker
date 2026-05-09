# 🗂 Office Memo Tracker

A desktop application for logging, managing, and tracking office memos, outgoing documents, and endorsement letters. Built for city hall use with offline functionality and a local database.

---

## 📋 Features

- **Three document tabs** — Incoming Memo, Outgoing Memo, and Endorsements
- **Log today's records** with auto-filled date and time
- **Log past records** with manual date entry (accepts MM/DD/YYYY, YYYY-MM-DD, and other formats)
- **Edit and delete** any record via the ✏️ button
- **Click any row** to view the full record details in a popup
- **Sort records** by clicking any column header (ascending/descending)
- **Filter records** by Month, Year, and Department
- **Search** across all fields in real time
- **Export to CSV** — respects active filters so you export exactly what you see
- **Import from CSV or Excel (.xlsx)** — flexible column detection, skips duplicates, allows partial records
- **Department editor** — add, rename, and delete departments without touching the code
- **Monthly summary** — shows this month's incoming and outgoing memo count at a glance
- **Pagination** — displays 50 records per page for performance
- **Close confirmation** — prevents accidental data loss

---

## 🖥 Requirements

- Windows 10 or later
- Python 3.10 or later *(only needed if running from source)*

### Python Dependencies *(source only)*
```
customtkinter
openpyxl
```

Install with:
```bash
pip install -r requirements.txt
```

---

## 🚀 Getting Started

### Option A — Running the .exe *(recommended for office use)*

1. Place `Office Memo Tracker.exe` anywhere (Desktop recommended)
2. Double-click to run
3. The database is automatically created at:
   ```
   C:\Users\[your username]\Documents\OfficeMemoTracker\doc_tracker.db
   ```
4. That's it — no installation needed

### Option B — Running from source

```bash
python document_tracker.py
```

---

## 📁 Database Location

The app stores all data in:
```
C:\Users\[your username]\Documents\OfficeMemoTracker\doc_tracker.db
```

- This folder is created automatically on first run
- The `.exe` can be placed anywhere — the data is always saved in the same location
- **Do not delete `doc_tracker.db`** — this file contains all your records
- Backing up this file is enough to back up all your data

---

## 📥 Incoming Memo Tab

Log memos received from other departments.

| Field | Description |
|---|---|
| Date & Time | Auto-filled with current date and time |
| From Department | Select from the department list |
| Subject / Memo Topic | Brief description of the memo |
| Received By | Name of the staff who received it |

---

## 📤 Outgoing Memo Tab

Log memos and documents sent out from your office.

| Field | Description |
|---|---|
| Date & Time | Auto-filled with current date and time |
| Subject / Title | Brief description of the document |
| Department Sent To | Select from the department list |
| Remarks | Optional notes (e.g. Via courier) |

---

## 📋 Endorsements Tab

Log endorsement letters issued by the office.

| Field | Description |
|---|---|
| Date & Time | Auto-filled with current date and time |
| Name | Name of the person being endorsed |
| Subject | Purpose of the endorsement |
| Endorsed To | Person the letter is addressed to |
| Office | Office or institution of the addressee |
| Remarks | Optional notes |

---

## 🕓 Logging Past Records

Each tab has a **Log Past [Type]** button for entering older records.

1. Click **🕓 Log Past Memo / Log Past Endorsement**
2. Enter the original date — accepted formats:
   - `07/14/2025` *(MM/DD/YYYY — recommended)*
   - `2025-07-14` *(YYYY-MM-DD)*
   - `July 14, 2025`
   - `Jul 14, 2025`
3. Fill in the remaining fields
4. Click **Save Past Entry**

---

## 📊 Filtering Records

The filter bar above the records table lets you narrow down what you see:

- **Month** — filter by a specific month
- **Year** — filter by a specific year
- **Department** — filter by sender or recipient department *(Incoming and Outgoing only)*
- **Reset** — clears all active filters

> 💡 The **Export to CSV** button exports only the records currently shown — so filters apply to exports too.

---

## ↓ Exporting Records

1. Apply any filters you want (optional)
2. Click **↓ Export to CSV**
3. Choose where to save the file
4. Open in Excel or share with others

Exported CSV columns per tab:

| Tab | Columns |
|---|---|
| Incoming Memo | Date & Time, Department, Subject, Received By |
| Outgoing Memo | Date, Subject, Sent To Dept., Remarks |
| Endorsements | Date, Name, Subject, Endorsed To, Office, Remarks |

---

## ↑ Importing Records

Import records from a CSV or Excel file exported from another machine or prepared manually.

1. Click **↑ Import from CSV**
2. Select a `.csv` or `.xlsx` file
3. The app will automatically detect which column is which based on the column names
4. A summary will show how many records were imported and how many were skipped

**Rules:**
- Rows that are completely empty are skipped
- Rows with missing required fields are still imported with blank values — you can edit them later
- Duplicate records (same date + department + subject) are automatically skipped
- Extra columns the app doesn't recognize are ignored

**Accepted date formats in imported files:**
- `07/14/2025`, `2025-07-14`, `July 14, 2025`, `Jul 14, 2025`, `07-14-2025`

---

## 🏢 Managing Departments

The **Departments** tab lets you manage the department list without editing any code.

- **Add** — type a department name and click ＋ Add Department
- **Rename** — click on any department name in the list, edit it, then press Enter or click away
- **Delete** — click the 🗑 button next to any department

Changes immediately reflect in all department dropdowns across all tabs.

---

## 🔄 Sharing Data Between Machines

Since the app works offline with a local database, use the export/import workflow to sync records between machines:

1. **Machine A** exports records to CSV
2. Send the CSV file to **Machine B** (via email, USB, shared folder, etc.)
3. **Machine B** imports the CSV — duplicates are automatically skipped
4. Re-import the same file anytime — it won't create duplicates

---

## 🏗 Building the .exe from Source

```bash
pip install pyinstaller
pip install openpyxl
pip install customtkinter

pyinstaller --onefile --windowed --name "Office Memo Tracker" --collect-all customtkinter document_tracker.py
```

The `.exe` will be inside the `dist` folder after the build completes.

---

## 🛠 Troubleshooting

**App doesn't open after double-clicking the .exe**
- Try right-clicking → Run as Administrator
- Make sure your antivirus isn't blocking it (PyInstaller .exe files sometimes trigger false positives — add an exception)

**Import says "0 records imported"**
- Check that your column names contain keywords the app recognizes (date, department, subject, name, etc.)
- Make sure the file isn't completely empty or has only blank rows at the top

**Records not showing after import**
- Check if a filter is active — click **Reset** in the filter bar to clear all filters

**App opens but looks wrong / missing fonts**
- Make sure you built with `--collect-all customtkinter` in the PyInstaller command

---

## 📂 File Structure

```
Office Memo Tracker.exe       ← The app (can be on Desktop)

Documents/
└── OfficeMemoTracker/
    └── doc_tracker.db        ← All your data (back this up!)
```

---

*Built for the Information and Communication Technology Office — City Hall*

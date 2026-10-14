# 📜 User Prompt History & Long-Term Memory Log

This file is a permanent record of all user prompts, instructions, and actions taken in this project (**Personal Expense Tracker**). 
Even if you return after a month or a year, the AI assistant will read this file and know exactly what you requested and what was built.

---

## 📌 Log Entries

### Entry #1: Initial Project Creation & Setup
- **Date / Time:** 2026-10-14 (Initial Project Setup)
- **User Request / Prompt:**
  > Build a Personal Expense & Budget Tracker command-line application backed by SQLite with data validation, analytics, automated unit tests, and GitHub Actions CI/CD.
- **Actions Taken:**
  - Created `expense_tracker.py` with SQLite database (`expenses.db`), CRUD operations, category statistics, and terminal ASCII visual spending charts.
  - Created `test_expense_tracker.py` containing 8 isolated pytest unit tests using SQLite fixtures.
  - Created `.github/workflows/ci.yml` for automated CI/CD pipeline on push/pull requests.
  - Created `requirements.txt` and `README.md`.
  - Initialized git repository and pushed initial commit to GitHub.

---

### Entry #2: Application Demo Execution
- **Date / Time:** 2026-10-14 17:11
- **User Action:**
  > Executed `python3 expense_tracker.py`
- **Result / Output:**
  - Added sample demo expenses (Internet Bill, Coffee, Metro Pass, Groceries).
  - Displayed formatted transaction table (Total $133.25).
  - Rendered category breakdown (Food 43.7%, Bills 41.3%, Transport 15.0%).

---

### Entry #3: Long-Term Prompt Memory Setup
- **Date / Time:** 2026-10-14 17:14 (Local time)
- **User Request / Prompt:**
  > *"save in file what promt i give to you means next time after year or month i ask you about that you will know that"*
- **Actions Taken:**
  - Created `PROMPT_HISTORY.md` to permanently store all historical and future prompts with timestamps and actions taken.
  - Created `GEMINI.md` (AI Project Memory & Instructions) at workspace root so that whenever any session opens (even after months or years), the assistant automatically reads this log and updates it with every new prompt.

---

### Entry #4: Excel Document Integration (Daily Entry & Weekly/Monthly/Yearly Calculations)
- **Date / Time:** 2026-10-14 17:20 (Local time)
- **User Request / Prompt:**
  > *"i want some changes in personam_expense_tracker project create excel document where i can also inter my daily expence and in new tab calculate total weekly,monthly,yrarly expensece"*
- **Actions Taken:**
  - Added `openpyxl==3.1.5` to `requirements.txt` for Excel (.xlsx) file generation.
  - Implemented `export_to_excel()` in `expense_tracker.py` with two dedicated tabs:
    1. **Tab 1: `Daily Expenses`**: Pre-fills existing records, provides 200 formatted entry rows with Category dropdown validation (`Food`, `Transport`, `Bills`, etc.), and automatic helper formulas for ISO week number, month key (`YYYY-MM`), and year.
    2. **Tab 2: `Summary & Analytics`**: Dynamic dashboard featuring real-time KPI overview cards, a 52-week Weekly Expenses table (`=SUMIF(...)`), a 12-month Monthly Expenses table (`=SUMIF(...)`), a multi-year Yearly Expenses table (`=SUMIF(...)`), and a Category spending breakdown with percentage shares.
  - Implemented `import_from_excel()` allowing bidirectional synchronization: newly entered rows in Excel are imported back into SQLite with `python3 expense_tracker.py --import-excel`.
  - Added CLI flags `--export-excel` and `--import-excel`, and updated `run_demo()` to automatically generate/update `expense_tracker.xlsx`.
  - Created new unit tests in `test_expense_tracker.py` (`test_export_to_excel` and `test_import_from_excel`), bringing total passing test suite to 10 tests (100% pass rate).
  - Updated `README.md` with complete documentation on using the Excel document and synchronization commands.


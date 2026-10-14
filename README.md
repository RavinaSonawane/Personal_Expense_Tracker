# 💰 Personal Expense & Budget Tracker (CLI + SQLite + Excel)

A lightweight, robust command-line application backed by **SQLite** and **Microsoft Excel (.xlsx)** to manage personal finances, track spending categories, enter daily expenses, and analyze weekly, monthly, and yearly budgets.

---

## 📌 Features

- **SQLite Database Persistence:** Stores transactions with `id`, `title`, `amount`, `category`, and `date`.
- **Excel Spreadsheet Integration (`expense_tracker.xlsx`):**
  - **Tab 1: `Daily Expenses`**:
    - Pre-populated with existing expenses.
    - Designed for direct daily data entry in Excel, Google Sheets, or LibreOffice.
    - Category dropdown validation (`Food`, `Transport`, `Bills`, `Shopping`, `Entertainment`, `Health`, `Education`, `Other`).
    - Auto-calculating `Week Num`, `Month`, and `Year` helper columns.
  - **Tab 2: `Summary & Analytics`**:
    - **Total Weekly Expenses**: Full schedule of 52 weeks with dynamic live `=SUMIF(...)` formulas.
    - **Total Monthly Expenses**: 12-month breakdown (Jan – Dec) with live formulas and full-year totals.
    - **Total Yearly Expenses**: Multi-year totals (2024 – 2028+) calculating dynamically.
    - **Spending by Category**: Category totals with percentage share formulas (`0.0%`).
    - **Live KPI Header Cards**: Real-time spending cards for Grand Total, Total Transactions, This Week's Total, This Month's Total, and This Year's Total.
- **Bidirectional Sync:**
  - Export SQLite database to Excel (`--export-excel`).
  - Import newly entered expenses from Excel back into SQLite (`--import-excel`).
- **Data Validation:** Enforces positive amounts, non-empty titles, and valid dates.
- **Automated Unit Tests:** 10 isolated tests with `pytest` testing SQLite operations, Excel generation, formulas, and bidirectional synchronization.
- **CI/CD Pipeline:** Automated GitHub Actions workflow on every push and pull request.

---

## 🚀 How to Run Locally

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the Application & Generate Excel File
```bash
python3 expense_tracker.py
```
This runs the CLI demo and automatically creates/refreshes `expense_tracker.xlsx`.

### 3. Work with the Excel File
- Open `expense_tracker.xlsx` in Excel, LibreOffice, or Google Sheets.
- Enter your daily expenses in the **Daily Expenses** tab (Date, Category, Description, Amount).
- Switch to the **Summary & Analytics** tab to view your live weekly, monthly, and yearly totals.

### 4. Sync Excel Entries Back to SQLite
If you entered new expenses in `expense_tracker.xlsx`, you can import them into the database anytime:
```bash
python3 expense_tracker.py --import-excel
```

### 5. Export Database to Custom Excel Path
```bash
python3 expense_tracker.py --export-excel my_budget.xlsx
```

### 6. Run Automated Tests
```bash
pytest -v
```

---

## 🔄 CI/CD Workflow (`.github/workflows/ci.yml`)

Every push or pull request to the `main` branch automatically:
1. Spawns an Ubuntu environment with Python 3.12.
2. Installs dependencies (`pytest`, `openpyxl`).
3. Runs `python expense_tracker.py` as a smoke test.
4. Executes the full `pytest` test suite (10 tests).
5. Reports build status.

---

## 👤 Author
**Ravina Sonawane**

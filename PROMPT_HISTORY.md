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

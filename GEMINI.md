# 🧠 AI Context & Long-Term Memory Instructions

This repository is **Personal Expense Tracker** (Author: **Ravina Sonawane**).

## ⚠️ Mandatory Memory Protocol

1. **Always Read Prompt History First:**
   - At the beginning of every interaction or session, read [PROMPT_HISTORY.md](file:///home/hitesh/ravina_project/Personal_Expense_Tracker/PROMPT_HISTORY.md) to recall all past user prompts, requirements, decisions, and actions taken.

2. **Save Every New User Prompt:**
   - Whenever the user gives a new prompt, task, or request, append a new entry to [PROMPT_HISTORY.md](file:///home/hitesh/ravina_project/Personal_Expense_Tracker/PROMPT_HISTORY.md) with:
     - Date and timestamp
     - The exact user prompt / request
     - The summary of actions taken or response provided
   - This guarantees that no context is ever lost, even after months or years.

## 📌 Project Architecture
- **Language / Runtime:** Python 3.12
- **Main Application:** `expense_tracker.py` (CLI + SQLite database `expenses.db`)
- **Unit Tests:** `test_expense_tracker.py` (pytest with isolated SQLite database fixtures)
- **CI/CD:** `.github/workflows/ci.yml` (GitHub Actions workflow)
- **Dependencies:** `requirements.txt`

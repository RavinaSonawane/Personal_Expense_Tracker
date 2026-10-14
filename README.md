# 💰 Personal Expense & Budget Tracker (CLI + SQLite)

A lightweight, robust command-line application backed by **SQLite** to manage personal finances, track spending categories, and enforce clean data validation.

---

## 📌 Features

- **SQLite Database Persistence:** Stores transactions with `id`, `title`, `amount`, `category`, and `date`.
- **Data Validation:** Prevents negative/zero amounts and empty titles or categories.
- **Spending Analytics:**
  - Calculate total expenditures.
  - Category breakdown with percentages and visual progress bars (`████`).
- **CRUD Operations:** Add expenses, list all records, filter by category, and delete records by ID.
- **Automated Unit Tests:** 8 isolated tests with `pytest` utilizing fixtures to run on temporary databases without touching production data.
- **CI/CD Pipeline:** Automated GitHub Actions workflow on every push and pull request.

---

## 🚀 How to Run Locally

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the Application
```bash
python3 expense_tracker.py
```

### 3. Run Automated Tests
```bash
pytest -v
```

---

## 🔄 CI/CD Workflow (`.github/workflows/ci.yml`)

Every push or pull request to the `main` branch automatically:
1. Spawns an Ubuntu environment with Python 3.12.
2. Installs required dependencies (`pytest`).
3. Runs `python expense_tracker.py` as a smoke test.
4. Executes the full `pytest` test suite.
5. Reports build status.

---

## 👤 Author
**Ravina Sonawane**

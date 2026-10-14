# =========================================================
#  Project: Personal Expense & Budget Tracker (CLI + SQLite)
#  Author: Ravina Sonawane
#  Description: Manage expenses, track budget, and analyze spending
# =========================================================

import sqlite3
from datetime import date as dt_date


DEFAULT_DB = "expenses.db"


def get_connection(db_path=DEFAULT_DB):
    """Return a sqlite3 database connection with Row factory."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path=DEFAULT_DB):
    """Create the expenses table if it does not already exist."""
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                amount REAL NOT NULL,
                category TEXT NOT NULL,
                date TEXT NOT NULL
            )
            """
        )
        conn.commit()


def add_expense(title, amount, category, expense_date=None, db_path=DEFAULT_DB):
    """
    Add a new expense.
    Validates that title is not empty, amount > 0, and category is provided.
    """
    if not title or not title.strip():
        raise ValueError("Expense title cannot be empty.")

    try:
        amount = float(amount)
    except (TypeError, ValueError):
        raise ValueError("Amount must be a valid number.")

    if amount <= 0:
        raise ValueError("Expense amount must be greater than zero.")

    if not category or not category.strip():
        raise ValueError("Category cannot be empty.")

    if expense_date is None:
        expense_date = dt_date.today().isoformat()

    init_db(db_path)

    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO expenses (title, amount, category, date)
            VALUES (?, ?, ?, ?)
            """,
            (title.strip(), round(amount, 2), category.strip().capitalize(), str(expense_date)),
        )
        conn.commit()
        return cursor.lastrowid


def get_all_expenses(db_path=DEFAULT_DB):
    """Return a list of all expenses as dictionaries."""
    init_db(db_path)
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, amount, category, date FROM expenses ORDER BY id DESC")
        rows = cursor.fetchall()
        return [dict(row) for row in rows]


def get_total_expenses(db_path=DEFAULT_DB):
    """Return the sum of all expenses."""
    init_db(db_path)
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COALESCE(SUM(amount), 0.0) FROM expenses")
        (total,) = cursor.fetchone()
        return round(float(total), 2)


def get_category_summary(db_path=DEFAULT_DB):
    """Return a dictionary of category -> total amount spent."""
    init_db(db_path)
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT category, ROUND(SUM(amount), 2) as total
            FROM expenses
            GROUP BY category
            ORDER BY total DESC
            """
        )
        rows = cursor.fetchall()
        return {row["category"]: row["total"] for row in rows}


def delete_expense(expense_id, db_path=DEFAULT_DB):
    """Delete an expense by ID. Returns True if deleted, False if not found."""
    init_db(db_path)
    with get_connection(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM expenses WHERE id = ?", (expense_id,))
        conn.commit()
        return cursor.rowcount > 0


def display_expenses(expenses):
    """Format and print expenses in a clean table format."""
    if not expenses:
        print("\nNo expenses found.")
        return

    print("\n" + "=" * 65)
    print(f"{'ID':<5} {'Date':<12} {'Category':<15} {'Title':<20} {'Amount':>10}")
    print("-" * 65)
    for exp in expenses:
        print(
            f"{exp['id']:<5} {exp['date']:<12} {exp['category']:<15} "
            f"{exp['title'][:18]:<20} ${exp['amount']:>9.2f}"
        )
    print("=" * 65)


def run_demo(db_path=DEFAULT_DB):
    """Run an automated demonstration of the expense tracker."""
    print("=" * 55)
    print("   PERSONAL EXPENSE TRACKER - DEMO")
    print("=" * 55)

    init_db(db_path)

    # Add sample expenses if empty
    if not get_all_expenses(db_path):
        add_expense("Groceries", 45.50, "Food", "2026-10-10", db_path)
        add_expense("Metro Pass", 20.00, "Transport", "2026-10-11", db_path)
        add_expense("Coffee with Friends", 12.75, "Food", "2026-10-12", db_path)
        add_expense("Internet Bill", 55.00, "Bills", "2026-10-13", db_path)
        print(">> Added initial demo expenses.")

    expenses = get_all_expenses(db_path)
    display_expenses(expenses)

    total = get_total_expenses(db_path)
    print(f"\n💰 Total Spending: ${total:.2f}")

    summary = get_category_summary(db_path)
    print("\n📊 Spending by Category:")
    for cat, amt in summary.items():
        percentage = (amt / total * 100) if total > 0 else 0
        bar = "█" * int(percentage // 10)
        print(f"  {cat:<12}: ${amt:>7.2f} ({percentage:>5.1f}%) {bar}")
    print("=" * 55)


if __name__ == "__main__":
    run_demo()

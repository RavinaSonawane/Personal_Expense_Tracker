# =========================================================
#  Tests for Personal Expense & Budget Tracker
# =========================================================

import pytest
from expense_tracker import (
    init_db,
    add_expense,
    get_all_expenses,
    get_total_expenses,
    get_category_summary,
    delete_expense,
)


@pytest.fixture
def temp_db(tmp_path):
    """Fixture providing a fresh isolated database file for each test."""
    db_file = str(tmp_path / "test_expenses.db")
    init_db(db_file)
    return db_file


def test_init_db(temp_db):
    """Test initializing database and creating table."""
    expenses = get_all_expenses(temp_db)
    assert expenses == []
    assert get_total_expenses(temp_db) == 0.0


def test_add_expense_valid(temp_db):
    """Test adding a valid expense."""
    exp_id = add_expense("Lunch", 15.50, "Food", "2026-10-14", temp_db)
    assert exp_id is not None
    assert exp_id > 0

    all_exp = get_all_expenses(temp_db)
    assert len(all_exp) == 1
    assert all_exp[0]["title"] == "Lunch"
    assert all_exp[0]["amount"] == 15.50
    assert all_exp[0]["category"] == "Food"
    assert all_exp[0]["date"] == "2026-10-14"


def test_add_expense_invalid_amount(temp_db):
    """Test adding expense with negative or zero amount raises ValueError."""
    with pytest.raises(ValueError, match="Expense amount must be greater than zero."):
        add_expense("Free Item", 0, "Other", db_path=temp_db)

    with pytest.raises(ValueError, match="Expense amount must be greater than zero."):
        add_expense("Refund", -10, "Other", db_path=temp_db)


def test_add_expense_invalid_title(temp_db):
    """Test adding expense with empty title raises ValueError."""
    with pytest.raises(ValueError, match="Expense title cannot be empty."):
        add_expense("", 25.0, "Food", db_path=temp_db)

    with pytest.raises(ValueError, match="Expense title cannot be empty."):
        add_expense("   ", 25.0, "Food", db_path=temp_db)


def test_add_expense_invalid_category(temp_db):
    """Test adding expense with empty category raises ValueError."""
    with pytest.raises(ValueError, match="Category cannot be empty."):
        add_expense("Snack", 5.0, "", db_path=temp_db)


def test_get_total_expenses(temp_db):
    """Test calculation of total expenses."""
    add_expense("Coffee", 4.50, "Food", db_path=temp_db)
    add_expense("Bus", 2.50, "Transport", db_path=temp_db)
    add_expense("Book", 13.00, "Education", db_path=temp_db)

    total = get_total_expenses(temp_db)
    assert total == 20.00


def test_get_category_summary(temp_db):
    """Test summary grouped by category."""
    add_expense("Breakfast", 10.00, "Food", db_path=temp_db)
    add_expense("Dinner", 25.00, "Food", db_path=temp_db)
    add_expense("Train", 15.00, "Transport", db_path=temp_db)

    summary = get_category_summary(temp_db)
    assert summary["Food"] == 35.00
    assert summary["Transport"] == 15.00


def test_delete_expense(temp_db):
    """Test deleting an expense by ID."""
    exp_id = add_expense("Unwanted item", 50.0, "Misc", db_path=temp_db)
    assert get_total_expenses(temp_db) == 50.0

    # Delete existing
    deleted = delete_expense(exp_id, temp_db)
    assert deleted is True
    assert get_total_expenses(temp_db) == 0.0

    # Delete non-existent ID
    deleted_again = delete_expense(9999, temp_db)
    assert deleted_again is False

"""Sheet shapes that pass validation, ported from stress_test.py."""
import datetime as dt

from excel_to_sql import excel_to_sql

from sheets import OK


def test_clean_sheet_validates(make_sheet, engine):
    # Arrange
    p = make_sheet("clean", [OK])

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert n == 1


def test_ragged_and_blank_rows_pass(make_sheet, engine):
    # Arrange: short row, blank row in the middle, all-None row
    p = make_sheet("ragged", [OK, [], [2, "A", 1], OK, [None] * 9])

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert n == 3


def test_header_offset_by_empty_rows_and_cols(make_sheet, engine):
    # Arrange
    p = make_sheet("offset", [OK], pre_rows=3, pre_cols=2)

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert n == 1


def test_unicode_quotes_and_sql_in_note_pass(make_sheet, engine):
    # Arrange
    row = [1, "ส้มแม่สิน", 1, None, None, None, None, None,
           "it's a \"test\"\nline2 🍊 ' ; DROP TABLE orders;--"]
    p = make_sheet("unicode", [row])

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert n == 1


def test_float_edge_values_pass(make_sheet, engine):
    # Arrange: inf reaches the loader as an empty cell
    p = make_sheet("floats", [
        [1, "A", 1, 0.1 + 0.2, None, None, None, None, None],
        [2, "A", 1, 1.5e-9, None, None, None, None, None],
        [3, "A", 1, float("inf"), None, None, None, None, None],
    ])

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert n == 3


def test_header_only_sheet_validates_zero_rows(make_sheet, engine):
    # Arrange
    p = make_sheet("headeronly", [])

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert n == 0

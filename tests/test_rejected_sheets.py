"""Sheet shapes that must be rejected, ported from stress_test.py."""
import datetime as dt
import zipfile

import pytest
from openpyxl import Workbook
from sqlalchemy import exc as sa_exc

from excel_to_sql import IngestionError, excel_to_sql

from sheets import HEADER, OK


def test_header_case_mismatch_names_the_real_column(make_sheet, engine):
    # Arrange
    p = make_sheet("case", [OK],
                   header=[" id ", "Customer", "qty", "price", "ship_date",
                           "created_at", "pickup", "paid", "note"])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "[Sheet1!B1]" in str(e.value)
    assert "did you mean 'customer'?" in str(e.value)


def test_duplicate_header_reports_first_position(make_sheet, engine):
    # Arrange
    p = make_sheet("dup", [OK + ["dup"]], header=HEADER + ["qty"])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "[Sheet1!J1]" in str(e.value)
    assert "already used at C1" in str(e.value)


def test_unknown_header_lists_the_table_columns(make_sheet, engine):
    # Arrange
    p = make_sheet("unknown", [OK + ["x"]], header=HEADER + ["discount"])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "column 'discount'" in str(e.value)
    assert "table has: id, customer, qty" in str(e.value)


def test_missing_not_null_column_is_rejected(make_sheet, engine):
    # Arrange
    p = make_sheet("missing", [[1, 10, 250.5]], header=["id", "qty", "price"])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "column 'customer'" in str(e.value)
    assert "column is NOT NULL with no default" in str(e.value)


def test_number_stored_as_text_rejects_both_cells(make_sheet, engine):
    # Arrange
    p = make_sheet("textnum", [[1, "ACME", "10", "250.5", dt.date(2026, 1, 5),
                                None, None, None, None]])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "[Sheet1!C2]" in str(e.value)
    assert "[Sheet1!D2]" in str(e.value)


def test_date_as_text_and_as_serial_group_into_one_problem(make_sheet, engine):
    # Arrange
    p = make_sheet("datetext", [
        [1, "ACME", 1, None, "2026-01-05", None, None, None, None],
        [2, "ACME", 1, None, 46027, None, None, None, None],
    ])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "[Sheet1!E2, Sheet1!E3]" in str(e.value)
    assert "got 2 cells" in str(e.value)


def test_datetime_with_time_component_rejected_in_date_column(make_sheet, engine):
    # Arrange
    p = make_sheet("mixup", [[1, "ACME", 1, None, dt.datetime(2026, 1, 5, 9, 0),
                              dt.date(2026, 1, 5), None, None, None]])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "[Sheet1!E2]" in str(e.value)
    assert "cell carries a time component" in str(e.value)


def test_integer_beyond_2p53_rejected(make_sheet, engine):
    # Arrange
    p = make_sheet("bigint", [
        [1234567890123456789, "ACME", 1, None, None, None, None, None, None],
        [9007199254740993, "ACME", 1, None, None, None, None, None, None],
    ])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "[Sheet1!A2, Sheet1!A3]" in str(e.value)
    assert "2^53" in str(e.value)


def test_non_boolean_into_bool_column_rejected(make_sheet, engine):
    # Arrange: row 2 is a valid boolean, rows 3 and 4 are 1 and "TRUE"
    p = make_sheet("bools", [
        [1, "A", 1, None, None, None, None, True, None],
        [2, "A", 1, None, None, None, None, 1, None],
        [3, "A", 1, None, None, None, None, "TRUE", None],
    ])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "[Sheet1!H3, Sheet1!H4]" in str(e.value)
    assert "format the cell as TRUE/FALSE" in str(e.value)


def test_bool_into_int_column_rejected(make_sheet, engine):
    # Arrange
    p = make_sheet("boolint", [[1, "A", True, None, None, None, None, None, None]])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "[Sheet1!C2]" in str(e.value)
    assert "got boolean" in str(e.value)


def test_blank_in_not_null_column_rejected(make_sheet, engine):
    # Arrange: whitespace-only and empty string are both blank
    p = make_sheet("blankish", [
        [1, "   ", 1, None, None, None, None, None, None],
        [2, "", 1, None, None, None, None, None, None],
    ])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "[Sheet1!B2, Sheet1!B3]" in str(e.value)
    assert "column has no default" in str(e.value)


def test_varchar_overflow_rejected(make_sheet, engine):
    # Arrange: customer is VARCHAR(10), this is 16 characters
    p = make_sheet("toolong", [[1, "ACME" * 4, 1, None, None, None, None, None, None]])

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "expected VARCHAR(10)" in str(e.value)
    assert "text of length 16" in str(e.value)
    assert "value is too long" in str(e.value)

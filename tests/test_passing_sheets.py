"""Sheet shapes that pass validation, ported from stress_test.py."""
import datetime as dt

from sqlalchemy import (Column, Date, DateTime, Integer, MetaData, Table,
                        select, text)

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


def test_int_column_accepts_whole_float(engine, make_sheet):
    # Arrange: xlsx stores every number as a double, so 5 and 5.0 are one cell
    p = make_sheet("wholefloat", [[1, "A", 5.0, None, None, None, None, None, None]])

    # Act
    n = excel_to_sql(p, "orders", engine)

    # Assert
    assert n == 1
    with engine.connect() as conn:
        qty = conn.execute(text("select qty from orders")).scalar()
    assert qty == 5


def test_date_and_datetime_widen_and_narrow(engine, make_sheet):
    # Arrange: a date widens to midnight datetime, a midnight datetime narrows to date
    p = make_sheet("datemix", [
        [1, "A", 1, None, dt.date(2026, 1, 5), dt.date(2026, 1, 6),
         None, None, None],
        [2, "A", 1, None, dt.datetime(2026, 1, 7, 0, 0),
         dt.datetime(2026, 1, 8, 9, 30), None, None, None],
    ])

    # Act
    n = excel_to_sql(p, "orders", engine)

    # Assert
    assert n == 2
    md = MetaData()
    orders = Table("orders", md,
                   Column("id", Integer, primary_key=True),
                   Column("ship_date", Date),
                   Column("created_at", DateTime))
    with engine.connect() as conn:
        rows = conn.execute(
            select(orders.c.ship_date, orders.c.created_at)
            .order_by(orders.c.id)).fetchall()
    assert rows[0].ship_date == dt.date(2026, 1, 5)
    assert rows[0].created_at == dt.datetime(2026, 1, 6)
    assert rows[1].ship_date == dt.date(2026, 1, 7)
    assert rows[1].created_at == dt.datetime(2026, 1, 8, 9, 30)

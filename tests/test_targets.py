"""The target parameter's three accepted shapes, as the README promises."""
import datetime as dt

from sqlalchemy import text

from excel_to_sql import excel_to_sql

from sheets import OK


def test_url_string_target(make_sheet, engine, tmp_path):
    # Arrange: schema exists via the engine fixture, URL points at the same file
    p = make_sheet("clean", [OK])

    # Act
    n = excel_to_sql(p, "orders", f"sqlite:///{tmp_path}/orders.db", dry_run=True)

    # Assert
    assert n == 1


def test_dry_run_writes_nothing(make_sheet, engine):
    # Arrange
    p = make_sheet("clean", [OK])

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert n == 1
    with engine.connect() as conn:
        assert conn.execute(text("select count(*) from orders")).scalar() == 0


def test_connection_target_caller_commits(make_sheet, engine):
    # Arrange
    p = make_sheet("clean", [OK])

    # Act
    with engine.connect() as conn:
        n = excel_to_sql(p, "orders", conn)

        # Assert: rows exist inside the caller's transaction only
        with engine.connect() as other:
            assert other.execute(text("select count(*) from orders")).scalar() == 0
        conn.commit()

    with engine.connect() as other:
        assert other.execute(text("select count(*) from orders")).scalar() == 1

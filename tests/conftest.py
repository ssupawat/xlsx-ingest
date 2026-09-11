from pathlib import Path

import pytest
from sqlalchemy import (BigInteger, Boolean, Column, Date, DateTime, Integer,
                        MetaData, Numeric, String, Table, Time, create_engine)

from sheets import HEADER, sheet as build_sheet


@pytest.fixture
def engine(tmp_path):
    eng = create_engine(f"sqlite:///{tmp_path}/orders.db")
    md = MetaData()
    Table(
        "orders", md,
        Column("id", BigInteger, primary_key=True),
        Column("customer", String(10), nullable=False),
        Column("qty", Integer, nullable=False),
        Column("price", Numeric),
        Column("ship_date", Date),
        Column("created_at", DateTime),
        Column("pickup", Time),
        Column("paid", Boolean),
        Column("note", String),
    )
    md.create_all(eng)
    yield eng
    eng.dispose()


@pytest.fixture
def make_sheet(tmp_path):
    def _make(name, rows, header=HEADER, pre_rows=0, pre_cols=0, sheets=None):
        return build_sheet(Path(tmp_path) / f"{name}.xlsx", rows,
                           header=header, pre_rows=pre_rows,
                           pre_cols=pre_cols, sheets=sheets)
    return _make

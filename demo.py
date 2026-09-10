"""Runnable proof: build a dirty xlsx, point it at a real sqlite table."""
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from sqlalchemy import (Column, Date, Integer, MetaData, Numeric, String,
                        Table, create_engine, select)

from excel_to_sql import IngestionError, excel_to_sql

# ---- 1. real table = the contract ----------------------------------------
engine = create_engine("sqlite:///demo.db")
md = MetaData()
orders = Table(
    "orders", md,
    Column("id", Integer, primary_key=True),
    Column("customer", String(10), nullable=False),
    Column("qty", Integer, nullable=False),
    Column("price", Numeric),
    Column("ship_date", Date),
)
md.drop_all(engine)
md.create_all(engine)

# ---- 2. a "clean" sheet that still breaks the contract -------------------
wb = Workbook()
ws = wb.active
ws.append(["id", "customer", "qty", "price", "ship_date"])
ws.append([1, "ACME", 10, 250.5, dt.date(2026, 1, 5)])       # ok
ws.append([2, "ACME", 2.5, 99, dt.date(2026, 1, 6)])         # qty fractional
ws.append([3, "ACME", 5, 10, "2026-01-07"])                  # date typed as text
ws.append([4, None, 5, None, None])                          # customer NOT NULL
ws.append([5, "A-VERY-LONG-NAME", 1, 1, None])               # exceeds VARCHAR(10)
wb.save("dirty.xlsx")

try:
    excel_to_sql("dirty.xlsx", "orders", engine, dry_run=True)
except IngestionError as e:
    print(e, "\n")

# ---- 3. clean sheet -> execute -------------------------------------------
wb = Workbook()
ws = wb.active
ws.append(["id", "customer", "qty", "price", "ship_date"])
ws.append([1, "ACME", 10, 250.5, dt.date(2026, 1, 5)])
ws.append([2, "TOYO", 3, None, dt.date(2026, 1, 6)])         # blank -> NULL
wb.save("clean.xlsx")

n = excel_to_sql("clean.xlsx", "orders", engine, batch_size=1)
print("inserted:", n)
with engine.connect() as c:
    for row in c.execute(select(orders)):
        print(" ", row)

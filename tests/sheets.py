"""Shared .xlsx builders for the test suite."""
import datetime as dt
from pathlib import Path

from openpyxl import Workbook

HEADER = ["id", "customer", "qty", "price", "ship_date",
          "created_at", "pickup", "paid", "note"]

OK = [1, "ACME", 10, 250.5, dt.date(2026, 1, 5),
      dt.datetime(2026, 1, 5, 9, 30), dt.time(14, 0), True, "ok"]


def sheet(path, rows, header=HEADER, pre_rows=0, pre_cols=0, sheets=None):
    """One workbook on `path`: optional padding, header, data rows, extra sheets."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheets[0] if sheets else "Sheet1"
    pad = [None] * pre_cols
    for _ in range(pre_rows):
        ws.append([])
    if header is not None:
        ws.append(pad + header)
    for r in rows:
        ws.append(pad + list(r))
    for extra in (sheets or [])[1:]:
        wb.create_sheet(extra).append(header)
    wb.save(path)
    return Path(path)

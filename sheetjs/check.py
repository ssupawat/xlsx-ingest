"""Does the checker still work on a workbook that SheetJS re-serialised?

  python3 build_source.py     # an .xlsx with percent, currency, dates, a merge,
  node roundtrip.js           #   #N/A and an uncached formula -> two SheetJS copies
  python3 check.py
"""
from pathlib import Path
import sys, pathlib; sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from sqlalchemy import create_engine
from excel_to_sql import IngestionError, excel_to_sql

engine = create_engine("sqlite:///stress.db")
for f in ["rt_orig.xlsx", "rt_default.xlsx", "rt_cellStylescellDates.xlsx"]:
    print(f"\n== {f}")
    try:
        print("  -> PASS rows", excel_to_sql(f, "orders", engine,
                                             sheet_name="Data", dry_run=True))
    except IngestionError as e:
        print("  ->", str(e).splitlines()[0])
        for line in str(e).splitlines()[1:6]:
            print("    ", line.strip()[:118])
    except ValueError as e:
        print("  -> ValueError:", str(e)[:300])

"""Stress test: throw every realistic Excel shape at excel_to_sql()."""
import datetime as dt
import shutil
import traceback
import zipfile
from pathlib import Path

from openpyxl import Workbook
from sqlalchemy import (BigInteger, Boolean, Column, Date, DateTime, Integer,
                        MetaData, Numeric, String, Table, Time, create_engine,
                        select)

from excel_to_sql import IngestionError, excel_to_sql

TMP = Path("stress")
if TMP.exists():
    shutil.rmtree(TMP)
TMP.mkdir()

engine = create_engine("sqlite:///stress.db")
md = MetaData()
t = Table(
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
md.drop_all(engine)
md.create_all(engine)

HEADER = ["id", "customer", "qty", "price", "ship_date",
          "created_at", "pickup", "paid", "note"]


def sheet(name, rows, header=HEADER, pre_rows=0, pre_cols=0, sheets=None):
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
    p = TMP / f"{name}.xlsx"
    wb.save(p)
    return p


def run(label, path, **kw):
    print(f"\n--- {label}")
    try:
        n = excel_to_sql(path, "orders", engine, dry_run=True, **kw)
        print(f"    PASS  rows={n}")
    except IngestionError as e:
        for line in str(e).splitlines()[:6]:
            print("   ", line)
    except Exception as e:
        print(f"    {type(e).__name__}: {str(e).splitlines()[0][:150]}")


OK = [1, "ACME", 10, 250.5, dt.date(2026, 1, 5),
      dt.datetime(2026, 1, 5, 9, 30), dt.time(14, 0), True, "ok"]

# 1 ── baseline
run("1 baseline", sheet("s1", [OK]))

# 2 ── ragged rows (short row, trailing empties, blank row in the middle)
run("2 ragged / blank rows",
    sheet("s2", [OK, [], [2, "A", 1], OK, [None] * 9]))

# 3 ── header offset by empty rows/cols (Excel exports love this)
run("3 offset header", sheet("s3", [OK], pre_rows=3, pre_cols=2))

# 4 ── header whitespace + case
run("4 header ' Customer '",
    sheet("s4", [OK], header=[" id ", "Customer", "qty", "price", "ship_date",
                              "created_at", "pickup", "paid", "note"]))

# 5 ── duplicate header
run("5 duplicate header",
    sheet("s5", [OK + ["dup"]], header=HEADER + ["qty"]))

# 6 ── unknown column
run("6 unknown column",
    sheet("s6", [OK + ["x"]], header=HEADER + ["discount"]))

# 7 ── missing NOT NULL column
run("7 missing NOT NULL col",
    sheet("s7", [[1, 10, 250.5]], header=["id", "qty", "price"]))

# 8 ── numbers as text
run("8 number stored as text",
    sheet("s8", [[1, "ACME", "10", "250.5", dt.date(2026, 1, 5),
                  None, None, None, None]]))

# 9 ── date as text / date as serial number
run("9 date as text & serial",
    sheet("s9", [[1, "ACME", 1, None, "2026-01-05", None, None, None, None],
                 [2, "ACME", 1, None, 46027, None, None, None, None]]))

# 10 ── datetime in a DATE column, date in a DATETIME column
run("10 date/datetime mixups",
    sheet("s10", [[1, "ACME", 1, None, dt.datetime(2026, 1, 5, 9, 0),
                   dt.date(2026, 1, 5), None, None, None]]))

# 11 ── big integers beyond float53
run("11 bigint precision",
    sheet("s11", [[1234567890123456789, "ACME", 1, None, None, None, None, None, None],
                  [9007199254740993, "ACME", 1, None, None, None, None, None, None]]))

# 12 ── booleans / 0-1 / "TRUE"
run("12 boolean variants",
    sheet("s12", [[1, "A", 1, None, None, None, None, True, None],
                  [2, "A", 1, None, None, None, None, 1, None],
                  [3, "A", 1, None, None, None, None, "TRUE", None]]))

# 13 ── bool into INT column
run("13 bool -> INTEGER",
    sheet("s13", [[1, "A", True, None, None, None, None, None, None]]))

# 14 ── whitespace-only / empty string in NOT NULL
run("14 blank-ish NOT NULL",
    sheet("s14", [[1, "   ", 1, None, None, None, None, None, None],
                  [2, "", 1, None, None, None, None, None, None]]))

# 15 ── unicode, quotes, newline, emoji, long text
run("15 unicode & quotes",
    sheet("s15", [[1, "ส้มแม่สิน", 1, None, None, None, None, None,
                   "it's a \"test\"\nline2 🍊 ' ; DROP TABLE orders;--"]]))

# 16 ── float precision & scientific notation
run("16 float edge",
    sheet("s16", [[1, "A", 1, 0.1 + 0.2, None, None, None, None, None],
                  [2, "A", 1, 1.5e-9, None, None, None, None, None],
                  [3, "A", 1, float("inf"), None, None, None, None, None]]))

# 17 ── header only / totally empty sheet
run("17 header only", sheet("s17", []))
wb = Workbook(); wb.save(TMP / "s18.xlsx")
run("18 empty sheet", TMP / "s18.xlsx")

# 19 ── multi sheet: default first, by name, bad name
p = sheet("s19", [OK], sheets=["Data", "Summary"])
run("19a first sheet", p)
run("19b sheet_name='Summary'", p, sheet_name="Summary")
run("19c sheet_name='Nope'", p, sheet_name="Nope")

# 20 ── merged cells
wb = Workbook(); ws = wb.active
ws.append(HEADER)
ws.append(OK)
ws.append([2, "ACME", 3, None, None, None, None, None, None])
ws.merge_cells("B2:B3")
wb.save(TMP / "s20.xlsx")
run("20 merged customer B2:B3", TMP / "s20.xlsx")

# 21 ── formula without cached value (openpyxl writes no <v>)
wb = Workbook(); ws = wb.active
ws.append(HEADER)
ws.append([1, "ACME", "=5+5", None, None, None, None, None, None])
wb.save(TMP / "s21.xlsx")
run("21 formula, no cached value", TMP / "s21.xlsx")

# 22 ── formula WITH cached value + error cell (#N/A), patched into the XML
src = TMP / "s22_base.xlsx"
wb = Workbook(); ws = wb.active
ws.append(HEADER)
ws.append([1, "ACME", 1, None, None, None, None, None, "z"])
ws.append([2, "ACME", 1, None, None, None, None, None, "z"])
wb.save(src)
dst = TMP / "s22.xlsx"
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "xl/worksheets/sheet1.xml":
            x = data.decode()
            x = x.replace('<c r="C2" t="n"><v>1</v></c>',
                          '<c r="C2"><f>2+3</f><v>5</v></c>')
            x = x.replace('<c r="C3" t="n"><v>1</v></c>',
                          '<c r="C3" t="e"><f>NA()</f><v>#N/A</v></c>')
            data = x.encode()
        zout.writestr(item, data)
run("22 cached formula + #N/A cell", dst)

# 23 ── real insert path: duplicate PK
p = sheet("s23", [OK, OK])
print("\n--- 23 duplicate primary key (real insert)")
try:
    print("    inserted:", excel_to_sql(p, "orders", engine))
except Exception as e:
    print(f"    {type(e).__name__}: {str(e).splitlines()[0][:120]}")
with engine.connect() as c:
    print("    rows now in table:", len(list(c.execute(select(t)))))

# 24 ── table that does not exist
run("24 unknown table", sheet("s24", [OK]))
try:
    excel_to_sql(TMP / "s24.xlsx", "nope", engine, dry_run=True)
except Exception as e:
    print(f"    unknown table -> {type(e).__name__}")

# 25 ── csv / xls through the same reader
csvp = TMP / "s25.csv"
csvp.write_text("id,customer,qty\n1,ACME,10\n")
run("25 .csv input", csvp)

# 26 ── display formats that lie about the stored value
wb = Workbook(); ws = wb.active
ws.append(HEADER)
for i, fmt in enumerate(['0%', '"฿"#,##0.00', '0.00E+00', ';;;', '# ?/?'], start=1):
    ws.append([i, "A", 1, 0.15, None, None, None, None, None])
    ws.cell(row=i + 1, column=4).number_format = fmt
ws.append([6, "A", 1, 1234.5, None, None, None, None, None])       # plain
ws.cell(row=7, column=4).number_format = '#,##0.00'                # separators only
ws.append([7, "A", 1, 1234.5, None, None, None, None, None])       # General
wb.save(TMP / "s26.xlsx")
run("26 number formats", TMP / "s26.xlsx")

# 27 ── percent applied to a whole column, date format on a text column
wb = Workbook(); ws = wb.active
ws.append(HEADER)
ws.append([1, "A", 1, None, None, None, None, None, 0.5])
ws.cell(row=2, column=9).number_format = '0%'
wb.save(TMP / "s27.xlsx")
run("27 percent into a TEXT column", TMP / "s27.xlsx")

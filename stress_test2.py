"""Round 2 — the assumptions _bad_styles() makes about Excel's XML.

openpyxl never emits these shapes, so each file here is patched by hand.
"""
import datetime as dt
import re
import shutil
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from openpyxl import Workbook
from sqlalchemy import (BigInteger, Boolean, Column, Date, DateTime, Integer,
                        MetaData, Numeric, String, Table, Time, create_engine)

from excel_to_sql import IngestionError, excel_to_sql

TMP = Path("stress2")
if TMP.exists():
    shutil.rmtree(TMP)
TMP.mkdir()

engine = create_engine("sqlite:///stress.db")
md = MetaData()
Table("orders", md,
      Column("id", BigInteger, primary_key=True),
      Column("customer", String(10), nullable=False),
      Column("qty", Integer, nullable=False),
      Column("price", Numeric),
      Column("ship_date", Date),
      Column("created_at", DateTime),
      Column("pickup", Time),
      Column("paid", Boolean),
      Column("note", String))
md.create_all(engine)

HEADER = ["id", "customer", "qty", "price", "ship_date",
          "created_at", "pickup", "paid", "note"]
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def base(name: str) -> Path:
    """One row, D2 = 0.15 formatted as 0%  -> gives us a real percent xf."""
    wb = Workbook()
    ws = wb.active
    ws.append(HEADER)
    ws.append([1, "ACME", 1, 0.15, None, None, None, None, "note"])
    ws.cell(row=2, column=4).number_format = "0%"
    p = TMP / f"{name}.xlsx"
    wb.save(p)
    return p


def patch(src: Path, dst: Path, **files) -> Path:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            fn = files.get(item.filename.replace("/", "_").replace(".", "_"))
            if fn:
                data = fn(data.decode()).encode()
            zout.writestr(item, data)
    return dst


def percent_xf(path: Path) -> int:
    """index of the cellXf that carries numFmtId 9."""
    with zipfile.ZipFile(path) as z:
        xfs = ET.fromstring(z.read("xl/styles.xml")).find(f"{M}cellXfs")
    return next(i for i, xf in enumerate(xfs) if xf.get("numFmtId") == "9")


def run(label, path, expect, **kw):
    try:
        excel_to_sql(path, "orders", engine, dry_run=True, **kw)
        got, detail = "PASS", ""
    except IngestionError as e:
        got, detail = "REJECT", str(e).splitlines()[1].strip()
    except Exception as e:
        got, detail = type(e).__name__, str(e).splitlines()[0][:110]
    flag = "ok " if got == expect else "BUG"
    print(f"[{flag}] {label:<44} expect={expect:<7} got={got}")
    if detail:
        print(f"        {detail}")


print("A/B/C/G probe whether a percent format is seen when Excel expresses it "
      "in a way openpyxl never does.\n")

# A ── xf carries numFmtId=9 but is told not to apply it -> must NOT fire
src = base("a")
k = percent_xf(src)
run("A applyNumberFormat=0 (format not in use)",
    patch(src, TMP / "a2.xlsx",
          xl_styles_xml=lambda x: re.sub(
              r'<xf numFmtId="9"[^>]*/>',
              '<xf numFmtId="9" applyNumberFormat="0" xfId="0"/>', x, count=1)),
    "PASS")

# B ── format lives on the column, cells carry no style of their own
src = base("b")
k = percent_xf(src)
run("B <col style> percent, cells unstyled",
    patch(src, TMP / "b2.xlsx",
          xl_worksheets_sheet1_xml=lambda x: x
          .replace(f'<c r="D2" s="{k}"', '<c r="D2"')
          .replace("<sheetData>", f'<cols><col min="4" max="4" style="{k}"/></cols><sheetData>')),
    "REJECT")

# C ── format inherited from a named cell style via xfId
src = base("c")
k = percent_xf(src)


def add_parent(x: str) -> str:
    x = x.replace("<cellStyleXfs count=\"1\">", "<cellStyleXfs count=\"2\">")
    x = re.sub(r"(</cellStyleXfs>)", r'<xf numFmtId="9"/>\1', x, count=1)
    return x


run("C inherited from cellStyleXfs via xfId",
    patch(src, TMP / "c2.xlsx",
          xl_styles_xml=lambda x: add_parent(x).replace(
              f'<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
              '<xf xfId="1"/>', 1)),
    "REJECT")

# G ── format lives on the row
src = base("g")
k = percent_xf(src)
run("G <row customFormat> percent, cells unstyled",
    patch(src, TMP / "g2.xlsx",
          xl_worksheets_sheet1_xml=lambda x: x
          .replace(f'<c r="D2" s="{k}"', '<c r="D2"')
          .replace('<row r="2"', f'<row s="{k}" customFormat="1" r="2"', 1)),
    "REJECT")

# D ── a percent style on a text cell: the string still displays as itself
src = base("d")
k = percent_xf(src)
run("D percent style on a TEXT cell (no lie)",
    patch(src, TMP / "d2.xlsx",
          xl_worksheets_sheet1_xml=lambda x: x
          .replace(f'<c r="D2" s="{k}"', '<c r="D2"')
          .replace('<c r="I2" t="inlineStr"', f'<c r="I2" s="{k}" t="inlineStr"')),
    "PASS")

# E ── a file we cannot inspect at all must not sail through unchecked
p = TMP / "legacy.xls"
p.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 512)  # OLE2 header
run("E legacy .xls (uninspectable)", p, "ValueError")

# F ── traps on a sheet that is not sheet1
wb = Workbook()
wb.active.title = "Summary"
wb.active.append(["ignore me"])
ws = wb.create_sheet("Data")
ws.append(HEADER)
ws.append([1, "ACME", 1, None, None, None, None, None, "z"])
p = TMP / "f.xlsx"
wb.save(p)
run("F #N/A on the 2nd sheet",
    patch(p, TMP / "f2.xlsx",
          xl_worksheets_sheet2_xml=lambda x: x.replace(
              '<c r="C2" t="n"><v>1</v></c>',
              '<c r="C2" t="e"><f>NA()</f><v>#N/A</v></c>')),
    "REJECT", sheet_name="Data")

# H ── same but selected as the default (first) sheet, to prove rels mapping
run("H clean 2nd sheet by name", p, "PASS", sheet_name="Data")

# I ── container check: .xlsx and nothing else
for label, fname, payload in [
    ("I1 .xlsm (macro workbook)", "book.xlsm", None),
    ("I2 .zip renamed to .xlsx", "fake.xlsx", b"PK\x03\x04not-a-workbook"),
    ("I3 .csv", "data.csv", b"id,customer\n1,ACME\n"),
]:
    q = TMP / fname
    if payload is None:
        shutil.copy(base("src_for_xlsm"), q)
    elif payload.startswith(b"PK"):
        with zipfile.ZipFile(q, "w") as z:
            z.writestr("hello.txt", "not a workbook")
    else:
        q.write_bytes(payload)
    run(label, q, "ValueError")

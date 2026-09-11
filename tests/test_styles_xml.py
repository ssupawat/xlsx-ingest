"""Style-inheritance XML shapes openpyxl never emits, ported from stress_test2.py.

openpyxl writes one canonical percent xf; each test patches the workbook XML by
hand to express the same format the way Excel itself can.
"""
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile

import pytest
from openpyxl import Workbook

from excel_to_sql import IngestionError, excel_to_sql

from sheets import HEADER

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def base(tmp_path, name):
    """One row, D2 = 0.15 formatted as 0%, giving a real percent xf."""
    wb = Workbook()
    ws = wb.active
    ws.append(HEADER)
    ws.append([1, "ACME", 1, 0.15, None, None, None, None, "note"])
    ws.cell(row=2, column=4).number_format = "0%"
    p = tmp_path / f"{name}.xlsx"
    wb.save(p)
    return p


def patch(src, dst, **files):
    """Rewrite zip members; keys are member names with / and . replaced by _."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            fn = files.get(item.filename.replace("/", "_").replace(".", "_"))
            if fn:
                data = fn(data.decode()).encode()
            zout.writestr(item, data)
    return dst


def percent_xf(path):
    """Index of the cellXf that carries numFmtId 9."""
    with zipfile.ZipFile(path) as z:
        xfs = ET.fromstring(z.read("xl/styles.xml")).find(f"{M}cellXfs")
    return next(i for i, xf in enumerate(xfs) if xf.get("numFmtId") == "9")


def add_parent(x):
    """Add a named cell style carrying the percent format."""
    x = x.replace('<cellStyleXfs count="1">', '<cellStyleXfs count="2">')
    return re.sub(r"(</cellStyleXfs>)", r'<xf numFmtId="9"/>\1', x, count=1)


def test_percent_xf_with_apply_number_format_zero_passes(tmp_path, engine):
    # Arrange: the xf carries numFmtId=9 but is told not to apply it
    src = base(tmp_path, "a")
    k = percent_xf(src)
    p = patch(src, tmp_path / "a2.xlsx", xl_styles_xml=lambda x: re.sub(
        r'<xf numFmtId="9"[^>]*/>',
        '<xf numFmtId="9" applyNumberFormat="0" xfId="0"/>', x, count=1))

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert n == 1


def test_percent_on_column_style_rejected(tmp_path, engine):
    # Arrange: format lives on the column, cells carry no style of their own
    src = base(tmp_path, "b")
    k = percent_xf(src)
    p = patch(src, tmp_path / "b2.xlsx", xl_worksheets_sheet1_xml=lambda x: x
              .replace(f'<c r="D2" s="{k}"', '<c r="D2"')
              .replace("<sheetData>",
                       f'<cols><col min="4" max="4" style="{k}"/></cols><sheetData>'))

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "percent format" in str(e.value)


def test_percent_inherited_from_named_style_rejected(tmp_path, engine):
    # Arrange: the cell xf inherits numFmtId from a named style via xfId
    src = base(tmp_path, "c")
    k = percent_xf(src)
    p = patch(src, tmp_path / "c2.xlsx", xl_styles_xml=lambda x: add_parent(x).replace(
        '<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" '
        'applyNumberFormat="1"/>', '<xf xfId="1"/>', 1))

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "percent format" in str(e.value)


def test_percent_on_row_style_rejected(tmp_path, engine):
    # Arrange: format lives on the row
    src = base(tmp_path, "g")
    k = percent_xf(src)
    p = patch(src, tmp_path / "g2.xlsx", xl_worksheets_sheet1_xml=lambda x: x
              .replace(f'<c r="D2" s="{k}"', '<c r="D2"')
              .replace('<row r="2"', f'<row s="{k}" customFormat="1" r="2"', 1))

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert "percent format" in str(e.value)


def test_percent_style_on_text_cell_passes(tmp_path, engine):
    # Arrange: the string still displays as itself, so nothing lies
    src = base(tmp_path, "d")
    k = percent_xf(src)
    p = patch(src, tmp_path / "d2.xlsx", xl_worksheets_sheet1_xml=lambda x: x
              .replace(f'<c r="D2" s="{k}"', '<c r="D2"')
              .replace('<c r="I2" t="inlineStr"',
                       f'<c r="I2" s="{k}" t="inlineStr"'))

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True)

    # Assert
    assert n == 1


def test_legacy_xls_rejected(tmp_path, engine):
    # Arrange: an OLE2 container we cannot inspect must not sail through
    p = tmp_path / "legacy.xls"
    p.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 512)

    # Act / Assert
    with pytest.raises(ValueError, match="only .xlsx is accepted"):
        excel_to_sql(p, "orders", engine, dry_run=True)


def test_na_cell_on_second_sheet_rejected(tmp_path, engine):
    # Arrange: an #N/A cell on sheet2, reached through the rels mapping
    wb = Workbook()
    wb.active.title = "Summary"
    wb.active.append(["ignore me"])
    ws = wb.create_sheet("Data")
    ws.append(HEADER)
    ws.append([1, "ACME", 1, None, None, None, None, None, "z"])
    src = tmp_path / "f.xlsx"
    wb.save(src)
    p = patch(src, tmp_path / "f2.xlsx", xl_worksheets_sheet2_xml=lambda x: x.replace(
        '<c r="C2" t="n"><v>1</v></c>',
        '<c r="C2" t="e"><f>NA()</f><v>#N/A</v></c>'))

    # Act
    with pytest.raises(IngestionError) as e:
        excel_to_sql(p, "orders", engine, dry_run=True, sheet_name="Data")

    # Assert
    assert "Excel error value #N/A" in str(e.value)


def test_clean_second_sheet_by_name_passes(tmp_path, engine):
    # Arrange: same file, its second sheet untouched
    wb = Workbook()
    wb.active.title = "Summary"
    wb.active.append(["ignore me"])
    ws = wb.create_sheet("Data")
    ws.append(HEADER)
    ws.append([1, "ACME", 1, None, None, None, None, None, "z"])
    p = tmp_path / "h.xlsx"
    wb.save(p)

    # Act
    n = excel_to_sql(p, "orders", engine, dry_run=True, sheet_name="Data")

    # Assert
    assert n == 1


def test_xlsm_suffix_rejected(tmp_path, engine):
    # Arrange
    src = base(tmp_path, "xlsm_src")
    p = tmp_path / "book.xlsm"
    shutil.copy(src, p)

    # Act / Assert
    with pytest.raises(ValueError, match="only .xlsx is accepted"):
        excel_to_sql(p, "orders", engine, dry_run=True)


def test_zip_renamed_to_xlsx_rejected(tmp_path, engine):
    # Arrange
    p = tmp_path / "fake.xlsx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("hello.txt", "not a workbook")

    # Act / Assert
    with pytest.raises(ValueError, match="not a readable .xlsx"):
        excel_to_sql(p, "orders", engine, dry_run=True)

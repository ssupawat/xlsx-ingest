"""
WYSIWYG Data Ingestion — PoC

1. No magic   : calamine reads each cell as-is. No column-level dtype inference,
                no NaN, no "1" -> 1.0, no silent str/date parsing.
2. Runtime contract : the live DB schema (SQLAlchemy reflection) is the validator.
                No pre-written model / no pydantic layer.
3. Actionable feedback : every rejection carries sheet + A1 coordinate + column
                + expected vs. got, so the user fixes the .xlsx and re-runs.
"""

from __future__ import annotations

import datetime as dt
import decimal
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional, Union

from python_calamine import CalamineWorkbook
from sqlalchemy import MetaData, Table, create_engine
from sqlalchemy.engine import Connection, Engine

# --------------------------------------------------------------------------- #
# Kinds — the only vocabulary we validate against
# --------------------------------------------------------------------------- #

INT, FLOAT, TEXT, DATE, DATETIME, TIME, BOOL, OPAQUE = (
    "INT", "FLOAT", "TEXT", "DATE", "DATETIME", "TIME", "BOOL", "OPAQUE",
)

_PY_TO_KIND = {
    bool: BOOL,
    int: INT,
    float: FLOAT,
    decimal.Decimal: FLOAT,
    str: TEXT,
    dt.datetime: DATETIME,
    dt.date: DATE,
    dt.time: TIME,
}


@dataclass(frozen=True)
class ColumnSpec:
    name: str
    kind: str
    nullable: bool
    max_length: Optional[int]
    has_default: bool


@dataclass(frozen=True)
class CellError:
    sheet: str
    coord: str          # "C7"
    column: str         # DB column name
    expected: str
    got: str
    detail: str = ""

    def __str__(self) -> str:
        base = (f"[{self.sheet}!{self.coord}] column '{self.column}': "
                f"expected {self.expected}, got {self.got}")
        return f"{base}; {self.detail}" if self.detail else base


class SourceFileError(ValueError):
    """The file itself is unusable, for a reason no coordinate can express and no
    user can fix by editing their data. Aimed at whoever built the pipeline."""


class IngestionError(ValueError):
    """Carries every violation found in the sheet, not just the first one."""

    def __init__(self, errors: list[CellError], shown: int = 20):
        self.errors = errors
        # One cause repeated down a column is one thing to fix, not 50,000 —
        # printing it per row buries every other problem in the sheet.
        groups: dict[tuple, list[CellError]] = {}
        for e in errors:
            groups.setdefault((e.column, e.expected, e.detail), []).append(e)
        lines = []
        for (column, expected, detail), hits in list(groups.items())[:shown]:
            first = ", ".join(f"{h.sheet}!{h.coord}" for h in hits[:3])
            where = f"{first} and {len(hits) - 3} more" if len(hits) > 3 else first
            got = hits[0].got if len(hits) == 1 else f"{len(hits)} cells"
            lines.append(f"  [{where}] column '{column}': expected {expected}, "
                         f"got {got}" + (f"; {detail}" if detail else ""))
        more = (f"\n  ... and {len(groups) - shown} more problems"
                if len(groups) > shown else "")
        super().__init__(
            f"{len(errors)} cell(s) rejected in {len(groups)} distinct problem(s):\n"
            + "\n".join(lines) + more)


# --------------------------------------------------------------------------- #
# Runtime contract: read the real schema
# --------------------------------------------------------------------------- #

def _kind_of(sa_type: Any) -> str:
    try:
        py = sa_type.python_type
    except (NotImplementedError, AttributeError):
        return OPAQUE
    return _PY_TO_KIND.get(py, OPAQUE)


def _reflect(table_name: str, bind: Union[Engine, Connection]) -> tuple[Table, dict[str, ColumnSpec]]:
    md = MetaData()
    table = Table(table_name, md, autoload_with=bind)
    specs = {}
    for col in table.columns:
        specs[col.name] = ColumnSpec(
            name=col.name,
            kind=_kind_of(col.type),
            nullable=bool(col.nullable),
            max_length=getattr(col.type, "length", None),
            has_default=col.default is not None
            or col.server_default is not None
            or bool(col.autoincrement is True and col.primary_key),
        )
    return table, specs


# --------------------------------------------------------------------------- #
# WYSIWYG cell validation — no coercion, only rejection
# --------------------------------------------------------------------------- #

_EXACT_INT = 2 ** 53


def _finite(v: float) -> bool:
    return v == v and v not in (float("inf"), float("-inf"))


def _blank(v: Any) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def _typename(v: Any) -> str:
    if isinstance(v, bool):
        return "boolean"
    if isinstance(v, int):
        return f"number({v})"
    if isinstance(v, float):
        return f"number({v!r})"
    if isinstance(v, str):
        return f"text({v[:24]!r})"
    if isinstance(v, dt.datetime):
        return "datetime"
    if isinstance(v, dt.date):
        return "date"
    if isinstance(v, dt.time):
        return "time"
    if isinstance(v, dt.timedelta):
        return "duration"
    return type(v).__name__


def _validate(value: Any, spec: ColumnSpec) -> tuple[Any, Optional[tuple[str, str, str]]]:
    """Return (value_to_bind, None) or (None, (expected, got, detail))."""
    if _blank(value):
        if spec.nullable or spec.has_default:
            return None, None
        return None, ("NOT NULL", "blank cell", "column has no default")

    k = spec.kind

    if k is OPAQUE:
        return value, None

    if k == BOOL:
        if isinstance(value, bool):
            return value, None
        return None, ("BOOLEAN", _typename(value), "format the cell as TRUE/FALSE")

    if k == INT:
        if isinstance(value, bool):
            return None, ("INTEGER", "boolean", "")
        if isinstance(value, float) and not _finite(value):
            return None, ("INTEGER", _typename(value), "not a finite number")
        if isinstance(value, int) or (isinstance(value, float) and value.is_integer()):
            # xlsx stores every number as a double: 5 and 5.0 are the same cell.
            # This is the only normalisation here, and only while it stays lossless.
            if abs(value) >= _EXACT_INT:
                return None, ("INTEGER", _typename(value),
                              "beyond Excel's exact integer range (2^53), the file "
                              "already lost digits; store this column as Text")
            return int(value), None
        if isinstance(value, float):
            return None, ("INTEGER", _typename(value), "value has a fractional part")
        return None, ("INTEGER", _typename(value), "cell is not numeric in Excel")

    if k == FLOAT:
        if isinstance(value, bool):
            return None, ("NUMERIC", "boolean", "")
        if isinstance(value, (int, float)):
            if isinstance(value, float) and not _finite(value):
                return None, ("NUMERIC", _typename(value), "not a finite number")
            return value, None
        return None, ("NUMERIC", _typename(value), "cell is not numeric in Excel")

    if k == TEXT:
        if isinstance(value, str):
            if spec.max_length is not None and len(value) > spec.max_length:
                return None, (f"VARCHAR({spec.max_length})",
                              f"text of length {len(value)}", "value is too long")
            return value, None
        return None, ("TEXT", _typename(value), "format the cell as Text in Excel")

    if k == DATETIME:
        if isinstance(value, dt.datetime):
            return value, None
        if isinstance(value, dt.date):
            return dt.datetime(value.year, value.month, value.day), None
        return None, ("DATETIME", _typename(value), "format the cell as Date in Excel")

    if k == DATE:
        if isinstance(value, dt.datetime):
            if value.time() != dt.time(0, 0):
                return None, ("DATE", "datetime", "cell carries a time component")
            return value.date(), None
        if isinstance(value, dt.date):
            return value, None
        return None, ("DATE", _typename(value), "format the cell as Date in Excel")

    if k == TIME:
        if isinstance(value, dt.time):
            return value, None
        if isinstance(value, dt.timedelta):
            return None, ("TIME", "duration", "cell is formatted as elapsed time")
        return None, ("TIME", _typename(value), "format the cell as Time in Excel")

    return value, None


# --------------------------------------------------------------------------- #
# Coordinates
# --------------------------------------------------------------------------- #

def _col_letter(idx0: int) -> str:
    s, n = "", idx0 + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _a1(row0: int, col0: int) -> str:
    return f"{_col_letter(col0)}{row0 + 1}"


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #

def excel_to_sql(
    file_path: Union[str, Path],
    table_name: str,
    target: Union[Engine, Connection, str],
    sheet_name: Optional[str] = None,
    batch_size: int = 1000,
    dry_run: bool = False,
) -> int:
    """
    Read an .xlsx file and insert its rows, validated against the live DB schema.

    target : Engine, active Connection, or a SQLAlchemy URL string.
    dry_run: validate only, nothing is executed.

    Returns the number of rows inserted (or, with dry_run, validated).
    Validation is eager and exhaustive: IngestionError lists every bad cell
    with its A1 coordinate before a single row is written.
    """
    bind, owns = _resolve_bind(target)
    try:
        table, specs = _reflect(table_name, bind)
        header, rows, sheet_label, traps = _read_sheet(file_path, sheet_name)
        mapping = _match_header(header, specs, sheet_label)
        records, errors = _build(rows, mapping, specs, sheet_label, traps)
        if errors:
            raise IngestionError(errors)
        if dry_run:
            return len(records)
        return _execute(records, table, target, bind, batch_size)
    finally:
        if owns:
            bind.dispose()


def _resolve_bind(target) -> tuple[Union[Engine, Connection], bool]:
    if isinstance(target, (Engine, Connection)):
        return target, False
    if isinstance(target, str) and "://" in target:
        return create_engine(target), True
    raise TypeError(
        f"target must be an Engine, a Connection, or a SQLAlchemy URL; got {target!r}"
    )


def _require_xlsx(path: Path) -> None:
    """.xlsx only, and only one that still carries its formatting.

    Every other container hides error cells, uncached formulas and display
    formats from the checks below, and silently skipping those checks would be
    worse than refusing the file.
    """
    import xml.etree.ElementTree as ET
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"{path.name}: only .xlsx is accepted; re-save the file.")
    if not zipfile.is_zipfile(path):
        raise ValueError(f"{path.name}: not a readable .xlsx (wrong contents for the extension).")
    with zipfile.ZipFile(path) as z:
        if "xl/workbook.xml" not in z.namelist():
            raise ValueError(
                f"{path.name}: not a readable .xlsx (wrong contents for the extension).")
        try:
            app = z.read("docProps/app.xml").decode("utf-8", "ignore")
            xfs = ET.fromstring(z.read("xl/styles.xml")).find("m:cellXfs", _NS)
        except KeyError:
            return
        # A workbook re-serialised by SheetJS without cellStyles carries exactly
        # one, default, cell format. SheetJS still computed the display text at
        # read time (w: "2026-01-05", w: "15%") and a viewer will happily show it,
        # but the file it wrote holds 46027 and 0.15 with nothing to mark them.
        # Whatever the user approved on screen, this file can no longer prove it.
        if "SheetJS" in app and xfs is not None and all(
                (xf.get("numFmtId") or "0") == "0" for xf in xfs):
            raise SourceFileError(
                f"{path.name}: written by SheetJS with the number formats stripped. "
                "Dates are bare serial numbers and percent/currency cells are bare "
                "values, so nothing here can be checked against what was shown on "
                "screen. Fix the generator (XLSX.read(buf, {cellDates: true}) and "
                "XLSX.write(wb, {cellStyles: true})), or pass the original upload "
                "through untouched."
            )


def _read_sheet(file_path, sheet_name) -> tuple[list, list, str, dict[str, tuple[str, str]]]:
    path = Path(file_path)
    _require_xlsx(path)
    wb = CalamineWorkbook.from_path(str(path))
    sheet = (wb.get_sheet_by_name(sheet_name) if sheet_name
             else wb.get_sheet_by_index(0))
    # skip_empty_area=False keeps real A1 coordinates intact.
    grid = sheet.to_python(skip_empty_area=False)
    if not grid:
        raise ValueError(f"sheet '{sheet.name}' is empty")

    # Cells calamine hands over as "" but which are NOT blank in the file.
    # Without this they would slide into the DB as NULL and nobody would know.
    traps = _suspect_cells(path, sheet.name)
    for (r0, c0), (r1, c1) in sheet.merged_cell_ranges:
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                if (r, c) != (r0, c0):
                    traps.setdefault(_a1(r, c),
                                     ("a literal value",
                                      f"part of merged range at {_a1(r0, c0)}"))

    for i, row in enumerate(grid):
        if any(not _blank(c) for c in row):
            return row, grid[i + 1:], sheet.name, traps
    raise ValueError(f"sheet '{sheet.name}' has no header row")


_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
       "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
       "p": "http://schemas.openxmlformats.org/package/2006/relationships"}


# Built-in number formats whose display differs from the stored value.
# 5-8 & 37-44: currency / accounting / negatives in parentheses.
# 9,10: percent (stored value is 100x smaller than what the eye reads).
# 11,48: scientific.  12,13: fractions.
_BAD_BUILTIN = {5, 6, 7, 8, 9, 10, 11, 12, 13,
                37, 38, 39, 40, 41, 42, 43, 44, 48}
_DATE_BUILTIN = set(range(14, 23)) | {45, 46, 47}


def _format_problem(code: str) -> Optional[str]:
    """Why this number format makes the cell lie. None = it shows what it stores."""
    import re
    bare = re.sub(r"\[[^\]]*\]", "", code)          # colours, [$฿-41E], conditions
    if not bare.strip("; "):
        return "hidden by the ';;;' format; the cell looks empty but holds a value"
    if "%" in bare:
        return "percent format; the value stored is 100x smaller than the digits shown"
    if re.search(r"[Ee][+-]", bare):
        return "scientific format"
    if "?" in bare and "/" in bare:
        return "fraction format"
    if "[$" in code or re.search(r'["\u00a4$€£¥฿]', code):
        return "currency format; the symbol is not part of the stored number"
    if "(" in bare:
        return "negatives shown in parentheses"
    return None


def _why_bad(fid: int, custom: dict) -> Optional[str]:
    if fid in custom:
        return _format_problem(custom[fid])
    if fid in _BAD_BUILTIN:
        return _format_problem(_BUILTIN_CODE.get(fid, "")) or "display format"
    return None                                    # General, plain, text, date


def _bad_styles(z, ET) -> dict[int, str]:
    """style index -> reason, for every cell style that displays something other
    than what it stores. Almost always empty, which keeps the sheet scan skippable.

    A cellXf may carry a numFmtId it does not use (applyNumberFormat="0") or use
    one it does not carry (inherited from the named style at xfId), so both
    directions have to be resolved or we get false positives and false negatives.
    """
    try:
        styles = ET.fromstring(z.read("xl/styles.xml"))
    except KeyError:
        return {}
    custom = {int(n.get("numFmtId")): n.get("formatCode") or ""
              for n in styles.iter(f"{{{_NS['m']}}}numFmt")}

    parents = [int(xf.get("numFmtId") or 0)
               for xf in (styles.find("m:cellStyleXfs", _NS) or [])]

    out = {}
    for i, xf in enumerate(styles.find("m:cellXfs", _NS) or []):
        apply_ = xf.get("applyNumberFormat")
        local = xf.get("numFmtId")
        if apply_ == "0" or (apply_ is None and local is None):
            xf_id = int(xf.get("xfId") or -1)
            fid = parents[xf_id] if 0 <= xf_id < len(parents) else 0
        else:
            fid = int(local or 0)
        why = _why_bad(fid, custom)
        if why:
            out[i] = why
    return out


_BUILTIN_CODE = {5: '"$"#,##0', 6: '"$"#,##0', 7: '"$"#,##0.00', 8: '"$"#,##0.00',
                 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
                 37: "#,##0 ;(#,##0)", 38: "#,##0 ;(#,##0)",
                 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;(#,##0.00)",
                 41: '_("$"* #,##0_)', 42: '_("$"* #,##0_)',
                 43: '_("$"* #,##0.00_)', 44: '_("$"* #,##0.00_)',
                 48: "##0.0E+0"}


def _col_of(coord: str) -> str:
    return "".join(ch for ch in coord if ch.isalpha())


def _column_styles(z, sheet_xml: str, ET) -> dict[str, str]:
    """column letter -> style index, from <cols><col style=.../></cols>."""
    out = {}
    m = f"{{{_NS['m']}}}"
    for event, el in ET.iterparse(z.open(sheet_xml), events=("start",)):
        if el.tag == m + "sheetData":
            break                                   # <cols> always precedes it
        if el.tag == m + "col" and el.get("style"):
            for c in range(int(el.get("min", 1)) - 1, int(el.get("max", 0))):
                out[_col_letter(c)] = el.get("style")
    return out


def _suspect_cells(path: Path, sheet_name: str) -> dict[str, tuple[str, str]]:
    """Every cell whose stored value is not what the sheet shows:

    - error values (#N/A, #REF!) and formulas with no cached result — calamine
      hands both over as "", i.e. they would slip into the DB as NULL;
    - cells wearing a display format (%, currency, scientific, hidden) — a clean
      file has none of these, so we reject rather than reinterpret.
    """
    import xml.etree.ElementTree as ET
    if not zipfile.is_zipfile(path):
        return {}                                   # unreachable: _read_sheet guards
    out: dict[str, tuple[str, str]] = {}
    with zipfile.ZipFile(path) as z:
        try:
            wbx = ET.fromstring(z.read("xl/workbook.xml"))
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        except KeyError:
            return {}
        rid = {rel.get("Id"): rel.get("Target") for rel in rels}
        target = None
        for s in wbx.findall("m:sheets/m:sheet", _NS):
            if s.get("name") == sheet_name:
                target = rid.get(s.get(f"{{{_NS['r']}}}id"))
        if not target:
            return {}
        name = "xl/" + target.lstrip("/").removeprefix("xl/")
        if name not in z.namelist():
            return {}

        bad = _bad_styles(z, ET)
        # Fast path: no lying styles anywhere in the workbook, no formulas and no
        # error cells => nothing to find. Scanning bytes beats building a DOM
        # over a million cells by ~100x, and this is the common case.
        markers = [b'<f>', b'<f ', b't="e"']
        for i in bad:                               # cell, row and column styles
            markers += [f's="{i}"'.encode(), f'style="{i}"'.encode()]
        with z.open(name) as fh:
            tail, hit = b"", False
            while chunk := fh.read(1 << 20):
                buf = tail + chunk
                if any(mk in buf for mk in markers):
                    hit = True
                    break
                tail = buf[-16:]
        if not hit:
            return out

        m = f"{{{_NS['m']}}}"
        col_style = _column_styles(z, name, ET)     # <col style="..."> defaults
        row_style = None
        for event, cell in ET.iterparse(z.open(name), events=("start", "end")):
            if cell.tag == m + "row" and event == "start":
                row_style = cell.get("s") if cell.get("customFormat") == "1" else None
                continue
            if cell.tag != m + "c" or event != "end":
                continue
            coord, value = cell.get("r"), cell.find(m + "v")
            if not coord:
                continue
            # a cell with no style of its own still inherits the row's, then the
            # column's — "select column D, press %" is exactly that case
            sid = cell.get("s") or row_style or col_style.get(_col_of(coord))
            style = bad.get(int(sid)) if sid is not None else None
            if cell.get("t") == "e":
                out[coord] = ("a literal value",
                              f"Excel error value {value.text if value is not None else ''}".strip())
            elif cell.find(m + "f") is not None and (
                    value is None or not (value.text or "").strip()):
                out[coord] = ("a literal value",
                              "formula with no cached result; open and save in Excel")
            elif style and cell.get("t") in (None, "n"):
                out[coord] = ("an unformatted cell", style + "; clear the number format in Excel")
            cell.clear()
    return out


def _match_header(header, specs, sheet_label) -> list[tuple[int, str]]:
    """[(excel_col_index, db_column_name)] — header text must match column names."""
    mapping, errors, seen = [], [], {}
    for idx, cell in enumerate(header):
        if _blank(cell):
            continue
        name = str(cell).strip()
        if name in seen:
            errors.append(CellError(sheet_label, _a1(0, idx), name, "a unique header",
                                    "duplicate",
                                    f"already used at {_a1(0, seen[name])}"))
            continue
        if name not in specs:
            hint = _near(name, specs)
            errors.append(CellError(sheet_label, _a1(0, idx), name, "a real column",
                                    "unknown header",
                                    hint or f"table has: {', '.join(specs)}"))
            continue
        mapping.append((idx, name))
        seen[name] = idx
    for name, spec in specs.items():
        if name not in seen and not spec.nullable and not spec.has_default:
            errors.append(CellError(sheet_label, "row 1", name, "a header column",
                                    "missing", "column is NOT NULL with no default"))
    if errors:
        raise IngestionError(errors)
    return mapping


def _near(name: str, specs) -> str:
    import difflib
    hit = difflib.get_close_matches(name.lower(),
                                    [c.lower() for c in specs], n=1, cutoff=0.75)
    return f"did you mean '{hit[0]}'?" if hit else ""


def _build(rows, mapping, specs, sheet_label, traps) -> tuple[list[dict], list[CellError]]:
    records, errors = [], []
    for r, row in enumerate(rows, start=1):          # r=1 -> excel row 2
        if all(_blank(c) for c in row):
            continue
        rec = {}
        for idx, name in mapping:
            raw = row[idx] if idx < len(row) else None
            coord = _a1(r, idx)
            trap = traps.get(coord)
            if trap:
                expected, detail = trap
                errors.append(CellError(sheet_label, coord, name, expected,
                                        _typename(raw) if not _blank(raw) else "unreadable cell",
                                        detail))
                continue
            value, err = _validate(raw, specs[name])
            if err:
                expected, got, detail = err
                errors.append(CellError(sheet_label, coord, name,
                                        expected, got, detail))
            else:
                rec[name] = value
        records.append(rec)
    return records, errors


def _execute(records, table, target, bind, batch_size) -> int:
    """Batched inserts inside a SINGLE transaction — all rows land or none do."""
    if not records:
        return 0

    def push(conn) -> int:
        total = 0
        for start in range(0, len(records), batch_size):
            chunk = records[start:start + batch_size]
            conn.execute(table.insert(), chunk)
            total += len(chunk)
        return total

    if isinstance(target, Connection):
        return push(target)                      # caller owns the transaction
    engine = bind if isinstance(bind, Engine) else target
    with engine.begin() as conn:
        return push(conn)

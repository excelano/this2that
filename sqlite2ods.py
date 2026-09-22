#!/usr/bin/env python3
"""sqlite2ods - Convert a SQLite database to an OpenDocument .ods spreadsheet.

Usage: sqlite2ods <input.db> [output.ods] [--table NAME]

Writes one sheet per table, named after the table, with a bold header row.
Pass --table to write a single-sheet document from just that table. Values
keep their SQLite type (integer, real, text); BLOB columns are written as
hex strings, since ODS has no binary cell type. Unlike sqlite2xlsx, sheets
here carry no autofilter or frozen header -- odfpy has no straightforward
API for either.
"""
import argparse
import re
import sqlite3
import sys

from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import Style, TextProperties
from odf.table import Table, TableCell, TableRow
from odf.text import P

# Stamped from packaging/control by install.sh and packaging/build-deb.sh.
# A checkout is not stamped, so running from the clone reports the dev
# placeholder rather than claiming a release it is not.
__version__ = "0.0.0+dev"

# ODS sheet names tolerate more than Excel's 31-char limit, but keep the
# family's sheet-naming rules consistent across sqlite2xlsx and sqlite2ods.
_BAD_TITLE = re.compile(r"[\\/*?:\[\]']")


def fmt(value):
    """Render a SQLite value as an ODS cell value."""
    if isinstance(value, bytes):
        return value.hex()
    return value


def table_names(conn):
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
    )
    return [r[0] for r in rows]


def sheet_title(name, used):
    """A unique, ODS-legal sheet title from a table name."""
    title = _BAD_TITLE.sub("_", name)[:31] or "Sheet"
    base, n = title, 2
    while title.lower() in used:
        suffix = f"_{n}"
        title = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(title.lower())
    return title


def cell_for(value, bold_style=None):
    """Build one TableCell, typed as a float or as text."""
    v = fmt(value)
    if isinstance(v, (int, float)):
        cell = TableCell(valuetype="float", value=v, stylename=bold_style)
        text = str(v)
    else:
        cell = TableCell(valuetype="string", stylename=bold_style)
        text = "" if v is None else str(v)
    cell.addElement(P(text=text))
    return cell


def write_sheet(doc, conn, table, sheet_name, bold_style):
    sheet = Table(name=sheet_name)
    cur = conn.execute(f'SELECT * FROM "{table}"')
    columns = [d[0] for d in cur.description]

    header = TableRow()
    for name in columns:
        header.addElement(cell_for(name, bold_style))
    sheet.addElement(header)

    for row in cur:
        tr = TableRow()
        for value in row:
            tr.addElement(cell_for(value))
        sheet.addElement(tr)

    doc.spreadsheet.addElement(sheet)


def main():
    ap = argparse.ArgumentParser(
        prog="sqlite2ods", description="Convert a SQLite database to an OpenDocument .ods spreadsheet."
    )
    ap.add_argument("--version", "-V", action="version", version=f"%(prog)s {__version__}")
    ap.add_argument("input", help="input SQLite database file")
    ap.add_argument("output", nargs="?", help="output .ods (default: input name with .ods)")
    ap.add_argument("--table", help="export only this table, as a single-sheet document")
    args = ap.parse_args()

    try:
        # mode=ro so a bad path surfaces as an error here instead of
        # sqlite3 silently creating a new, empty database file.
        conn = sqlite3.connect(f"file:{args.input}?mode=ro", uri=True)
        conn.execute("SELECT 1")
    except sqlite3.OperationalError as e:
        sys.exit(f"Error: {e}")

    tables = table_names(conn)
    if not tables:
        sys.exit("Error: no tables found")

    if args.table:
        if args.table not in tables:
            sys.exit(f"Error: no table named {args.table!r}. Available: {', '.join(tables)}")
        tables = [args.table]

    doc = OpenDocumentSpreadsheet()
    bold_style = Style(name="Bold", family="table-cell")
    bold_style.addElement(TextProperties(fontweight="bold"))
    doc.automaticstyles.addElement(bold_style)

    used = set()
    for table in tables:
        write_sheet(doc, conn, table, sheet_title(table, used), bold_style)

    out = args.output or re.sub(r"\.[^.]*$", "", args.input) + ".ods"
    doc.save(out)
    print(f"Wrote: {out}")


if __name__ == "__main__":
    main()

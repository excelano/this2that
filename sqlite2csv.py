#!/usr/bin/env python3
# Author: David M. Anderson
# Built with AI assistance (Claude, Anthropic)
"""sqlite2csv - Dump SQLite tables to RFC 4180 CSV.

Usage: sqlite2csv <input.db> [output.csv] [--table NAME] [--list]
       sqlite2csv <input.db> --all [-o DIR]

Converts one table by default: with a single-table database that table is
used, otherwise pass --table to pick one. --all dumps every table to its own
CSV, named after the table, in the output directory (default: the current
directory). --list prints the table names and exits.
"""
import argparse
import csv
import re
import sqlite3
import sys
from pathlib import Path

# Stamped from packaging/control by install.sh and packaging/build-deb.sh.
# A checkout is not stamped, so running from the clone reports the dev
# placeholder rather than claiming a release it is not.
__version__ = "0.0.0+dev"


def fmt(value):
    """Render a cell value as a CSV string."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.hex()
    return str(value)


def table_names(conn):
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
    )
    return [r[0] for r in rows]


def dump_table(conn, table, out_path):
    cur = conn.execute(f'SELECT * FROM "{table}"')
    columns = [d[0] for d in cur.description]
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(columns)
        for row in cur:
            writer.writerow([fmt(v) for v in row])
    print(f"Wrote: {out_path}")


def main():
    ap = argparse.ArgumentParser(prog="sqlite2csv", description="Dump SQLite tables to RFC 4180 CSV.")
    ap.add_argument("--version", "-V", action="version", version=f"%(prog)s {__version__}")
    ap.add_argument("input", help="input SQLite database file")
    ap.add_argument("output", nargs="?", help="output .csv (default: input name with .csv); ignored with --all")
    ap.add_argument("--table", help="table to export (default: the only table, if there is exactly one)")
    ap.add_argument("--all", action="store_true", help="dump every table to its own CSV")
    ap.add_argument("-o", "--outdir", default=".", help="output directory for --all (default: current directory)")
    ap.add_argument("--list", action="store_true", help="list table names and exit")
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

    if args.list:
        print("\n".join(tables))
        return

    if args.all:
        outdir = Path(args.outdir)
        outdir.mkdir(parents=True, exist_ok=True)
        for table in tables:
            dump_table(conn, table, outdir / f"{table}.csv")
        return

    if args.table:
        if args.table not in tables:
            sys.exit(f"Error: no table named {args.table!r}. Available: {', '.join(tables)}")
        table = args.table
    elif len(tables) == 1:
        table = tables[0]
    else:
        sys.exit(
            f"Error: database has {len(tables)} tables, pass --table to pick one or --all to dump every table. "
            f"Available: {', '.join(tables)}"
        )

    out = args.output or re.sub(r"\.[^.]*$", "", args.input) + ".csv"
    dump_table(conn, table, out)


if __name__ == "__main__":
    main()

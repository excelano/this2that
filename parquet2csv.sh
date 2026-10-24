#!/bin/bash
# parquet2csv - Convert a Parquet file to CSV via the DuckDB CLI.
#
# Author: David M. Anderson
# Built with AI assistance (Claude, Anthropic)
#
# Usage: parquet2csv <input.parquet> [output.csv]
#
# Output defaults to the input name with a .csv extension. Nested columns
# (struct, map, union, list) are written as JSON text and BLOB columns as hex,
# same as sqlite2csv.

set -euo pipefail

usage() {
    echo "Usage: parquet2csv <input.parquet> [output.csv]"
    exit "${1:-1}"
}

case "${1:-}" in
    -h|--help) usage 0 ;;
    "") usage ;;
esac

command -v duckdb >/dev/null || {
    echo "Error: parquet2csv needs the DuckDB CLI, which Debian does not package; see https://duckdb.org/docs/installation" >&2
    exit 1
}

IN="$1"
[ -f "$IN" ] || { echo "Error: file not found: $IN" >&2; exit 1; }
OUT="${2:-${IN%.*}.csv}"

sql_string() { printf "'%s'" "${1//\'/\'\'}"; }
SRC="read_parquet($(sql_string "$IN"))"
SCHEMA=$(duckdb -noheader -list -separator $'\t' -c "DESCRIBE SELECT * FROM $SRC")

cols=""
while IFS=$'\t' read -r name type _; do
    ident="\"${name//\"/\"\"}\""
    case "$type" in
        STRUCT*|MAP*|UNION*|*\]) expr="to_json($ident) AS $ident" ;;
        BLOB) expr="hex($ident) AS $ident" ;;
        *) expr="$ident" ;;
    esac
    cols="${cols:+$cols, }$expr"
done <<< "$SCHEMA"

duckdb -noheader -list -c "COPY (SELECT $cols FROM $SRC) TO $(sql_string "$OUT") (HEADER)" >/dev/null
echo "Wrote: $OUT"

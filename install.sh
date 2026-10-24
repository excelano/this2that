#!/bin/bash
# install.sh - Install the office-convert commands into ~/bin.
#
# Author: David M. Anderson
# Built with AI assistance (Claude, Anthropic)
#
# Copies each converter to ~/bin without its extension, so you invoke them as
# md2docx, docx2md, csv2xlsx, xlsx2csv, sqlite2csv, sqlite2xlsx, sqlite2ods,
# parquet2csv, and md2pptx.
#
# Requires: pandoc (md <-> docx, md -> pptx), python3-openpyxl (csv <-> xlsx,
# sqlite -> xlsx), and python3-odf (sqlite -> ods). sqlite2csv uses only the
# Python standard library. parquet2csv needs the DuckDB CLI, which Debian does
# not package (https://duckdb.org/docs/installation).
#   sudo apt install pandoc python3-openpyxl python3-odf

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# packaging/control is where the version is written down; the installed
# copies are stamped from it so the number lives in exactly one place.
VER=$(awk '/^Version:/{print $2}' "$SCRIPT_DIR/packaging/control")
[ -n "$VER" ] || { echo "no Version: in packaging/control" >&2; exit 1; }

mkdir -p "$HOME/bin"

# The sed stamps __version__ in the Python converters. The shell wrappers
# carry no such line and pass through unchanged.
install_one() {
    local src="$1" name="$2"
    sed "s/^__version__ = .*/__version__ = \"$VER\"/" "$SCRIPT_DIR/$src" > "$HOME/bin/$name"
    chmod +x "$HOME/bin/$name"
    echo "Installed: ~/bin/$name"
}

install_one md2docx.sh  md2docx
install_one docx2md.sh  docx2md
install_one csv2xlsx.py csv2xlsx
install_one xlsx2csv.py xlsx2csv
install_one sqlite2csv.py sqlite2csv
install_one sqlite2xlsx.py sqlite2xlsx
install_one sqlite2ods.py sqlite2ods
install_one parquet2csv.sh parquet2csv
install_one md2pptx.sh  md2pptx

if [[ ":$PATH:" != *":$HOME/bin:"* ]]; then
    echo 'export PATH="$HOME/bin:$PATH"' >> "$HOME/.bashrc"
    echo "Added ~/bin to PATH in .bashrc (restart shell or: source ~/.bashrc)"
fi

command -v pandoc >/dev/null || echo "Warning: pandoc not found (needed for md/docx/pptx)."
python3 -c "import openpyxl" 2>/dev/null || echo "Warning: python3-openpyxl not found (needed for csv/xlsx/sqlite2xlsx)."
python3 -c "import odf" 2>/dev/null || echo "Warning: python3-odf not found (needed for sqlite2ods)."
command -v duckdb >/dev/null || echo "Warning: duckdb not found (needed for parquet2csv)."

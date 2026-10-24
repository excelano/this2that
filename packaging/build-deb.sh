#!/bin/bash
# build-deb.sh — build the this2that .deb. Prints the .deb path.
#
# Author: David M. Anderson
# Built with AI assistance (Claude, Anthropic)
#
# The scripts are architecture-independent, so this produces one _all.deb that
# serves every machine. Version comes from packaging/control, which is the only
# place it is written down.
#
# Publish it from the excelano-apt repo:
#   ./ship-local.sh ~/this2that/packaging/build-deb.sh

set -euo pipefail

# Without this the builder's umask reaches the package: a 002 umask ships
# group-writable directories, which is wrong for anything under /usr.
umask 022

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
CONTROL="$ROOT/packaging/control"

VER=$(awk '/^Version:/{print $2}' "$CONTROL")
[ -n "$VER" ] || { echo "no Version: in $CONTROL" >&2; exit 1; }

STAGE=$(mktemp -d)
trap 'rm -rf "$STAGE"' EXIT
chmod 0755 "$STAGE"

mkdir -p "$STAGE/DEBIAN" "$STAGE/usr/bin" "$STAGE/usr/share/doc/this2that"
cp "$CONTROL" "$STAGE/DEBIAN/control"

# Same nine commands install.sh puts in ~/bin, under the same short names.
# The sed stamps __version__ in the Python converters from the same control
# file $VER came from. The shell wrappers carry no such line and pass
# through unchanged.
install_one() {
    sed "s/^__version__ = .*/__version__ = \"$VER\"/" "$ROOT/$1" > "$STAGE/usr/bin/$2"
    chmod 0755 "$STAGE/usr/bin/$2"
}
install_one md2docx.sh     md2docx
install_one docx2md.sh     docx2md
install_one csv2xlsx.py    csv2xlsx
install_one xlsx2csv.py    xlsx2csv
install_one sqlite2csv.py  sqlite2csv
install_one sqlite2xlsx.py sqlite2xlsx
install_one sqlite2ods.py  sqlite2ods
install_one parquet2csv.sh parquet2csv
install_one md2pptx.sh     md2pptx

install -m 0644 "$ROOT/LICENSE"   "$STAGE/usr/share/doc/this2that/copyright"
install -m 0644 "$ROOT/README.md" "$STAGE/usr/share/doc/this2that/README.md"

OUT="/tmp/this2that_${VER}_all.deb"
dpkg-deb --build --root-owner-group "$STAGE" "$OUT" >/dev/null
echo "$OUT"

#!/bin/bash
# docx2odt - Convert a Word .docx file to an OpenDocument .odt via pandoc.
#
# Author: David M. Anderson
# Built with AI assistance (Claude, Anthropic)
#
# Usage: docx2odt <input.docx> [output.odt]
#
# Output defaults to the input name with a .odt extension. The document's
# structure comes across (headings, lists, tables, images, links, footnotes)
# but not its page layout or custom styles. To brand the result, set
# REFERENCE_DOC to a styled .odt, or drop a reference.odt next to the input;
# pandoc uses it as the style template.

set -euo pipefail

usage() {
    echo "Usage: docx2odt <input.docx> [output.odt]"
    echo "  Set REFERENCE_DOC=<file.odt> to apply a style template."
    exit "${1:-1}"
}

case "${1:-}" in
    -h|--help) usage 0 ;;
    "") usage ;;
esac

IN="$1"
[ -f "$IN" ] || { echo "Error: file not found: $IN" >&2; exit 1; }
OUT="${2:-${IN%.*}.odt}"

ref_args=()
if [ -n "${REFERENCE_DOC:-}" ]; then
    ref_args=(--reference-doc "$REFERENCE_DOC")
elif [ -f "$(dirname "$IN")/reference.odt" ]; then
    ref_args=(--reference-doc "$(dirname "$IN")/reference.odt")
fi

pandoc "$IN" --from=docx --to=odt "${ref_args[@]}" -o "$OUT"
echo "Wrote: $OUT"

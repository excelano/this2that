#!/usr/bin/env python3
# Author: David M. Anderson
# Built with AI assistance (Claude, Anthropic)
"""odt2docx - Convert an OpenDocument .odt file to a Word .docx via pandoc.

Usage: odt2docx <input.odt> [output.docx]

The reverse of docx2odt. Output defaults to the input name with a .docx
extension. The document's structure comes across (headings, lists, tables,
images, links, footnotes) but not its page layout or custom styles. To brand
the result, set REFERENCE_DOC to a styled .docx, or drop a reference.docx
next to the input; pandoc uses it as the style template.

pandoc's ODT reader drops a table's heading rows and skips merged-away cells,
which loses the first row of most Writer tables and shifts the cells after a
merge into the wrong column. The input is rewritten first so the heading rows
read as ordinary first rows and each merged-away cell as an empty one; the
table's text and column positions survive, its heading styling and merges do
not.
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

# Stamped from packaging/control by install.sh and packaging/build-deb.sh.
# A checkout is not stamped, so running from the clone reports the dev
# placeholder rather than claiming a release it is not.
__version__ = "0.0.0+dev"


def flatten_tables(content):
    """content.xml with heading rows unwrapped and covered cells made plain empty cells."""
    content = re.sub(r"</?table:table-header-rows>", "", content)
    content = re.sub(r"<table:covered-table-cell\b[^>]*/>", "<table:table-cell/>", content)
    content = re.sub(r"<table:covered-table-cell\b[^>]*>", "<table:table-cell>", content)
    return content.replace("</table:covered-table-cell>", "</table:table-cell>")


def rewrite(src, dst):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item)
            if item.filename == "content.xml":
                data = flatten_tables(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)


def main():
    ap = argparse.ArgumentParser(prog="odt2docx", description="Convert an OpenDocument .odt file to a Word .docx.")
    ap.add_argument("--version", "-V", action="version", version=f"%(prog)s {__version__}")
    ap.add_argument("input", help="input .odt file")
    ap.add_argument("output", nargs="?", help="output .docx (default: input name with .docx)")
    args = ap.parse_args()

    if not os.path.isfile(args.input):
        sys.exit(f"Error: file not found: {args.input}")
    if not shutil.which("pandoc"):
        sys.exit("Error: odt2docx needs pandoc (sudo apt install pandoc)")

    beside = os.path.join(os.path.dirname(args.input), "reference.docx")
    reference = os.environ.get("REFERENCE_DOC") or (beside if os.path.isfile(beside) else None)
    ref_args = ["--reference-doc", reference] if reference else []
    out = args.output or re.sub(r"\.[^.]*$", "", args.input) + ".docx"

    with tempfile.TemporaryDirectory() as tmp:
        flat = os.path.join(tmp, "input.odt")
        try:
            rewrite(args.input, flat)
        except (zipfile.BadZipFile, KeyError) as e:
            sys.exit(f"Error: not an OpenDocument file: {args.input} ({e})")
        result = subprocess.run(["pandoc", flat, "--from=odt", "--to=docx", *ref_args, "-o", out])
    if result.returncode:
        sys.exit(result.returncode)
    print(f"Wrote: {out}")


if __name__ == "__main__":
    main()

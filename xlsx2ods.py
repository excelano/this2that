#!/usr/bin/env python3
# Author: David M. Anderson
# Built with AI assistance (Claude, Anthropic)
"""xlsx2ods - Convert an Excel .xlsx workbook to an OpenDocument .ods spreadsheet.

Usage: xlsx2ods <input.xlsx> [output.ods]

Every sheet comes across with typed cells (number, percentage, date, time,
boolean, text), merged cells, workbook-level named ranges, and formulas
rewritten from Excel syntax into OpenFormula, each stored with the result
Excel last saved. A formula that cannot be rewritten (a function OpenDocument
has no equivalent for, a structured table reference, a link to another
workbook) keeps its saved value and is reported on stderr by cell. Styling,
column widths, charts, conditional formatting, and pivot tables are not
carried.
"""
import argparse
import re
import sys
import zipfile
from datetime import date, datetime, time, timedelta

from odf.number import DateStyle, Day, Hours, Minutes, Month, Number, PercentageStyle, Seconds, TimeStyle, Year
from odf.number import Text as NumberText
from odf.opendocument import OpenDocumentSpreadsheet
from odf.style import Style
from odf.table import (CalculationSettings, CoveredTableCell, Iteration, NamedExpressions, NamedRange, Table,
                       TableCell, TableRow)
from odf.text import P
from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Token, Tokenizer, TokenizerError
from openpyxl.utils import column_index_from_string
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

# Stamped from packaging/control by install.sh and packaging/build-deb.sh.
# A checkout is not stamped, so running from the clone reports the dev
# placeholder rather than claiming a release it is not.
__version__ = "0.0.0+dev"

# Excel functions OpenFormula carries under a COM.MICROSOFT. name: the set
# LibreOffice 25.2 reads and writes, so the result opens in it unchanged.
COM_MICROSOFT = frozenset("""
AGGREGATE BAHTTEXT BETA.DIST BETA.INV BINOM.DIST BINOM.INV CEILING CEILING.MATH
CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CONCAT
CONFIDENCE.NORM CONFIDENCE.T COVARIANCE.P COVARIANCE.S ENCODEURL ERFC.PRECISE
ERF.PRECISE EXPON.DIST F.DIST F.DIST.RT FILTER FILTERXML F.INV F.INV.RT FLOOR
FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY
FORECAST.ETS.STAT FORECAST.LINEAR F.TEST GAMMA.DIST GAMMA.INV GAMMALN.PRECISE
HYPGEOM.DIST IFS ISO.CEILING LET LOGNORM.DIST LOGNORM.INV MAXIFS MINIFS MODE.MULT
MODE.SNGL NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV
PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC POISSON.DIST
QUARTILE.EXC QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ SEQUENCE SORT SORTBY STDEV.P
STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT TEXTJOIN T.INV T.INV.2T T.TEST UNIQUE VAR.P
VAR.S WEBSERVICE WEIBULL.DIST WORKDAY.INTL XLOOKUP XMATCH Z.TEST
""".split())

# Functions Excel writes with an _xlfn. prefix that are plain OpenFormula.
NATIVE_XLFN = frozenset("""
ACOT ACOTH ARABIC BASE BITAND BITLSHIFT BITOR BITRSHIFT BITXOR COMBINA COT COTH CSC
CSCH DAYS DECIMAL GAMMA GAUSS IFNA ISFORMULA ISOWEEKNUM MUNIT NUMBERVALUE PDURATION
PERMUTATIONA PHI RRI SEC SECH SHEET SHEETS UNICHAR UNICODE XOR
""".split())

RENAMED = {
    "FORMULATEXT": "FORMULA",
    "SKEW.P": "SKEWP",
    **{f: f"LEGACY.{f}" for f in "CHIDIST CHIINV CHITEST FDIST FINV NORMSDIST NORMSINV TDIST".split()},
}

# Legacy Excel functions with no OpenFormula counterpart.
UNSUPPORTED = frozenset("""
CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT
CUBEVALUE ECMA.CEILING PHONETIC RTD
""".split())

_CELL = re.compile(r"\$?[A-Za-z]{1,3}\$?\d+")
_REF_PART = re.compile(r"(\$?[A-Za-z]{1,3}\$?\d+|\$?[A-Za-z]{1,3}|\$?\d+)")
_SHEET = re.compile(r"('(?:[^']|'')+'|[^'!:]+)!")


class Untranslatable(Exception):
    pass


def odf_sheet(name):
    if re.fullmatch(r"[A-Za-z_]\w*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name):
        return "$" + name
    return "$'" + name.replace("'", "''") + "'"


def split_sheet(text):
    """Split an optional `Sheet!` prefix off a reference part."""
    m = _SHEET.match(text)
    if not m:
        return None, text
    sheet = m.group(1)
    if sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, text[m.end():]


def ref_to_odf(text, bracket=True):
    """An Excel reference (A1, Sheet!A1:B2, A:A, 1:2) as an OpenFormula one, or None."""
    sheet, rest = split_sheet(text)
    parts = rest.split(":")
    if len(parts) > 2:
        return None
    if len(parts) == 2:
        sheet2, parts[1] = split_sheet(parts[1])
        if sheet2 is not None and sheet2 != sheet:
            return None
    # A whole column or row (A:A, 1:2) is only a reference as a range.
    if len(parts) == 1 and not _CELL.fullmatch(parts[0]):
        return None
    if not all(_REF_PART.fullmatch(p) for p in parts):
        return None
    prefix = odf_sheet(sheet) if sheet is not None else ""
    body = prefix + "." + ":.".join(p.upper() for p in parts)
    return f"[{body}]" if bracket else body


def function_name(raw):
    """Map an Excel function name onto its OpenFormula name."""
    prefixed = raw.upper().startswith("_XLFN.")
    name = re.sub(r"^(_xlfn\.)?(_xlws\.)?", "", raw, flags=re.I).upper()
    if name in RENAMED:
        return RENAMED[name]
    if name in COM_MICROSOFT:
        return "COM.MICROSOFT." + name
    if name in UNSUPPORTED or (prefixed and name not in NATIVE_XLFN):
        raise Untranslatable(f"no OpenFormula equivalent for {name}")
    return name


def operand(token, names):
    if token.subtype == Token.LOGICAL:
        return token.value.upper() + "()"
    if token.subtype != Token.RANGE:
        return token.value
    ref = ref_to_odf(token.value)
    if ref:
        return ref
    if token.value.startswith("_xlpm.") or token.value in names:
        return token.value
    raise Untranslatable(f"cannot rewrite reference {token.value}")


def to_openformula(formula, names):
    """Rewrite an Excel formula ("=...") as an OpenFormula one ("of:=...")."""
    try:
        tokens = Tokenizer(formula).items
    except TokenizerError as e:
        raise Untranslatable(str(e)) from None

    def neighbour(i, step):
        i += step
        while 0 <= i < len(tokens) and tokens[i].type == Token.WSPACE:
            i += step
        return tokens[i] if 0 <= i < len(tokens) else None

    out = []
    for i, t in enumerate(tokens):
        if t.type == Token.OPERAND:
            out.append(operand(t, names))
        elif t.type == Token.FUNC and t.subtype == Token.OPEN:
            # The tokenizer fuses a range operator onto a following function:
            # A1:INDEX(...) arrives as one token, "A1:INDEX(".
            head, _, func = t.value[:-1].rpartition(":")
            if head:
                ref = ref_to_odf(head)
                if not ref:
                    raise Untranslatable(f"cannot rewrite reference {head}")
                out.append(ref + ":")
            out.append(function_name(func) + "(")
        elif t.type == Token.SEP:
            out.append(";" if t.subtype == Token.ARG else "|")
        elif t.type == Token.OP_IN and t.value == ",":
            out.append("~")
        elif t.type == Token.WSPACE:
            # Whitespace between two operands is Excel's intersection operator.
            before, after = neighbour(i, -1), neighbour(i, 1)
            ends = before is not None and (before.type == Token.OPERAND or before.subtype == Token.CLOSE)
            starts = after is not None and (after.type == Token.OPERAND or after.subtype == Token.OPEN)
            out.append("!" if ends and starts else t.value)
        else:
            out.append(t.value)
    return "of:=" + "".join(out)


class Styles:
    """Cell styles carrying the number formats a typed cell needs to display."""

    def __init__(self, doc):
        self.doc = doc
        self.cache = {}

    def _add(self, key, data_style):
        self.doc.automaticstyles.addElement(data_style)
        cell = Style(name=f"ce-{key}", family="table-cell", datastylename=data_style.getAttribute("name"))
        self.doc.automaticstyles.addElement(cell)
        self.cache[key] = cell
        return cell

    def date(self, with_time):
        key = "datetime" if with_time else "date"
        if key in self.cache:
            return self.cache[key]
        ds = DateStyle(name=f"N-{key}")
        for el in (Year(style="long"), NumberText(text="-"), Month(style="long"), NumberText(text="-"),
                   Day(style="long")):
            ds.addElement(el)
        if with_time:
            for el in (NumberText(text=" "), Hours(style="long"), NumberText(text=":"), Minutes(style="long"),
                       NumberText(text=":"), Seconds(style="long")):
                ds.addElement(el)
        return self._add(key, ds)

    def time(self):
        if "time" in self.cache:
            return self.cache["time"]
        ts = TimeStyle(name="N-time", truncateonoverflow="false")
        for el in (Hours(style="long"), NumberText(text=":"), Minutes(style="long"), NumberText(text=":"),
                   Seconds(style="long")):
            ts.addElement(el)
        return self._add("time", ts)

    def percent(self, decimals):
        key = f"pct{decimals}"
        if key in self.cache:
            return self.cache[key]
        ps = PercentageStyle(name=f"N-{key}")
        ps.addElement(Number(decimalplaces=decimals, minintegerdigits=1))
        ps.addElement(NumberText(text="%"))
        return self._add(key, ps)


def percent_decimals(number_format):
    """Decimal places of a percentage format such as 0.00%, or None if not one."""
    if "%" not in number_format:
        return None
    m = re.search(r"\.(0+)%", number_format)
    return len(m.group(1)) if m else 0


def iso_duration(seconds):
    seconds = round(seconds)
    return f"PT{seconds // 3600:02d}H{seconds % 3600 // 60:02d}M{seconds % 60:02d}S"


def typed_cell(value, number_format, styles, **attrs):
    """A TableCell holding value, typed and styled so it displays as Excel showed it."""
    if value is None:
        return TableCell(**attrs)
    if isinstance(value, bool):
        cell = TableCell(valuetype="boolean", booleanvalue=str(value).lower(), **attrs)
        text = str(value).upper()
    elif isinstance(value, (int, float)):
        decimals = percent_decimals(number_format)
        if decimals is None:
            cell = TableCell(valuetype="float", value=value, **attrs)
            text = str(value)
        else:
            cell = TableCell(valuetype="percentage", value=value, stylename=styles.percent(decimals), **attrs)
            text = f"{value * 100:.{decimals}f}%"
    elif isinstance(value, datetime):
        with_time = value.time() != time(0)
        text = value.isoformat(sep=" ", timespec="seconds") if with_time else value.date().isoformat()
        cell = TableCell(valuetype="date", datevalue=value.isoformat(timespec="seconds") if with_time else text,
                         stylename=styles.date(with_time), **attrs)
    elif isinstance(value, date):
        text = value.isoformat()
        cell = TableCell(valuetype="date", datevalue=text, stylename=styles.date(False), **attrs)
    elif isinstance(value, (time, timedelta)):
        seconds = (value.total_seconds() if isinstance(value, timedelta)
                   else value.hour * 3600 + value.minute * 60 + value.second)
        text = iso_duration(seconds)[2:].replace("H", ":").replace("M", ":").rstrip("S")
        cell = TableCell(valuetype="time", timevalue=iso_duration(seconds), stylename=styles.time(), **attrs)
    else:
        cell = TableCell(valuetype="string", **attrs)
        text = str(value)
    for line in text.split("\n"):
        cell.addElement(P(text=line))
    return cell


def formula_attrs(cell, names, warn):
    """The table:formula (and matrix span) attributes for a formula cell, or {}."""
    raw, where = cell.value, f"{cell.parent.title}!{cell.coordinate}"
    if isinstance(raw, ArrayFormula):
        text, span = raw.text, raw.ref
    elif isinstance(raw, DataTableFormula):
        warn(f"{where}: kept the value; what-if data tables are not converted")
        return {}
    elif cell.data_type == "f":
        text, span = raw, None
    else:
        return {}
    try:
        attrs = {"formula": to_openformula(text, names)}
    except Untranslatable as e:
        warn(f"{where}: kept the value; {e}")
        return {}
    if span and ":" in span:
        (c1, r1), (c2, r2) = (re.match(r"([A-Z]+)(\d+)", p).groups() for p in span.replace("$", "").split(":"))
        attrs["numbermatrixcolumnsspanned"] = column_index_from_string(c2) - column_index_from_string(c1) + 1
        attrs["numbermatrixrowsspanned"] = int(r2) - int(r1) + 1
    return attrs


def write_sheet(doc, ws_f, ws_v, styles, names, warn):
    sheet = Table(name=ws_f.title)
    spans, covered = {}, set()
    for rng in ws_f.merged_cells.ranges:
        spans[(rng.min_row, rng.min_col)] = (rng.max_col - rng.min_col + 1, rng.max_row - rng.min_row + 1)
        covered.update((r, c) for r in range(rng.min_row, rng.max_row + 1)
                       for c in range(rng.min_col, rng.max_col + 1))
        covered.discard((rng.min_row, rng.min_col))

    empty_rows = 0
    max_row, max_column = ws_f.max_row, ws_f.max_column
    for r in range(1, max_row + 1):
        cells, blanks = [], 0
        for c in range(1, max_column + 1):
            if (r, c) in covered:
                if blanks:
                    cells.append(TableCell(numbercolumnsrepeated=blanks))
                    blanks = 0
                cells.append(CoveredTableCell())
                continue
            cell = ws_f.cell(r, c)
            value = ws_v.cell(r, c).value
            attrs = formula_attrs(cell, names, warn)
            if (r, c) in spans:
                attrs["numbercolumnsspanned"], attrs["numberrowsspanned"] = spans[(r, c)]
            if value is None and not attrs:
                blanks += 1
                continue
            if blanks:
                cells.append(TableCell(numbercolumnsrepeated=blanks))
                blanks = 0
            cells.append(typed_cell(value, cell.number_format, styles, **attrs))
        if not cells:
            empty_rows += 1
            continue
        if empty_rows:
            row = TableRow(numberrowsrepeated=empty_rows)
            row.addElement(TableCell())
            sheet.addElement(row)
            empty_rows = 0
        row = TableRow()
        for cell in cells:
            row.addElement(cell)
        sheet.addElement(row)
    if not sheet.childNodes:
        row = TableRow()
        row.addElement(TableCell())
        sheet.addElement(row)
    doc.spreadsheet.addElement(sheet)


def named_ranges(wb, warn):
    """Workbook-level defined names that are plain references, as ODF named ranges."""
    ranges = {}
    for name, defn in wb.defined_names.items():
        if name.startswith("_xlnm."):
            continue
        address = ref_to_odf(defn.attr_text, bracket=False)
        if address and address.startswith("$"):
            ranges[name] = address
        else:
            warn(f"name {name}: not converted; only a single-sheet reference becomes a named range")
    return ranges


def main():
    ap = argparse.ArgumentParser(
        prog="xlsx2ods", description="Convert an Excel .xlsx workbook to an OpenDocument .ods spreadsheet."
    )
    ap.add_argument("--version", "-V", action="version", version=f"%(prog)s {__version__}")
    ap.add_argument("input", help="input .xlsx file")
    ap.add_argument("output", nargs="?", help="output .ods (default: input name with .ods)")
    args = ap.parse_args()

    try:
        wb_f = load_workbook(args.input)
        wb_v = load_workbook(args.input, data_only=True)
    except (OSError, ValueError, KeyError, InvalidFileException, zipfile.BadZipFile) as e:
        sys.exit(f"Error: {e}")

    warnings = []
    warn = warnings.append
    doc = OpenDocumentSpreadsheet()
    # Excel matches criteria (COUNTIF, MATCH, ...) with wildcards and ignores
    # case; an ODF document that does not say so uses case-sensitive regexes.
    settings = CalculationSettings(casesensitive="false", automaticfindlabels="false",
                                   useregularexpressions="false", usewildcards="true")
    calc = wb_f.calculation
    if calc.iterate:
        settings.addElement(Iteration(status="enable", steps=calc.iterateCount or 100,
                                      maximumdifference=calc.iterateDelta or 0.001))
    doc.spreadsheet.addElement(settings)
    styles = Styles(doc)
    names = named_ranges(wb_f, warn)
    for ws_f in wb_f.worksheets:
        write_sheet(doc, ws_f, wb_v[ws_f.title], styles, names, warn)

    if names:
        expressions = NamedExpressions()
        first = odf_sheet(wb_f.worksheets[0].title)
        for name, address in names.items():
            expressions.addElement(NamedRange(name=name, cellrangeaddress=address, basecelladdress=f"{first}.$A$1"))
        doc.spreadsheet.addElement(expressions)

    out = args.output or re.sub(r"\.[^.]*$", "", args.input) + ".ods"
    doc.save(out)
    for w in warnings:
        print(f"Warning: {w}", file=sys.stderr)
    print(f"Wrote: {out}")


if __name__ == "__main__":
    main()

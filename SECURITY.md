# Security Policy

## Reporting a vulnerability

Please report suspected vulnerabilities privately through GitHub Security Advisories at https://github.com/excelano/this2that/security/advisories/new. If you would rather not use GitHub, email david.anderson@excelano.com instead. I aim to respond within seven days.

Please do not open public issues for security problems.

## Supported versions

Security fixes ship in a new tagged release. Update with `sudo apt upgrade this2that` (if installed from the Excelano apt repo) or by pulling the repository and re-running `./install.sh`. There are no maintained release branches; only the latest release is supported.

## What this2that does

this2that is a set of command-line converters that run locally. Each reads an input file, writes an output file, and does nothing else: no configuration, no state, no telemetry, and no network connections. The scripts are thin wrappers — pandoc does the Markdown/Word/PowerPoint and Word/OpenDocument text conversion, openpyxl the Excel side, odfpy the OpenDocument spreadsheet side, Python's `sqlite3` the SQLite reads, and the DuckDB CLI the Parquet reads — so the security posture of a conversion is largely theirs.

**The input files are the attack surface.** `docx2md`, `docx2odt`, `odt2docx`, `xlsx2csv`, `xlsx2ods`, `ods2xlsx`, the `sqlite2*` converters, and `parquet2csv` parse files you may have received from someone else, and a malicious `.docx`, `.odt`, `.xlsx`, `.ods`, SQLite database, or `.parquet` is a file handed to a parser rather than a command handed to a shell. Converting an untrusted file exposes you to whatever pandoc, openpyxl, odfpy, SQLite, or DuckDB would be exposed to; keep them current. Formulas are rewritten as text and never evaluated: `xlsx2ods` and `ods2xlsx` carry a formula from one file into another, where it runs only when a spreadsheet application opens the result, as it would have run in the original. The SQLite converters open the database read-only. `parquet2csv` passes the input and output paths to DuckDB as quoted SQL string literals, so a crafted filename cannot change the query. `md2docx`, `md2pptx`, and `csv2xlsx` read plain text you authored and are correspondingly less interesting.

A reference document is also an input. `md2docx` and `md2pptx` apply a pandoc reference `.docx`/`.pptx` as a style template, taken from `REFERENCE_DOC` or picked up automatically from a `reference.docx`/`reference.pptx` sitting next to the input. That file is parsed like any other document, so a reference template from an untrusted source deserves the same caution as an untrusted input.

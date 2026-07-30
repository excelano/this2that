# Security Policy

## Reporting a vulnerability

Please report suspected vulnerabilities privately through GitHub Security Advisories at https://github.com/excelano/office-convert/security/advisories/new. If you would rather not use GitHub, email david.anderson@excelano.com instead. I aim to respond within seven days.

Please do not open public issues for security problems.

## Supported versions

Security fixes ship in a new tagged release. Update with `sudo apt upgrade office-convert` (if installed from the Excelano apt repo) or by pulling the repository and re-running `./install.sh`. There are no maintained release branches; only the latest release is supported.

## What office-convert does

office-convert is five command-line converters that run locally. Each reads an input file, writes an output file, and does nothing else: no configuration, no state, no telemetry, and no network connections. The scripts are thin wrappers — pandoc does the Markdown/Word/PowerPoint conversion and openpyxl does the CSV/Excel conversion, so the security posture of a conversion is largely theirs.

**The input files are the attack surface.** `docx2md` and `xlsx2csv` parse file formats you may have received from someone else, and a malicious `.docx` or `.xlsx` is a document handed to a parser rather than a command handed to a shell. Converting an untrusted document exposes you to whatever pandoc or openpyxl would be exposed to; keep both current through your distribution. `md2docx`, `md2pptx`, and `csv2xlsx` read plain text you authored and are correspondingly less interesting.

A reference document is also an input. `md2docx` and `md2pptx` apply a pandoc reference `.docx`/`.pptx` as a style template, taken from `REFERENCE_DOC` or picked up automatically from a `reference.docx`/`reference.pptx` sitting next to the input. That file is parsed like any other document, so a reference template from an untrusted source deserves the same caution as an untrusted input.

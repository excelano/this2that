# office-convert

Small command-line converters between plain-text source formats and their
Microsoft Office (or Office-adjacent) counterparts. The plain-text side
(Markdown, CSV) is the one you author and version; the other side is the
export you hand to people who live in Word, Excel, PowerPoint, or LibreOffice
Calc. The SQLite and Parquet converters run the other way, turning a
database or data file into CSV or a spreadsheet someone can open.

| Command | Direction | Engine |
|---|---|---|
| `md2docx`    | Markdown → Word `.docx`      | pandoc |
| `docx2md`    | Word `.docx` → Markdown      | pandoc |
| `csv2xlsx`   | CSV → Excel `.xlsx`          | openpyxl |
| `xlsx2csv`   | Excel `.xlsx` → CSV          | openpyxl |
| `sqlite2csv` | SQLite table(s) → CSV        | stdlib `sqlite3` |
| `sqlite2xlsx`| SQLite table(s) → Excel `.xlsx` | openpyxl |
| `sqlite2ods` | SQLite table(s) → OpenDocument `.ods` | odfpy |
| `parquet2csv`| Parquet → CSV                | DuckDB CLI |
| `md2pptx`    | Markdown → PowerPoint `.pptx` | pandoc |

PowerPoint is export-only: pandoc writes `.pptx` but cannot read it, so there
is no `pptx2md`.

## Install

From the Excelano apt repository, which pulls in pandoc, openpyxl, and odfpy
for you:

```bash
curl -fsSL https://excelano.com/apt/setup.sh | sudo sh   # one-time
sudo apt install office-convert
```

Or from a clone, which puts each script in `~/bin` under its short name:

```bash
sudo apt install pandoc python3-openpyxl python3-odf
./install.sh
```

`parquet2csv` also needs the DuckDB CLI, which Debian does not package, so
neither route installs it; get it from
[duckdb.org](https://duckdb.org/docs/installation). The other converters
work without it.

Uninstall the clone-installed copies with `./uninstall.sh`. To build the
package yourself, run `packaging/build-deb.sh`; the scripts are
architecture-independent, so one `_all.deb` serves every machine. This repo
has no release automation of its own, so it publishes through
`excelano-apt`'s `ship-local.sh packaging/build-deb.sh`, which builds, pools,
prunes, and re-signs the indices in one run.

## Usage

Every command takes an input file and an optional output path; the output
defaults to the input name with the new extension.

```bash
md2docx report.md                 # -> report.docx
md2docx report.md final.docx      # explicit output
docx2md proposal.docx             # -> proposal.md (images to ./media/)
csv2xlsx data.csv                 # -> data.xlsx (styled, typed)
csv2xlsx a.csv b.csv -o book.xlsx # -> one sheet per CSV, named by file stem
xlsx2csv workbook.xlsx --sheet Q3 # -> workbook.csv from the Q3 sheet
sqlite2csv app.db                 # -> app.csv, if app.db has exactly one table
sqlite2csv app.db --table users   # -> app.csv from the users table
sqlite2csv app.db --all -o out/   # -> out/users.csv, out/orders.csv, ...
sqlite2xlsx app.db                # -> app.xlsx, one sheet per table
sqlite2xlsx app.db --table users  # -> app.xlsx, single sheet from users
sqlite2ods app.db                 # -> app.ods, one sheet per table
parquet2csv events.parquet        # -> events.csv
md2pptx deck.md                   # -> deck.pptx
```

`csv2xlsx`, `xlsx2csv`, `sqlite2csv`, `sqlite2xlsx`, and `sqlite2ods` report
their version with `--version` or `-V`. The number is stamped from
`packaging/control` when the package is built or `install.sh` runs, so a
converter run straight out of a clone reports `0.0.0+dev` rather than
claiming a release it is not. The pandoc wrappers and `parquet2csv` are
short enough to sit inside the one-page exemption in the CLI flag standard
and carry no version flag.

### Branding the Office output

`md2docx` and `md2pptx` accept a pandoc reference document as a style template.
Set `REFERENCE_DOC`, or drop a `reference.docx` / `reference.pptx` next to the
input and it is picked up automatically.

```bash
REFERENCE_DOC=~/templates/excelano.docx md2docx report.md
```

### CSV typing

`csv2xlsx` conservatively types cells so Excel treats numbers as numbers and
dates as dates, but only when the conversion is lossless. Values Excel famously
mangles are kept as text: leading-zero codes (`007`, zip codes), phone-style
strings (`+1 555...`), and digit runs longer than 15. Pass `--text` to disable
inference entirely and write every cell verbatim.

Given several CSVs and an `-o` output, `csv2xlsx` writes one workbook with a
sheet per CSV, titled after each input's filename stem (sanitized and de-duped
for Excel's sheet-name rules), in the order given. A single input keeps the
plain `csv2xlsx input [output]` form; only two or more inputs need `-o`.

### SQLite export (sqlite2csv)

`sqlite2csv` uses only the Python standard library — no extra dependency.
It opens the database read-only, so a bad path errors out instead of
sqlite3 silently creating a new, empty file. With no `--table`, it exports
the database's only table if there is exactly one; with more than one it
errors and lists the table names, same as `xlsx2csv` does for sheets.
`--all` dumps every table to its own CSV in the output directory (`-o`,
default the current directory), and `--list` prints the table names and
exits. BLOB columns are written as hex strings.

### SQLite export (sqlite2xlsx, sqlite2ods)

Both write one workbook, one sheet per table, named after the table
(sanitized and de-duped for the format's sheet-name rules); `--table` writes
a single-sheet document from just that table instead. BLOB columns are
written as hex strings, same as `sqlite2csv`. `sqlite2xlsx` sheets get the
same bold frozen header, autofilter, and auto-fitted column widths as
`csv2xlsx`; `sqlite2ods` sheets get a bold header only, since odfpy has no
straightforward API for freeze panes or autofilter.

### Parquet export (parquet2csv)

`parquet2csv` hands the reading and writing to DuckDB and only decides how
each column lands in a CSV cell. Scalar columns are written as DuckDB
formats them; nested columns (struct, map, union, list) are written as JSON
text, and BLOB columns as hex strings, same as the SQLite converters.

### Slide division (md2pptx)

Slides break on a heading at the slide level or a horizontal rule (`---`), the
same rule the Axe Markdown slide viewer uses — so one `.md` both presents in the
browser and exports to PowerPoint here.

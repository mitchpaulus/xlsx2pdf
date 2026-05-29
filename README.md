# xlsx2pdf

A small Windows command-line utility for converting Excel workbooks to PDF. It
drives Excel via COM automation, so a local installation of Microsoft Excel is
required. The tool is intended as a building block for mass-converting `.xlsx`
files to PDF from scripts or pipelines.

## Contents

- `xlsx2pdf.cpp` — primary implementation, a native C++ executable.
- `xlsx2pdf.ps1` — equivalent PowerShell script, useful as a reference or for
  environments where rebuilding the C++ binary is inconvenient.
- `xlsx2pdf.sln` / `xlsx2pdf.vcxproj` — Visual Studio solution and project for
  building the C++ executable.

## Requirements

- Windows
- Microsoft Excel (the tool launches Excel via COM)
- Visual Studio 2019/2022 with the **Desktop development with C++** workload to
  build from source, **or** PowerShell 5+ to run `xlsx2pdf.ps1`

## Building

Open `xlsx2pdf.sln` in Visual Studio and build, or from a Developer Command
Prompt:

```
msbuild xlsx2pdf.sln /p:Configuration=Release /p:Platform=x64
```

The resulting executable is written under `bin\`.

## Usage

```
xlsx2pdf [options] <input-path> [worksheet-name]
xlsx2pdf batch [options]
```

- `input-path` — path to the `.xlsx` (or other Excel-readable) workbook.
- `worksheet-name` — optional. If omitted, the first worksheet is exported.
  If supplied and not found, the tool prints the list of available sheet names
  and exits with a non-zero status.

Options:

- `--landscape`, `-l` — export the page in landscape orientation.
- `--portrait`, `-p` — export the page in portrait orientation (default).
- `--fit-to-page`, `-f` — scale the worksheet to fit on a single page
  (sets `PageSetup.FitToPagesWide = 1` and `FitToPagesTall = 1`).
- `--skip-exists`, `-s` — skip the conversion when the target PDF already
  exists. In `batch` mode this skips per row, making an interrupted run cheap to
  resume (Excel is not even launched if every row is skipped).
- `--output`, `-o <path>` — write the PDF to `<path>` instead of the default
  location. Relative paths are resolved against the current directory, and the
  target directory must already exist.

Without `--output`, the PDF is written to `%TMP%\xlsx.pdf`. Status messages are
written to stderr, so they can be separated from any stdout consumers in a
pipeline.

### Examples

Convert the first worksheet of a workbook:

```
xlsx2pdf C:\reports\january.xlsx
```

Convert a specific worksheet:

```
xlsx2pdf C:\reports\january.xlsx "Summary"
```

Landscape, scaled to a single page:

```
xlsx2pdf --landscape --fit-to-page C:\reports\january.xlsx "Summary"
```

Write the PDF to a specific location:

```
xlsx2pdf --output C:\reports\pdfs\january.pdf C:\reports\january.xlsx "Summary"
```

## Mass conversion (`batch`)

The `batch` subcommand converts many workbooks in one invocation, reusing a
**single** Excel instance for the whole run. This avoids paying Excel's startup
and shutdown cost (several seconds each) per file, so it is dramatically faster
than launching `xlsx2pdf` once per workbook.

It reads tab-separated rows from **stdin**, one conversion per line:

```
<input-path>	<output-pdf>	[worksheet-name]
```

- Column 1 — path to the input workbook (required).
- Column 2 — path to write the PDF (required). The target directory must exist.
- Column 3 — worksheet name (optional; defaults to the first sheet).

Blank lines are skipped. Input is read as UTF-8 (a leading byte-order mark is
ignored). The page options (`--landscape`, `--fit-to-page`, …) apply to every
row. Progress is reported to stderr with an `(i/N)` counter, where `N` is the
number of non-blank rows; each row prints a `Converting …` line and then its
result as `OK`, `SKIP`, or `FAIL`:

```
(1/3) Converting C:\reports\jan.xlsx -> C:\pdfs\jan.pdf
(1/3) OK
(2/3) SKIP	C:\pdfs\feb.pdf already exists
(3/3) Converting C:\reports\mar.xlsx -> C:\pdfs\mar.pdf
(3/3) OK
Batch complete: 2 succeeded, 0 failed, 1 skipped, 3 total.
```

The process exits non-zero if any row failed — a failing row does not abort the
rest. Pass `--skip-exists` to skip rows whose output PDF already exists, which
makes a resumed run cheap.

PowerShell — build the TSV and pipe it in:

```powershell
Get-ChildItem -Path .\workbooks -Filter *.xlsx | ForEach-Object {
    "{0}`t{1}" -f $_.FullName, ".\pdfs\$($_.BaseName).pdf"
} | xlsx2pdf batch
```

cmd.exe — from an existing list file:

```
type list.tsv | xlsx2pdf batch --fit-to-page
```

### Single-instance vs. per-file

You can still convert one file at a time (e.g. in a shell loop with
`--output`), but each invocation launches and tears down its own Excel process.
Prefer `batch` whenever you have more than a handful of files.

## PowerShell variant

`xlsx2pdf.ps1` mirrors the C++ behavior and accepts the same two arguments:

```powershell
.\xlsx2pdf.ps1 -InputPath C:\reports\january.xlsx -WorksheetName "Summary" -Landscape -FitToPage
```

## Notes and limitations

- Excel must be installed and licensed on the machine running the tool.
- The output path defaults to `%TMP%\xlsx.pdf`; pass `--output` to write
  elsewhere. The target directory must already exist.
- The tool launches Excel headlessly (`Visible = false` in the C++ build) and
  suppresses dialogs (`DisplayAlerts = false`). If Excel hangs, the tool will
  wait briefly and then terminate the process it started.
- Only one worksheet is exported per run. To export an entire workbook,
  invoke the tool once per sheet.

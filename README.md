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

The PDF is written to `%TMP%\xlsx.pdf`. Status messages are written to stderr,
so they can be separated from any stdout consumers in a pipeline.

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

## Mass conversion

`xlsx2pdf` itself processes a single workbook per invocation. To convert many
files, wrap it in a shell loop and move the output between runs (since each
run overwrites `%TMP%\xlsx.pdf`).

PowerShell:

```powershell
Get-ChildItem -Path .\workbooks -Filter *.xlsx | ForEach-Object {
    xlsx2pdf $_.FullName
    Move-Item -Force "$env:TMP\xlsx.pdf" ".\pdfs\$($_.BaseName).pdf"
}
```

cmd.exe:

```
for %F in (workbooks\*.xlsx) do (
    xlsx2pdf "%F" && move /Y "%TMP%\xlsx.pdf" "pdfs\%~nF.pdf"
)
```

## PowerShell variant

`xlsx2pdf.ps1` mirrors the C++ behavior and accepts the same two arguments:

```powershell
.\xlsx2pdf.ps1 -InputPath C:\reports\january.xlsx -WorksheetName "Summary" -Landscape -FitToPage
```

## Notes and limitations

- Excel must be installed and licensed on the machine running the tool.
- The output path is fixed at `%TMP%\xlsx.pdf`; callers are responsible for
  moving or renaming the file between conversions.
- The tool launches Excel headlessly (`Visible = false` in the C++ build) and
  suppresses dialogs (`DisplayAlerts = false`). If Excel hangs, the tool will
  wait briefly and then terminate the process it started.
- Only one worksheet is exported per run. To export an entire workbook,
  invoke the tool once per sheet.

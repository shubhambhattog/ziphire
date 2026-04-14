# ZipHire CLI - Resume Bundler

A Node.js CLI tool for TnP workflows that reads student data from Excel, downloads resumes from Google Drive links, renames each PDF to `ROLLNO_NAME.pdf`, and packages everything into a ZIP.

## Repository

- Name: ziphire
- Owner: shubhambhattog
- Branch: main

## Tech Stack

- Node.js (single-file CLI)
- `xlsx` for Excel parsing
- `archiver` for ZIP creation
- Built-in `https` for downloads

## What It Does

1. Reads an Excel file passed as CLI argument.
2. Processes all sheets in the workbook (any sheet names).
3. Auto-detects Roll, Name, and Resume columns by header keywords for each sheet.
4. Extracts Google Drive file IDs from links.
5. Downloads files from Drive.
6. Renames each file to `ROLLNO_NAME.pdf`.
7. Adds files to a ZIP archive inside per-sheet folders.
8. Writes `_download_log.txt` inside the ZIP with `OK`, `FAIL`, and `SKIP` lines.

## Installation

```bash
npm install
```

## Usage

```bash
node ziphire.js students.xlsx
```

## Input Expectations

Each sheet in the Excel file should contain student rows with headers similar to:

- Roll: `roll`, `enrollment`, `enroll`, `reg`
- Name: `name`, `student`, `candidate`
- Resume link: `resume`, `cv`, `drive`, `link`, `url`, `gdrive`

The script auto-detects these columns by keyword matching for each sheet. Sheets without matching columns are skipped and recorded in diagnostics.

## Supported Drive Link Formats

- `https://drive.google.com/file/d/FILE_ID/view?...`
- URLs containing `?id=FILE_ID`
- Bare file ID strings

## Output

A ZIP is created in the same directory as the Excel file:

- `<excel_file_name>_resumes.zip`

Inside the ZIP:

- One folder per sheet (for example `Btech/`, `Mtech/`, or any sheet name)
- Resume PDFs named `ROLLNO_NAME.pdf` inside the corresponding sheet folder
- `_download_log.txt` with status lines, for example:

```text
OK    [Btech] 231210014_Aniket_Kumar_Singh.pdf
FAIL  [Mtech] 231210034_Bharat_Kumar.pdf - invalid Drive link: <url>
SKIP  [Btech] 231220058_Shruti_Agarwal.pdf - missing data
```

## Common Failure Reasons

- Invalid or malformed Drive URL
- Missing roll/name/resume values in a row
- Drive permissions not public
- HTTP errors while downloading

## Notes

- The tool processes rows sequentially.
- File/folder links from Drive are not valid file downloads.
- Keep Node and dependencies updated for reliability.

## Scripts

From `package.json`:

```bash
npm start -- students.xlsx
```

# Fileman CLI

Console CLI to help with file organization.

Commands:

- `organize [dir]` — Group files by `--by ext|date`, supports `--recursive` and `--dry-run`.
- `tree [dir]` — Print a directory tree. Options: `--depth <n>`, `--dirs-first`.
- `clean-empty [dir]` — Remove empty directories recursively.
- `rename [dir]` — Batch rename files. Options: `--match <regex>`, `--replace <str>`, `--ext <.pdf>`, `--dry-run`.
- `rename [dir]` — Batch rename files. Options: `--path <dir>`, `--match <regex>`, `--replace <str>`, `--ext <.pdf>`, `--dry-run`.
- `dedupe [dir]` — Find duplicates by SHA-256; `--delete` removes duplicates, `--dry-run` previews.
- `audit-nonpdf [dir]` — List all non-.pdf files and show if a same-name `.pdf` exists in the same folder. Options: `--path <dir>`, `--recursive`.
- `excel2pdf [dir]` — Convert all Excel files in a directory to PDF (all sheets). Options: `--path <dir>`, `--recursive`, `--out <dir>`, `--overwrite`.
- `excel2pdf-list [list]` — Convert Excel files listed in a text file to PDF (all sheets). Options: `--list <file>`, `--out <dir>`, `--overwrite`.
- `excel-audit-sheets [dir]` — Find Excel files that have more than N sheets. Options: `--path <dir>`, `--recursive`, `--min <n>` (default 2).
  - Deletion: add `--delete-pdf` to remove same-name `.pdf` files next to matched Excel; combine with `--dry-run` to preview.
  - Conversion: add `--to-pdf` to convert matched Excel to PDF alongside (all sheets); use `--overwrite` to replace existing PDFs; `--dry-run` previews actions.
- `office2pdf-missing [dir]` — Convert Word/Excel files without same-name `.pdf` to PDF (alongside). Options: `--path <dir>`, `--recursive`, `--dry-run`, `--overwrite`.

## Install / Run

Inside the project directory:

- Run directly: `npm run fileman -- <command> [options]`
- Or: `node bin/fileman.js <command> [options]`

Optional global link: `npm link` then use `fileman` anywhere.

## Examples

- Organize by extension: `npm run fileman -- organize c:\\path\\to\\folder --by ext`
- Organize by date (YYYY/MM): `npm run fileman -- organize . --by date --recursive`
- Show tree depth 2: `npm run fileman -- tree . --depth 2 --dirs-first`
- Clean empty dirs: `npm run fileman -- clean-empty .`
- Rename PDFs in a specific folder: `npm run fileman -- rename --path c:\\path\\to\\folder --ext .pdf --match "-" --replace "_"`
- Find duplicates: `npm run fileman -- dedupe .`
- Audit non-PDFs: `npm run fileman -- audit-nonpdf --path c:\\path\\to\\folder --recursive`
- Excel to PDF: `npm run fileman -- excel2pdf --path c:\\path\\to\\excel-folder --recursive --out c:\\path\\to\\pdfs`
- Excel to PDF from list: `npm run fileman -- excel2pdf-list --list c:\\path\\to\\list.txt --out c:\\path\\to\\pdfs --overwrite`
- Audit Excel sheets: `npm run fileman -- excel-audit-sheets --path c:\\path\\to\\excel-folder --recursive --min 2`
  - With deletion preview: `npm run fileman -- excel-audit-sheets --path c:\\path\\to\\excel-folder --recursive --min 2 --delete-pdf --dry-run`
  - Apply deletion: `npm run fileman -- excel-audit-sheets --path c:\\path\\to\\excel-folder --recursive --min 2 --delete-pdf`
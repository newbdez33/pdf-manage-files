#!/usr/bin/env node
const { program } = require('commander');

program
  .name('fileman')
  .description('File organization CLI: organize, tree, clean-empty, rename, dedupe')
  .version('1.0.0');

// Organize
program
  .command('organize [dir]')
  .description('Organize files by extension or date into subfolders')
  .option('--by <mode>', 'Group mode: ext or date', 'ext')
  .option('--recursive', 'Scan subdirectories recursively', false)
  .option('--dry-run', 'Show actions without moving files', false)
  .action(async (dir = '.', opts) => {
    const organize = require('../src/commands/organize');
    await organize(dir, opts);
  });

// Tree
program
  .command('tree [dir]')
  .description('Print a directory tree view')
  .option('--depth <n>', 'Max depth to display', (v) => parseInt(v, 10))
  .option('--dirs-first', 'List directories before files', false)
  .action(async (dir = '.', opts) => {
    const tree = require('../src/commands/tree');
    await tree(dir, opts);
  });

// Clean empty
program
  .command('clean-empty [dir]')
  .description('Remove empty directories recursively')
  .action(async (dir = '.', opts) => {
    const cleanEmpty = require('../src/commands/cleanEmpty');
    await cleanEmpty(dir, opts);
  });

// Rename
program
  .command('rename [dir]')
  .description('Batch rename files by regex pattern')
  .option('--path <dir>', 'Directory containing files to rename')
  .option('--match <regex>', 'Regex pattern to match in filename', '')
  .option('--replace <str>', 'Replacement string', '')
  .option('--ext <ext>', 'Only rename files with extension (e.g., .pdf)')
  .option('--dry-run', 'Show actions without renaming', false)
  .action(async (dir = '.', opts) => {
    const rename = require('../src/commands/rename');
    await rename(dir, opts);
  });

// Dedupe
program
  .command('dedupe [dir]')
  .description('Find duplicate files by hash; optionally delete duplicates')
  .option('--delete', 'Delete duplicates, keeping first occurrence', false)
  .option('--dry-run', 'Show planned deletions without deleting', false)
  .action(async (dir = '.', opts) => {
    const dedupe = require('../src/commands/dedupe');
    await dedupe(dir, opts);
  });

// Audit non-PDF files and check .pdf counterpart existence
program
  .command('audit-nonpdf [dir]')
  .description('List non-.pdf files with status of same-name .pdf presence')
  .option('--path <dir>', 'Directory to scan')
  .option('--recursive', 'Scan subdirectories recursively', false)
  .action(async (dir = '.', opts) => {
    const audit = require('../src/commands/auditNonPdf');
    await audit(dir, opts);
  });

// Excel to PDF (all sheets)
program
  .command('excel2pdf [dir]')
  .description('Convert all Excel files in a directory to PDF (all sheets)')
  .option('--path <dir>', 'Directory containing Excel files')
  .option('--recursive', 'Scan subdirectories recursively', false)
  .option('--out <dir>', 'Output directory for PDFs (default: alongside)')
  .option('--overwrite', 'Overwrite existing PDFs if present', false)
  .action(async (dir = '.', opts) => {
    const excel2pdf = require('../src/commands/excelToPdf');
    await excel2pdf(dir, opts);
  });

// Excel to PDF from list file
program
  .command('excel2pdf-list [list]')
  .description('Convert Excel files listed in a text file to PDF (all sheets)')
  .option('--list <file>', 'Text file containing paths (one per line)')
  .option('--out <dir>', 'Output directory for PDFs (default: alongside)')
  .option('--overwrite', 'Overwrite existing PDFs if present', false)
  .action(async (list = '', opts) => {
    const excel2pdfList = require('../src/commands/excelToPdfList');
    await excel2pdfList(list, opts);
  });

// Audit Excel workbooks with more than N sheets
program
  .command('excel-audit-sheets [dir]')
  .description('Find Excel files having more than N sheets (default N=1)')
  .option('--path <dir>', 'Directory to scan')
  .option('--recursive', 'Scan subdirectories recursively', false)
  .option('--min <n>', 'Minimum sheet count threshold', (v) => parseInt(v, 10))
  .option('--delete-pdf', 'Delete same-name .pdf alongside matched Excel files', false)
  .option('--dry-run', 'Preview deletions without removing files', false)
  .option('--to-pdf', 'Convert matched Excel to PDF alongside (all sheets)', false)
  .option('--overwrite', 'Overwrite existing PDFs if present (with --to-pdf)', false)
  .action(async (dir = '.', opts) => {
    const auditSheets = require('../src/commands/excelAuditSheets');
    await auditSheets(dir, opts);
  });

// Convert missing PDF for Word/Excel (alongside)
program
  .command('office2pdf-missing [dir]')
  .description('Convert Word/Excel files without same-name PDF to PDF (alongside)')
  .option('--path <dir>', 'Directory to scan')
  .option('--recursive', 'Scan subdirectories recursively', false)
  .option('--overwrite', 'Also convert when PDF exists (overwrite)', false)
  .option('--dry-run', 'Preview conversions without exporting', false)
  .action(async (dir = '.', opts) => {
    const officeMissing = require('../src/commands/officeToPdfMissing');
    await officeMissing(dir, opts);
  });

program.parseAsync();
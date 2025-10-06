const path = require('path');
const fs = require('fs');
const { listDir, ensureDir } = require('../utils');
const { spawnSync } = require('child_process');

const EXCEL_EXTS = new Set(['.xls', '.xlsx', '.xlsm', '.xlsb']);

function toPdfPath(inputPath, outDir) {
  const base = path.basename(inputPath, path.extname(inputPath));
  const destDir = outDir ? path.resolve(outDir) : path.dirname(inputPath);
  ensureDir(destDir);
  return path.join(destDir, base + '.pdf');
}

function exportExcelToPdfViaPowerShell(inputPath, outputPath) {
  // PowerShell script to use Excel COM to export entire workbook to PDF.
  const psScript = `
  $ErrorActionPreference = 'Stop'
  $xlFixedFormatType = 0 # xlTypePDF
  $excel = New-Object -ComObject Excel.Application
  try {
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open("${inputPath.replace(/"/g, '""')}")
    try {
      # Export entire workbook to PDF
      $wb.ExportAsFixedFormat($xlFixedFormatType, "${outputPath.replace(/"/g, '""')}")
    } finally {
      $wb.Close($false)
    }
  } finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  `;
  const res = spawnSync('powershell', ['-NoProfile', '-Command', psScript], { encoding: 'utf8' });
  if (res.status !== 0) {
    throw new Error(res.stderr || res.stdout || 'Excel export failed');
  }
}

module.exports = async function excelToPdf(dir, opts) {
  const selected = opts.path ? opts.path : dir;
  const targetDir = path.resolve(selected);
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('目录未找到:', targetDir);
    process.exitCode = 1;
    return;
  }

  const recursive = !!opts.recursive;
  const outDir = opts.out || null;
  const overwrite = !!opts.overwrite;

  const entries = listDir(targetDir, recursive).filter(e => !e.isDir);
  const excelFiles = entries.filter(e => EXCEL_EXTS.has(path.extname(e.name).toLowerCase()));

  let converted = 0;
  let skipped = 0;
  let failed = 0;

  for (const f of excelFiles) {
    const inputPath = f.path;
    const outputPath = toPdfPath(inputPath, outDir);
    if (!overwrite && fs.existsSync(outputPath)) {
      console.log('[skip] exists:', outputPath);
      skipped++;
      continue;
    }
    try {
      exportExcelToPdfViaPowerShell(inputPath, outputPath);
      console.log('[ok] ', inputPath, '=>', outputPath);
      converted++;
    } catch (e) {
      console.error('[fail]', inputPath, e.message);
      failed++;
    }
  }

  console.log('——');
  console.log(`总计Excel文件: ${excelFiles.length}`);
  console.log(`成功转换: ${converted}`);
  console.log(`已存在跳过: ${skipped}`);
  console.log(`失败: ${failed}`);
  console.log('提示: 该功能依赖 Windows 下已安装的 Microsoft Excel。');
};
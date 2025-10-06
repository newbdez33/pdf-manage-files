const path = require('path');
const fs = require('fs');
const { listDir, ensureDir } = require('../utils');
const { spawnSync } = require('child_process');

const WORD_EXTS = new Set(['.doc', '.docx', '.docm']);
const EXCEL_EXTS = new Set(['.xls', '.xlsx', '.xlsm', '.xlsb']);

function toPdfPathAlongside(inputPath) {
  const base = path.basename(inputPath, path.extname(inputPath));
  const destDir = path.dirname(inputPath);
  ensureDir(destDir);
  return path.join(destDir, base + '.pdf');
}

function exportExcelToPdfViaPowerShell(inputPath, outputPath) {
  const psScript = `
  $ErrorActionPreference = 'Stop'
  $xlFixedFormatType = 0 # xlTypePDF
  $excel = New-Object -ComObject Excel.Application
  try {
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open("${inputPath.replace(/"/g, '""')}")
    try {
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

function exportWordToPdfViaPowerShell(inputPath, outputPath) {
  const psScript = `
  $ErrorActionPreference = 'Stop'
  $wdExportFormatPDF = 17
  $word = New-Object -ComObject Word.Application
  try {
    $word.Visible = $false
    $doc = $word.Documents.Open("${inputPath.replace(/"/g, '""')}")
    try {
      $doc.ExportAsFixedFormat("${outputPath.replace(/"/g, '""')}", $wdExportFormatPDF)
    } finally {
      $doc.Close($false)
    }
  } finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  }
  `;
  const res = spawnSync('powershell', ['-NoProfile', '-Command', psScript], { encoding: 'utf8' });
  if (res.status !== 0) {
    throw new Error(res.stderr || res.stdout || 'Word export failed');
  }
}

module.exports = async function officeToPdfMissing(dirArg, opts) {
  const selected = opts.path ? opts.path : dirArg;
  const targetDir = path.resolve(selected || '.');
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('目录未找到:', targetDir);
    process.exitCode = 1;
    return;
  }

  const recursive = !!opts.recursive;
  const overwrite = !!opts.overwrite;
  const dryRun = !!opts.dryRun;

  const entries = listDir(targetDir, recursive).filter(e => !e.isDir);
  const officeFiles = entries.filter(e => {
    const ext = path.extname(e.name).toLowerCase();
    return WORD_EXTS.has(ext) || EXCEL_EXTS.has(ext);
  });

  let candidates = 0;
  let converted = 0;
  let skippedExist = 0;
  let failed = 0;

  for (const f of officeFiles) {
    const ext = path.extname(f.name).toLowerCase();
    const base = path.basename(f.name, ext);
    const pdfPath = path.join(f.dir, `${base}.pdf`);

    const hasPdf = fs.existsSync(pdfPath);
    if (hasPdf && !overwrite) {
      skippedExist++;
      continue;
    }

    candidates++;
    if (dryRun) {
      console.log('[dry-run] 转为PDF:', f.path, '=>', pdfPath);
      continue;
    }
    try {
      if (EXCEL_EXTS.has(ext)) {
        exportExcelToPdfViaPowerShell(f.path, pdfPath);
      } else if (WORD_EXTS.has(ext)) {
        exportWordToPdfViaPowerShell(f.path, pdfPath);
      } else {
        continue;
      }
      console.log('[ok] ', f.path, '=>', pdfPath);
      converted++;
    } catch (e) {
      console.error('[fail]', f.path, e.message);
      failed++;
    }
  }

  console.log('——');
  console.log(`Office文件总计: ${officeFiles.length}`);
  console.log(`待处理(缺PDF或允许覆盖): ${candidates}`);
  console.log(`成功转换: ${converted}${dryRun ? ' (dry-run)' : ''}`);
  console.log(`已存在跳过: ${skippedExist}`);
  console.log(`失败: ${failed}`);
  console.log('提示: 该功能依赖 Windows 下已安装的 Microsoft Word 与 Microsoft Excel。');
}
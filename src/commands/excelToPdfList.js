const path = require('path');
const fs = require('fs');
const { ensureDir } = require('../utils');
const { spawnSync } = require('child_process');

const EXCEL_EXTS = new Set(['.xls', '.xlsx', '.xlsm', '.xlsb']);

function toPdfPath(inputPath, outDir) {
  const base = path.basename(inputPath, path.extname(inputPath));
  const destDir = outDir ? path.resolve(outDir) : path.dirname(inputPath);
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

function parseListFile(listFilePath) {
  const raw = fs.readFileSync(listFilePath, 'utf8');
  const lines = raw.split(/\r?\n/);
  const out = [];
  for (let line of lines) {
    line = line.trim();
    if (!line || line.startsWith('#')) continue;
    // Strip surrounding quotes if present
    if ((line.startsWith('"') && line.endsWith('"')) || (line.startsWith("'") && line.endsWith("'"))) {
      line = line.slice(1, -1);
    }
    out.push(line);
  }
  return out;
}

module.exports = async function excelToPdfList(listArg, opts) {
  const listFile = opts.list ? opts.list : listArg;
  if (!listFile) {
    console.error('请提供文件清单: 使用 --list <file> 或位置参数 [list]');
    process.exitCode = 1;
    return;
  }
  const listPath = path.resolve(listFile);
  if (!fs.existsSync(listPath) || !fs.statSync(listPath).isFile()) {
    console.error('清单文件未找到:', listPath);
    process.exitCode = 1;
    return;
  }

  const outDir = opts.out || null;
  const overwrite = !!opts.overwrite;

  const items = parseListFile(listPath);
  const seen = new Set();
  const files = [];
  for (const item of items) {
    const full = path.resolve(item);
    if (seen.has(full)) continue;
    seen.add(full);
    if (!fs.existsSync(full) || !fs.statSync(full).isFile()) {
      console.log('[skip] 不存在或不是文件:', full);
      continue;
    }
    const ext = path.extname(full).toLowerCase();
    if (!EXCEL_EXTS.has(ext)) {
      console.log('[skip] 非Excel文件:', full);
      continue;
    }
    files.push(full);
  }

  let converted = 0;
  let skipped = 0;
  let failed = 0;

  for (const inputPath of files) {
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
  console.log(`清单总计条目: ${items.length}`);
  console.log(`有效Excel文件: ${files.length}`);
  console.log(`成功转换: ${converted}`);
  console.log(`已存在跳过: ${skipped}`);
  console.log(`失败: ${failed}`);
  console.log('提示: 该功能依赖 Windows 下已安装的 Microsoft Excel。');
};
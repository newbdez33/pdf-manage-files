const path = require('path');
const fs = require('fs');
const { listDir, ensureDir } = require('../utils');
const { spawnSync } = require('child_process');

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

function getSheetCountViaPowerShell(inputPath) {
  const psScript = `
  $ErrorActionPreference = 'Stop'
  $excel = New-Object -ComObject Excel.Application
  try {
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open("${inputPath.replace(/"/g, '""')}")
    try {
      $count = $wb.Worksheets.Count
      Write-Output $count
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
    throw new Error(res.stderr || res.stdout || 'Excel sheet count failed');
  }
  const out = (res.stdout || '').trim();
  const num = parseInt(out, 10);
  if (Number.isNaN(num)) throw new Error('Invalid count: ' + out);
  return num;
}

module.exports = async function excelAuditSheets(dir, opts) {
  const selected = opts.path ? opts.path : dir;
  const targetDir = path.resolve(selected);
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('目录未找到:', targetDir);
    process.exitCode = 1;
    return;
  }

  const recursive = !!opts.recursive;
  const min = (opts.min !== undefined && !Number.isNaN(opts.min)) ? Number(opts.min) : 3; // 默认>2
  const deletePdf = !!opts.deletePdf;
  const dryRun = !!opts.dryRun;
  const toPdf = !!opts.toPdf;
  const overwrite = !!opts.overwrite;

  const entries = listDir(targetDir, recursive).filter(e => !e.isDir);
  const excelFiles = entries.filter(e => EXCEL_EXTS.has(path.extname(e.name).toLowerCase()));

  let total = excelFiles.length;
  let matched = 0;
  let failed = 0;
  let deleted = 0;
  let converted = 0;
  let convertSkipped = 0;

  for (const f of excelFiles) {
    try {
      const count = getSheetCountViaPowerShell(f.path);
      if (count >= min) {
        matched++;
        console.log(`${f.path} | 工作表数量: ${count}`);
        if (toPdf) {
          const outPdf = toPdfPathAlongside(f.path);
          if (!overwrite && fs.existsSync(outPdf)) {
            console.log('[skip] PDF已存在:', outPdf);
            convertSkipped++;
          } else {
            if (dryRun) {
              console.log('[dry-run] 转为PDF:', f.path, '=>', outPdf);
            } else {
              try {
                exportExcelToPdfViaPowerShell(f.path, outPdf);
                console.log('[ok] 转为PDF:', f.path, '=>', outPdf);
                converted++;
              } catch (e3) {
                console.error('[fail-convert]', f.path, e3.message);
              }
            }
          }
        }
        if (deletePdf) {
          const ext = path.extname(f.name);
          const base = path.basename(f.name, ext);
          const pdfPath = path.join(f.dir, `${base}.pdf`);
          if (fs.existsSync(pdfPath)) {
            if (dryRun) {
              console.log(`[dry-run] 删除PDF: ${pdfPath}`);
            } else {
              try {
                fs.unlinkSync(pdfPath);
                deleted++;
                console.log(`[deleted] ${pdfPath}`);
              } catch (e2) {
                console.error('[fail-delete]', pdfPath, e2.message);
              }
            }
          } else {
            console.log(`[skip] 无同名PDF: ${pdfPath}`);
          }
        }
      }
    } catch (e) {
      failed++;
      console.error('[fail]', f.path, e.message);
    }
  }

  console.log('——');
  console.log(`总计Excel文件: ${total}`);
  console.log(`满足(>=${min})的数量: ${matched}`);
  console.log(`失败: ${failed}`);
  if (deletePdf) {
    console.log(`删除的PDF数量: ${deleted}${dryRun ? ' (dry-run)' : ''}`);
  }
  if (toPdf) {
    console.log(`成功转换为PDF: ${converted}${dryRun ? ' (dry-run)' : ''}`);
    console.log(`转换跳过(已存在): ${convertSkipped}`);
  }
  console.log('提示: 该功能依赖 Windows 下已安装的 Microsoft Excel。');
};
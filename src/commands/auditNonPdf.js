const path = require('path');
const fs = require('fs');
const { listDir } = require('../utils');

module.exports = async function auditNonPdf(dir, opts) {
  const selected = opts.path ? opts.path : dir;
  const targetDir = path.resolve(selected);
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('目录未找到:', targetDir);
    process.exitCode = 1;
    return;
  }

  const recursive = !!opts.recursive;
  const entries = listDir(targetDir, recursive).filter(e => !e.isDir);

  let total = 0;
  let withPdfCounterpart = 0;

  for (const f of entries) {
    const ext = path.extname(f.name).toLowerCase();
    if (ext === '.pdf') continue; // 仅统计非 .pdf 文件
    total++;
    const base = path.basename(f.name, ext);
    const pdfCandidate = path.join(f.dir, `${base}.pdf`);
    const exists = fs.existsSync(pdfCandidate);
    if (exists) withPdfCounterpart++;
    const status = exists ? '有' : '无';
    console.log(`${f.path} | 同名PDF: ${status}`);
  }

  console.log('——');
  console.log(`非PDF文件数量: ${total}`);
  console.log(`其中存在同名PDF的数量: ${withPdfCounterpart}`);
};
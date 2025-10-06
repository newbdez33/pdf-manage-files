const path = require('path');
const fs = require('fs');
const { listDir, moveFile, ensureDir } = require('../utils');

module.exports = async function organize(dir, opts) {
  const targetDir = path.resolve(dir);
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('Directory not found:', targetDir);
    process.exitCode = 1;
    return;
  }

  const recursive = !!opts.recursive;
  const by = (opts.by || 'ext').toLowerCase();
  const items = listDir(targetDir, recursive).filter(e => !e.isDir);

  const organizedRoot = path.join(targetDir, 'organized');
  ensureDir(organizedRoot);

  let moved = 0;
  for (const f of items) {
    // Skip files already inside organized
    if (f.path.startsWith(organizedRoot)) continue;

    let dest;
    if (by === 'date') {
      const dt = f.stats.mtime;
      const year = String(dt.getFullYear());
      const month = String(dt.getMonth() + 1).padStart(2, '0');
      dest = path.join(organizedRoot, 'by-date', year, month, f.name);
    } else {
      const ext = path.extname(f.name).slice(1) || 'noext';
      dest = path.join(organizedRoot, 'by-ext', ext, f.name);
    }

    if (opts.dryRun) {
      console.log('[dry-run] move', f.path, '=>', dest);
    } else {
      try {
        moveFile(f.path, dest);
        moved++;
      } catch (e) {
        console.error('Failed to move:', f.path, e.message);
      }
    }
  }

  if (!opts.dryRun) console.log('Moved files:', moved);
};
const path = require('path');
const fs = require('fs');

module.exports = async function rename(dir, opts) {
  const selected = opts.path ? opts.path : dir;
  const targetDir = path.resolve(selected);
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('Directory not found:', targetDir);
    process.exitCode = 1;
    return;
  }

  const pattern = opts.match ? new RegExp(opts.match) : null;
  const replace = opts.replace || '';
  const filterExt = opts.ext || null;

  const entries = fs.readdirSync(targetDir, { withFileTypes: true });
  let renamed = 0;
  for (const ent of entries) {
    if (!ent.isFile()) continue;
    const oldPath = path.join(targetDir, ent.name);
    const ext = path.extname(ent.name);
    if (filterExt && ext !== filterExt) continue;
    let newName = ent.name;
    if (pattern) newName = newName.replace(pattern, replace);
    if (newName !== ent.name) {
      const newPath = path.join(targetDir, newName);
      if (opts.dryRun) {
        console.log('[dry-run] rename', oldPath, '=>', newPath);
      } else {
        try {
          fs.renameSync(oldPath, newPath);
          renamed++;
        } catch (e) {
          console.error('Failed to rename:', oldPath, e.message);
        }
      }
    }
  }
  if (!opts.dryRun) console.log('Renamed files:', renamed);
};
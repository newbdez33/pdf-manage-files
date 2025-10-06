const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

function hashFile(filePath) {
  return new Promise((resolve, reject) => {
    const hash = crypto.createHash('sha256');
    const stream = fs.createReadStream(filePath);
    stream.on('error', reject);
    stream.on('data', (d) => hash.update(d));
    stream.on('end', () => resolve(hash.digest('hex')));
  });
}

module.exports = async function dedupe(dir, opts) {
  const targetDir = path.resolve(dir);
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('Directory not found:', targetDir);
    process.exitCode = 1;
    return;
  }

  const entries = fs.readdirSync(targetDir, { withFileTypes: true })
    .filter(e => e.isFile());

  const map = new Map();
  for (const ent of entries) {
    const filePath = path.join(targetDir, ent.name);
    try {
      const hash = await hashFile(filePath);
      if (!map.has(hash)) map.set(hash, []);
      map.get(hash).push(filePath);
    } catch (e) {
      console.error('Hash failed:', filePath, e.message);
    }
  }

  let dupCount = 0;
  for (const [hash, files] of map.entries()) {
    if (files.length > 1) {
      console.log('Duplicate group (sha256=' + hash + '):');
      files.forEach((f, i) => console.log('  ' + (i === 0 ? '[keep] ' : '[dup]  ') + f));
      const toDelete = files.slice(1);
      if (opts.delete) {
        for (const f of toDelete) {
          if (opts.dryRun) {
            console.log('[dry-run] delete', f);
          } else {
            try {
              fs.unlinkSync(f);
              dupCount++;
            } catch (e) {
              console.error('Delete failed:', f, e.message);
            }
          }
        }
      }
    }
  }

  if (opts.delete && !opts.dryRun) console.log('Deleted duplicates:', dupCount);
};
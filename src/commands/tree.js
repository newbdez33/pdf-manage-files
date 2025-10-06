const path = require('path');
const fs = require('fs');

function printTree(dir, { depth, dirsFirst }) {
  function walk(current, level) {
    if (depth !== undefined && level > depth) return;
    let entries;
    try {
      entries = fs.readdirSync(current, { withFileTypes: true });
    } catch (e) {
      console.error('Cannot read:', current, e.message);
      return;
    }
    if (dirsFirst) {
      entries.sort((a, b) => (b.isDirectory() - a.isDirectory()) || a.name.localeCompare(b.name));
    } else {
      entries.sort((a, b) => a.name.localeCompare(b.name));
    }
    for (const ent of entries) {
      const prefix = '  '.repeat(level) + (ent.isDirectory() ? '📁 ' : '📄 ');
      console.log(prefix + ent.name);
      if (ent.isDirectory()) {
        walk(path.join(current, ent.name), level + 1);
      }
    }
  }
  console.log('📂', path.basename(dir));
  walk(dir, 1);
}

module.exports = async function tree(dir, opts) {
  const targetDir = path.resolve(dir);
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('Directory not found:', targetDir);
    process.exitCode = 1;
    return;
  }
  printTree(targetDir, { depth: opts.depth, dirsFirst: !!opts.dirsFirst });
};
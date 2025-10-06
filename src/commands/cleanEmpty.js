const path = require('path');
const fs = require('fs');

function removeEmptyDirs(dir) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const ent of entries) {
    if (ent.isDirectory()) {
      const full = path.join(dir, ent.name);
      removeEmptyDirs(full);
    }
  }
  // After processing children, check if empty
  const remaining = fs.readdirSync(dir);
  if (remaining.length === 0) {
    try {
      fs.rmdirSync(dir);
      console.log('Removed empty dir:', dir);
    } catch (e) {
      // ignore
    }
  }
}

module.exports = async function cleanEmpty(dir) {
  const targetDir = path.resolve(dir);
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('Directory not found:', targetDir);
    process.exitCode = 1;
    return;
  }
  removeEmptyDirs(targetDir);
};
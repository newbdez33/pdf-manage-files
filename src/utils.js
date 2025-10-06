const fs = require('fs');
const path = require('path');

function isDirectory(p) {
  try {
    return fs.statSync(p).isDirectory();
  } catch {
    return false;
  }
}

function listDir(dir, recursive = false) {
  const results = [];
  const stack = [dir];
  while (stack.length) {
    const current = stack.pop();
    const items = fs.readdirSync(current);
    for (const name of items) {
      const full = path.join(current, name);
      let stats;
      try {
        stats = fs.statSync(full);
      } catch {
        continue;
      }
      const entry = { path: full, name, dir: current, isDir: stats.isDirectory(), stats };
      results.push(entry);
      if (recursive && entry.isDir) {
        stack.push(full);
      }
    }
  }
  return results;
}

function ensureDir(p) {
  fs.mkdirSync(p, { recursive: true });
}

function moveFile(src, dest) {
  ensureDir(path.dirname(dest));
  fs.renameSync(src, dest);
}

function formatBytes(bytes) {
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  if (bytes === 0) return '0 B';
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return `${(bytes / Math.pow(1024, i)).toFixed(1)} ${sizes[i]}`;
}

module.exports = { isDirectory, listDir, ensureDir, moveFile, formatBytes };
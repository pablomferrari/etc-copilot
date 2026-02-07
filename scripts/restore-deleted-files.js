/**
 * Restores files in node_modules that were renamed by antivirus (e.g. index.js -> index.js.DELETE.xxxx).
 * Run after npm install if you see "Cannot find module" for files that exist as *.DELETE.*
 */
const fs = require('fs');
const path = require('path');

const nodeModules = path.join(__dirname, '..', 'node_modules');
const deletePattern = /^(.+)\.DELETE\.[a-f0-9]+$/;

function walkDir(dir) {
  if (!fs.existsSync(dir)) return;
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const e of entries) {
    const full = path.join(dir, e.name);
    if (e.isDirectory() && e.name !== '.bin') {
      walkDir(full);
    } else if (e.isFile() && deletePattern.test(e.name)) {
      const target = path.join(dir, e.name.replace(deletePattern, '$1'));
      if (!fs.existsSync(target)) {
        try {
          fs.copyFileSync(full, target);
          console.log('Restored:', path.relative(nodeModules, target));
        } catch (err) {
          console.warn('Failed to restore', target, err.message);
        }
      }
    }
  }
}

walkDir(nodeModules);
console.log('Done checking for antivirus-renamed files.');

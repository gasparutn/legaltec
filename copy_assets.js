const fs = require('fs');
const path = require('path');
const srcDir = path.join(__dirname, 'src');
const distDir = path.join(__dirname, 'dist');

if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir);
}

const files = fs.readdirSync(srcDir);
for (const file of files) {
  if (file.endsWith('.html') || file === 'appsscript.json') {
    fs.copyFileSync(path.join(srcDir, file), path.join(distDir, file));
  }
}
console.log('Assets copied to dist/');

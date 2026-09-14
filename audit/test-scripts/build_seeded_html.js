const fs = require('fs');
const path = require('path');
const N = process.argv[2];
const template = fs.readFileSync(path.join(__dirname, 'app.html'), 'utf8');
const payload = fs.readFileSync(path.join(__dirname, 'out', `synthetic_${N}.json`), 'utf8');
const out = template.replace(
  /<script type="application\/json" id="library-data">.*?<\/script>/s,
  `<script type="application/json" id="library-data">${payload.replace(/<\//g,'<\\/')}</script>`
);
const outPath = path.join(__dirname, 'out', `seeded_${N}.html`);
fs.writeFileSync(outPath, out);
console.log('Wrote', outPath, 'size MB:', (fs.statSync(outPath).size/1024/1024).toFixed(2));

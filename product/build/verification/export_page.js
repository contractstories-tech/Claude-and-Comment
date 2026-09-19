// Extracts the exact JavaScript that CLExport.bas writes into the exported
// index page, and runs it in a real DOM against realistic clause content.
//   npm install jsdom && node export_page.js
const fs = require('fs'), path = require('path'), { JSDOM } = require('jsdom');

const src = fs.readFileSync(path.join(__dirname, '..', '..', 'word', 'CLExport.bas'), 'utf8');
const body = src.slice(src.indexOf('Private Function CLIndexFoot'), src.indexOf('Private Function CLExportReadme'));
const js = (body.match(/"((?:[^"]|"")*)"/g) || [])
  .map(s => s.slice(1, -1).replace(/""/g, '"')).join('');
const script = js.slice(js.indexOf('<script>'));

const article = (title, meta, text) =>
  `<article><h2>${title}</h2><p class='meta'>${meta}</p><pre>${text}</pre></article>`;
const page = `<!doctype html><html><head><meta charset="utf-8"><title>t</title></head><body><main>
<label for='q'>Search</label><input id='q' type='search'><p id='status' role='status'></p>
${article('Mutual indemnity', 'Clause &middot; Liability &middot; indemnity', 'Each party shall indemnify the other against all Losses.')}
${article('Force majeure', 'Clause &middot; Risk', 'Neither party is liable for delay caused by a Force Majeure Event.')}
${article('Liability cap', 'Clause &middot; Liability &middot; cap', 'Total liability shall not exceed the Charges paid.')}
${script}`;

const dom = new JSDOM(page, { runScripts: 'dangerously' });
const d = dom.window.document, q = d.getElementById('q'), s = d.getElementById('status');
const visible = () => [...d.querySelectorAll('article')].filter(a => !a.hidden).map(a => a.querySelector('h2').textContent);
const run = v => { q.value = v; q.dispatchEvent(new dom.window.Event('input')); return { status: s.textContent, visible: visible() }; };

let failed = 0;
const check = (label, query, expected) => {
  const r = run(query), ok = JSON.stringify(r.visible) === JSON.stringify(expected);
  if (!ok) failed++;
  console.log(`  ${ok ? 'pass' : 'FAIL'}  ${label.padEnd(44)}${r.status.padEnd(24)}${JSON.stringify(r.visible)}`);
};

console.log(`initial: ${s.textContent} ${JSON.stringify(visible())}\n`);
check('empty search shows everything', '', ['Mutual indemnity', 'Force majeure', 'Liability cap']);
check('a quoted phrase works', '"force majeure"', ['Force majeure']);
check('the same phrase unquoted', 'force majeure', ['Force majeure']);
check('one word narrows', 'liability', ['Mutual indemnity', 'Liability cap']);
check('two words narrow further', 'liability cap', ['Liability cap']);
check('capitals make no difference', 'FORCE MAJEURE', ['Force majeure']);
check('the topic line is searched too', 'Risk', ['Force majeure']);
check('nothing matches', 'frustration', []);
console.log('\n' + (failed ? `${failed} FAILED` : 'all export-page checks passed'));
process.exit(failed ? 1 : 0);

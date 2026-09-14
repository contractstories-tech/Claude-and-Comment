const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const OUT = path.resolve(__dirname, 'out');

async function run(browser, n) {
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push(e.message));
  page.on('dialog', d => d.accept());
  await page.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });

  const filePath = 'file://' + path.join(OUT, `seeded_${n}.html`);
  const t0 = Date.now();
  await page.goto(filePath);
  await page.waitForSelector('.entryRow, .empty', { timeout: 60000 });
  const loadMs = Date.now() - t0;

  // search timing
  const t1 = Date.now();
  await page.fill('#search', 'liability');
  await page.waitForTimeout(200); // debounce is 120ms in app
  const searchMs = Date.now() - t1;
  const searchCount = await page.evaluate(() => document.getElementById('count').textContent);

  // clear + filter by category
  await page.click('#btnClearFilters');
  await page.waitForTimeout(50);
  const t2 = Date.now();
  await page.selectOption('#filterStatus', 'Preferred');
  await page.waitForTimeout(50);
  const filterMs = Date.now() - t2;
  const filterCount = await page.evaluate(() => document.getElementById('count').textContent);

  // click an entry (detail render)
  await page.click('#btnClearFilters');
  await page.waitForTimeout(50);
  const t3 = Date.now();
  await page.click('.entryRow');
  await page.waitForTimeout(50);
  const detailMs = Date.now() - t3;

  // Save As (download) timing
  const t4 = Date.now();
  const [dl] = await Promise.all([
    page.waitForEvent('download', { timeout: 60000 }),
    page.click('#btnSaveAs'),
  ]);
  const savePath = path.join(OUT, `resaved_${n}.html`);
  await dl.saveAs(savePath);
  const saveMs = Date.now() - t4;
  const savedSize = fs.statSync(savePath).size;

  await context.close();
  return { n, loadMs, searchMs, searchCount, filterMs, filterCount, detailMs, saveMs, savedSizeMB: (savedSize/1024/1024).toFixed(2), errs };
}

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const results = [];
  for (const n of [500, 2000, 5000]) {
    console.log('Running perf test for N=' + n);
    const r = await run(browser, n);
    results.push(r);
    console.log(JSON.stringify(r, null, 2));
  }
  fs.writeFileSync(path.join(OUT, 'perf_results.json'), JSON.stringify(results, null, 2));
  await browser.close();
})();

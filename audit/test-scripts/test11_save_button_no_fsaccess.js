const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const OUT = path.resolve(__dirname, 'out');
const appPath = path.resolve(__dirname, 'app.html');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();
  page.on('dialog', d => d.accept());
  await page.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
  await page.goto('file://' + appPath);
  await page.waitForTimeout(200);

  await page.click('#btnNew');
  await page.fill('.titleInput', 'Save Button Test v1');
  await page.waitForTimeout(150);

  // Click the main Save button (not Save As) - this is what a Firefox/Safari user would normally use
  const fileStatusBefore = await page.evaluate(() => document.getElementById('fileStatus').textContent);
  const [dl1] = await Promise.all([page.waitForEvent('download'), page.click('#btnSave')]);
  const p1 = path.join(OUT, 'save1_' + Date.now() + '.html');
  await dl1.saveAs(p1);
  const fileStatusAfter1 = await page.evaluate(() => ({
    text: document.getElementById('fileStatus').textContent,
    dirty: document.getElementById('fileStatus').classList.contains('dirty'),
  }));

  // Make another edit and click Save again
  await page.fill('.titleInput', 'Save Button Test v2 EDITED');
  await page.waitForTimeout(150);
  const [dl2] = await Promise.all([page.waitForEvent('download'), page.click('#btnSave')]);
  const p2 = path.join(OUT, 'save2_' + Date.now() + '.html');
  await dl2.saveAs(p2);
  const fileStatusAfter2 = await page.evaluate(() => ({
    text: document.getElementById('fileStatus').textContent,
    dirty: document.getElementById('fileStatus').classList.contains('dirty'),
  }));

  function extractPayload(htmlPath) {
    const html = fs.readFileSync(htmlPath, 'utf8');
    const m = html.match(/<script type="application\/json" id="library-data">(.*?)<\/script>/s);
    return JSON.parse(m[1]);
  }
  const d1 = extractPayload(p1);
  const d2 = extractPayload(p2);

  console.log(JSON.stringify({
    fileStatusBefore, fileStatusAfter1, fileStatusAfter2,
    file1: { revision: d1.revision, savedAt: d1.savedAt, title: d1.entries[0].title },
    file2: { revision: d2.revision, savedAt: d2.savedAt, title: d2.entries[0].title },
    SAME_REVISION_DIFFERENT_CONTENT: d1.revision === d2.revision && d1.entries[0].title !== d2.entries[0].title,
  }, null, 2));

  await browser.close();
})();

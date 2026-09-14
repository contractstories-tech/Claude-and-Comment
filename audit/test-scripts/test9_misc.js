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
  const errs = [];
  page.on('pageerror', e => errs.push(e.message));
  await page.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
  await page.goto('file://' + appPath);
  await page.waitForTimeout(200);

  // Inject an entry with rich XML directly into internal state via a JSONL-style inbox import (simulate Word capture)
  const q = '"';
  const line = JSON.stringify({
    app: 'ClauseCommentsLibraryInbox', recordVersion: 1, libraryId: '',
    id: 'wtest1', title: 'Rich Clause From Word', type: 'Clause', category: 'Other', status: 'Unclassified',
    position: 'none', favorite: false, clauseText: 'Original rich text.', commentText: '', contextText: '',
    wordOpenXml: '<w:document>fake xml</w:document>', richSourceText: 'Original rich text.',
    tags: [], source: 'Test.docx', needsReview: true, entryRevision: 1,
    createdAt: new Date().toISOString(), updatedAt: new Date().toISOString()
  });
  const jsonlPath = path.join(OUT, 'inbox_sim.jsonl');
  fs.writeFileSync(jsonlPath, line + '\n');
  await page.setInputFiles('#importInput', jsonlPath);
  await page.waitForTimeout(200);
  await page.click('#importMerge');
  await page.waitForTimeout(200);

  await page.click('#btnClearFilters');
  await page.waitForTimeout(100);
  // select the rich entry
  const rows = await page.$$('.entryRow');
  let clicked = false;
  for (const r of rows) {
    const t = await r.textContent();
    if (t.includes('Rich Clause From Word')) { await r.click(); clicked = true; break; }
  }
  await page.waitForTimeout(150);

  const freshBanner = await page.evaluate(() => {
    const banners = [...document.querySelectorAll('.reviewBanner')].map(b=>b.textContent);
    return banners;
  });

  // now edit clause text -> should go stale
  await page.fill('.clauseTA', 'Original rich text. EDITED');
  await page.waitForTimeout(150);
  await page.click('.rowTop .rowTitle'); // trigger blur alternative
  await page.locator('.clauseTA').blur();
  await page.waitForTimeout(150);

  const staleBanner = await page.evaluate(() => {
    const banners = [...document.querySelectorAll('.reviewBanner')].map(b=>b.textContent);
    const pills = [...document.querySelectorAll('.pill')].map(p=>p.textContent);
    return { banners, pills };
  });

  // keyboard shortcuts
  await page.keyboard.press('Control+f');
  const searchFocused = await page.evaluate(() => document.activeElement.id === 'search');
  await page.keyboard.press('Control+Shift+N');
  await page.waitForTimeout(100);
  const newEntryFocused = await page.evaluate(() => document.activeElement.classList.contains('titleInput'));

  console.log(JSON.stringify({ clicked, freshBanner, staleBanner, searchFocused, newEntryFocused, errs }, null, 2));
  await browser.close();
})();

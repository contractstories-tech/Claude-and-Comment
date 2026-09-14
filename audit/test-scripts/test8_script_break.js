const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const OUT = path.resolve(__dirname, 'out');
const appPath = path.resolve(__dirname, 'app.html');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();
  const dialogs = [];
  page.on('dialog', d => { dialogs.push(d.message()); d.accept(); });
  const pageErrors = [];
  page.on('pageerror', e => pageErrors.push(e.message));
  await page.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
  await page.goto('file://' + appPath);
  await page.waitForTimeout(200);

  await page.click('#btnNew');
  await page.fill('.titleInput', 'Script Break Test');
  await page.fill('.clauseTA', 'payload </script><script>window.__pwned=true;</script> more text');
  await page.waitForTimeout(150);

  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnSaveAs')]);
  const savedPath = path.join(OUT, 'scriptbreak.html');
  await dl.saveAs(savedPath);

  // Reopen it — if the escaping failed, the page would either fail to parse JSON, or __pwned would be set (script executed)
  const context2 = await browser.newContext();
  const page2 = await context2.newPage();
  const pageErrors2 = [];
  page2.on('pageerror', e => pageErrors2.push(e.message));
  await page2.goto('file://' + savedPath);
  await page2.waitForTimeout(300);
  const pwned = await page2.evaluate(() => window.__pwned === true);
  const entryCount = await page2.evaluate(() => document.querySelectorAll('.entryRow').length);
  const clauseTextRestored = await page2.evaluate(() => {
    const tag = document.getElementById('library-data');
    const d = JSON.parse(tag.textContent);
    return d.entries[0] ? d.entries[0].clauseText : null;
  });

  console.log(JSON.stringify({ pwned, entryCount, clauseTextRestored, pageErrors, pageErrors2, dialogs }, null, 2));
  await browser.close();
})();

const { chromium } = require('playwright');
const path = require('path');
const appPath = path.resolve(__dirname, 'app.html');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const context = await browser.newContext();
  const page = await context.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push(e.message));
  page.on('dialog', d => d.accept());
  await page.addInitScript(() => {
    delete window.showOpenFilePicker; delete window.showSaveFilePicker;
    Object.defineProperty(window, 'indexedDB', { value: undefined, configurable: true });
  });
  await page.goto('file://' + appPath);
  await page.waitForTimeout(300);
  await page.click('#btnNew');
  await page.fill('.titleInput', 'No-IndexedDB Test Entry');
  await page.waitForTimeout(600);
  const recoveryStatus1 = await page.evaluate(() => document.getElementById('recoveryStatus').textContent);

  await page.reload();
  await page.waitForTimeout(400);
  const after = await page.evaluate(() => ({
    count: document.querySelectorAll('.entryRow').length,
    recoveryStatus: document.getElementById('recoveryStatus').textContent,
  }));
  console.log(JSON.stringify({ recoveryStatus1, after, errs }, null, 2));
  await browser.close();
})();

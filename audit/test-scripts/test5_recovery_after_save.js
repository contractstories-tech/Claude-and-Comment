const { chromium } = require('playwright');
const path = require('path');
const savedHtmlPath = path.join(__dirname, 'out', 'clause-comments-library-copy.html');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const context = await browser.newContext();
  const page = await context.newPage();
  page.on('dialog', d => d.accept());

  await page.goto('file://' + savedHtmlPath);
  await page.waitForTimeout(300);
  await page.click('#btnNew');
  await page.fill('.titleInput', 'Post-Save Unsaved Addition');
  await page.waitForTimeout(600);

  await page.reload();
  await page.waitForTimeout(500);

  const finalState = await page.evaluate(() => ({
    count: document.querySelectorAll('.entryRow').length,
    recoveryStatus: document.getElementById('recoveryStatus').textContent,
    titles: [...document.querySelectorAll('.rowTitle')].map(e=>e.textContent),
  }));
  console.log('After reload on a PREVIOUSLY-SAVED file with new unsaved edit:', JSON.stringify(finalState, null, 2));

  await browser.close();
})();

const { chromium } = require('playwright');
const path = require('path');
const appPath = path.resolve(__dirname, 'app.html');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const context = await browser.newContext();
  const page = await context.newPage();
  page.on('dialog', d => d.accept());

  await page.goto('file://' + appPath);
  await page.waitForTimeout(300);
  const libId1 = await page.evaluate(() => window.__debugLibId || null);
  // libraryId isn't exposed globally; read via fileStatus won't show it. Instead inspect IndexedDB keys directly.
  await page.click('#btnNew');
  await page.fill('.titleInput', 'Crash Test Entry');
  await page.waitForTimeout(600);

  const idbKeysBefore = await page.evaluate(async () => {
    return await new Promise(resolve => {
      const req = indexedDB.open('ClauseCommentsLibraryDB', 1);
      req.onsuccess = () => {
        const db = req.result;
        const tx = db.transaction('kv');
        const store = tx.objectStore('kv');
        const keysReq = store.getAllKeys();
        keysReq.onsuccess = () => resolve(keysReq.result);
      };
    });
  });
  console.log('IDB keys BEFORE reload:', idbKeysBefore);

  await page.reload();
  await page.waitForTimeout(400);

  const idbKeysAfter = await page.evaluate(async () => {
    return await new Promise(resolve => {
      const req = indexedDB.open('ClauseCommentsLibraryDB', 1);
      req.onsuccess = () => {
        const db = req.result;
        const tx = db.transaction('kv');
        const store = tx.objectStore('kv');
        const keysReq = store.getAllKeys();
        keysReq.onsuccess = () => resolve(keysReq.result);
      };
    });
  });
  console.log('IDB keys AFTER reload:', idbKeysAfter);

  const finalState = await page.evaluate(() => ({
    count: document.querySelectorAll('.entryRow').length,
    recoveryStatus: document.getElementById('recoveryStatus').textContent,
  }));
  console.log('Final state after reload:', finalState);

  await browser.close();
})();

const { chromium } = require('playwright');
const path = require('path');
const OUT = path.resolve(__dirname, 'out');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  for (const n of [500, 2000, 5000]) {
    const context = await browser.newContext();
    const page = await context.newPage();
    await page.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
    await page.goto('file://' + path.join(OUT, `seeded_${n}.html`));
    await page.waitForSelector('.entryRow');
    // buildWordCache is inside an IIFE closure, not global. Reconstruct cache size by
    // reading the app's internal `data` indirectly isn't possible from outside; instead
    // approximate by calling the exact same logic used by connectWordCache: we can't call
    // private fn directly, so we simulate what a "Sync" would produce by checking the
    // JSON payload size as an upper-bound proxy plus base64 overhead (~33%) already included in wordOpenXml text.
    const approx = await page.evaluate(() => {
      const tag = document.getElementById('library-data');
      const data = JSON.parse(tag.textContent);
      let plainChars = 0, xmlChars = 0, entries = data.entries.length;
      for (const e of data.entries) {
        plainChars += (e.title||'').length + (e.clauseText||'').length + (e.commentText||'').length + (e.contextText||'').length;
        xmlChars += (e.wordOpenXml||'').length;
      }
      return { entries, plainChars, xmlChars };
    });
    console.log(n, JSON.stringify(approx));
    await context.close();
  }
  await browser.close();
})();

const { chromium } = require('playwright');
const path = require('path');

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' , args:['--enable-features=FileSystemAccessAPI']});
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();
  const consoleMsgs = [];
  page.on('console', msg => consoleMsgs.push(`[${msg.type()}] ${msg.text()}`));
  page.on('pageerror', err => consoleMsgs.push(`[pageerror] ${err.message}`));
  page.on('requestfailed', req => consoleMsgs.push(`[requestfailed] ${req.url()} ${req.failure()?.errorText}`));
  page.on('request', req => consoleMsgs.push(`[request] ${req.method()} ${req.url()}`));

  const filePath = 'file://' + path.resolve(__dirname, 'app.html');
  await page.goto(filePath);
  await page.waitForTimeout(500);

  const apiCheck = await page.evaluate(() => ({
    showOpenFilePicker: typeof window.showOpenFilePicker,
    showSaveFilePicker: typeof window.showSaveFilePicker,
    isSecureContext: window.isSecureContext,
    indexedDB: typeof window.indexedDB,
    origin: location.origin,
    protocol: location.protocol,
  }));
  console.log('API CHECK:', JSON.stringify(apiCheck, null, 2));

  console.log('--- Console/network messages after load ---');
  consoleMsgs.filter(m=>!m.startsWith('[request]')).forEach(m => console.log(m));

  // Check network requests (should be zero external)
  const externalReqs = consoleMsgs.filter(m => m.startsWith('[request]') && !m.includes('file://'));
  console.log('EXTERNAL REQUESTS:', externalReqs.length, externalReqs);

  await page.screenshot({ path: path.resolve(__dirname, 'shot_initial.png') });

  await browser.close();
})();

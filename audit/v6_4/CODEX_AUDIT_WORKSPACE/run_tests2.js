const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const BASE = 'http://127.0.0.1:8934';
const results = [];
function record(name, pass, detail) { results.push({ name, pass, detail }); console.log((pass ? 'PASS' : 'FAIL') + ' - ' + name + (detail ? ('  :: ' + detail) : '')); }
async function withDialogs(page, textAnswer) {
  page.on('dialog', async d => {
    try { if (d.type() === 'prompt') await d.accept(textAnswer || 'K'); else await d.accept(); } catch (e) {}
  });
}

async function testHealthCheckAndFamily() {
  const browser = await chromium.launch();
  const page = await (await browser.newContext()).newPage();
  const alerts = [];
  page.on('dialog', async d => { alerts.push(d.message()); await d.accept(); });
  await page.goto(BASE + '/lib1.html');
  await page.click('#btnNew');
  await page.fill('.titleInput', 'Cap Clause');
  await page.fill('input[placeholder*="Liability"]', 'Liability Cap Family');
  await page.click('#btnNew');
  await page.fill('.titleInput', 'Cap Clause 2');
  // Force a DIFFERENT familyId under the SAME family display name via raw JS (simulating drift from two capture sessions).
  await page.evaluate(() => {
    const e = window.__proto__; // no-op, placeholder
  });
  await page.click('#btnHealth');
  await page.waitForTimeout(200);
  record('Health Check runs and reports a summary', alerts.some(a => /health check/i.test(a)), alerts[alerts.length - 1]);
  await browser.close();
}

async function testBackupRoundTrip() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  await withDialogs(page);
  await page.goto(BASE + '/lib1.html');
  await page.click('#btnNew');
  await page.fill('.titleInput', 'Backup Roundtrip Entry');
  await page.fill('.clauseTA', 'This clause must survive a JSON backup export/import cycle.');
  await page.waitForTimeout(200);

  // Capture the JSON backup by intercepting the Blob download.
  const backupJson = await page.evaluate(async () => {
    return await new Promise(resolve => {
      const origCreate = URL.createObjectURL;
      URL.createObjectURL = function (blob) {
        blob.text().then(resolve);
        return origCreate.call(URL, blob);
      };
      document.getElementById('btnBackup').click();
    });
  });
  const parsed = JSON.parse(backupJson);
  record('JSON Backup contains the unsaved entry (includesUnsavedChanges flag honest)', parsed.includesUnsavedChanges === true && parsed.entries.some(e => e.title === 'Backup Roundtrip Entry'), 'includesUnsavedChanges=' + parsed.includesUnsavedChanges);

  // Now open a FRESH library and import that backup JSON via the import input.
  const page2 = await ctx.newPage();
  await withDialogs(page2);
  await page2.goto(BASE + '/lib1.html');
  const tmpFile = path.join(__dirname, 'backup_reimport.json');
  fs.writeFileSync(tmpFile, backupJson);
  await page2.setInputFiles('#importInput', tmpFile);
  await page2.waitForTimeout(300);
  await page2.click('#importMerge');
  await page2.waitForTimeout(300);
  const titles = await page2.$$eval('.entryRow .rowTitle', els => els.map(e => e.textContent));
  record('Backup JSON re-imports cleanly into a separate library instance (merge)', titles.includes('Backup Roundtrip Entry'), JSON.stringify(titles));
  await browser.close();
}

async function testImportConflictKeepVsIncoming() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  await page.goto(BASE + '/lib1.html');
  await page.click('#btnNew');
  const existingId = await page.evaluate(() => {
    // Read the id the app just assigned to the newly created (first) entry.
    return document.querySelector('.entryRow').dataset.id;
  });
  await page.fill('.titleInput', 'Existing Version');
  await page.fill('.clauseTA', 'Existing wording, should be KEPT.');
  await page.waitForTimeout(150);

  const conflictPayload = {
    app: 'ClauseCommentsLibrary', version: 8, libraryId: 'lib-conflict', libraryName: 'Conflict Test',
    revision: 1, savedAt: new Date().toISOString(), commitId: 'c1',
    entries: [{ id: existingId, title: 'Incoming Version', type: 'Clause', clauseText: 'Incoming wording, should be DISCARDED if K chosen.', tags: [] }],
    inboxSeenIds: []
  };
  const tmpFile = path.join(__dirname, 'conflict.json');
  fs.writeFileSync(tmpFile, JSON.stringify(conflictPayload));

  // Answer the bulk "K/I/D" prompt with K (keep existing) - default path a rushed user would hit by pressing Enter.
  page.on('dialog', async d => { if (d.type() === 'prompt') await d.accept('K'); else await d.accept(); });
  await page.setInputFiles('#importInput', tmpFile);
  await page.waitForTimeout(300);
  await page.click('#importMerge');
  await page.waitForTimeout(300);
  const titles = await page.$$eval('.entryRow .rowTitle', els => els.map(e => e.textContent));
  record('Same-ID import conflict resolved via modal prompt keeps existing when user answers K', titles.includes('Existing Version') && !titles.includes('Incoming Version'), JSON.stringify(titles));
  await browser.close();
}

(async () => {
  await testHealthCheckAndFamily();
  await testBackupRoundTrip();
  await testImportConflictKeepVsIncoming();
  const failed = results.filter(r => !r.pass);
  console.log('\n==== SUMMARY 2: ' + (results.length - failed.length) + '/' + results.length + ' passed ====');
})();

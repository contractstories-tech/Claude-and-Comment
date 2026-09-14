const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const OUT = path.resolve(__dirname, 'out');
const appPath = path.resolve(__dirname, 'app.html');
const savedHtmlPath = path.join(OUT, 'clause-comments-library-copy.html');

async function withPage(browser) {
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();
  const msgs = [];
  page.on('console', msg => msgs.push(`[console.${msg.type()}] ${msg.text()}`));
  page.on('pageerror', err => msgs.push(`[pageerror] ${err.message}`));
  page.on('dialog', async d => { msgs.push(`[dialog:${d.type()}] ${d.message()}`); await d.accept(); });
  await page.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
  return { context, page, msgs };
}

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const results = {};

  // === TEST B: Round-trip reopen of the saved HTML from test2 ===
  {
    const { context, page, msgs } = await withPage(browser);
    await page.goto('file://' + savedHtmlPath);
    await page.waitForTimeout(300);
    const state = await page.evaluate(() => ({
      count: document.querySelectorAll('.entryRow').length,
      fileStatus: document.getElementById('fileStatus').textContent,
      recoveryStatus: document.getElementById('recoveryStatus').textContent,
    }));
    results.B_roundtripReopen = { state, msgs };
    await context.close();
  }

  // === TEST C: Crash-recovery — make changes, DO NOT save, reload page (simulates crash), check recovery prompt ===
  {
    const { context, page, msgs } = await withPage(browser);
    await page.goto('file://' + appPath);
    await page.waitForTimeout(300);
    await page.click('#btnNew');
    await page.fill('.titleInput', 'Unsaved Crash-Test Entry');
    await page.waitForTimeout(500); // allow scheduleRecovery debounce (350ms) to fire idb write
    // Simulate "crash": reload same file without saving
    await page.reload();
    await page.waitForTimeout(400);
    const afterReload = await page.evaluate(() => ({
      count: document.querySelectorAll('.entryRow').length,
      recoveryStatus: document.getElementById('recoveryStatus').textContent,
      titles: [...document.querySelectorAll('.rowTitle')].map(e=>e.textContent),
    }));
    results.C_crashRecoveryOnReload = { afterReload, msgs };
    await context.close();
  }

  // === TEST D: Import merge with conflicting ID, then replace ===
  {
    const { context, page, msgs } = await withPage(browser);
    await page.goto('file://' + appPath);
    await page.waitForTimeout(300);
    await page.click('#btnNew');
    await page.fill('.titleInput', 'Original Entry For Conflict Test');
    await page.fill('.clauseTA', 'Original text v1');
    await page.waitForTimeout(200);
    const origId = await page.evaluate(() => document.querySelectorAll('.entryRow')[0].onclick.toString());
    // grab the real entry id from app internal state via DOM data isn't exposed; use JS eval on closures not possible.
    // Instead export backup, parse id
    const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#btnBackup')]);
    const p1 = path.join(OUT, 'conflict_base.json');
    await dl.saveAs(p1);
    const base = JSON.parse(fs.readFileSync(p1, 'utf8'));
    const entryId = base.entries[0].id;

    // Build a conflicting import JSON file: same id, different title/text
    const importPayload = {
      app: 'ClauseCommentsLibrary', version: 6, libraryId: 'lib-other', libraryName: 'x', revision: 0, savedAt: null,
      entries: [{ ...base.entries[0], title: 'CONFLICTING Entry Title', clauseText: 'Conflicting text v2' }],
      inboxSeenIds: []
    };
    const importPath = path.join(OUT, 'conflict_import.json');
    fs.writeFileSync(importPath, JSON.stringify(importPayload));

    await page.setInputFiles('#importInput', importPath);
    await page.waitForTimeout(200);
    const dialogVisible = await page.evaluate(() => document.getElementById('importDialog').open);
    const summary = await page.evaluate(() => document.getElementById('importSummary').textContent);
    results.D_importDialog = { dialogVisible, summary };

    // Click Merge -> triggers prompt() for conflict resolution -> our dialog handler auto-accepts (Keep existing, default 'K')
    await page.click('#importMerge');
    await page.waitForTimeout(200);
    const afterMerge = await page.evaluate(() => {
      const rows = [...document.querySelectorAll('.rowTitle')].map(e=>e.textContent);
      return { rows, count: rows.length };
    });
    results.D_afterMergeKeepExisting = afterMerge;
    results.D_msgsSoFar = msgs.slice();

    await context.close();
  }

  // === TEST E: Malformed / unrelated file import ===
  {
    const { context, page, msgs } = await withPage(browser);
    await page.goto('file://' + appPath);
    await page.waitForTimeout(300);
    const garbagePath = path.join(OUT, 'garbage.json');
    fs.writeFileSync(garbagePath, '{ not valid json ][');
    const before1 = msgs.length;
    await page.setInputFiles('#importInput', garbagePath);
    await page.waitForTimeout(300);
    results.E_malformedImport = { newMsgs: msgs.slice(before1) };

    // unrelated but valid JSON (e.g. random object)
    const unrelatedPath = path.join(OUT, 'unrelated.json');
    fs.writeFileSync(unrelatedPath, JSON.stringify({ foo: 'bar', numbers: [1,2,3] }));
    const before2 = msgs.length;
    await page.setInputFiles('#importInput', unrelatedPath);
    await page.waitForTimeout(300);
    results.E_unrelatedJsonImport = { newMsgs: msgs.slice(before2) };

    // malformed HTML file (foreign html, no library-data tag)
    const foreignHtmlPath = path.join(OUT, 'foreign.html');
    fs.writeFileSync(foreignHtmlPath, '<html><body><h1>Not a library</h1></body></html>');
    const before3 = msgs.length;
    await page.setInputFiles('#fileInput', foreignHtmlPath);
    await page.waitForTimeout(300);
    const stillUsable = await page.evaluate(() => !!document.getElementById('btnNew'));
    results.E_foreignHtmlOpen = { newMsgs: msgs.slice(before3), appStillUsableAfter: stillUsable };

    results.E_msgs = msgs;
    await context.close();
  }

  fs.writeFileSync(path.join(OUT, 'results_BCDE.json'), JSON.stringify(results, null, 2));
  console.log(JSON.stringify(results, null, 2));
  await browser.close();
})();

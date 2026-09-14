const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const OUT = path.resolve(__dirname, 'out');
fs.mkdirSync(OUT, { recursive: true });

function log(...a) { console.log(...a); }

async function withPage(browser, { killFsAccess } = {}) {
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();
  const msgs = [];
  page.on('console', msg => msgs.push(`[console.${msg.type()}] ${msg.text()}`));
  page.on('pageerror', err => msgs.push(`[pageerror] ${err.message}\n${err.stack||''}`));
  page.on('dialog', async d => { msgs.push(`[dialog:${d.type()}] ${d.message()}`); await d.accept(); });
  if (killFsAccess) {
    await page.addInitScript(() => { delete window.showOpenFilePicker; delete window.showSaveFilePicker; });
  }
  return { context, page, msgs };
}

async function loadApp(page, filePath) {
  await page.goto('file://' + filePath);
  await page.waitForTimeout(300);
}

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
  const appPath = path.resolve(__dirname, 'app.html');
  const results = {};

  // === TEST A: fsAccess disabled (simulates unsupported browser / download-only fallback) ===
  {
    const { context, page, msgs } = await withPage(browser, { killFsAccess: true });
    await loadApp(page, appPath);

    // basic UI state
    const initial = await page.evaluate(() => ({
      fileStatus: document.getElementById('fileStatus').textContent,
      inboxStatus: document.getElementById('inboxStatus').textContent,
      cacheStatus: document.getElementById('cacheStatus').textContent,
      recoveryStatus: document.getElementById('recoveryStatus').textContent,
      saveAsLabel: document.getElementById('btnSaveAs').textContent,
    }));
    results.A_initial = initial;

    // Create new entry
    await page.click('#btnNew');
    await page.waitForTimeout(100);
    await page.fill('.titleInput', 'Limitation of Liability — Cap at 12 Months Fees <script>alert(1)</script>');
    await page.fill('.clauseTA', 'Neither party’s aggregate liability shall exceed <b>12 months</b> of fees paid.\n\nLine 2 with "quotes" & ampersands & <tags> & — em dash & emoji 😀 & 中文.');
    await page.fill('.commentTA', 'Negotiation comment: push back if counterparty proposes uncapped liability.');
    await page.fill('.contextTA', 'Use when client is the service provider.');
    const tagsInput = await page.$('.tagsInput');
    await tagsInput.fill('liability, cap, msa');
    await page.waitForTimeout(150);

    const afterEdit = await page.evaluate(() => {
      const titleEl = document.querySelector('.rowTitle');
      return {
        rowTitleText: titleEl ? titleEl.textContent : null,
        rowTitleHTML: titleEl ? titleEl.innerHTML : null,
        dirty: document.getElementById('fileStatus').classList.contains('dirty'),
      };
    });
    results.A_afterEditXSSCheck = afterEdit;

    // toggle favorite via detail button
    await page.click('.detailBtns .miniBtn'); // first button is Favourite
    await page.waitForTimeout(100);

    // duplicate
    const beforeDupCount = await page.evaluate(() => document.querySelectorAll('.entryRow').length);
    await page.click('.detailBtns .miniBtn:nth-child(2)'); // Duplicate
    await page.waitForTimeout(150);
    const afterDupCount = await page.evaluate(() => document.querySelectorAll('.entryRow').length);
    results.A_duplicate = { beforeDupCount, afterDupCount };

    // search
    await page.fill('#search', 'liability');
    await page.waitForTimeout(250);
    results.A_searchCount = await page.evaluate(() => document.getElementById('count').textContent);

    // create a second entry: empty everything (edge case) then another with only comment
    await page.click('#btnNew');
    await page.waitForTimeout(100);
    // leave title blank, save nothing special - check it doesn't crash
    const emptyEntryOk = await page.evaluate(() => !!document.querySelector('.titleInput'));
    results.A_emptyEntryFormRenders = emptyEntryOk;

    // very long clause text
    const longText = 'Lorem ipsum dolor sit amet. '.repeat(3000); // ~87k chars
    await page.fill('.clauseTA', longText);
    await page.waitForTimeout(200);
    const longLen = await page.evaluate(() => document.querySelector('.clauseTA').value.length);
    results.A_longClauseLen = longLen;

    // JSON Backup download
    const [download1] = await Promise.all([
      page.waitForEvent('download'),
      page.click('#btnBackup'),
    ]);
    const backupPath = path.join(OUT, await download1.suggestedFilename());
    await download1.saveAs(backupPath);
    results.A_backupFile = backupPath;

    // Save As (download fallback since fsAccess killed)
    const [download2] = await Promise.all([
      page.waitForEvent('download'),
      page.click('#btnSaveAs'),
    ]);
    const savedHtmlPath = path.join(OUT, await download2.suggestedFilename());
    await download2.saveAs(savedHtmlPath);
    results.A_savedHtmlFile = savedHtmlPath;
    results.A_dirtyAfterDownloadSaveAs = await page.evaluate(() => document.getElementById('fileStatus').classList.contains('dirty'));

    // delete one entry
    const countBeforeDelete = await page.evaluate(() => document.querySelectorAll('.entryRow').length);
    await page.click('#btnClearFilters');
    await page.waitForTimeout(100);
    const firstRow = await page.$('.entryRow');
    if (firstRow) await firstRow.click();
    await page.waitForTimeout(100);
    await page.click('.detailBtns .miniBtn:nth-child(3)'); // Delete
    await page.waitForTimeout(150);
    const countAfterDelete = await page.evaluate(() => document.querySelectorAll('.entryRow').length);
    results.A_delete = { countBeforeDelete, countAfterDelete };

    results.A_consoleMsgs = msgs;
    await context.close();
  }

  fs.writeFileSync(path.join(OUT, 'results_A.json'), JSON.stringify(results, null, 2));
  log(JSON.stringify(results, null, 2));

  await browser.close();
})();

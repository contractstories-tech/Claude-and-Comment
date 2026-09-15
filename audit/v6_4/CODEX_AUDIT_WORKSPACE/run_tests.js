// Independent adversarial test harness for Clause & Comments Library v6.4 HTML app.
// Uses real Chromium via Playwright against the UNMODIFIED shipped HTML (served locally).
// All scratch files live under CODEX_AUDIT_WORKSPACE; production package is untouched.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE = 'http://127.0.0.1:8934';
const results = [];

function record(name, pass, detail) {
  results.push({ name, pass, detail });
  console.log((pass ? 'PASS' : 'FAIL') + ' - ' + name + (detail ? ('  :: ' + detail) : ''));
}

async function withDialogs(page, autoAccept = true) {
  page.on('dialog', async d => {
    try {
      if (d.type() === 'prompt') await d.accept('K');
      else if (autoAccept) await d.accept();
      else await d.dismiss();
    } catch (e) {}
  });
}

async function test1_consoleAndNetwork() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  const errors = [];
  const requests = [];
  page.on('pageerror', e => errors.push(String(e)));
  page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('request', r => requests.push(r.url()));
  await withDialogs(page);
  await page.goto(BASE + '/lib1.html', { waitUntil: 'networkidle' });
  await page.waitForTimeout(500);
  const external = requests.filter(u => !u.startsWith(BASE));
  record('Loads with no console/page errors', errors.length === 0, errors.join(' | '));
  record('No external/network requests beyond own origin (offline/no telemetry)', external.length === 0, JSON.stringify(external));
  await browser.close();
}

async function test2_createAndDirtyRecovery() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  await withDialogs(page);
  await page.goto(BASE + '/lib1.html');
  await page.click('#btnNew');
  await page.fill('.titleInput', 'Test Clause Alpha');
  await page.waitForTimeout(600); // allow scheduleRecovery debounce (350ms) to fire
  const fileStatus = await page.textContent('#fileStatus');
  const recoveryStatus = await page.textContent('#recoveryStatus');
  record('New entry marks fileStatus dirty', fileStatus.includes('unsaved') === false ? false : true, fileStatus);
  record('Recovery snapshot status updates after edit', /recovery/i.test(recoveryStatus) || /saved/i.test(recoveryStatus), recoveryStatus);
  await browser.close();
}

async function test3_crashRecoveryAcrossReload() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  await withDialogs(page, true); // accept confirm() => recover newest snapshot
  await page.goto(BASE + '/lib1.html');
  await page.click('#btnNew');
  await page.fill('.titleInput', 'Crash Recovery Candidate');
  await page.waitForTimeout(600);
  await page.close(); // simulate crash: no save, tab just closed
  const page2 = await ctx.newPage();
  await withDialogs(page2, true);
  await page2.goto(BASE + '/lib1.html'); // reload the pristine template again
  await page2.waitForTimeout(800);
  const entries = await page2.$$eval('.entryRow .rowTitle', els => els.map(e => e.textContent));
  record('Unsaved work recovered after simulated crash (tab close, same origin)', entries.includes('Crash Recovery Candidate'), JSON.stringify(entries));
  await browser.close();
}

async function test4_duplicateIdOpen() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  await withDialogs(page, true); // accept "Open in Repair Mode?" confirm
  await page.goto(BASE + '/lib1.html');

  // Build a malicious/duplicate-ID library payload and feed it through the hidden file input.
  const dup = {
    app: 'ClauseCommentsLibrary', version: 8, libraryId: 'lib-dup-test', libraryName: 'Dup Test',
    revision: 1, savedAt: new Date().toISOString(), commitId: 'c-test', entries: [
      { id: 'SAME-ID', title: 'First', type: 'Clause', category: 'Other', status: 'Preferred', clauseText: 'Text A', tags: [] },
      { id: 'SAME-ID', title: 'Second', type: 'Clause', category: 'Other', status: 'Preferred', clauseText: 'Text B', tags: [] }
    ], inboxSeenIds: []
  };
  const htmlTemplate = fs.readFileSync(path.join(__dirname, 'site', 'lib1.html'), 'utf8');
  const injected = htmlTemplate.replace(/<script type="application\/json" id="library-data">.*?<\/script>/s,
    `<script type="application/json" id="library-data">${JSON.stringify(dup).replace(/</g, '\\u003c')}</script>`);
  const tmpFile = path.join(__dirname, 'dup_ids.html');
  fs.writeFileSync(tmpFile, injected);

  await page.setInputFiles('#fileInput', tmpFile);
  await page.waitForTimeout(600);
  const entries = await page.$$eval('.entryRow .rowTitle', els => els.map(e => e.textContent));
  const countText = await page.textContent('#count');
  record('Duplicate-ID file triggers Repair Mode instead of crashing/silently dropping', entries.length === 2, 'entries=' + JSON.stringify(entries) + ' count=' + countText);
  const repairedTitles = entries.filter(t => /Repaired duplicate/i.test(t));
  record('Repaired duplicate is flagged in title / marked Needs review', repairedTitles.length === 1, JSON.stringify(entries));
  await browser.close();
}

async function test5_corruptedEmbeddedData() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  const alerts = [];
  page.on('dialog', async d => { alerts.push(d.message()); await d.accept(); });
  await page.goto(BASE + '/lib1.html');
  const htmlTemplate = fs.readFileSync(path.join(__dirname, 'site', 'lib1.html'), 'utf8');
  const corrupted = htmlTemplate.replace(/<script type="application\/json" id="library-data">.*?<\/script>/s,
    `<script type="application/json" id="library-data">{not valid json!!!</script>`);
  const tmpFile = path.join(__dirname, 'corrupt.html');
  fs.writeFileSync(tmpFile, corrupted);
  await page.setInputFiles('#fileInput', tmpFile);
  await page.waitForTimeout(400);
  record('Corrupted embedded JSON produces a clear alert (not a silent failure/blank load)', alerts.some(a => /corrupt/i.test(a)), JSON.stringify(alerts));
  await browser.close();
}

async function test6_futureVersionRejected() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  const alerts = [];
  page.on('dialog', async d => { alerts.push(d.message()); await d.accept(); });
  await page.goto(BASE + '/lib1.html');
  const htmlTemplate = fs.readFileSync(path.join(__dirname, 'site', 'lib1.html'), 'utf8');
  const future = {
    app: 'ClauseCommentsLibrary', version: 99, libraryId: 'lib-future', libraryName: 'Future',
    revision: 1, savedAt: new Date().toISOString(), commitId: 'c1',
    entries: [{ id: 'x1', title: 'Future entry', type: 'Clause', clauseText: 'abc', tags: [] }], inboxSeenIds: []
  };
  const injected = htmlTemplate.replace(/<script type="application\/json" id="library-data">.*?<\/script>/s,
    `<script type="application/json" id="library-data">${JSON.stringify(future)}</script>`);
  const tmpFile = path.join(__dirname, 'future.html');
  fs.writeFileSync(tmpFile, injected);
  await page.setInputFiles('#fileInput', tmpFile);
  await page.waitForTimeout(400);
  record('Newer-schema file is refused rather than silently corrupted/truncated', alerts.some(a => /newer version/i.test(a)), JSON.stringify(alerts));
  await browser.close();
}

async function test7_multiTabSecondaryInstance() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const pageA = await ctx.newPage();
  await withDialogs(pageA);
  await pageA.goto(BASE + '/lib1.html');
  await pageA.waitForTimeout(300);
  const pageB = await ctx.newPage();
  await withDialogs(pageB);
  await pageB.goto(BASE + '/lib1.html');
  await pageB.waitForTimeout(700);
  const bannerB = await pageB.$eval('#instanceBanner', el => el.className + '|' + el.textContent).catch(() => 'MISSING');
  const bannerA = await pageA.$eval('#instanceBanner', el => el.className + '|' + el.textContent).catch(() => 'MISSING');
  record('Second tab of same (pristine) library is marked read-only/secondary via BroadcastChannel', bannerB.includes('show'), 'B=' + bannerB + ' A=' + bannerA);
  // Verify second tab's Save button actually disabled
  const saveDisabled = await pageB.$eval('#btnSave', el => el.disabled);
  record('Secondary-instance tab has Save disabled (cannot silently fork/overwrite)', saveDisabled === true, 'disabled=' + saveDisabled);
  await browser.close();
}

async function test8_mockedSaveAsBasename() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  await withDialogs(page);
  // Mock the File System Access API before the app script runs.
  await page.addInitScript(() => {
    window.__savedFiles = {};
    let calls = [];
    window.__pickerCalls = calls;
    function makeHandle(name) {
      let buf = '';
      return {
        name,
        kind: 'file',
        async getFile() {
          return { size: buf.length, text: async () => buf, name };
        },
        async createWritable() {
          return {
            async write(text) { buf = text; },
            async close() { window.__savedFiles[name] = buf; }
          };
        },
        async queryPermission() { return 'granted'; },
        async requestPermission() { return 'granted'; }
      };
    }
    window.showSaveFilePicker = async (opts) => {
      calls.push(opts);
      return makeHandle(opts.suggestedName || 'clause-comments-library.html');
    };
  });
  await page.goto(BASE + '/lib1.html?openedname=Clause_Comments_Library_TEST.html');
  // currentOpenFileName() only triggers for file: protocol; over http it will be blank.
  // Instead verify Save As suggests the fileName variable (null on fresh load) -> generic default.
  await page.click('#btnNew');
  await page.fill('.titleInput', 'Save Test Entry');
  await page.click('#btnSaveAs');
  await page.waitForTimeout(300);
  const calls = await page.evaluate(() => window.__pickerCalls);
  const saved = await page.evaluate(() => window.__savedFiles);
  const statusAfter = await page.textContent('#fileStatus');
  record('Save As invokes showSaveFilePicker with a concrete suggestedName', calls.length === 1 && !!calls[0].suggestedName, JSON.stringify(calls));
  record('After mocked Save, status switches from Opened to Master', /Master:/.test(statusAfter), statusAfter);
  record('Saved HTML re-embeds updated entries (round-trippable master)', Object.values(saved).some(html => html.includes('Save Test Entry')), 'keys=' + Object.keys(saved));
  await browser.close();
}

async function test9_diskConflictProtection() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  await withDialogs(page, false); // dismiss confirm => should NOT silently overwrite
  let externalDiskContent = null;
  await page.addInitScript(() => {
    window.__handle = (() => {
      let buf = null;
      return {
        name: 'shared-master.html',
        async getFile() { return { size: buf ? buf.length : 0, text: async () => buf || '' }; },
        async createWritable() { return { async write(t) { buf = t; }, async close() {} }; },
        async queryPermission() { return 'granted'; },
        async requestPermission() { return 'granted'; },
        __setBuf(v) { buf = v; },
        __getBuf() { return buf; }
      };
    })();
    window.showSaveFilePicker = async (opts) => window.__handle;
  });
  await page.goto(BASE + '/lib1.html');
  await page.click('#btnNew');
  await page.fill('.titleInput', 'Original Save');
  await page.click('#btnSaveAs');
  await page.waitForTimeout(300);
  // Simulate a second writer changing the on-disk file after this window's last read.
  await page.evaluate(() => {
    const original = window.__handle.__getBuf();
    const mutated = original.replace('"revision":1', '"revision":5').replace('"commitId":"c-', '"commitId":"external-c-');
    window.__handle.__setBuf(mutated);
  });
  await page.click('.titleInput');
  await page.fill('.titleInput', 'Original Save Edited Locally');
  await page.click('#btnSave');
  await page.waitForTimeout(400);
  const bufAfter = await page.evaluate(() => window.__handle.__getBuf());
  const stillExternal = bufAfter.includes('external-c-');
  record('Save refuses to silently clobber a master changed on disk since open (stale-write guard)', stillExternal, 'external marker still present=' + stillExternal);
  await browser.close();
}

async function test10_performanceAtScale() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  await withDialogs(page);
  const N = 2000;
  const entries = [];
  for (let i = 0; i < N; i++) {
    entries.push({
      id: 'perf-' + i, title: 'Synthetic Clause ' + i, type: 'Clause', category: (i % 5 === 0) ? 'Indemnity' : 'Other',
      status: ['Preferred', 'Fallback', 'Retired', 'Unclassified'][i % 4], position: ['provider', 'balanced', 'customer', 'none'][i % 4],
      clauseText: 'The parties agree that clause number ' + i + ' governs liability arising under section ' + (i % 20) + '.',
      commentText: i % 3 === 0 ? ('Negotiation note for clause ' + i) : '', tags: ['tag' + (i % 15), 'group' + (i % 7)],
      needsReview: i % 6 === 0, favorite: i % 11 === 0, source: 'Synthetic Matter ' + (i % 40)
    });
  }
  const payload = { app: 'ClauseCommentsLibrary', version: 8, libraryId: 'lib-perf', libraryName: 'Perf Test', revision: 1, savedAt: new Date().toISOString(), commitId: 'c-perf', entries, inboxSeenIds: [] };
  const htmlTemplate = fs.readFileSync(path.join(__dirname, 'site', 'lib1.html'), 'utf8');
  const injected = htmlTemplate.replace(/<script type="application\/json" id="library-data">.*?<\/script>/s,
    `<script type="application/json" id="library-data">${JSON.stringify(payload)}</script>`);
  fs.writeFileSync(path.join(__dirname, 'site', 'perf.html'), injected);

  const t0 = Date.now();
  await page.goto(BASE + '/perf.html');
  await page.waitForSelector('.entryRow');
  const loadMs = Date.now() - t0;

  const searchMs = await page.evaluate(async () => {
    const input = document.getElementById('search');
    const t = performance.now();
    input.value = 'liability section 5';
    input.dispatchEvent(new Event('input'));
    await new Promise(r => setTimeout(r, 200)); // debounce is 120ms
    return performance.now() - t;
  });
  const countText = await page.textContent('#count');
  record('App loads and renders with ' + N + ' entries', true, 'load=' + loadMs + 'ms');
  record('Search-as-you-type over ' + N + ' entries completes promptly (<500ms incl. debounce)', searchMs < 500, searchMs.toFixed(1) + 'ms, ' + countText);
  await browser.close();
}

async function test11_richStaleDetection() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  await withDialogs(page);
  const entries = [{
    id: 'rich-1', title: 'Rich Clause', type: 'Clause', category: 'Other', status: 'Preferred',
    clauseText: 'Original wording that was captured from Word.',
    wordOpenXml: '<pkg:package>FAKE</pkg:package>',
    richSourceText: 'Original wording that was captured from Word.',
    richSourceKnown: true, tags: []
  }];
  const payload = { app: 'ClauseCommentsLibrary', version: 8, libraryId: 'lib-rich', libraryName: 'Rich Test', revision: 1, savedAt: new Date().toISOString(), commitId: 'c1', entries, inboxSeenIds: [] };
  const htmlTemplate = fs.readFileSync(path.join(__dirname, 'site', 'lib1.html'), 'utf8');
  const injected = htmlTemplate.replace(/<script type="application\/json" id="library-data">.*?<\/script>/s,
    `<script type="application/json" id="library-data">${JSON.stringify(payload)}</script>`);
  fs.writeFileSync(path.join(__dirname, 'site', 'rich.html'), injected);
  await page.goto(BASE + '/rich.html');
  await page.click('.entryRow');
  const before = await page.textContent('#detailInner');
  record('Fresh rich-text entry shows "can be used for Word insertion" banner', /can be used for Word insertion/i.test(before));
  // Now edit the clause text so it diverges from richSourceText -> should become stale.
  await page.fill('.clauseTA', 'Original wording that was EDITED after Word capture.');
  await page.click('.titleInput'); // blur to trigger onblur re-render
  await page.waitForTimeout(200);
  const after = await page.textContent('#detailInner');
  record('Editing clause text after Word capture marks rich formatting stale (won\'t insert old formatting over new wording)', /stale/i.test(after), after.slice(0, 200));
  await browser.close();
}

(async () => {
  await test1_consoleAndNetwork();
  await test2_createAndDirtyRecovery();
  await test3_crashRecoveryAcrossReload();
  await test4_duplicateIdOpen();
  await test5_corruptedEmbeddedData();
  await test6_futureVersionRejected();
  await test7_multiTabSecondaryInstance();
  await test8_mockedSaveAsBasename();
  await test9_diskConflictProtection();
  await test10_performanceAtScale();
  await test11_richStaleDetection();

  const failed = results.filter(r => !r.pass);
  console.log('\n==== SUMMARY: ' + (results.length - failed.length) + '/' + results.length + ' passed ====');
  if (failed.length) {
    console.log('FAILED:');
    failed.forEach(f => console.log(' - ' + f.name + ' :: ' + f.detail));
  }
  fs.writeFileSync(path.join(__dirname, 'results.json'), JSON.stringify(results, null, 2));
})();

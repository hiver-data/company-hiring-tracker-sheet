// Offline tests: runs apps-script/Code.gs with mocked Apps Script globals. No network, no Apify run.
import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const here = path.dirname(fileURLToPath(import.meta.url));
const code = fs.readFileSync(path.join(here, '..', 'apps-script', 'Code.gs'), 'utf8');
const items = JSON.parse(fs.readFileSync(path.join(here, 'fixtures', 'dataset-items.json'), 'utf8'));
const matches = JSON.parse(fs.readFileSync(path.join(here, 'fixtures', 'company-matches.json'), 'utf8'));

function makeEnv({ statuses = ['RUNNING', 'SUCCEEDED'], pageSize = 1000, tabs = {}, props = { APIFY_TOKEN: 'tok' } } = {}) {
  const calls = [], toasts = [], triggers = [];
  let clock = 0, polls = 0;
  const sheets = {};
  const mkSheet = (name, rows = []) => {
    const s = { name, rows: rows.map(r => r.slice()), frozen: 0,
      getDataRange: () => ({ getValues: () => s.rows.map(r => r.slice()) }),
      getLastRow: () => s.rows.length,
      clear: () => { s.rows = []; },
      setFrozenRows: n => { s.frozen = n; },
      getRange: (r, c, nr, nc) => ({ setValues: v => { v.forEach((row, i) => { s.rows[r - 1 + i] = s.rows[r - 1 + i] || []; row.forEach((x, j) => { s.rows[r - 1 + i][c - 1 + j] = x; }); }); } }) };
    sheets[name] = s; return s;
  };
  Object.entries(tabs).forEach(([n, r]) => mkSheet(n, r));
  const store = { ...props };
  const ctx = {
    console, JSON, Date: class extends Date { static now() { return 1.79e12 + clock; } }, Math, parseInt, parseFloat, String, Object, Array, Error,
    Utilities: { sleep: ms => { clock += ms; }, formatDate: (d, tz, f) => d.toISOString().slice(0, 10) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => store[k] ?? null, setProperty: (k, v) => { store[k] = v; }, deleteProperty: k => { delete store[k]; } }) },
    ScriptApp: { newTrigger: fn => ({ timeBased: () => ({ after: ms => ({ create: () => triggers.push({ fn, after: ms }) }), everyDays: n => ({ atHour: h => ({ create: () => triggers.push({ fn, everyDays: n, hour: h }) }) }) }) }),
      getProjectTriggers: () => triggers.map(t => ({ getHandlerFunction: () => t.fn, t })), deleteTrigger: x => { const i = triggers.indexOf(x.t); if (i >= 0) triggers.splice(i, 1); } },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: n => sheets[n] || null, insertSheet: n => mkSheet(n), toast: m => toasts.push(m) }) },
    UrlFetchApp: { fetch: (url, opts) => {
      calls.push({ url, method: opts.method, auth: opts.headers && opts.headers.Authorization, body: opts.payload && JSON.parse(opts.payload) });
      const resp = (code, obj) => ({ getResponseCode: () => code, getContentText: () => JSON.stringify(obj) });
      if (opts.method === 'POST' && url.includes('/acts/hiver~all-open-jobs-by-company/runs')) return resp(201, { data: { id: 'RUN1', status: 'READY' } });
      if (url.includes('/actor-runs/RUN1')) { const st = statuses[Math.min(polls++, statuses.length - 1)]; return resp(200, { data: { id: 'RUN1', status: st, statusMessage: 'x', defaultDatasetId: 'DS1', defaultKeyValueStoreId: 'KV1' } }); }
      if (url.includes('/datasets/DS1/items')) { const off = +/offset=(\d+)/.exec(url)[1]; return resp(200, items.slice(off, off + pageSize)); }
      if (url.includes('/records/COMPANY_MATCHES')) return resp(200, matches);
      return resp(404, {});
    } },
  };
  vm.createContext(ctx);
  vm.runInContext(code, ctx);
  return { ctx, calls, toasts, triggers, sheets, store, clock: () => clock };
}

const companies = [['company'], ['stripe.com'], ['Spotify'], ['spotify'], ['']];
let n = 0;
function t(name, fn) { fn(); n++; console.log('ok -', name); }

t('async run: POST /runs (never run-sync), polls, writes Jobs + Matches', () => {
  const e = makeEnv({ tabs: { Companies: companies } });
  assert.equal(e.ctx.fetchOpenJobs(), 'done');
  const post = e.calls[0];
  assert.equal(post.method, 'POST');
  assert.match(post.url, /\/acts\/hiver~all-open-jobs-by-company\/runs\?maxTotalChargeUsd=1$/);
  assert.ok(!e.calls.some(c => c.url.includes('run-sync')));
  assert.deepEqual(post.body.companies, ['stripe.com', 'Spotify']); // header, blank and duplicate dropped
  assert.equal(post.body.maxJobs, 500);
  assert.equal(post.body.onlyNewSinceLastRun, undefined);
  assert.equal(e.calls[0].auth, 'Bearer tok');
  const jobs = e.sheets.Jobs.rows;
  assert.equal(jobs.length, items.length + 1);
  assert.equal(jobs[0].join(','), 'company,source,title,location,workplaceType,isRemote,department,employmentType,postedAt,salaryMin,salaryMax,salaryCurrency,salaryInterval,url,applyUrl,inputCompany');
  assert.equal(jobs[1][2], items[0].title);
  assert.equal(jobs[1][13], items[0].url);
  assert.equal(e.sheets.Matches.rows.length, Object.keys(matches).length + 1);
  const tide = e.sheets.Matches.rows.find(r => r[0] === 'Tide');
  assert.equal(tide[1], 'not matched');
  assert.match(tide[3], /no board/);
  assert.equal(e.store.RUN_ID, undefined);
});

t('settings become Actor input; cap defaults 500 jobs / $1.00', () => {
  const e = makeEnv({ tabs: { Companies: companies, Settings: [['setting', 'value'], ['titleKeywords', 'engineer, data'], ['locations', 'London'], ['remoteOnly', 'TRUE'], ['postedWithinDays', 14], ['maxJobs', 120], ['maxChargeUsd', 0.25]] } });
  e.ctx.fetchOpenJobs();
  const b = e.calls[0].body;
  assert.deepEqual(b.titleKeywords, ['engineer', 'data']);
  assert.deepEqual(b.locations, ['London']);
  assert.equal(b.remoteOnly, true);
  assert.equal(b.postedWithinDays, 14);
  assert.equal(b.maxJobs, 120);
  assert.match(e.calls[0].url, /maxTotalChargeUsd=0.25$/);
});

t('token is read from Script Properties; missing token fails clearly', () => {
  const e = makeEnv({ tabs: { Companies: companies }, props: {} });
  assert.throws(() => e.ctx.fetchOpenJobs(), /Set Apify token/);
  assert.equal(e.calls.length, 0);
});

t('long run: stops polling before the 6 min limit, schedules a resume, resumes and finishes', () => {
  const e = makeEnv({ tabs: { Companies: companies }, statuses: Array(70).fill('RUNNING').concat(['SUCCEEDED']) });
  assert.equal(e.ctx.fetchOpenJobs(), 'pending');
  assert.ok(e.clock() <= 5 * 60 * 1000, 'slice must stay under 5 min: ' + e.clock());
  assert.equal(e.store.RUN_ID, 'RUN1');
  assert.deepEqual(e.triggers.map(x => x.fn), ['resumeRun']);
  assert.equal(e.ctx.fetchOpenJobs(), 'done'); // second call does not start a new run
  assert.equal(e.calls.filter(c => c.method === 'POST').length, 1);
  assert.equal(e.triggers.length, 0);
  assert.equal(e.sheets.Jobs.rows.length, items.length + 1);
});

t('dataset pages are all read', () => {
  const e = makeEnv({ tabs: { Companies: companies }, pageSize: 10 });
  e.ctx.fetchOpenJobs();
  assert.equal(e.sheets.Jobs.rows.length, items.length + 1);
});

t('daily mode: onlyNewSinceLastRun, appends to "New jobs" with a date column, keeps earlier rows', () => {
  const e = makeEnv({ tabs: { Companies: companies, 'New jobs': [['foundOn', 'company'], ['2026-10-01', 'old']] } });
  e.ctx.dailyRun();
  assert.equal(e.calls[0].body.onlyNewSinceLastRun, true);
  const r = e.sheets['New jobs'].rows;
  assert.equal(r[1][1], 'old');
  assert.equal(r.length, 2 + items.length);
  assert.match(r[2][0], /^\d{4}-\d{2}-\d{2}$/);
  assert.equal(r[2][3], items[0].title);
});

t('failed run clears state and reports the status message', () => {
  const e = makeEnv({ tabs: { Companies: companies }, statuses: ['FAILED'] });
  assert.throws(() => e.ctx.fetchOpenJobs(), /Apify run failed/);
  assert.equal(e.store.RUN_ID, undefined);
});

t('daily trigger on / off', () => {
  const e = makeEnv({ tabs: { Companies: companies } });
  e.ctx.enableDaily(); e.ctx.enableDaily();
  assert.equal(e.triggers.filter(x => x.fn === 'dailyRun').length, 1);
  e.ctx.disableDaily();
  assert.equal(e.triggers.length, 0);
});

t('no token value is ever written to a sheet', () => {
  const e = makeEnv({ tabs: { Companies: companies }, props: { APIFY_TOKEN: 'SECRET123' } });
  e.ctx.fetchOpenJobs();
  assert.ok(!JSON.stringify(Object.values(e.sheets).map(s => s.rows)).includes('SECRET123'));
});

console.log(`${n} tests passed`);

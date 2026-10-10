/**
 * hiver > Company hiring tracker
 * Every open job at the companies listed in the "Companies" tab, via the
 * hiver/all-open-jobs-by-company Apify Actor. Async run + polling (no run-sync),
 * resumes itself when a run takes longer than one script execution.
 * Your Apify token lives in Script Properties, never in a cell.
 */
var ACTOR_ID = 'hiver~all-open-jobs-by-company';
var API = 'https://api.apify.com/v2';
var JOB_COLUMNS = ['company', 'source', 'title', 'location', 'workplaceType', 'isRemote', 'department',
  'employmentType', 'postedAt', 'salaryMin', 'salaryMax', 'salaryCurrency', 'salaryInterval', 'url', 'applyUrl', 'inputCompany'];
var MATCH_COLUMNS = ['company', 'status', 'board', 'jobs this run', 'evidence or reason', 'failed boards'];
var DEFAULT_SETTINGS = [
  ['titleKeywords', '', 'comma separated; keep jobs whose title contains one of these words'],
  ['excludeTitleKeywords', '', 'comma separated; drop jobs whose title contains one of these words'],
  ['locations', '', 'comma separated; keep jobs whose location contains one of these words'],
  ['remoteOnly', 'FALSE', 'TRUE keeps only jobs the company marks as remote'],
  ['postedWithinDays', '', 'whole number, empty = all open jobs'],
  ['maxJobs', '500', 'hard cap on jobs delivered AND charged per run ($1.50 per 1,000 jobs)'],
  ['maxJobsPerCompany', '25', 'whole number, default 25, 0 = no cap per company. Keeps one big employer from using up maxJobs before the companies below it are read. maxJobs should be at least companies x this number'],
  ['maxChargeUsd', '1', 'the run stops once it has cost this many dollars']
];
var POLL_MS = 5000;
var SLICE_MS = 4.5 * 60 * 1000; // stay well inside Apps Script's 6 minute limit
var RESUME_AFTER_MIN = 1;

function now_() { return Date.now(); }
function sleep_(ms) { Utilities.sleep(ms); }
function props_() { return PropertiesService.getScriptProperties(); }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function onOpen() {
  SpreadsheetApp.getUi().createMenu('hiver')
    .addItem('Fetch open jobs', 'fetchOpenJobs')
    .addItem('Check / resume running fetch', 'resumeRun')
    .addSeparator()
    .addItem('Set Apify token', 'setToken')
    .addItem('Turn on daily new-jobs check', 'enableDaily')
    .addItem('Turn off daily check', 'disableDaily')
    .addItem('Create Settings tab', 'ensureSettings')
    .addToUi();
}

function setToken() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Apify API token', 'Apify Console > Settings > API & Integrations. Stored in Script Properties, not in the sheet.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() === ui.Button.OK && r.getResponseText().trim()) {
    props_().setProperty('APIFY_TOKEN', r.getResponseText().trim());
    toast_('Token saved.');
  }
}

function toast_(msg) { try { ss_().toast(msg, 'hiver', 8); } catch (e) { /* ignore */ } }

function sheet_(name) {
  var s = ss_().getSheetByName(name);
  return s || ss_().insertSheet(name);
}

function ensureSettings() {
  var s = ss_().getSheetByName('Settings');
  if (s) {
    var have = {};
    s.getDataRange().getValues().forEach(function (row) { have[String(row[0]).trim()] = 1; });
    var missing = DEFAULT_SETTINGS.filter(function (d) { return !have[d[0]]; });
    if (missing.length) s.getRange(s.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
    return s;
  }
  s = ss_().insertSheet('Settings');
  s.getRange(1, 1, 1, 3).setValues([['setting', 'value', 'what it does']]);
  s.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
  return s;
}

function readCompanies_() {
  var s = ss_().getSheetByName('Companies');
  if (!s) throw new Error('Add a "Companies" tab with company names or websites in column A.');
  var vals = s.getDataRange().getValues(), out = [], seen = {};
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0] == null ? '' : vals[i][0]).trim();
    if (!v || (i === 0 && /^(company|companies|name|website)/i.test(v))) continue;
    if (!seen[v.toLowerCase()]) { seen[v.toLowerCase()] = 1; out.push(v); }
  }
  if (!out.length) throw new Error('No companies in column A of the Companies tab.');
  return out;
}

function readSettings_() {
  var s = ss_().getSheetByName('Settings'), cfg = {};
  DEFAULT_SETTINGS.forEach(function (d) { cfg[d[0]] = d[1]; });
  if (s) {
    s.getDataRange().getValues().forEach(function (row) {
      var k = String(row[0]).trim();
      if (k in cfg && row[1] !== '' && row[1] != null) cfg[k] = row[1];
    });
  }
  return cfg;
}

function list_(v) {
  return String(v == null ? '' : v).split(',').map(function (x) { return x.trim(); }).filter(Boolean);
}

/** Build the Actor input from the Companies and Settings tabs. onlyNew = daily mode. */
function buildInput_(onlyNew) {
  var cfg = readSettings_();
  var input = { companies: readCompanies_(), includeDescription: false };
  var maxJobs = parseInt(cfg.maxJobs, 10);
  input.maxJobs = maxJobs > 0 ? maxJobs : 500;
  var perCompany = parseInt(cfg.maxJobsPerCompany, 10);
  if (perCompany > 0) input.maxJobsPerCompany = perCompany;
  if (list_(cfg.titleKeywords).length) input.titleKeywords = list_(cfg.titleKeywords);
  if (list_(cfg.excludeTitleKeywords).length) input.excludeTitleKeywords = list_(cfg.excludeTitleKeywords);
  if (list_(cfg.locations).length) input.locations = list_(cfg.locations);
  if (String(cfg.remoteOnly).toUpperCase() === 'TRUE') input.remoteOnly = true;
  var days = parseInt(cfg.postedWithinDays, 10);
  if (days > 0) input.postedWithinDays = days;
  if (onlyNew) input.onlyNewSinceLastRun = true;
  return input;
}

function token_() {
  var t = props_().getProperty('APIFY_TOKEN');
  if (!t) throw new Error('No Apify token yet: menu hiver > Set Apify token.');
  return t;
}

function api_(method, path, body) {
  var opts = { method: method, headers: { Authorization: 'Bearer ' + token_() }, muteHttpExceptions: true };
  if (body !== undefined) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  var res = UrlFetchApp.fetch(API + path, opts);
  var code = res.getResponseCode();
  if (code === 404 && method === 'GET') return null;
  if (code >= 300) throw new Error('Apify API ' + code + ' on ' + method + ' ' + path.split('?')[0] + ': ' + String(res.getContentText()).slice(0, 200));
  return JSON.parse(res.getContentText());
}

function maxCharge_() {
  var c = parseFloat(readSettings_().maxChargeUsd);
  return c > 0 ? c : 1;
}

/** Menu: start an async run, then poll. */
function fetchOpenJobs() { return startRun_(false); }

/** Time trigger target: only jobs new since the last run, appended to "New jobs". */
function dailyRun() { return startRun_(true); }

function startRun_(onlyNew) {
  var p = props_();
  if (p.getProperty('RUN_ID')) {
    toast_('A fetch is already running; checking it.');
    return resumeRun();
  }
  var input = buildInput_(onlyNew);
  var run = api_('POST', '/acts/' + ACTOR_ID + '/runs?maxTotalChargeUsd=' + maxCharge_(), input).data;
  p.setProperty('RUN_ID', run.id);
  p.setProperty('RUN_MODE', onlyNew ? 'new' : 'all');
  var startMsg = 'Started run ' + run.id + ' for ' + input.companies.length + ' companies.';
  if (input.maxJobsPerCompany && input.maxJobs < input.companies.length * input.maxJobsPerCompany) startMsg += ' maxJobs (' + input.maxJobs + ') is less than companies x maxJobsPerCompany (' + input.companies.length * input.maxJobsPerCompany + '): later companies may get fewer rows. Raise maxJobs to cover them all.';
  toast_(startMsg);
  return pollRun_();
}

/** Menu / trigger: continue polling a stored run. */
function resumeRun() {
  if (!props_().getProperty('RUN_ID')) { toast_('No fetch is running.'); return 'idle'; }
  return pollRun_();
}

function pollRun_() {
  var p = props_(), id = p.getProperty('RUN_ID'), mode = p.getProperty('RUN_MODE') || 'all';
  var started = now_();
  while (true) {
    var run = api_('GET', '/actor-runs/' + id);
    if (!run) { clearRun_(); throw new Error('Run ' + id + ' not found.'); }
    run = run.data;
    if (run.status === 'SUCCEEDED' || run.status === 'ABORTED' || run.status === 'TIMED-OUT') {
      clearRun_();
      deleteResumeTriggers_();
      var n = writeResults_(run, mode);
      var msg = run.status === 'SUCCEEDED' ? n + ' jobs written.' : 'Run ' + run.status + ' - wrote the ' + n + ' jobs it delivered.';
      if (hitTotalCap_(n)) msg += ' Stopped at maxJobs: companies lower in the list may have 0 rows (see Matches). Raise maxJobs or set maxJobsPerCompany in Settings.';
      toast_(msg);
      return 'done';
    }
    if (run.status === 'FAILED') {
      clearRun_(); deleteResumeTriggers_();
      throw new Error('Apify run failed: ' + (run.statusMessage || 'see the run in Apify Console') + ' (' + id + ')');
    }
    if (now_() - started > SLICE_MS) {
      scheduleResume_();
      toast_('Still running (' + (run.statusMessage || run.status) + '). It will continue by itself in about a minute.');
      return 'pending';
    }
    sleep_(POLL_MS);
  }
}

function clearRun_() { props_().deleteProperty('RUN_ID'); props_().deleteProperty('RUN_MODE'); }

function scheduleResume_() {
  deleteResumeTriggers_();
  ScriptApp.newTrigger('resumeRun').timeBased().after(RESUME_AFTER_MIN * 60 * 1000).create();
}

function deleteResumeTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'resumeRun') ScriptApp.deleteTrigger(t);
  });
}

function fetchItems_(datasetId) {
  var all = [], offset = 0, limit = 1000;
  while (true) {
    var page = api_('GET', '/datasets/' + datasetId + '/items?format=json&clean=true&offset=' + offset + '&limit=' + limit);
    if (!page || !page.length) break;
    all = all.concat(page);
    offset += page.length;
  }
  return all;
}

function jobRow_(job) {
  return JOB_COLUMNS.map(function (c) { var v = job[c]; return v == null ? '' : v; });
}

function matchRows_(report, counts) {
  counts = counts || {};
  return Object.keys(report || {}).map(function (k) {
    var m = report[k], n = counts[k] || 0;
    if (m.status === 'matched') return [k, 'matched', (m.boards || []).join(', '), n, m.evidence || '', (m.failedBoards || []).join(', ')];
    return [k, 'not matched', '', n, m.reason || '', ''];
  });
}

/** Jobs written per company as typed in column A (the Actor's inputCompany field). */
function countByCompany_(items) {
  var c = {};
  items.forEach(function (j) { var k = j.inputCompany; if (k != null) c[k] = (c[k] || 0) + 1; });
  return c;
}

function writeTable_(name, header, rows) {
  var s = sheet_(name);
  s.clear();
  s.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) s.getRange(2, 1, rows.length, header.length).setValues(rows);
  s.setFrozenRows(1);
}

function writeResults_(run, mode) {
  var items = fetchItems_(run.defaultDatasetId).filter(function (j) { return j && j.url; });
  var rec = api_('GET', '/key-value-stores/' + run.defaultKeyValueStoreId + '/records/COMPANY_MATCHES');
  var rows = items.map(jobRow_);
  if (mode === 'new') {
    var today = Utilities.formatDate(new Date(now_()), 'UTC', 'yyyy-MM-dd');
    var s = sheet_('New jobs');
    if (s.getLastRow() === 0) s.getRange(1, 1, 1, JOB_COLUMNS.length + 1).setValues([['foundOn'].concat(JOB_COLUMNS)]);
    if (rows.length) s.getRange(s.getLastRow() + 1, 1, rows.length, JOB_COLUMNS.length + 1)
      .setValues(rows.map(function (r) { return [today].concat(r); }));
    s.setFrozenRows(1);
  } else {
    writeTable_('Jobs', JOB_COLUMNS, rows);
  }
  if (rec) writeTable_('Matches', MATCH_COLUMNS, matchRows_(rec, countByCompany_(items)));
  return rows.length;
}

/** True when the run stopped at the run-wide maxJobs cap (companies lower in the list may have 0 rows). */
function hitTotalCap_(n) {
  var cfg = readSettings_(), cap = parseInt(cfg.maxJobs, 10);
  if (!(cap > 0)) cap = 500;
  return n >= cap && !(parseInt(cfg.maxJobsPerCompany, 10) > 0);
}

function enableDaily() {
  disableDaily();
  ScriptApp.newTrigger('dailyRun').timeBased().everyDays(1).atHour(7).create();
  toast_('Daily check on (about 07:00 script time zone). New jobs are appended to the "New jobs" tab.');
}

function disableDaily() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyRun') ScriptApp.deleteTrigger(t);
  });
}

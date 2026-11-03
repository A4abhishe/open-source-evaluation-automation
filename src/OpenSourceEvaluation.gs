/**
 * Open Source Evaluation – Mid Evaluation automation
 * Spreadsheet: OpenSource Evaluation Registration (Responses)
 *
 * HOW TO INSTALL (one time, about 3 minutes)
 * 1. Open the spreadsheet > Extensions > Apps Script.
 * 2. Add a script file (Files + > Script), paste this whole file, click Save.
 * 3. In the function dropdown pick "setupOpenSourceEvaluation" and click Run.
 *    Approve the permissions (Forms, Sheets, external requests, triggers).
 * 4. Reload the spreadsheet. A "Open Source Evaluation Tools" menu appears.
 *    The form links are written to the "Eval Config" tab (rows 39–40).
 * 5. Optional but recommended: Open Source Evaluation Tools > Set GitHub token
 *    (a GitHub personal access token with public read access; without it
 *    GitHub allows only 60 requests per hour).
 *
 * What it does
 * - Creates the mid evaluation Google Form. The College ID dropdown is built
 *   from 'Attendance Master'!B12:C (roll number + name).
 * - Every submission is copied into the "Pulse Check Responses" tab in the
 *   exact column layout the score formulas expect (College ID kept as text).
 * - Opens and closes the form automatically at the times in 'Eval Config'!C10:D10.
 *   While the form is open, syncResponses runs every 5 minutes as a safety net.
 * - "Fetch GitHub data" pulls PR details and CI status into the "GitHub Data" tab.
 * - "Add new students" appends late joiners to the score tabs without reordering.
 */

const CFG = {
  roster: { sheet: 'Attendance Master', firstRow: 12, idCol: 2, nameCol: 3 },
  responses: 'Pulse Check Responses',
  github: 'GitHub Data',
  config: 'Eval Config',
  scoreTabs: ['Mid Evaluation Scores', 'Final Evaluation'],
  windowOpenCell: 'C10',
  windowCloseCell: 'D10',
  formLinkCells: { label: 'A39', value: 'B39', editLabel: 'A40', editValue: 'B40' },
  checkpoint: '1',
};

/*
 * Form title, description and closed message are built from the
 * Eval Config window, so the times shown to students always match
 * the times the triggers actually use.
 */
function formTitle_() {
  const w = formWindow_();
  return 'Open Source Evaluation – Mid Evaluation' +
    (w.close ? ' (closes ' + formatWindowTime_(w.close, 'd MMM, h:mm a') + ')' : '');
}

function formDescription_() {
  return 'Mid evaluation for the Open Source Evaluation. Fill this ' + formWindowText_() +
    ' (it takes about 10 minutes).\n\n' +
    'How your answers are used: the evaluation system reads this form, downloads the code from your fork, runs the ORIGINAL tests of your ' +
    'repository and reads your Markdown files (approach.md, challenges.md, optimizations.md, TESTING.md and your notes file). ' +
    'Marks are out of 100: Codebase Understanding 15, Problem Solving 25, Working Prototype 25, Code Quality 15, Testing 10, ' +
    'Demo & Communication 10.\n\n' +
    'Before you start: push your latest code to your fork, and keep the output of "python -m pytest -q" ready.\n\n' +
    'Answer honestly. Your answers are compared with your code, and mentors will ask you about them at the final evaluation. ' +
    'If you submit more than once, only your latest response counts.';
}

const PS = {
  OrderOps: {
    roster: 'Orderops',
    behaviours: [
      'Process all 7 event types (payment_confirmed … refund_processed)',
      'Enforce valid transitions (cancel after dispatch does not become CANCELLED)',
      'event_id idempotency: a replayed event changes nothing',
      'Integer paise; refund positive, within outstanding amount, only for paid cancelled/rejected orders',
      'refund_due_paise and refunded_paise tracked separately; REFUNDED only when fully refunded',
      'Reconciliation reports missing records, amount mismatches and status mismatches separately',
      'Safe handling of empty input, missing fields, unknown IDs, invalid event types and duplicates',
    ],
  },
  RewardBook: {
    roster: 'Reward Book',
    behaviours: [
      'Points formula with integer arithmetic (base, bonus, cap applied last)',
      'Invalid or negative amounts/rules give 0 points; cap of 0 means 0 points',
      'Only captured purchases earn points (declined, pending, unknown do not)',
      'Each event_id processed once (replays do not change balances)',
      'Proportional refund reversal; never more than awarded; balance never negative',
      'Auditable ledger entry for each accepted purchase/refund',
      'Repeat alerts by account, date and amount (review only, never block)',
      'Safe handling of empty input, missing IDs, unknown transactions and boundary values',
    ],
  },
  SmartDispatch: {
    roster: 'Smart Dispatch',
    behaviours: [
      'Validate order ID, restaurant ID, destination zone and non-negative times',
      'Only active partners below capacity; missing travel data means no reliable ETA',
      'Pickup time and ETA calculated as per the formula',
      'Lowest ETA wins; ties broken by partner ID (ascending)',
      'Missing travel data never treated as zero; unassigned orders explain why',
      'Duplicate cancellation does not release workload/capacity twice',
      'Deterministic results; safe with empty lists, unknown IDs, malformed records, negative times',
    ],
  },
};

const T = {
  id: 'College ID and Name',
  email: 'College Email',
  ps: 'Problem Statement',
  beh: ' – required behaviours completed',
  user: 'GitHub Username',
  fork: 'Fork URL',
  branch: 'Branch Name',
  pr: 'Pull Request URL',
  pytest: 'Last line of python -m pytest -q',
  main: 'Does python -m main run without errors?',
  files: 'Files completed',
  flow: 'Explain the data/state flow in your own words',
  bug: 'Main bug you fixed: file, function, root cause and fix',
  edges: 'Edge cases you handled',
  blockers: 'Blockers and plan until the final evaluation',
  decl: 'Declaration',
};

/* ---------------- Menu ---------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Open Source Evaluation Tools')
    .addItem('Set up form (run once)', 'setupOpenSourceEvaluation')
    .addSeparator()
    .addItem('Fetch GitHub data', 'fetchGitHubData')
    .addItem('Rebuild responses from form', 'rebuildResponses')
    .addItem('Add new students to score tabs', 'syncRoster')
    .addSeparator()
    .addItem('Open form now', 'openPulseForm')
    .addItem('Close form now', 'closePulseForm')
    .addItem('Re-schedule form open/close from Eval Config', 'scheduleFormWindow')
    .addItem('Set GitHub token', 'setGitHubToken')
    .addToUi();
}

/* ---------------- Setup ---------------- */

function setupOpenSourceEvaluation() {
  const props = PropertiesService.getScriptProperties();
  let form;
  const existing = props.getProperty('FORM_ID');
  if (existing) {
    form = FormApp.openById(existing);
  } else {
    form = buildForm_();
    props.setProperty('FORM_ID', form.getId());
  }
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onPulseSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onPulseSubmit').forForm(form).onFormSubmit().create();
  scheduleFormWindow();

  const cfg = SpreadsheetApp.getActive().getSheetByName(CFG.config);
  cfg.getRange(CFG.formLinkCells.label).setValue('Form link for students');
  cfg.getRange(CFG.formLinkCells.value).setValue(form.getPublishedUrl());
  cfg.getRange(CFG.formLinkCells.editLabel).setValue('Form edit link');
  cfg.getRange(CFG.formLinkCells.editValue).setValue(form.getEditUrl());
  onOpen();
  safeAlert_('Setup complete.\n\nStudent link:\n' + form.getPublishedUrl() +
    '\n\nThe form opens and closes automatically at the times in Eval Config.');
}

/** Updates the title and description of the existing form (safe to run any time). */
function refreshFormText() {
  getForm_().setTitle(formTitle_()).setDescription(formDescription_());
  safeAlert_('Form title and description updated.');
}

function buildForm_() {
  const ss = SpreadsheetApp.getActive();
  const form = FormApp.create(formTitle_());
  form.setDescription(formDescription_());
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false); // no sign-in, re-submission allowed (latest counts)
  form.setAllowResponseEdits(false);
  form.setProgressBar(true);
  form.setConfirmationMessage('Thank you. Your mid evaluation response is recorded. Keep working on your PR; the final evaluation is with mentors.');

  // Section 1: identity
  const roster = readRoster_();
  form.addListItem()
    .setTitle(T.id)
    .setHelpText('Pick your own roll number. Your name, batch, mobile and email are taken from the attendance roster, so choose carefully. If your roll number is missing, tell your batch mentor before submitting.')
    .setChoiceValues(roster.map(r => r.id + ' - ' + r.name))
    .setRequired(true);
  form.addTextItem()
    .setTitle(T.email)
    .setHelpText('Your email as given in the attendance roster (normally your college email).')
    .setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a valid email address.')
      .requireTextIsEmail().build())
    .setRequired(true);
  const psItem = form.addListItem()
    .setTitle(T.ps)
    .setHelpText('The problem statement you are working on. The next page shows its checklist.')
    .setRequired(true);

  // One checklist page per problem statement
  const names = Object.keys(PS);
  const pages = {};
  names.forEach(name => {
    pages[name] = form.addPageBreakItem()
      .setTitle(name + ' checklist')
      .setHelpText('Tick only the behaviours that work in the code you have pushed. These come from section 2 (Required behaviour) of your problem statement. Each tick is checked against your code.');
    form.addCheckboxItem()
      .setTitle(name + T.beh)
      .setChoiceValues(PS[name].behaviours.concat(['None yet']))
      .setRequired(true);
  });

  // GitHub page
  const ghPage = form.addPageBreakItem()
    .setTitle('Your GitHub work')
    .setHelpText('Your code is downloaded from here for evaluation. Make sure the fork is public and your latest commit is pushed.');
  // After each checklist page, jump to the GitHub page
  names.forEach((name, i) => {
    if (i < names.length - 1) pages[names[i + 1]].setGoToPage(ghPage);
  });
  psItem.setChoices(names.map(n => psItem.createChoice(n, pages[n])));

  form.addTextItem().setTitle(T.user)
    .setHelpText('Only the username, for example: rahul-dev (not the full link).')
    .setRequired(true);
  form.addTextItem().setTitle(T.fork)
    .setHelpText('Link to YOUR fork, for example: https://github.com/<your-username>/OrderOps-Open-Source-Evaluation')
    .setValidation(FormApp.createTextValidation()
      .setHelpText('Paste a github.com link.')
      .requireTextContainsPattern('github.com/').build())
    .setRequired(true);
  form.addTextItem().setTitle(T.branch)
    .setHelpText('The branch that has your latest work, for example: main or fix-refunds.')
    .setRequired(true);
  form.addTextItem().setTitle(T.pr)
    .setHelpText('Link to your pull request on the original repository, for example: https://github.com/<org>/OrderOps-Open-Source-Evaluation/pull/12. If you have not opened it yet, type: Not yet')
    .setRequired(true);

  // Progress page
  form.addPageBreakItem().setTitle('Progress')
    .setHelpText('Run the commands from your repository folder and copy the real output.');
  form.addTextItem().setTitle(T.pytest)
    .setHelpText('Copy the final summary line exactly, for example: 3 failed, 209 passed in 0.45s')
    .setRequired(true);
  form.addMultipleChoiceItem().setTitle(T.main)
    .setHelpText('Run: python -m main')
    .setChoiceValues(['Yes', 'No'])
    .setRequired(true);
  form.addCheckboxItem().setTitle(T.files)
    .setHelpText('Tick only files that are filled with your own work and pushed. Empty templates do not count.')
    .setChoiceValues(['approach.md', 'challenges.md', 'optimizations.md', 'TESTING.md',
      'Notes file (ORDER_FLOW_NOTES.md / REWARD_RULE_NOTES.md / DISPATCH_NOTES.md)',
      'New regression tests in tests/', 'None yet'])
    .setRequired(true);

  // Understanding page
  form.addPageBreakItem().setTitle('Your understanding')
    .setHelpText('Write in your own words. Short and specific answers score better than long generic ones.');
  form.addParagraphTextItem().setTitle(T.flow)
    .setHelpText('Max 150 words. How does data move through main.py, service.py and models.py? Name the key functions. (Codebase Understanding)')
    .setValidation(FormApp.createParagraphTextValidation().requireTextLengthLessThanOrEqualTo(1200).build())
    .setRequired(true);
  form.addParagraphTextItem().setTitle(T.bug)
    .setHelpText('Format: file > function > what was wrong (root cause) > what you changed > which test proves it. (Problem Solving)')
    .setRequired(true);
  form.addParagraphTextItem().setTitle(T.edges)
    .setHelpText('List the edge cases you handled, for example: duplicate event_id, empty input, refund larger than amount. (Testing)')
    .setRequired(true);
  form.addParagraphTextItem().setTitle(T.blockers)
    .setHelpText('Optional. Where are you stuck? Mentors use this to help you. Not scored.')
    .setRequired(false);

  // Declaration page
  form.addPageBreakItem().setTitle('Declaration');
  form.addCheckboxItem().setTitle(T.decl)
    .setChoiceValues(['I confirm this is my own individual work, I followed the AI-use policy, and I can explain my changes and make a small live change if asked.'])
    .setRequired(true);

  // Keep the form file next to the spreadsheet
  try {
    const parent = DriveApp.getFileById(ss.getId()).getParents();
    if (parent.hasNext()) DriveApp.getFileById(form.getId()).moveTo(parent.next());
  } catch (e) { /* not critical */ }
  return form;
}

function readRoster_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.roster.sheet);
  const last = sh.getLastRow();
  if (last < CFG.roster.firstRow) return []; // empty roster
  const vals = sh.getRange(CFG.roster.firstRow, CFG.roster.idCol, last - CFG.roster.firstRow + 1, 2).getDisplayValues();
  const seen = {};
  return vals
    .map(r => ({ id: String(r[0]).trim(), name: String(r[1]).trim() }))
    .filter(r => /^\d{6,}$/.test(r.id) && !seen[r.id] && (seen[r.id] = true))
    .sort((a, b) => a.id.localeCompare(b.id));
}

/* ---------------- Responses ---------------- */

function onPulseSubmit(e) {
  // Never throws: if anything goes wrong, syncResponses (every 5 min and at close) adds the response.
  let row;
  try { row = responseToRow_(e.response); } catch (err) { Logger.log('onPulseSubmit: could not read response, syncResponses will add it. ' + err); return; }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(300000)) { Logger.log('onPulseSubmit: sheet busy for 5 min, syncResponses will add it.'); return; }
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.responses);
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log('onPulseSubmit: write failed, syncResponses will add it. ' + err);
  } finally {
    lock.releaseLock();
  }
}

/** Safety net: makes Pulse Check Responses match every response stored in the form. */
function syncResponses() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(120000)) { Logger.log('syncResponses: busy, next run will retry.'); return; }
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.responses);
    const rows = getForm_().getResponses().map(responseToRow_);
    const current = Math.max(sh.getLastRow() - 1, 0);
    if (rows.length === current) { Logger.log('syncResponses: ' + current + ' rows, nothing missing.'); return; }
    if (current) sh.getRange(2, 1, current, 21).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 21).setNumberFormat('@').setValues(rows);
    SpreadsheetApp.flush();
    Logger.log('syncResponses: sheet had ' + current + ' rows, now ' + rows.length + '.');
  } finally {
    lock.releaseLock();
  }
}

/** Runs syncResponses every 5 minutes while the form is open (removed again when the form closes). */
function startSafetyNet() {
  stopSafetyNet_();
  ScriptApp.newTrigger('syncResponses').timeBased().everyMinutes(5).create();
  Logger.log('Safety net on: syncResponses every 5 minutes.');
}

function stopSafetyNet_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncResponses').forEach(t => ScriptApp.deleteTrigger(t));
}

function rebuildResponses() {
  const form = getForm_();
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.responses);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 21).clearContent();
  const rows = form.getResponses().map(responseToRow_);
  if (rows.length) sh.getRange(2, 1, rows.length, 21).setNumberFormat('@').setValues(rows);
  safeAlert_('Rebuilt ' + rows.length + ' responses from the form.');
}

function responseToRow_(resp) {
  const a = {};
  resp.getItemResponses().forEach(ir => {
    const v = ir.getResponse();
    a[ir.getItem().getTitle()] = Array.isArray(v) ? v.join('; ') : String(v);
  });
  const idName = a[T.id] || '';
  const id = (idName.match(/^\d+/) || [''])[0];
  const name = idName.replace(/^\d+\s*-\s*/, '');
  const ps = a[T.ps] || '';
  const ts = resp.getTimestamp();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return [
    CFG.checkpoint + '|' + id,
    Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm:ss'),
    CFG.checkpoint,
    id,
    name,
    a[T.email] || '',
    ps,
    a[ps + T.beh] || '',
    (a[T.user] || '').replace(/^https?:\/\/github\.com\//i, '').replace(/\/.*$/, '').trim(),
    a[T.fork] || '',
    a[T.branch] || '',
    a[T.pr] || '',
    a[T.pytest] || '',
    a[T.main] || '',
    a[T.files] || '',
    a[T.flow] || '',
    a[T.bug] || '',
    a[T.edges] || '',
    a[T.blockers] || '',
    a[T.decl] ? 'Confirmed' : '',
    windowStatus_(ts),
  ];
}

function windowStatus_(ts) {
  const w = formWindow_();
  if (!w.open || !w.close) return 'Yes';
  if (ts < w.open) return 'Early';
  if (ts > new Date(w.close.getTime() + 2 * 60 * 1000)) return 'Late'; // 2 min grace
  return 'Yes';
}

/* ---------------- Form open / close ---------------- */

/** Reads the form window from Eval Config. Missing or invalid times are returned as null. */
function formWindow_() {
  const cfg = SpreadsheetApp.getActive().getSheetByName(CFG.config);
  const open = cfg.getRange(CFG.windowOpenCell).getValue();
  const close = cfg.getRange(CFG.windowCloseCell).getValue();
  return {
    open: open instanceof Date ? open : null,
    close: close instanceof Date ? close : null,
  };
}

function formatWindowTime_(date, pattern) {
  return Utilities.formatDate(date, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), pattern);
}

function formWindowText_() {
  const w = formWindow_();
  if (w.open && w.close) {
    return 'between ' + formatWindowTime_(w.open, 'd MMM, h:mm a') +
      ' and ' + formatWindowTime_(w.close, 'd MMM, h:mm a');
  }
  return 'while the form is open';
}

function scheduleFormWindow() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['openPulseForm', 'closePulseForm'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  const w = formWindow_();
  const now = new Date();

  // The window is already over: keep the form closed instead of reopening it.
  if (w.close && w.close <= now) {
    closePulseForm(true);
    return;
  }

  if (w.open && w.open > now) {
    ScriptApp.newTrigger('openPulseForm').timeBased().at(w.open).create();
    closePulseForm(true);
  } else {
    openPulseForm(true);
  }
  if (w.close) {
    ScriptApp.newTrigger('closePulseForm').timeBased().at(w.close).create();
  }
}

function openPulseForm(silent) {
  try { refreshRosterChoices_(); } catch (e) { Logger.log('Roster dropdown not refreshed: ' + e); }
  getForm_().setAcceptingResponses(true);
  startSafetyNet();
  if (silent !== true) safeAlert_('Form is now OPEN.');
}

function closePulseForm(silent) {
  const f = getForm_();
  f.setAcceptingResponses(false);
  try {
    f.setCustomClosedFormMessage('The mid evaluation form is closed. It was open ' + formWindowText_() + '. Contact your batch mentor if you have a problem.');
  } catch (e) { Logger.log('Closed message not set: ' + e); }
  try { syncResponses(); } catch (e) { Logger.log('Final sync failed: ' + e); }
  stopSafetyNet_();
  if (silent !== true) safeAlert_('Form is now CLOSED.');
}

function getForm_() {
  const id = PropertiesService.getScriptProperties().getProperty('FORM_ID');
  if (!id) throw new Error('Form not created yet. Run Open Source Evaluation Tools > Set up form first.');
  return FormApp.openById(id);
}

/* ---------------- GitHub ---------------- */

function setGitHubToken() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('GitHub token', 'Paste a GitHub personal access token (read-only public access is enough). It is stored in Script Properties, not in the sheet.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', res.getResponseText().trim());
    ui.alert('Token saved.');
  }
}

function ghHeaders_() {
  const h = { Accept: 'application/vnd.github+json', 'User-Agent': 'opensource-eval-automation' };
  const tok = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (tok) h.Authorization = 'Bearer ' + tok;
  return h;
}

function fetchGitHubData() {
  const ss = SpreadsheetApp.getActive();
  const resp = ss.getSheetByName(CFG.responses);
  const out = ss.getSheetByName(CFG.github);
  const last = resp.getLastRow();
  if (last < 2) { safeAlert_('No responses yet.'); return; }
  const data = resp.getRange(2, 1, last - 1, 21).getDisplayValues();

  // latest response per key
  const latest = {};
  data.forEach(r => { if (r[0]) latest[r[0]] = r; });
  const items = Object.keys(latest).map(k => {
    const r = latest[k];
    const m = String(r[11]).match(/github\.com\/([^\/\s]+)\/([^\/\s]+)\/pull\/(\d+)/i);
    return { key: k, cp: r[2], id: r[3], prUrl: r[11], m: m };
  });

  const headers = ghHeaders_();
  const withPr = items.filter(it => it.m);
  const BATCH = 25;
  for (let i = 0; i < withPr.length; i += BATCH) {
    const chunk = withPr.slice(i, i + BATCH);
    const reqs = [];
    chunk.forEach(it => {
      const base = 'https://api.github.com/repos/' + it.m[1] + '/' + it.m[2] + '/pulls/' + it.m[3];
      reqs.push({ url: base, headers: headers, muteHttpExceptions: true });
      reqs.push({ url: base + '/files?per_page=100', headers: headers, muteHttpExceptions: true });
    });
    const res = UrlFetchApp.fetchAll(reqs);
    chunk.forEach((it, j) => {
      const pr = res[2 * j], files = res[2 * j + 1];
      if (pr.getResponseCode() !== 200) { it.error = 'PR ' + pr.getResponseCode(); return; }
      it.pr = JSON.parse(pr.getContentText());
      it.files = files.getResponseCode() === 200 ? JSON.parse(files.getContentText()) : [];
    });
    // CI status
    const ciReqs = chunk.filter(it => it.pr).map(it => ({
      url: 'https://api.github.com/repos/' + it.pr.head.repo.full_name + '/commits/' + it.pr.head.sha + '/check-runs?per_page=50',
      headers: headers, muteHttpExceptions: true,
    }));
    const ciRes = UrlFetchApp.fetchAll(ciReqs);
    chunk.filter(it => it.pr).forEach((it, j) => {
      try {
        const runs = JSON.parse(ciRes[j].getContentText()).check_runs || [];
        if (!runs.length) it.ci = 'No checks';
        else if (runs.some(r => r.status !== 'completed')) it.ci = 'Pending';
        else if (runs.every(r => r.conclusion === 'success' || r.conclusion === 'skipped')) it.ci = 'Success';
        else it.ci = 'Failure';
      } catch (e) { it.ci = 'Unknown'; }
    });
  }

  const tz = ss.getSpreadsheetTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  const rows = items.map(it => {
    if (!it.m) return [it.key, it.cp, it.id, it.prUrl, 'No PR', '', '', '', '', '', '', '', '', '', now, (!it.prUrl || /not\s*yet/i.test(it.prUrl)) ? '' : 'PR link not recognised'];
    if (!it.pr) return [it.key, it.cp, it.id, it.prUrl, '', '', '', '', '', '', '', '', '', '', now, it.error || 'Not fetched'];
    const f = it.files || [];
    const testChanged = f.some(x => /^tests\//.test(x.filename) && (x.status === 'removed' || x.deletions > 0)) ? 'Yes' : 'No';
    const wfChanged = f.some(x => /^\.github\/workflows\//.test(x.filename)) ? 'Yes' : 'No';
    return [it.key, it.cp, it.id, it.prUrl, it.pr.merged ? 'merged' : it.pr.state, it.pr.title,
      it.pr.head.ref, it.pr.head.sha, f.map(x => x.filename).join('; '), it.pr.additions, it.pr.deletions,
      testChanged, wfChanged, it.ci || '', now, ''];
  });

  if (out.getLastRow() > 1) out.getRange(2, 1, out.getLastRow() - 1, 16).clearContent();
  if (rows.length) out.getRange(2, 1, rows.length, 16).setNumberFormat('@').setValues(rows);
  safeAlert_('GitHub data fetched for ' + rows.length + ' students (' + withPr.length + ' with a PR link).');
}

/* ---------------- Roster sync ---------------- */

function syncRoster() {
  const ss = SpreadsheetApp.getActive();
  const roster = readRoster_().map(r => r.id);
  let added = 0;
  CFG.scoreTabs.forEach(name => {
    const sh = ss.getSheetByName(name);
    const ids = sh.getRange(2, 2, sh.getMaxRows() - 1, 1).getDisplayValues().map(r => r[0]).filter(String);
    const missing = roster.filter(id => ids.indexOf(id) < 0);
    if (missing.length) {
      sh.getRange(ids.length + 2, 2, missing.length, 1).setNumberFormat('@').setValues(missing.map(id => [id]));
      added += missing.length;
    }
  });
  const n = refreshRosterChoices_();
  safeAlert_((added ? 'Added ' + added + ' new rows across the score tabs.' : 'All roster students are already in the score tabs.') +
    '\nForm College ID list now has ' + n + ' students.');
}

/** Rebuilds the College ID dropdown of the form from the attendance roster. */
function refreshRosterChoices_() {
  const roster = readRoster_();
  const item = getForm_().getItems(FormApp.ItemType.LIST).filter(i => i.getTitle() === T.id)[0];
  if (!item) throw new Error('College ID question not found in the form.');
  item.asListItem().setChoiceValues(roster.map(r => r.id + ' - ' + r.name));
  Logger.log('College ID list refreshed: ' + roster.length + ' students.');
  return roster.length;
}

/* ---------------- Utils ---------------- */

function safeAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

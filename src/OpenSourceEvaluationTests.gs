/**
 * Test helpers for the Open Source Evaluation automation.
 * Safe to keep in the project; nothing here runs on its own.
 *
 * Replace the ids below with roll numbers that exist in your
 * 'Attendance Master' roster before running a test.
 */

const TEST_STUDENTS = [
  { id: '1000000001', ps: 'OrderOps', pr: 'https://github.com/your-org/OrderOps-Open-Source-Evaluation/pull/1' },
  { id: '1000000002', ps: 'RewardBook', pr: 'https://github.com/your-org/RewardBook-Open-Source-Evaluation/pull/1' },
  { id: '1000000003', ps: 'SmartDispatch', pr: 'Not yet' },
  { id: '1000000004', ps: 'SmartDispatch', pr: 'https://github.com/your-org/SmartDispatch-Open-Source-Evaluation/pull/1' },
];

/** Accept any valid email (some roster emails are not on the college domain). */
function maint_relaxEmailValidation() {
  const form = getForm_();
  const item = form.getItems().filter(i => i.getTitle() === T.email)[0];
  item.asTextItem()
    .setHelpText('Your email as given in the attendance roster (normally your college email).')
    .setValidation(FormApp.createTextValidation().setHelpText('Enter a valid email address.').requireTextIsEmail().build());
  form.setLimitOneResponsePerUser(false);
  Logger.log('Email validation relaxed. One-response limit: ' + form.hasLimitOneResponsePerUser() + ', collect email: ' + form.collectsEmail());
}

function buildTestResponse_(form, s, tag) {
  const byTitle = {};
  form.getItems().forEach(it => { byTitle[it.getTitle()] = it; });
  const roster = readRoster_();
  const r = roster.filter(x => x.id === s.id)[0];
  if (!r) throw new Error('Test student ' + s.id + ' is not in the roster. Update TEST_STUDENTS.');
  const resp = form.createResponse();
  const add = (title, val) => {
    const it = byTitle[title];
    if (!it) throw new Error('Missing form item: ' + title);
    let typed;
    switch (it.getType()) {
      case FormApp.ItemType.LIST: typed = it.asListItem().createResponse(val); break;
      case FormApp.ItemType.TEXT: typed = it.asTextItem().createResponse(val); break;
      case FormApp.ItemType.PARAGRAPH_TEXT: typed = it.asParagraphTextItem().createResponse(val); break;
      case FormApp.ItemType.MULTIPLE_CHOICE: typed = it.asMultipleChoiceItem().createResponse(val); break;
      case FormApp.ItemType.CHECKBOX: typed = it.asCheckboxItem().createResponse(val); break;
      default: throw new Error('Unhandled type for ' + title);
    }
    resp.withItemResponse(typed);
  };
  add(T.id, r.id + ' - ' + r.name);
  add(T.email, 'test.' + s.id + '@example.edu');
  add(T.ps, s.ps);
  add(s.ps + T.beh, [PS[s.ps].behaviours[0], PS[s.ps].behaviours[2]]);
  add(T.user, 'test-user-' + s.id);
  add(T.fork, 'https://github.com/test-user-' + s.id + '/' + s.ps + '-Open-Source-Evaluation');
  add(T.branch, 'main');
  add(T.pr, s.pr);
  add(T.pytest, '10 failed, 200 passed in 0.5s');
  add(T.main, 'Yes');
  add(T.files, ['approach.md', 'TESTING.md']);
  add(T.flow, 'TEST ' + tag + ': main.py loads records, service.py processes events.');
  add(T.bug, 'TEST ' + tag + ': service.py > process_event > duplicate event_id handled twice.');
  add(T.edges, 'TEST ' + tag + ': duplicate event, empty input.');
  add(T.blockers, 'TEST ' + tag);
  add(T.decl, [form.getItems().filter(i => i.getTitle() === T.decl && i.getType() === FormApp.ItemType.CHECKBOX)[0].asCheckboxItem().getChoices()[0].getValue()]);
  return resp;
}

/** Test 1: one response per test student. */
function test1_submitFour() {
  const form = getForm_();
  const was = form.isAcceptingResponses();
  form.setAcceptingResponses(true);
  try {
    TEST_STUDENTS.forEach(s => buildTestResponse_(form, s, 'T1').submit());
  } finally {
    if (!was) form.setAcceptingResponses(false);
  }
  Logger.log('Submitted ' + TEST_STUDENTS.length + ' test responses.');
}

/** Test 2: burst of 12 submissions back to back (checks no row is overwritten). */
function test2_burst() {
  const form = getForm_();
  const was = form.isAcceptingResponses();
  form.setAcceptingResponses(true);
  try {
    for (let i = 0; i < 12; i++) buildTestResponse_(form, TEST_STUDENTS[i % TEST_STUDENTS.length], 'T2-' + i).submit();
  } finally {
    if (!was) form.setAcceptingResponses(false);
  }
  Logger.log('Submitted 12 burst responses.');
}

/**
 * Test 3: schedule the form to open in 2 minutes and close in 4 minutes.
 * This overwrites the real window in Eval Config: set it back afterwards.
 */
function test3_scheduleSoon() {
  const cfg = SpreadsheetApp.getActive().getSheetByName(CFG.config);
  const now = Date.now();
  cfg.getRange(CFG.windowOpenCell).setValue(new Date(now + 2 * 60 * 1000));
  cfg.getRange(CFG.windowCloseCell).setValue(new Date(now + 4 * 60 * 1000));
  scheduleFormWindow();
  Logger.log('Window set to open ' + new Date(now + 120000) + ' and close ' + new Date(now + 240000));
}

/** Prints the current state of form, triggers and tabs. */
function test_status() {
  const form = getForm_();
  const ss = SpreadsheetApp.getActive();
  const cfg = ss.getSheetByName(CFG.config);
  Logger.log('Accepting responses: ' + form.isAcceptingResponses());
  Logger.log('Form responses stored: ' + form.getResponses().length);
  Logger.log('Pulse Check Responses rows: ' + (ss.getSheetByName(CFG.responses).getLastRow() - 1));
  Logger.log('GitHub Data rows: ' + (ss.getSheetByName(CFG.github).getLastRow() - 1));
  Logger.log('Window: ' + cfg.getRange(CFG.windowOpenCell).getDisplayValue() + ' -> ' + cfg.getRange(CFG.windowCloseCell).getDisplayValue());
  ScriptApp.getProjectTriggers().forEach(t => Logger.log('Trigger: ' + t.getHandlerFunction() + ' / ' + t.getEventType()));
}

/** Removes EVERY response and test row. Run only before the real evaluation starts. */
function test_cleanup() {
  const form = getForm_();
  form.deleteAllResponses();
  const ss = SpreadsheetApp.getActive();
  [[CFG.responses, 21], [CFG.github, 16], ['Auto Marks', 23]].forEach(([name, cols]) => {
    const sh = ss.getSheetByName(name);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, cols).clearContent();
  });
  Logger.log('Cleanup done. Form responses now: ' + form.getResponses().length);
}


/** Removes ONLY test responses (email starting with "test.") from the form, then re-syncs the sheet. Safe while real students are submitting. */
function test_cleanupTestOnly() {
  const form = getForm_();
  let removed = 0;
  form.getResponses().forEach(r => {
    const isTest = r.getItemResponses().some(ir => ir.getItem().getTitle() === T.email && String(ir.getResponse()).indexOf('test.') === 0);
    if (isTest) { form.deleteResponse(r.getId()); removed++; }
  });
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.responses);
  const lock = LockService.getScriptLock(); lock.waitLock(120000);
  try {
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 21).clearContent();
    const rows = form.getResponses().map(responseToRow_);
    if (rows.length) sh.getRange(2, 1, rows.length, 21).setNumberFormat('@').setValues(rows);
  } finally { lock.releaseLock(); }
  Logger.log('Removed ' + removed + ' test responses. Real responses kept: ' + form.getResponses().length);
}

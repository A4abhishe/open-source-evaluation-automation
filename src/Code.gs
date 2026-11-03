/**
 * Auto-assigns a problem statement to every new form response.
 * - Picks the problem statement with the fewest students so far (random tie-break),
 *   so the split stays balanced.
 * - If the same student (same contact number) submits again, they keep their
 *   existing problem statement.
 * - Values are written as plain text, so they never change afterwards.
 */
const SHEET_NAME = 'Registration Details';
const HEADER = 'Problem Statement';
const PROBLEMS = ['Orderops', 'Reward Book', 'Smart Dispatch'];

function onFormSubmitAssign(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    assignMissing();
  } finally {
    lock.releaseLock();
  }
}

function assignMissing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Registration sheet not found');
  SpreadsheetApp.flush();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const psCol = headers.indexOf(HEADER) + 1;
  const nameCol = headers.indexOf('Name') + 1;
  const phoneCol = headers.indexOf('Contact Number') + 1;
  if (!psCol || !nameCol) throw new Error('Header not found');

  const n = lastRow - 1;
  const ps = sh.getRange(2, psCol, n, 1).getValues().map(r => String(r[0]).trim());
  const names = sh.getRange(2, nameCol, n, 1).getValues().map(r => String(r[0]).trim());
  const phones = phoneCol ? sh.getRange(2, phoneCol, n, 1).getValues().map(r => String(r[0]).trim()) : [];

  const counts = {};
  PROBLEMS.forEach(p => counts[p] = 0);
  const byPhone = {};
  ps.forEach((p, i) => {
    if (counts[p] !== undefined) {
      counts[p]++;
      if (phones[i]) byPhone[phones[i]] = p;
    }
  });

  ps.forEach((p, i) => {
    if (p !== '' || names[i] === '') return;
    let pick = phones[i] && byPhone[phones[i]];
    if (!pick) {
      const min = Math.min(...PROBLEMS.map(x => counts[x]));
      const options = PROBLEMS.filter(x => counts[x] === min);
      pick = options[Math.floor(Math.random() * options.length)];
      counts[pick]++;
      if (phones[i]) byPhone[phones[i]] = pick;
    }
    sh.getRange(i + 2, psCol).setValue(pick);
  });
}

/** Run this once to switch the automation on. */
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmitAssign')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onFormSubmitAssign')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  assignMissing();
}

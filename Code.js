/**
 * Zinia Experimental Lab — PARSEC Whitefield
 * Apps Script web app.
 *
 * WHY THE PROJECT IS SPLIT INTO EIGHT FILES
 * The kiosk is one 350 KB page, most of it the Zinion artwork stored as text.
 * The Apps Script editor silently truncates very large pastes, which leaves a
 * page that loads but whose script never runs — the name box does nothing and
 * the characters never appear. Each file here is small enough to paste safely.
 *
 * FILES TO CREATE  (+ button in the editor)
 *   Code.gs     .gs    this file
 *   Index       HTML   the page
 *   Styles      HTML   the CSS
 *   Zin1 … Zin5 HTML   one Zinion each
 *   App         HTML   the program
 *
 * Names must match exactly, capital letter included.
 *
 * DEPLOY
 *   Deploy > New deployment > Web app
 *     Execute as:     Me
 *     Who has access: Anyone
 *
 * AFTER EVERY LATER EDIT
 *   Deploy > Manage deployments > pencil > Version: New version > Deploy.
 *   Editing a file alone changes nothing on the live URL.
 *
 * ONE THING APPS SCRIPT CANNOT DO
 *   It serves every web app inside a sandboxed frame, and a framed page is
 *   never granted the microphone. Scream Lab falls back to tap-to-jump here.
 *   For the shouting to work the page needs its own https:// address —
 *   GitHub Pages, or a local server on the kiosk machine.
 */

/* ?view=dash opens straight onto "Today's board", filled from the results Sheet.
 * Only the private "Dashboard" deployment (Who has access: Only myself) can fill
 * it — see DASH_ENABLED. */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.view = (e && e.parameter && e.parameter.view === 'dash') ? 'dash' : '';
  // the refresh button reloads the whole web app, so the kiosk picks up the newest deployment
  t.url = ScriptApp.getService().getUrl() + (t.view ? '?view=' + t.view : '');
  return t
    .evaluate()
    .setTitle('Zinia Experimental Lab')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Pulls one HTML file into another. Used by the <?!= include('X'); ?> lines. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ================== DAILY QUIZ REPORT ==================
 * Each finished gallery-quiz round is sent here by the kiosk (logResult) and
 * appended to a Google Sheet in the owner's Drive. Once a day, sendDailyReport
 * turns that day's rows into a Google Doc and emails it as a PDF.
 *
 * ONE-TIME SETUP: open this project in the Apps Script editor, pick
 * setupDailyReport in the function menu, press Run, and approve the
 * permissions. That creates the Sheet and the daily 8 pm trigger.
 */
const REPORT_TO     = 'nikhil.dhruva@paraminnovation.org';
const REPORT_HOUR   = 20;              // 8 pm, Asia/Kolkata (Apps Script runs it within that hour)
const REPORT_TZ     = 'Asia/Kolkata';
const REPORT_FOLDER = 'Zinia Lab Daily Reports';
const STAGES        = ['Citizen', 'Seeker', 'Solver', 'Innovator'];
const ANSWER_STAGES = ['Innovator', 'Solver', 'Seeker', 'Citizen'];   // the order the cascade asks them
// Every gallery tab starts with these; the columns after them depend on the gallery
// (Cymatics: one per cascade stage; Traffic: one per question).
const BASE_HEADERS  = ['Date', 'Time', 'Name', 'Gallery', 'Stage reached'];
const LOG_HEADERS   = BASE_HEADERS.concat(ANSWER_STAGES.map(s => `${s} answer`));   // Cymatics
const SCORE_HEADERS = ['Score', 'Right', 'Wrong'];   // always the last columns of a gallery tab
/* The dashboard's data feed. false in every public version, so the kiosk link can
 * never hand out visitor names. To update the private "Dashboard" deployment: set
 * this true and appsscript.json "access" to "MYSELF", push, redeploy that deployment
 * only, then put both back and push before touching the public deployment. */
const DASH_ENABLED  = false;

/** The project's results spreadsheet: one tab per gallery (Cymatics, Rain Stick, …),
 *  one row per visitor. The account that owns this script needs edit access to it. */
const LOG_SHEET_ID = '1ZaVu8JGjFomYui3OTajSiU8cVzoHTEptYP0mFgLoj3E';

function logBook_() {
  return SpreadsheetApp.openById(LOG_SHEET_ID);
}

/** The tab for one gallery, made (with its header row) the first time that gallery is played.
 *  Any of `headers` the tab doesn't have yet are added as new columns on the right. */
function galleryTab_(ss, gallery, headers) {
  // Sheet tab names can't hold : \ / ? * [ ] and are capped at 100 characters.
  let name = (String(gallery || 'Other').replace(/[:\\/?*\[\]]/g, ' ').trim() || 'Other').slice(0, 100);
  let tab = ss.getSheetByName(name);
  // A same-named tab that wasn't made here (has content, first cell isn't "Date") is left alone.
  if (tab && tab.getLastRow() > 0 && tab.getRange(1, 1).getDisplayValue() !== BASE_HEADERS[0]) {
    name = `${name} (quiz)`.slice(0, 100);
    tab = ss.getSheetByName(name);
  }
  if (!tab) {
    // Before tabs were per gallery, every row went into one "Results" tab. If all of its rows
    // are this gallery's, it simply becomes this gallery's tab, keeping those rows together.
    const old = ss.getSheetByName('Results');
    if (old && old.getRange(1, 1).getDisplayValue() === BASE_HEADERS[0]) {
      const galleries = old.getDataRange().getDisplayValues().slice(1).map(r => r[3]);
      if (galleries.every(g => g === name)) tab = old.setName(name);
    }
  }
  if (!tab) {
    // Reuse the spreadsheet's untouched starter tab rather than leaving it empty beside the galleries.
    const blank = ss.getSheets().find(t => t.getName() === 'Sheet1' && t.getLastRow() === 0);
    tab = blank ? blank.setName(name) : ss.insertSheet(name);
  }
  const have = tab.getLastRow() ? tab.getRange(1, 1, 1, tab.getLastColumn()).getDisplayValues()[0] : [];
  const missing = headers.filter(h => have.indexOf(h) < 0);
  if (missing.length) {
    tab.getRange(1, have.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    tab.setFrozenRows(1);
  }
  return { tab: tab, headers: have.concat(missing) };
}

/** Every visitor row across all gallery tabs (any tab whose header starts with Date). */
function allRows_() {
  return logBook_().getSheets()
    .map(t => t.getDataRange().getDisplayValues())
    .filter(v => v.length && v[0][0] === LOG_HEADERS[0])
    .reduce((all, v) => all.concat(v.slice(1)), []);
}

/** Today's visitors from every gallery tab, for "Today's board". Owner only:
 *  visitor names must not reach anyone who opens the public kiosk link. */
function getDashData() {
  if (!DASH_ENABLED) throw new Error('The dashboard only opens from its private link.');
  const day = Utilities.formatDate(new Date(), REPORT_TZ, 'yyyy-MM-dd');
  const out = [];
  logBook_().getSheets().forEach(t => {
    const v = t.getDataRange().getDisplayValues();
    if (!v.length || v[0][0] !== BASE_HEADERS[0]) return;
    const col = {};
    v[0].forEach((h, i) => col[h] = i);
    const cell = (r, h) => col[h] === undefined ? '' : r[col[h]];
    const num = (r, h) => cell(r, h) === '' ? null : Number(cell(r, h));
    v.slice(1).filter(r => r[0] === day).forEach(r => out.push({
      name: cell(r, 'Name'), zone: cell(r, 'Gallery'), at: cell(r, 'Time'),
      c2iStage: cell(r, 'Stage reached'),
      score: num(r, 'Score'), right: num(r, 'Right'), wrong: num(r, 'Wrong')
    }));
  });
  return out;
}

/* The hidden stats screen on the kiosk (double-tap Argon, then the password).
 * It is asked for from the public link, so it checks the password here and sends
 * back only day, gallery, stage and visit for each finished quiz: never a name.
 * Only the password's SHA-256 is kept. */
const STATS_PASS_HASH = 'a25ef4e636a84f16d0df491e2c265438b2e99859313e0601eb1716f5cdfd8a64';

function checkPassword_(password) {
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password || ''), Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
  if (hash !== STATS_PASS_HASH) throw new Error('That password isn’t right.');
}

/* Zetta's secret door (double-tap Zetta): the same password opens today's game
 * leaders. The leaders live on the kiosk itself, so this only checks the password. */
function checkLeadersPassword(password) {
  checkPassword_(password);
  return true;
}

/* Axon's secret door (double-tap or hold Axon): every visit's questions and answers,
 * with day, time, gallery and stage. Same password; names are never sent. */
function getAnswers(password) {
  checkPassword_(password);
  return answerRows_(false);
}

/* Mach's secret door (double-tap or hold Mach): the same visits, each with the
 * visitor's name, for the gallery-by-gallery visitor list. Names leave the server
 * only after the password is checked here. */
function getVisitors(password) {
  checkPassword_(password);
  return answerRows_(true);
}

function answerRows_(withNames) {
  const out = [];
  logBook_().getSheets().forEach(t => {
    const v = t.getDataRange().getDisplayValues();
    if (!v.length || v[0][0] !== BASE_HEADERS[0]) return;
    const head = v[0];
    const col = {};
    head.forEach((h, i) => col[h] = i);
    const cell = (r, h) => col[h] === undefined ? '' : r[col[h]];
    // "Question N asked" + "Question N answer" pairs; older columns are just "<something> answer"
    const pairs = [];
    if (col['Visited before'] !== undefined) pairs.push({ q: 'Visited before?', a: col['Visited before'] });
    head.forEach((h, i) => {
      const m = /^Question (\d+) asked$/.exec(h);
      if (m && col[`Question ${m[1]} answer`] !== undefined) pairs.push({ qCol: i, a: col[`Question ${m[1]} answer`] });
      else if (/ answer$/.test(h) && !/^Question \d+ answer$/.test(h)) pairs.push({ q: h.replace(/ answer$/, ''), a: i });
    });
    v.slice(1).forEach(r => {
      const qa = pairs.map(p => [p.q || r[p.qCol], r[p.a]]).filter(x => x[0] && x[1]);
      if (!qa.length) return;
      const row = { day: cell(r, 'Date'), time: cell(r, 'Time'), gallery: cell(r, 'Gallery'), stage: cell(r, 'Stage reached'), qa: qa };
      if (withNames) row.name = cell(r, 'Name');
      out.push(row);
    });
  });
  return out;
}

function getStageStats(password) {
  checkPassword_(password);
  const out = [];
  logBook_().getSheets().forEach(t => {
    const v = t.getDataRange().getDisplayValues();
    if (!v.length || v[0][0] !== BASE_HEADERS[0]) return;
    const col = {};
    v[0].forEach((h, i) => col[h] = i);
    const cell = (r, h) => col[h] === undefined ? '' : r[col[h]];
    v.slice(1).forEach(r => {
      const stage = cell(r, 'Stage reached');
      if (STAGES.indexOf(stage) < 0) return;
      out.push({ day: cell(r, 'Date'), gallery: cell(r, 'Gallery'), stage: stage, visited: cell(r, 'Visited before') });
    });
  });
  return out;
}

function reportFolder_() {
  const it = DriveApp.getFoldersByName(REPORT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(REPORT_FOLDER);
}

/** Called from the kiosk (google.script.run) when a quiz round finishes.
 *  Adds one row to that gallery's tab in the results Sheet and one row to the "Zinian Data" Doc.
 *  entry.fields is a list of [column header, value] for the gallery's own questions. */
function logResult(entry) {
  if (!entry || typeof entry !== 'object') return;
  const clean = (v, max) => String(v == null ? '' : v).slice(0, max || 80);
  let fields = Array.isArray(entry.fields) ? entry.fields : null;
  if (!fields) {
    // Older kiosk pages sent { answers: { Innovator: … } } for the Cymatics cascade.
    const answers = (entry.answers && typeof entry.answers === 'object') ? entry.answers : {};
    fields = ANSWER_STAGES.map(st => [`${st} answer`, answers[st]]);
  }
  fields = fields.filter(f => Array.isArray(f) && f[0]).slice(0, 20)
    .map(f => [clean(f[0], 100), clean(f[1], 1000)]);

  const now = new Date();
  const values = {
    'Date': Utilities.formatDate(now, REPORT_TZ, 'yyyy-MM-dd'),
    'Time': Utilities.formatDate(now, REPORT_TZ, 'HH:mm'),
    'Name': clean(entry.name),
    'Gallery': clean(entry.zone),
    'Stage reached': STAGES.indexOf(entry.stage) >= 0 ? entry.stage : ''
  };
  fields.forEach(f => values[f[0]] = f[1]);
  const docHeaders = BASE_HEADERS.concat(fields.map(f => f[0]));
  // the dashboard's points and accuracy; Sheet only, so the Doc's tables keep their columns
  const num = v => (v === '' || v == null || isNaN(Number(v))) ? '' : String(Number(v));
  values['Score'] = num(entry.score);
  values['Right'] = num(entry.right);
  values['Wrong'] = num(entry.wrong);
  const headers = docHeaders.concat(SCORE_HEADERS);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const t = galleryTab_(logBook_(), values['Gallery'], headers);
    // Leading = + - @ would make Sheets treat a visitor's text as a formula; the ' keeps it plain text.
    t.tab.appendRow(t.headers.map(h => String(values[h] || '').replace(/^[=+\-@]/, "'$&")));
    try { appendToDataDoc_(docHeaders, docHeaders.map(h => values[h] || '')); }
    catch (e) { console.error('Zinian Data doc: ' + e); }
  } finally {
    lock.releaseLock();
  }
}

/** The running record in the "Zinian Data" Google Doc: one table, one row per visitor. */
const DATA_DOC_ID = '1LURjSxi0QpdKldZpUpJqf6w2_mcYYEAkF1L_2bF-cGw';

function appendToDataDoc_(headers, row) {
  const doc = DocumentApp.openById(DATA_DOC_ID);
  const body = doc.getBody();
  const key = headers.join('|');
  const headerOf = t => {
    if (!t.getNumRows()) return '';
    const r = t.getRow(0), out = [];
    for (let c = 0; c < r.getNumCells(); c++) out.push(r.getCell(c).getText());
    return out.join('|');
  };
  let table = body.getTables().find(t => headerOf(t) === key);
  if (!table) {
    // Table first, so a visitor row can always be written; the headings and layout are extras.
    table = body.appendTable([headers]);
    try {
      table.getRow(0).editAsText().setBold(true);
      const title = 'Zinia Experimental Lab — Gallery Quiz Data';
      if (!body.getParagraphs().some(p => p.getText() === title)) {
        body.insertParagraph(0, title).setHeading(DocumentApp.ParagraphHeading.TITLE);
        body.insertParagraph(1, 'PARSEC Whitefield · one row per visitor, added automatically by the kiosk.');
      }
      body.insertParagraph(body.getChildIndex(table), row[3] || 'Gallery')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.setPageWidth(842).setPageHeight(595);   // A4 landscape, so the columns fit
    } catch (e) { console.error('Zinian Data doc layout: ' + e); }
  }
  // Show the date the Foundation way (DD MMM YYYY) in the Doc; the Sheet keeps yyyy-MM-dd for filtering.
  const shown = row.slice();
  const [y, m, d] = row[0].split('-').map(Number);
  shown[0] = Utilities.formatDate(new Date(y, m - 1, d, 12), REPORT_TZ, 'dd MMM yyyy');
  const tr = table.appendTableRow();
  shown.forEach(v => tr.appendTableCell(v));
  doc.saveAndClose();
}

/** Run from the editor to check the Doc link works: adds one clearly marked test row. */
function testDataDoc() {
  const today = Utilities.formatDate(new Date(), REPORT_TZ, 'yyyy-MM-dd');
  const now = Utilities.formatDate(new Date(), REPORT_TZ, 'HH:mm');
  appendToDataDoc_(LOG_HEADERS, [today, now, 'TEST ROW', 'Water Cymatics', 'Solver', '', '', 'Test — delete me', '']);
}

/* Staff test plays, removed from the Sheet and the Doc. Run from the editor (or clasp run):
 *   removeStaffRows()      lists the matching rows, deletes nothing
 *   removeStaffRows(true)  deletes them
 * A row matches when its Name contains one of these, any case. Only someone signed in can
 * run it, so the public kiosk link (anonymous visitors) can never call it. */
const STAFF_NAMES = /nikhil|sitesh/i;

function removeStaffRows(apply) {
  if (!Session.getActiveUser().getEmail()) throw new Error('Run this from the Apps Script editor.');
  const found = [];
  logBook_().getSheets().forEach(t => {
    const v = t.getDataRange().getDisplayValues();
    if (!v.length || v[0][0] !== BASE_HEADERS[0]) return;
    const nameCol = v[0].indexOf('Name');
    if (nameCol < 0) return;
    // bottom up, so deleting a row doesn't shift the ones still to check
    for (let i = v.length - 1; i >= 1; i--) {
      if (!STAFF_NAMES.test(v[i][nameCol])) continue;
      found.push(`Sheet · ${t.getName()} · ${v[i][0]} ${v[i][1]} · ${v[i][nameCol]}`);
      if (apply) t.deleteRow(i + 1);
    }
  });
  const doc = DocumentApp.openById(DATA_DOC_ID);
  doc.getBody().getTables().forEach(table => {
    if (!table.getNumRows() || table.getRow(0).getNumCells() < 4 || table.getRow(0).getCell(2).getText() !== 'Name') return;
    for (let i = table.getNumRows() - 1; i >= 1; i--) {
      const row = table.getRow(i);
      if (row.getNumCells() < 4 || !STAFF_NAMES.test(row.getCell(2).getText())) continue;
      found.push(`Doc · ${row.getCell(3).getText()} · ${row.getCell(0).getText()} ${row.getCell(1).getText()} · ${row.getCell(2).getText()}`);
      if (apply) table.removeRow(i);
    }
  });
  if (apply) doc.saveAndClose();
  console.log(`${apply ? 'Removed' : 'Would remove'} ${found.length} rows:\n` + found.join('\n'));
  return found;
}

/** Builds today's report as a Google Doc and emails it as a PDF. Run daily by the trigger. */
function sendDailyReport() {
  const now = new Date();
  const dayKey = Utilities.formatDate(now, REPORT_TZ, 'yyyy-MM-dd');
  const dayLabel = Utilities.formatDate(now, REPORT_TZ, 'dd MMM yyyy');

  const rows = allRows_()
    .filter(r => r[0] === dayKey);

  const counts = {};
  STAGES.forEach(s => counts[s] = 0);
  rows.forEach(r => { if (counts.hasOwnProperty(r[4])) counts[r[4]]++; });

  const doc = DocumentApp.create(`Zinia Lab — Gallery Quiz — ${dayLabel}`);
  const body = doc.getBody();
  body.getParagraphs()[0].setText('Zinia Experimental Lab — Gallery Quiz Report').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(`PARSEC Whitefield · ${dayLabel}`).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);

  body.appendParagraph('Summary').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Visitors who finished the quiz: ${rows.length}`);
  body.appendTable([['Stage', 'Visitors']].concat(STAGES.map(s => [s, String(counts[s])])))
    .getRow(0).editAsText().setBold(true);

  body.appendParagraph('All visitors').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  if (rows.length) {
    body.appendTable([['Time', 'Name', 'Gallery', 'Stage reached']].concat(rows.map(r => [r[1], r[2], r[3], r[4]])))
      .getRow(0).editAsText().setBold(true);
  } else {
    body.appendParagraph('No visitors finished the quiz today.');
  }
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  reportFolder_().addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  MailApp.sendEmail({
    to: REPORT_TO,
    subject: `Zinia Lab — Gallery Quiz report — ${dayLabel}`,
    body: `Today's Gallery Quiz report is attached as a PDF.\n\n` +
          `Visitors: ${rows.length}\n` +
          STAGES.map(s => `${s}: ${counts[s]}`).join('\n') +
          `\n\nGoogle Doc: ${doc.getUrl()}\n`,
    attachments: [file.getAs(MimeType.PDF).setName(`Zinia Lab Gallery Quiz ${dayLabel}.pdf`)]
  });
}

/** Run once from the editor: creates the results Sheet and the daily trigger. */
function setupDailyReport() {
  console.log('Results Sheet: ' + logBook_().getUrl());
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyReport')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased().everyDays(1).atHour(REPORT_HOUR).inTimezone(REPORT_TZ)
    .create();
}

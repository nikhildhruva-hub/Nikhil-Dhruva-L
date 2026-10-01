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

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
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
const LOG_HEADERS   = ['Date', 'Time', 'Name', 'Gallery', 'Stage reached']
  .concat(ANSWER_STAGES.map(s => `${s} answer`));

/** The project's results spreadsheet: one tab per gallery (Cymatics, Rain Stick, …),
 *  one row per visitor. The account that owns this script needs edit access to it. */
const LOG_SHEET_ID = '1ZaVu8JGjFomYui3OTajSiU8cVzoHTEptYP0mFgLoj3E';

function logBook_() {
  return SpreadsheetApp.openById(LOG_SHEET_ID);
}

/** The tab for one gallery, made (with its header row) the first time that gallery is played. */
function galleryTab_(ss, gallery) {
  // Sheet tab names can't hold : \ / ? * [ ] and are capped at 100 characters.
  let name = (String(gallery || 'Other').replace(/[:\\/?*\[\]]/g, ' ').trim() || 'Other').slice(0, 100);
  let tab = ss.getSheetByName(name);
  // A same-named tab that wasn't made here (has content, first cell isn't "Date") is left alone.
  if (tab && tab.getLastRow() > 0 && tab.getRange(1, 1).getDisplayValue() !== LOG_HEADERS[0]) {
    name = `${name} (quiz)`.slice(0, 100);
    tab = ss.getSheetByName(name);
  }
  if (tab && tab.getLastRow() > 0) return tab;

  if (!tab) {
    // Before tabs were per gallery, every row went into one "Results" tab. If all of its rows
    // are this gallery's, it simply becomes this gallery's tab, keeping those rows together.
    const old = ss.getSheetByName('Results');
    if (old && old.getRange(1, 1).getDisplayValue() === LOG_HEADERS[0]) {
      const galleries = old.getDataRange().getDisplayValues().slice(1).map(r => r[3]);
      if (galleries.every(g => g === name)) return old.setName(name);
    }
  }
  if (!tab) {
    // Reuse the spreadsheet's untouched starter tab rather than leaving it empty beside the galleries.
    const blank = ss.getSheets().find(t => t.getName() === 'Sheet1' && t.getLastRow() === 0);
    tab = blank ? blank.setName(name) : ss.insertSheet(name);
  }
  tab.appendRow(LOG_HEADERS);
  tab.setFrozenRows(1);
  tab.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
  return tab;
}

/** Every visitor row across all gallery tabs (any tab whose header starts with Date). */
function allRows_() {
  return logBook_().getSheets()
    .map(t => t.getDataRange().getDisplayValues())
    .filter(v => v.length && v[0][0] === LOG_HEADERS[0])
    .reduce((all, v) => all.concat(v.slice(1)), []);
}

function reportFolder_() {
  const it = DriveApp.getFoldersByName(REPORT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(REPORT_FOLDER);
}

/** Called from the kiosk (google.script.run) when a quiz round finishes.
 *  Adds one row to that gallery's tab in the results Sheet and one row to the "Zinian Data" Doc. */
function logResult(entry) {
  if (!entry || typeof entry !== 'object') return;
  const clean = (v, max) => String(v == null ? '' : v).slice(0, max || 80);
  const answers = (entry.answers && typeof entry.answers === 'object') ? entry.answers : {};
  const now = new Date();
  const row = [
    Utilities.formatDate(now, REPORT_TZ, 'yyyy-MM-dd'),
    Utilities.formatDate(now, REPORT_TZ, 'HH:mm'),
    clean(entry.name),
    clean(entry.zone),
    STAGES.indexOf(entry.stage) >= 0 ? entry.stage : ''
  ].concat(ANSWER_STAGES.map(s => clean(answers[s], 1000)));

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // Leading = + - @ would make Sheets treat a visitor's text as a formula; the ' keeps it plain text.
    galleryTab_(logBook_(), row[3]).appendRow(row.map(v => v.replace(/^[=+\-@]/, "'$&")));
    try { appendToDataDoc_(row); } catch (e) { console.error('Zinian Data doc: ' + e); }
  } finally {
    lock.releaseLock();
  }
}

/** The running record in the "Zinian Data" Google Doc: one table, one row per visitor. */
const DATA_DOC_ID = '1LURjSxi0QpdKldZpUpJqf6w2_mcYYEAkF1L_2bF-cGw';

function appendToDataDoc_(row) {
  const doc = DocumentApp.openById(DATA_DOC_ID);
  const body = doc.getBody();
  let table = body.getTables().find(t => t.getNumRows() && t.getCell(0, 0).getText() === LOG_HEADERS[0]);
  if (!table) {
    // Table first, so a visitor row can always be written; the title and layout are extras.
    table = body.appendTable([LOG_HEADERS]);
    try {
      table.getRow(0).editAsText().setBold(true);
      const hasTitle = body.getParagraphs().some(p => p.getText() === 'Zinia Experimental Lab — Gallery Quiz Data');
      if (!hasTitle) {
        body.insertParagraph(0, 'Zinia Experimental Lab — Gallery Quiz Data').setHeading(DocumentApp.ParagraphHeading.TITLE);
      }
      body.insertParagraph(body.getChildIndex(table), 'PARSEC Whitefield · one row per visitor, added automatically by the kiosk.');
      body.setPageWidth(842).setPageHeight(595);   // A4 landscape, so the nine columns fit
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
  appendToDataDoc_([today, now, 'TEST ROW', 'Cymatics', 'Solver', '', '', 'Test — delete me', '']);
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

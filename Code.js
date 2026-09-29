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
const LOG_HEADERS   = ['Date', 'Time', 'Name', 'Gallery', 'Stage reached'];
const STAGES        = ['Citizen', 'Seeker', 'Solver', 'Innovator'];

/** The results Sheet, created on first use; its ID is kept in Script Properties. */
function logSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('LOG_SHEET_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Zinia Lab — Gallery Quiz Results');
    ss.getSheets()[0].setName('Results').appendRow(LOG_HEADERS);
    ss.getSheets()[0].setFrozenRows(1);
    reportFolder_().addFile(DriveApp.getFileById(ss.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(ss.getId()));
    props.setProperty('LOG_SHEET_ID', ss.getId());
  }
  return ss.getSheetByName('Results');
}

function reportFolder_() {
  const it = DriveApp.getFoldersByName(REPORT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(REPORT_FOLDER);
}

/** Called from the kiosk (google.script.run) when a quiz round finishes. */
function logResult(entry) {
  if (!entry || typeof entry !== 'object') return;
  const clean = v => String(v == null ? '' : v).slice(0, 80);
  const now = new Date();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    logSheet_().appendRow([
      Utilities.formatDate(now, REPORT_TZ, 'yyyy-MM-dd'),
      Utilities.formatDate(now, REPORT_TZ, 'HH:mm'),
      clean(entry.name),
      clean(entry.zone),
      STAGES.indexOf(entry.stage) >= 0 ? entry.stage : ''
    ]);
  } finally {
    lock.releaseLock();
  }
}

/** Builds today's report as a Google Doc and emails it as a PDF. Run daily by the trigger. */
function sendDailyReport() {
  const now = new Date();
  const dayKey = Utilities.formatDate(now, REPORT_TZ, 'yyyy-MM-dd');
  const dayLabel = Utilities.formatDate(now, REPORT_TZ, 'dd MMM yyyy');

  const rows = logSheet_().getDataRange().getDisplayValues().slice(1)
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
  logSheet_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyReport')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased().everyDays(1).atHour(REPORT_HOUR).inTimezone(REPORT_TZ)
    .create();
}

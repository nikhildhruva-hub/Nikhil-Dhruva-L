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
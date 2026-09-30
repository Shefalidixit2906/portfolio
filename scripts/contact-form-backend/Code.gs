/**
 * Contact form backend for the portfolio site.
 *
 * SETUP (one-time):
 *   1. Create a new Google Sheet (e.g. "Portfolio Contact Submissions").
 *   2. In row 1, add these exact headers: Timestamp | Name | Email | Message
 *   3. Extensions -> Apps Script. Delete any starter code, paste this whole
 *      file in, and save.
 *   4. Deploy -> New deployment -> type "Web app".
 *        Execute as:  Me
 *        Who has access:  Anyone
 *   5. Click Deploy, authorize the requested permissions (this script only
 *      writes to this one Sheet), and copy the Web app URL it gives you.
 *   6. Send that URL back — it gets pasted into the site's contact form so
 *      submissions land in this Sheet automatically.
 *
 * Whenever you edit this file after the first deploy, use
 * Deploy -> Manage deployments -> edit (pencil) -> New version, otherwise
 * the live URL keeps serving the old code.
 */

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var params = e.parameter || {};

  sheet.appendRow([
    new Date(),
    params.name || "",
    params.email || "",
    params.message || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

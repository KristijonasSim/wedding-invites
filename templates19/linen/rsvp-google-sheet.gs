/**
 * RSVP -> Google Sheet (one per client)
 * 1. Create a new Google Sheet for the couple.
 * 2. Extensions -> Apps Script, paste this code, Save.
 * 3. Deploy -> New deployment -> type "Web app"
 *    Execute as: Me | Who has access: Anyone -> Deploy -> copy the URL.
 * 4. Paste that URL into CONFIG.rsvpUrl in index.html.
 * 5. Share the Sheet (view only) with the couple.
 */
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Time", "Name", "Email", "Attending", "Guests", "Meal", "Message"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
  }
  const p = e.parameter;
  const yes = p.attending === "Yes";
  sheet.appendRow([new Date(), p.name || "", p.email || "", p.attending || "",
                   yes ? (p.guests || "") : 0, yes ? (p.meal || "") : "", p.message || ""]);
  return ContentService.createTextOutput("ok");
}

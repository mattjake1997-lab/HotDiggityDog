/**
 * Hot Diggity Dog - online orders → Google Sheet
 *
 * 1. Paste this whole file into Extensions > Apps Script (replace what's there).
 * 2. Put your email below to get an email for every new order (or leave "" for none).
 * 3. Run "setup" once (it asks for permission the first time).
 * 4. Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
 * 5. Copy the Web app URL into the site's settings: ordering.sheetUrl
 */

const NOTIFY_EMAIL = "mattjake1997@gmail.com"; // "" to turn off order emails
const SHEET_NAME = "Orders";
const FIRST_ORDER_NUMBER = 101;
const HEADERS = ["Order #", "Received", "Status", "Name", "Phone", "Items", "Notes", "Total", "Payment"];
const STATUSES = ["NEW", "MAKING", "READY", "PICKED UP", "CANCELED"];

/** Run once: builds the Orders tab with headers, status dropdown and colors. */
function setup() {
  const sh = getSheet_();
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
  sh.getRange("C2:C").setDataValidation(rule);
  const range = sh.getRange("A2:I");
  const colors = { "NEW": "#FFE9A8", "MAKING": "#FFD3B0", "READY": "#C9EBD9", "PICKED UP": "#E6E6E6", "CANCELED": "#F4C7C7" };
  const rules = Object.keys(colors).map(s =>
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="' + s + '"').setBackground(colors[s]).setRanges([range]).build());
  sh.setConditionalFormatRules(rules);
  sh.setColumnWidth(6, 380);
  sh.getRange("F:F").setWrap(true);
  // Ask once for email permission so new-order emails work.
  if (NOTIFY_EMAIL) MailApp.getRemainingDailyQuota();
}

/** Receives an order from the website. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = getSheet_();
    const num = FIRST_ORDER_NUMBER + Math.max(0, sh.getLastRow() - 1);
    const items = (d.items || []).map(String).join("\n");
    sh.appendRow([num, new Date(), "NEW", clean_(d.name), "'" + clean_(d.phone), items,
                  clean_(d.notes), clean_(d.total), clean_(d.payment || "Pay at window")]);
    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, "🌭 New order #" + num + " – " + clean_(d.name),
        "Order #" + num + "\nName: " + clean_(d.name) + "\nPhone: " + clean_(d.phone) +
        "\n\n" + items + (d.notes ? "\n\nNotes: " + clean_(d.notes) : "") + "\n\nTotal: " + clean_(d.total) +
        " (" + clean_(d.payment || "Pay at window") + ")");
    }
    return json_({ ok: true, num: num });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Lets you open the Web app URL in a browser to check it's live. */
function doGet() {
  return json_({ ok: true, message: "Hot Diggity Dog order receiver is running." });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#D62828").setFontColor("#FFFFFF");
  }
  return sh;
}

function clean_(v) {
  // Stops anything typed into the order form from being treated as a spreadsheet formula.
  const s = String(v == null ? "" : v).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

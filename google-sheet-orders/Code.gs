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

// The "This Week" tab: where the truck will be. Blank rows don't show on the website.
const WEEK_SHEET = "This Week";
const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

// The "Settings" tab: the owner changes these from the Google Sheets app.
const SETTINGS_SHEET = "Settings";
const SETTINGS = [
  // [label shown in column A, key the website reads, starting value, dropdown choices or null]
  ["Online ordering", "onlineOrdering", "ON", ["ON", "OFF"]],
  ["Pay now with card", "payNow", "ON", ["ON", "OFF"]],
  ["Today's location", "todayLocation", "", null],
  ["Today's hours", "todayHours", "", null],
  ["Message when ordering is OFF", "closedMessage", "Online ordering is closed right now. Come see us at the window!", null],
];

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
  setupSettings_();
  setupWeek_();
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

/** The website reads the Settings tab through this. */
function doGet() {
  const settings = {};
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
    if (sh) {
      const vals = sh.getRange(2, 1, SETTINGS.length, 2).getDisplayValues();
      SETTINGS.forEach(([label, key], i) => { settings[key] = vals[i][1]; });
    }
  } catch (e) {}
  const week = [];
  try {
    const wk = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WEEK_SHEET);
    if (wk) {
      wk.getRange(2, 1, DAYS.length, 3).getDisplayValues().forEach(([day, where, hours]) => {
        if (String(where).trim()) week.push({ day: day, where: where, when: hours });
      });
    }
  } catch (e) {}
  return json_({ ok: true, message: "Hot Diggity Dog order receiver is running.", settings: settings, week: week });
}

function setupWeek_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(WEEK_SHEET);
  if (!sh) sh = ss.insertSheet(WEEK_SHEET);
  sh.getRange(1, 1, 1, 3).setValues([["Day", "Where (leave blank if not out)", "Hours"]])
    .setFontWeight("bold").setBackground("#D62828").setFontColor("#FFFFFF");
  DAYS.forEach((d, i) => sh.getRange(i + 2, 1).setValue(d).setFontWeight("bold"));
  sh.getRange(10, 1).setValue("Tip: update this at the start of each week. Blank days don't show on the website.")
    .setFontColor("#6B5546").setFontStyle("italic");
  sh.setColumnWidth(1, 120); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 160);
  sh.setFrozenRows(1);
}

function setupSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) sh = ss.insertSheet(SETTINGS_SHEET);
  sh.getRange(1, 1, 1, 3).setValues([["Setting", "Value", "How to use"]])
    .setFontWeight("bold").setBackground("#1F6B57").setFontColor("#FFFFFF");
  const help = {
    onlineOrdering: "OFF hides all order buttons on the website",
    payNow: "OFF = customers can only pay at the window",
    todayLocation: "Type an address or place. Leave blank to hide",
    todayHours: "Example: 11am - 2pm",
    closedMessage: "Shown on the website when ordering is OFF",
  };
  SETTINGS.forEach(([label, key, start, choices], i) => {
    const row = i + 2;
    sh.getRange(row, 1).setValue(label).setFontWeight("bold");
    const cell = sh.getRange(row, 2);
    if (cell.getDisplayValue() === "") cell.setValue(start);
    if (choices) cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(choices, true).build());
    sh.getRange(row, 3).setValue(help[key]).setFontColor("#6B5546");
  });
  sh.setColumnWidth(1, 210); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 330);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("ON").setBackground("#C9EBD9").setRanges([sh.getRange("B2:B3")]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("OFF").setBackground("#F4C7C7").setRanges([sh.getRange("B2:B3")]).build(),
  ];
  sh.setConditionalFormatRules(rules);
  sh.setFrozenRows(1);
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

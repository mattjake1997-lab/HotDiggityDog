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

// The "Menu" tab: he changes prices and what can be ordered online.
const MENU_SHEET = "Menu";
const STATUS_CHOICES = ["ORDER ONLINE", "ASK AT WINDOW", "HIDE"];
const MENU_DEFAULT = [
  ["Hot dog styles", "Firecracker", 7, "ORDER ONLINE"],
  ["Hot dog styles", "South of the Border", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Fiery Diggity", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Sweet & Spicy", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Spanish Dog", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Chili Cheese Dog", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Cali Dog", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Everything Dog", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Taco Dog", 7, "ORDER ONLINE"],
  ["Hot dog styles", "German Dog", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Chicago Dog", 7, "ORDER ONLINE"],
  ["Hot dog styles", "BLT Dog", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Street Dog", 7, "ORDER ONLINE"],
  ["Hot dog styles", "Build Your Own Dog", 7, "ORDER ONLINE"],
  ["Combos", "Any Hot Dog Combo", 10, "ORDER ONLINE"],
  ["Combos", "Crispy Chicken Sandwich Combo", 11, "ASK AT WINDOW"],
  ["Combos", "BBQ Sandwich Combo", 11, "ORDER ONLINE"],
  ["Combos", "Mozzarella Sticks Combo", 12, "ORDER ONLINE"],
  ["Combos", "Any Dog + Lemon Shake-Up", 12, "ASK AT WINDOW"],
  ["Chicken & sandwiches", "Chicken Tenders (3 pc)", 7, "ORDER ONLINE"],
  ["Chicken & sandwiches", "Chicken Tenders (5 pc)", 9, "ORDER ONLINE"],
  ["Chicken & sandwiches", "BBQ Pulled Pork or Chicken Sandwich", 8, "ORDER ONLINE"],
  ["Chicken & sandwiches", "Crispy Chicken Sandwich", 8, "ASK AT WINDOW"],
  ["Chicken & sandwiches", "Corn Dog", 7, "ORDER ONLINE"],
  ["Chicken & sandwiches", "Pork Parfait", 9, "ASK AT WINDOW"],
  ["Nachos & fries", "Loaded Nachos", 10, "ORDER ONLINE"],
  ["Nachos & fries", "Regular Nachos", 7, "ORDER ONLINE"],
  ["Nachos & fries", "Chili Cheese Fries", 7, "ORDER ONLINE"],
  ["Nachos & fries", "Loaded Fries", 10, "ORDER ONLINE"],
  ["Sides & drinks", "Fries", 4, "ORDER ONLINE"],
  ["Sides & drinks", "Chips", 2, "ORDER ONLINE"],
  ["Sides & drinks", "Coleslaw", 2, "ORDER ONLINE"],
  ["Sides & drinks", "Mozzarella Sticks (6 pc)", 7, "ORDER ONLINE"],
  ["Sides & drinks", "Side of Cheese Sauce", 2, "ORDER ONLINE"],
  ["Sides & drinks", "Side of Chili Sauce", 2, "ORDER ONLINE"],
  ["Sides & drinks", "Drink", 2, "ORDER ONLINE"],
  ["Sides & drinks", "Lemon Shake-Up (24 oz)", 7, "ASK AT WINDOW"],
  ["Sides & drinks", "Lemon Shake-Up Refill", 4, "ASK AT WINDOW"],
  ["Specials (wings sign)", "Boneless Wings (6 pc)", 7, "HIDE"],
  ["Specials (wings sign)", "Bone-In Wings (5 pc)", 10, "HIDE"],
  ["Specials (wings sign)", "Chicken Drumsticks (3 pc)", 7, "HIDE"],
  ["Specials (wings sign)", "Chicken Drumsticks (5 pc)", 9, "HIDE"],
  ["West African", "Peanut Sauce Chicken", 10, "ASK AT WINDOW"]
];

// The "Settings" tab: the owner changes these from the Google Sheets app.
const SETTINGS_SHEET = "Settings";
const SETTINGS = [
  // [label shown in column A, key the website reads, starting value, dropdown choices or null]
  ["Online ordering", "onlineOrdering", "ON", ["ON", "OFF"]],
  ["Pay now with card", "payNow", "ON", ["ON", "OFF"]],
  ["Today's location", "todayLocation", "", null],
  ["Today's hours", "todayHours", "", null],
  ["Message when ordering is OFF", "closedMessage", "Online ordering is closed right now. Come see us at the window!", null],
  ["Send order emails to", "notifyEmail", NOTIFY_EMAIL, null],
];

/** Adds a "Hot Diggity Dog" menu to the sheet (on a computer). */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("🌭 Hot Diggity Dog")
    .addItem("Reset orders (start over at #101)", "resetOrders")
    .addToUi();
}

/** Moves every order to the "Archive" tab, clears Orders, and restarts numbering at #101. */
function resetOrders() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert("Reset orders?",
    "All current orders will be moved to the Archive tab and order numbers will start again at #" + FIRST_ORDER_NUMBER + ".",
    ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getSheet_();
  const rows = sh.getLastRow() - 1;
  if (rows > 0) {
    let ar = ss.getSheetByName("Archive");
    if (!ar) {
      ar = ss.insertSheet("Archive");
      ar.appendRow(["Archived on"].concat(HEADERS));
      ar.getRange(1, 1, 1, HEADERS.length + 1).setFontWeight("bold").setBackground("#6B5546").setFontColor("#FFFFFF");
      ar.setFrozenRows(1);
    }
    const data = sh.getRange(2, 1, rows, HEADERS.length).getValues();
    const stamp = new Date();
    ar.getRange(ar.getLastRow() + 1, 1, rows, HEADERS.length + 1).setValues(data.map(r => [stamp].concat(r)));
    sh.deleteRows(2, rows);
  }
  ui.alert("Done", "Orders cleared. The next order will be #" + FIRST_ORDER_NUMBER + ".", ui.ButtonSet.OK);
}

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
  getEventSheet_();
  setupMenu_();
  setupHowTo_();
  // Ask once for email permission so new-order emails work.
  if (NOTIFY_EMAIL) MailApp.getRemainingDailyQuota();
}

/** Receives an order from the website. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.type === "event") return saveEvent_(d);
    const sh = getSheet_();
    const num = FIRST_ORDER_NUMBER + Math.max(0, sh.getLastRow() - 1);
    const items = (d.items || []).map(String).join("\n");
    sh.appendRow([num, new Date(), "NEW", clean_(d.name), "'" + clean_(d.phone), items,
                  clean_(d.notes), clean_(d.total), clean_(d.payment || "Pay at window")]);
    const to = notifyEmail_();
    if (to) {
      MailApp.sendEmail(to, "🌭 New order #" + num + " – " + clean_(d.name),
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

/** Saves an event booking request to the "Event Requests" tab. */
const EVENT_SHEET = "Event Requests";
function getEventSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(EVENT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(EVENT_SHEET);
    sh.appendRow(["Received", "Status", "Name", "Phone", "Event date", "Guests", "Details"]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#F2711C").setFontColor("#FFFFFF");
    sh.getRange("B2:B").setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(["NEW", "CALLED", "BOOKED", "DECLINED"], true).build());
    sh.setColumnWidth(7, 380);
    sh.getRange("G:G").setWrap(true);
  }
  return sh;
}

function saveEvent_(d) {
  const sh = getEventSheet_();
  sh.appendRow([new Date(), "NEW", clean_(d.name), "'" + clean_(d.phone), clean_(d.date),
                clean_(d.guests || "not sure"), clean_(d.details || "")]);
  const to = notifyEmail_();
  if (to) {
    MailApp.sendEmail(to, "🎉 Event request – " + clean_(d.name) + " (" + clean_(d.date) + ")",
      "Event booking request\nName: " + clean_(d.name) + "\nPhone: " + clean_(d.phone) +
      "\nDate: " + clean_(d.date) + "\nGuests: " + clean_(d.guests || "not sure") +
      "\nDetails: " + clean_(d.details || "-"));
  }
  return json_({ ok: true });
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
  const menu = [];
  try {
    const mn = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MENU_SHEET);
    if (mn && mn.getLastRow() > 1) {
      mn.getRange(2, 1, mn.getLastRow() - 1, 4).getDisplayValues().forEach(([section, item, price, status]) => {
        if (String(item).trim()) menu.push({ item: item, price: price, status: status });
      });
    }
  } catch (e) {}
  return json_({ ok: true, message: "Hot Diggity Dog order receiver is running.", settings: settings, week: week, menu: menu });
}

function setupMenu_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MENU_SHEET);
  const isNew = !sh;
  if (!sh) sh = ss.insertSheet(MENU_SHEET);
  if (isNew || sh.getLastRow() < 2) {
    sh.clear();
    sh.appendRow(["Section", "Item (don't change the name)", "Price", "Online status"]);
    sh.getRange(2, 1, MENU_DEFAULT.length, 4).setValues(MENU_DEFAULT);
  } else {
    // Add any new items that aren't in his sheet yet, without touching his changes.
    const have = new Set(sh.getRange(2, 2, sh.getLastRow() - 1, 1).getDisplayValues().map(r => String(r[0]).trim()));
    MENU_DEFAULT.filter(r => !have.has(r[1])).forEach(r => sh.appendRow(r));
  }
  sh.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#1F6B57").setFontColor("#FFFFFF");
  sh.setFrozenRows(1);
  const n = sh.getLastRow() - 1;
  sh.getRange(2, 3, n, 1).setNumberFormat("$0.00");
  sh.getRange(2, 4, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS_CHOICES, true).build());
  const rng = sh.getRange(2, 1, n, 4);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$D2="ORDER ONLINE"').setBackground("#C9EBD9").setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$D2="ASK AT WINDOW"').setBackground("#FFE9A8").setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$D2="HIDE"').setBackground("#E6E6E6").setRanges([rng]).build(),
  ]);
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 300); sh.setColumnWidth(3, 90); sh.setColumnWidth(4, 160);
  // Warn (not block) if someone edits the Section or Item name columns.
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  sh.getRange(2, 1, n, 2).protect().setDescription("Item names must match the website").setWarningOnly(true);
}

function setupHowTo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("📖 How To");
  if (!sh) sh = ss.insertSheet("📖 How To", 0);
  sh.clear();
  const lines = [
    ["📖 HOW TO RUN YOUR WEBSITE", "title"],
    ["Everything below is done in this Google Sheet. Changes show on the website the next time someone opens it.", "note"],
    ["", ""],
    ["🧾 SEE NEW ORDERS", "head"],
    ["1. Tap the ORDERS tab at the bottom.", ""],
    ["2. New orders show at the bottom of the list with the word NEW.", ""],
    ["3. Tap the NEW box to change it: MAKING → READY → PICKED UP.", ""],
    ["4. Customers show you their ORDER NUMBER at the window. Match it to the first column.", ""],
    ["5. \"Pay at window\" = take their money. \"Paying online\" = they already paid. Check the Square app.", ""],
    ["", ""],
    ["⛔ TURN ONLINE ORDERS OFF (busy, closed, or something is wrong)", "head"],
    ["1. Tap the SETTINGS tab.", ""],
    ["2. Next to \"Online ordering\" tap ON and change it to OFF.", ""],
    ["3. The website now shows just the menu. Turn it back ON when you're ready.", ""],
    ["", ""],
    ["💳 TURN CARD PAYMENTS OFF (keep pay at window)", "head"],
    ["1. SETTINGS tab → \"Pay now with card\" → change ON to OFF.", ""],
    ["", ""],
    ["📍 WHERE THE TRUCK IS TODAY", "head"],
    ["1. SETTINGS tab → type the address next to \"Today's location\".", ""],
    ["2. Type the hours next to \"Today's hours\" (example: 11am - 2pm).", ""],
    ["3. Erase them when you're done for the day.", ""],
    ["", ""],
    ["📅 THIS WEEK'S SCHEDULE (do this every Monday)", "head"],
    ["1. Tap the THIS WEEK tab.", ""],
    ["2. Next to each day, type where you'll be and the hours.", ""],
    ["3. Leave a day blank if you're not going out. Blank days don't show.", ""],
    ["", ""],
    ["🌭 CHANGE A PRICE OR WHAT PEOPLE CAN ORDER", "head"],
    ["1. Tap the MENU tab.", ""],
    ["2. To change a price: tap the price and type the new one.", ""],
    ["3. Change the Online status column:", ""],
    ["      ORDER ONLINE = people can order and pay for it on the website", ""],
    ["      ASK AT WINDOW = it shows on the menu, but they have to ask you", ""],
    ["      HIDE = it does not show on the website at all", ""],
    ["4. Don't change the item NAMES. The website needs them spelled exactly the same.", ""],
    ["   Printed menus and signs don't change by themselves. Ask Matt for new ones.", "note"],
    ["", ""],
    ["🎉 EVENT REQUESTS", "head"],
    ["1. Tap the EVENT REQUESTS tab. Call the customer, then change NEW to CALLED, BOOKED, or DECLINED.", ""],
    ["", ""],
    ["💵 REFUND A CARD ORDER", "head"],
    ["1. Open the Square app → Transactions → tap the payment → Issue refund.", ""],
    ["   No refunds for orders that aren't picked up (this is on the website).", "note"],
    ["", ""],
    ["🆘 IF SOMETHING LOOKS WRONG", "head"],
    ["1. SETTINGS → Online ordering → OFF. Customers order at the window like normal.", ""],
    ["2. Call or text Matt.", ""],
    ["", ""],
    ["⚠️ PLEASE DON'T: delete tabs, rename tabs, or open Extensions → Apps Script.", "warn"],
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines.map(l => [l[0]]));
  sh.setColumnWidth(1, 900);
  sh.getRange(1, 1, lines.length, 1).setFontSize(14).setWrap(true).setVerticalAlignment("middle");
  lines.forEach((l, i) => {
    const c = sh.getRange(i + 1, 1);
    if (l[1] === "title") c.setFontSize(22).setFontWeight("bold").setBackground("#F6C619");
    if (l[1] === "head") c.setFontSize(16).setFontWeight("bold").setBackground("#D62828").setFontColor("#FFFFFF");
    if (l[1] === "note") c.setFontStyle("italic").setFontColor("#6B5546");
    if (l[1] === "warn") c.setFontWeight("bold").setBackground("#F4C7C7");
  });
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(1);
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
    notifyEmail: "Who gets an email for each new order (blank = no emails)",
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

function notifyEmail_() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
    if (sh) {
      const i = SETTINGS.findIndex(r => r[1] === "notifyEmail");
      if (i >= 0) return String(sh.getRange(i + 2, 2).getDisplayValue()).trim();
    }
  } catch (e) {}
  return NOTIFY_EMAIL;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

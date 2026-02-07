/**
 * @author nathaliacauas
 */

/**
 * Reminder by date on Google Sheets (D-1 e D0)
 * - Sends emails on the day and on the day before the schedule
 * - Different messages for D-1 and D0
 * - Body using data on the respective row
 * - Adds status update to prevent repetitive emails
 *
 * IMPORTANT (TIME AND TIMEZONE):
 * - For the hour field, I use getDisplayValues() to get exatcly what was filled in the form
 * and then convert it correctly.
 */

const CONFIG = {
  sheetName: "ADD_SHEET_NAME",
  headerRow: "ADD_ROW_NUMEBER",
  timezone: "ADD_TIMEZONE",

  emailTo: "ADD_MAIN_MAIL",

  // You can add on multiline format
  emailBcc: `
ADD_OTHER_EMAIL
`,
  // "Control column"
  statusHeader: "ADD_NAME",

  // Column displaying the dates
  eventDateHeader: "ADD_NAME",

  // Hour (0–23) for the daily trigger 
  triggerHour: 6
};

const FIELDS = {
  emailAddress: "Email Address",
  any_x: "Any",
  other_x: "Other",
  column_x: "Column",
  you_x: "You",
  need_x: "Need"
};

function sendD1D0() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error(`Tab "${CONFIG.sheetName}" not found.`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= CONFIG.headerRow) return;

  // Header
  const headers = sheet.getRange(CONFIG.headerRow, 1, 1, lastCol).getValues()[0];
  const headerMap = buildHeaderMap(headers);

  // Control column setting
  let statusCol = headerMap[normalize(CONFIG.statusHeader)];
  if (!statusCol) {
    statusCol = lastCol + 1;
    sheet.getRange(CONFIG.headerRow, statusCol).setValue(CONFIG.statusHeader);
    headerMap[normalize(CONFIG.statusHeader)] = statusCol;
  }

  // Date column
  const eventDateCol = headerMap[normalize(CONFIG.eventDateHeader)];
  if (!eventDateCol) {
    throw new Error(
      `Couldn't find: "${CONFIG.eventDateHeader}". ` +
      `Update CONFIG.eventDateHeader with the correct name.`
    );
  }

  // Index (0 based)
  const idx = {
    emailAddress: getColIndex(headerMap, FIELDS.emailAddress),
    any_x: getColIndex(headerMap, FIELDS.any_x),
    other_x: getColIndex(headerMap, FIELDS.other_x),
    column_x: getColIndex(headerMap, FIELDS.column_x),
    you_x: getColIndex(headerMap, FIELDS.you_x),
    need_x: getColIndex(headerMap, FIELDS.need_x)
  };

  // Reads all rows
  const effectiveLastCol = Math.max(lastCol, statusCol);
  const numRows = lastRow - CONFIG.headerRow;
  const range = sheet.getRange(CONFIG.headerRow + 1, 1, numRows, effectiveLastCol);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();

  // Set correct information based on timezone (SETUP ON CONFIG)
  const today = startOfDay(new Date(), CONFIG.timezone);
  const tomorrow = addDays(today, 1);

  // Set up receivers
  const to = String(CONFIG.emailTo || "").trim();
  const bcc = getBccList(); 

  let sent = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowDisplay = displayValues[i];

    // Set up dates
    const eventRaw = row[eventDateCol - 1];
    const eventDateVal = parseDateIfNeeded(eventRaw);
    if (!eventDateVal) continue;

    const eventDay = startOfDay(eventDateVal, CONFIG.timezone);

    // Verify control column
    const statusVal = String(row[statusCol - 1] ?? "").trim().toUpperCase();

    const isD0 = sameDay(eventDay, today);
    const isD1 = sameDay(eventDay, tomorrow);

    if (!to) continue;

    const msgBase = createMessage(row, rowDisplay, idx);

    // D-1
    if (isD1 && !statusVal.includes("D1SENT")) {
      const subject = "Reminder: event is tomorrow";
      const body =
        `Hi,\n\n` +
        `This is a reminder (${formatDate(eventDay, CONFIG.timezone)}).\n\n` +
        `${msgBase}\n\n` +
        `— Email sent automatically.`;

      MailApp.sendEmail({
        to,
        bcc,     
        subject,
        body
      });

      sheet.getRange(CONFIG.headerRow + 1 + i, statusCol)
        .setValue(appendStatus(statusVal, "D1SENT"));
      sent++;
      continue;
    }

    // D0
    if (isD0 && !statusVal.includes("D0SENT")) {
      const subject = "Reminder: the event is today.";
      const body =
        `Hi,\n\n` +
        `This is a reminder (${formatDate(eventDay, CONFIG.timezone)}).\n\n` +
        `${msgBase}\n\n` +
        `— Email sent automatically.`;

      MailApp.sendEmail({
        to,
        bcc,      
        subject,
        body
      });

      sheet.getRange(CONFIG.headerRow + 1 + i, statusCol)
        .setValue(appendStatus(statusVal, "D0SENT"));
      sent++;
    }
  }

  Logger.log(`E-mails sent: ${sent}`);
}

/**
 * Format the body of the email
 * OBS: For the hour I use (rowDisplay) and covert to 24h-format.
 */
function createMessage(row, rowDisplay, idx) {
  const any_x = safe(row[idx.any_x]);
  const other_x = safe(row[idx.other_x]);
  const column_x = safe(row[idx.column_x]);
  const you_x = safe(row[idx.you_x]);

  const hourDisplay = String(rowDisplay[idx.need_x] ?? "").trim();
  const finalHour = parseTimeTo24h(hourDisplay) || hourDisplay;

  return (
    `Shcedule for ['${any_x}'].\n` +
    `Using that ['${other_x}'] with that ['${column_x}'] and ['${you_x}'] ` +
    `considering ['${finalHour}'].`
  );
}

/* =========================
   Utils
========================= */

function getBccList() {

  return String(CONFIG.emailBcc || "")
    .split(/\r?\n|,/)         
    .map(e => e.trim())
    .filter(e => e.length > 0)
    .join(",");
}

function buildHeaderMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[normalize(String(h))] = i + 1; 
  });
  return map;
}

function normalize(s) {
  return String(s).trim().toLowerCase();
}

// Retorns index 0 for array row[]
function getColIndex(headerMap, headerName) {
  const col = headerMap[normalize(headerName)];
  if (!col) throw new Error(`Header not found: "${headerName}"`);
  return col - 1;
}

function safe(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.timezone, "MM-dd-yyyy");
  return String(v).trim();
}

function formatDate(d, tz) {
  return Utilities.formatDate(d, tz, "MM-dd-yyyy");
}

/**
 * Normalize data: ALWAYS uses yyyy-MM-dd
 */
function startOfDay(d, tz) {
  const s = Utilities.formatDate(d, tz, "yyyy-MM-dd");
  const [y, m, day] = s.split("-").map(n => parseInt(n, 10));
  return new Date(y, m - 1, day);
}

function addDays(d, days) {
  const x = new Date(d.getTime());
  x.setDate(x.getDate() + days);
  return x;
}

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear() &&
         a.getMonth() === b.getMonth() &&
         a.getDate() === b.getDate();
}

function appendStatus(existingUpper, flag) {
  const base = String(existingUpper || "").trim();
  if (!base) return flag;
  if (base.includes(flag)) return base;
  return `${base} | ${flag}`;
}

/**
 * Accepts dates as Date or Text on the formats:
 * - M-d-yyyy / MM-dd-yyyy
 * - yyyy-MM-dd
 * - dd/MM/yyyy ou MM/dd/yyyy 
 * - fallback: new Date(string)
 */
function parseDateIfNeeded(v) {
  if (v instanceof Date) return v;

  const s = String(v ?? "").trim();
  if (!s) return null;

  // M-d-yyyy ou MM-dd-yyyy
  let m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2]));

  // yyyy-MM-dd
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  // dd/MM/yyyy ou MM/dd/yyyy
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const a = Number(m[1]);
    const b = Number(m[2]);
    const y = Number(m[3]);
    if (a > 12) return new Date(y, b - 1, a); // dd/MM/yyyy
    return new Date(y, a - 1, b);             // MM/dd/yyyy
  }

  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return parsed;

  return null;
}

/**
 * Converts to 24h:
 * - "6:00:00 PM" -> "18:00"
 * - "6:00 PM"    -> "18:00"
 * - "18:00:00"   -> "18:00"
 * - "18:00"      -> "18:00"
 */
function parseTimeTo24h(s) {
  if (!s) return "";

  // "6:00:00 PM" ou "6:00 PM"
  let m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)$/i);
  if (m) {
    let h = Number(m[1]);
    const min = Number(m[2]);
    const ap = m[4].toUpperCase();
    if (ap === "PM" && h !== 12) h += 12;
    if (ap === "AM" && h === 12) h = 0;
    return `${String(h).padStart(2, "0")}:${String(min).padStart(2, "0")}`;
  }

  // "18:00:00" ou "18:00"
  m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (m) {
    return `${String(Number(m[1])).padStart(2, "0")}:${m[2]}`;
  }

  return "";
}

/**
 * Run ONCE to create or update the daily trigger diário.
 */
function criarGatilhoDiario() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "sendD1D0") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("sendD1D0")
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.triggerHour)
    .create();

  Logger.log("Daily trigger created.");
}
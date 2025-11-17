// ===== CONFIGURATION =====
const YEAR = 2025;

const MONTH_SHEETS = [
  "January","February","March","April","May","June",
  "July","August","September","October","November","December"
];

const ADMIN_EMAIL = "admin@example.com";

const EMPLOYEE_ACCESS = {
  4: { name: "John", email: "john@example.com" },   // Column D
  5: { name: "Mary", email: "mary@example.com" },   // Column E
  6: { name: "Alex", email: "alex@example.com" },   // Column F
  7: { name: "Lisa", email: "lisa@example.com" }    // Column G
};

const CLOSED_WEEKDAYS = [0]; // Sundays
const HOLIDAYS = [
  "2025-01-01", // New Year's Day
  "2025-04-25", // Example
  "2025-12-25"  // Christmas
];
const TOTAL_LEAVE_DAYS = 20;
const MAX_PEOPLE_PER_DAY = 1;
const SUMMARY_SHEET = "Summary";
const BACKUP_FOLDER_ID = "folder_id"; // Folder ID in Google Drive to save the back ups
const MAX_BACKUPS = 3;

// ===== CREATE SUMMARY SHEET =====
function createSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let summary = ss.getSheetByName(SUMMARY_SHEET);
  if (!summary) {
    summary = ss.insertSheet(SUMMARY_SHEET);
  }
  summary.clear();

  summary.appendRow(["Employee Name", "Days Taken", "Remaining Leave Days"]);

  for (const { name } of Object.values(EMPLOYEE_ACCESS)) {
    summary.appendRow([name, 0, TOTAL_LEAVE_DAYS]);
  }
  summary.autoResizeColumns(1, 3);
}

// ===== UPDATE SUMMARY =====
function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName(SUMMARY_SHEET);
  if (!summary) return;
  const leaveCounts = {};
  for (const { name } of Object.values(EMPLOYEE_ACCESS)) {
    leaveCounts[name] = 0;
  }

  for (const monthName of MONTH_SHEETS) {
    const sheet = ss.getSheetByName(monthName);
    if (!sheet) continue;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    for (const [col, emp] of Object.entries(EMPLOYEE_ACCESS)) {
      const values = sheet.getRange(2, Number(col), lastRow - 1, 1).getValues().flat();
      leaveCounts[emp.name] += values.filter(v => v === true).length;
    }
  }

  // Update existing summary
  const lastRow = summary.getLastRow();
  const names = summary.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (!leaveCounts.hasOwnProperty(name)) continue;
    const taken = leaveCounts[name];
    const remaining = TOTAL_LEAVE_DAYS - taken;

    summary.getRange(i + 2, 2).setValue(taken);
    summary.getRange(i + 2, 3).setValue(remaining);
  }
}

function protectSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName(SUMMARY_SHEET);
  if (!summary) return;

  // Remove old protections (so it doesn't stack each time)
  const protections = summary.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(p => p.remove());

  // Protect the whole sheet (view only for employees)
  const protection = summary.protect().setDescription("Summary view-only");
  protection.removeEditors(protection.getEditors());
  protection.addEditor(ADMIN_EMAIL);
  protection.setWarningOnly(false);
}

// ===== ON EDIT: Check for FULL days =====
function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Ignore summary or other sheets
  if (!MONTH_SHEETS.includes(sheetName)) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  // Only run for employee columns
  if (!EMPLOYEE_ACCESS[col]) return;

  // Skip header or closed days
  if (row === 1) return;
  const statusCell = sheet.getRange(row, 3);
  if (statusCell.getValue() === "CLOSED") return;

  // Count how many are checked in this row
  const values = sheet.getRange(row, FIRST_EMPLOYEE_COL, 1, LAST_EMPLOYEE_COL - FIRST_EMPLOYEE_COL + 1).getValues()[0];

  const checkedCount = values.filter(v => v === true).length;

  // Check if the employee still has remaining leave
  const emp = EMPLOYEE_ACCESS[col];
  if (e.value === "TRUE" && !employeeHasLeaveRemaining(emp.name)) {
    e.range.setValue(false);
    SpreadsheetApp.getActiveSpreadsheet().toast(`${emp.name} has no remaining leave days!`, "Leave Denied");
    return;
  }


  // Update the status column
  if (checkedCount > MAX_PEOPLE_PER_DAY) {
    // Revert user’s change
    e.range.setValue(false);
    SpreadsheetApp.getActiveSpreadsheet().toast("This day is already full!", "Leave Denied");
    statusCell.setValue("FULL");
  } else {
    statusCell.setValue(checkedCount >= MAX_PEOPLE_PER_DAY ? "FULL" : "");
  }
  // Update summary after any valid edit
  updateSummary();
}

// ===== MAIN FUNCTION =====
function generateMonthlySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheets = ss.getSheets().map(s => s.getName());

  // Start by generating January (index 0)
  generateMonthByIndex(0, ss, existingSheets);
  shareSpreadsheetWithEmployees();
}

// ===== MONTH-BY-MONTH EXECUTION =====
function generateMonthByIndex(index, ss, existingSheets) {
  if (index >= MONTH_SHEETS.length) {
    Logger.log("All months generated.");
    createSummarySheet();
    protectSummarySheet();
    return;
  }
  const monthName = MONTH_SHEETS[index];
  const daysInMonth = new Date(YEAR, index + 1, 0).getDate();
  // Create or reset the sheet
  let sheet;
  if (existingSheets.includes(monthName)) {
    sheet = ss.getSheetByName(monthName);
    sheet.clear();
  } else {
    sheet = ss.insertSheet(monthName);
  }
  // Remove old protections
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  // Headers
  const headers = ["Date", "Day", "Status", ...Object.values(EMPLOYEE_ACCESS).map(e => e.name)];
  sheet.appendRow(headers);
  // Fill days
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(YEAR, index, d);
    const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-yyyy");
    const dayName = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][date.getDay()];
    const iso = date.toISOString().slice(0, 10);
    const isSunday = CLOSED_WEEKDAYS.includes(date.getDay());
    const isHoliday = HOLIDAYS.includes(iso);
    const closed = isSunday || isHoliday;
    const row = [formattedDate, dayName, closed ? "CLOSED" : ""];
    sheet.appendRow(row);
    if (closed) {
      sheet.getRange(d + 1, 1, 1, LAST_EMPLOYEE_COL).setBackground("#f4cccc");
    } else {
      // Insert checkboxes for open days
      sheet
        .getRange(d + 1, FIRST_EMPLOYEE_COL, 1, LAST_EMPLOYEE_COL - FIRST_EMPLOYEE_COL + 1)
        .insertCheckboxes();
    }
  }
  // Apply protections
  applySheetProtections(sheet, monthName, daysInMonth);
  // Resize columns neatly
  sheet.autoResizeColumns(1, LAST_EMPLOYEE_COL);
  // Schedule next month after this one completes
  if (index + 1 < MONTH_SHEETS.length) {
    ScriptApp.newTrigger("continueMonthGeneration")
      .timeBased()
      .after(60 * 1000) // 1 minute delay between months, used so that google does not time out the execution
      .create();
    // Store progress
    PropertiesService.getScriptProperties().setProperty("LAST_MONTH_INDEX", index + 1);
  } else {
    Logger.log("All months done. Creating summary...");
    createSummarySheet();
    protectSummarySheet();
  }
}

// ===== CONTINUATION TRIGGER =====
function continueMonthGeneration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheets = ss.getSheets().map(s => s.getName());
  const index = Number(PropertiesService.getScriptProperties().getProperty("LAST_MONTH_INDEX") || 0);
  generateMonthByIndex(index, ss, existingSheets);
}

// ===== PROTECTION HELPER =====
function applySheetProtections(sheet, monthName, daysInMonth) {
  const FIRST_EMPLOYEE_COL = Math.min(...Object.keys(EMPLOYEE_ACCESS));
  const LAST_EMPLOYEE_COL = Math.max(...Object.keys(EMPLOYEE_ACCESS));
  // Protect entire sheet (admin-only)
  const sheetProtection = sheet.protect().setDescription(`${monthName} - Admin only`);
  sheetProtection.removeEditors(sheetProtection.getEditors());
  sheetProtection.addEditor(ADMIN_EMAIL);
  sheetProtection.setWarningOnly(false);
  // Allow employees only on their checkboxes
  const unprotectedRanges = [];
  for (const [colStr, emp] of Object.entries(EMPLOYEE_ACCESS)) {
    const col = Number(colStr);
    const openDayRanges = [];
    for (let d = 1; d <= daysInMonth; d++) {
      const status = sheet.getRange(d + 1, 3).getValue();
      if (status === "CLOSED") continue;
      openDayRanges.push(sheet.getRange(d + 1, col));
    }
    if (openDayRanges.length > 0) {
      unprotectedRanges.push(...openDayRanges);
    }
  }
  // Apply unprotected ranges for sheet protection
  if (unprotectedRanges.length > 0) {
    sheetProtection.setUnprotectedRanges(unprotectedRanges);
  }
  // Individual range protections (per-cell)
  for (const [colStr, emp] of Object.entries(EMPLOYEE_ACCESS)) {
    const col = Number(colStr);
    if (!emp || !emp.email) continue;
    for (let d = 1; d <= daysInMonth; d++) {
      const status = sheet.getRange(d + 1, 3).getValue();
      if (status === "CLOSED") continue;
      try {
        const range = sheet.getRange(d + 1, col);
        const prot = range.protect().setDescription(`${monthName} - ${emp.name} (Row ${d + 1})`);
        prot.removeEditors(prot.getEditors());
        prot.addEditors([ADMIN_EMAIL, emp.email]);
        prot.setWarningOnly(false);
      } catch (err) {
        Logger.log(`Skipping invalid range at row ${d + 1}, col ${col}: ${err}`);
      }
    }
  }
}

function employeeHasLeaveRemaining(employeeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName(SUMMARY_SHEET);
  if (!summary) return true; // If summary missing, fail open

  const lastRow = summary.getLastRow();
  if (lastRow < 2) return true;
  const names = summary.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const remainingDays = summary.getRange(2, 3, lastRow - 1, 1).getValues().flat();
  for (let i = 0; i < names.length; i++) {
    if (names[i] === employeeName) {
      return remainingDays[i] > 0;
    }
  }

  return true; // If name not found, fail open
}

function shareSpreadsheetWithEmployees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emails = Object.values(EMPLOYEE_ACCESS).map(e => e.email);
  ss.addEditors(emails);  // Gives edit access to the whole file
  ss.addEditor(ADMIN_EMAIL); // Ensure admin always has full access
}

function backupSpreadsheetDaily() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalFile = DriveApp.getFileById(ss.getId());
  const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const backupName = `LeaveSystemBackup_${timestamp}`;
  const backupFile = originalFile.makeCopy(backupName, folder);
  logBackupAction(`Backup created: ${backupName}`);

  cleanupOldBackups(); // Clean out old backups after each run
}

function cleanupOldBackups() {
  const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  const files = folder.getFiles();
  // Collect all files
  const backups = [];
  while (files.hasNext()) {
    let f = files.next();
    backups.push({ file: f, date: f.getDateCreated() });
  }
  // Sort newest → oldest
  backups.sort((a, b) => b.date - a.date);
  // Delete anything beyond MAX_BACKUPS
  for (let i = MAX_BACKUPS; i < backups.length; i++) {
    backups[i].file.setTrashed(true);
    logBackupAction(`Deleted old backup: ${backups[i].file.getName()}`);
  }
}

function logBackupAction(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Backup_Log";
  let logSheet = ss.getSheetByName(sheetName);
  // Create if missing
  if (!logSheet) {
    logSheet = ss.insertSheet(sheetName);
    logSheet.appendRow(["Timestamp", "Message"]);
    // Protect sheet: admin only
    const protection = logSheet.protect().setDescription("Admin Only");
    protection.removeEditors(protection.getEditors());
    protection.addEditor(ADMIN_EMAIL);  // Use your predefined admin email
    logSheet.hideSheet();
  }
  // Append log entry
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );
  logSheet.appendRow([timestamp, message]);
}



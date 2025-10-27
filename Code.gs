// -----------------------------------------------------------------------------
// Configuration
// -----------------------------------------------------------------------------

const DATABASE_ID = '1LIQRnQb7lL-6hdSpsOwq-XVwRM6Go6u4q6RkSB2ZCXo';

// This is the CORRECT column mapping for COC_Balance_Detail
const DETAIL_COLS = {
  ENTRY_ID: 0,            // Entry ID
  EMPLOYEE_ID: 1,         // Employee ID
  EMPLOYEE_NAME: 2,       // Employee Name
  CERTIFICATE_ID: 3,      // Certificate ID (NEW - links to monthly certificate)
  RECORD_ID: 4,           // Record ID (from COC_Records)
  MONTH_YEAR: 5,          // Month-Year (e.g., "2025-10")
  DATE_EARNED: 6,         // Date Rendered/Earned
  DAY_TYPE: 7,            // Day Type (Weekday, Weekend, Regular Holiday, etc.)
  HOURS_EARNED: 8,        // Hours Earned
  HOURS_USED: 9,          // Hours Used (for FIFO tracking)
  HOURS_REMAINING: 10,    // Hours Remaining
  CERTIFICATE_ISSUE_DATE: 11, // Certificate Issue Date (NEW - when certificate was issued)
  EXPIRATION_DATE: 12,    // Expiration Date (Certificate Issue Date + 1 year - 1 day)
  STATUS: 13,             // Status (Active, Used, Expired, Cancelled)
  DATE_CREATED: 14,       // Date Created
  CREATED_BY: 15,         // Created By
  LAST_UPDATED: 16,       // Last Updated
  NOTES: 17               // Notes/Remarks
};

// COC_Records sheet column mapping (22 columns)
const RECORD_COLS = {
  RECORD_ID: 0,           // Record ID (unique identifier)
  EMPLOYEE_ID: 1,         // Employee ID
  EMPLOYEE_NAME: 2,       // Employee Name
  MONTH_YEAR: 3,          // Month-Year (e.g., "2025-10")
  DATE_RENDERED: 4,       // Date Rendered (date COC was earned)
  DAY_TYPE: 5,            // Day Type (Weekday, Weekend, Holiday, etc.)
  AM_IN: 6,               // AM In time
  AM_OUT: 7,              // AM Out time
  PM_IN: 8,               // PM In time
  PM_OUT: 9,              // PM Out time
  HOURS_WORKED: 10,       // Hours Worked
  MULTIPLIER: 11,         // Multiplier (based on day type)
  COC_EARNED: 12,         // COC Earned (hours)
  CERTIFICATE_ID: 13,     // Certificate ID (linked to COC_Certificates)
  DATE_RECORDED: 14,      // Date Recorded
  EXPIRATION_DATE: 15,    // Expiration Date
  STATUS: 16,             // Status (Pending, Active, Used, Expired, Cancelled)
  APPROVED_BY: 17,        // Approved By
  APPROVED_DATE: 18,      // Approved Date
  CREATED_BY: 19,         // Created By
  LAST_MODIFIED: 20,      // Last Modified
  MODIFIED_BY: 21         // Modified By
};

// COC_Certificates sheet column mapping (13 columns)
const CERT_COLS = {
  CERTIFICATE_ID: 0,      // Certificate ID (unique identifier)
  EMPLOYEE_ID: 1,         // Employee ID
  EMPLOYEE_NAME: 2,       // Employee Name
  MONTH_YEAR: 3,          // Month-Year (e.g., "2025-10")
  TOTAL_COC_EARNED: 4,    // Total COC Earned (hours)
  NUMBER_OF_RECORDS: 5,   // Number of Records (count)
  ISSUE_DATE: 6,          // Issue Date
  EXPIRATION_DATE: 7,     // Expiration Date
  CERTIFICATE_URL: 8,     // Certificate URL (Google Doc)
  PDF_URL: 9,             // PDF URL (exported PDF)
  STATUS: 10,             // Status (Active, Cancelled, etc.)
  CREATED_DATE: 11,       // Created Date
  CREATED_BY: 12          // Created By
};

// Employees sheet column mapping
const EMP_COLS = {
  EMPLOYEE_ID: 0,         // Employee ID
  FIRST_NAME: 1,          // First Name
  MIDDLE_INITIAL: 2,      // Middle Initial
  LAST_NAME: 3,           // Last Name
  SUFFIX: 4,              // Suffix (Jr., Sr., etc.)
  POSITION: 5,            // Position
  OFFICE: 6,              // Office
  STATUS: 7               // Status (Active, Inactive, etc.)
};

// COC_Ledger sheet column mapping
const LEDGER_COLS = {
  LEDGER_ID: 0,           // Ledger ID (unique identifier)
  EMPLOYEE_ID: 1,         // Employee ID
  EMPLOYEE_NAME: 2,       // Employee Name
  TRANSACTION_DATE: 3,    // Transaction Date
  TRANSACTION_TYPE: 4,    // Transaction Type (Earned, Used, Expired, Adjusted)
  REFERENCE_ID: 5,        // Reference ID (Certificate ID, CTO ID, etc.)
  BALANCE_BEFORE: 6,      // Balance Before
  COC_EARNED: 7,          // COC Earned (hours)
  CTO_USED: 8,            // CTO Used (hours)
  COC_EXPIRED: 9,         // COC Expired (hours)
  BALANCE_ADJUSTMENT: 10, // Balance Adjustment (hours)
  BALANCE_AFTER: 11,      // Balance After
  MONTH_YEAR_EARNED: 12,  // Month-Year Earned (e.g., "2025-10")
  EXPIRATION_DATE: 13,    // Expiration Date
  PROCESSED_BY: 14,       // Processed By
  PROCESSED_DATE: 15,     // Processed Date
  REMARKS: 16             // Remarks/Notes
};

// Status constants
const STATUS_PENDING = 'Pending';
const STATUS_ACTIVE = 'Active';
const STATUS_USED = 'Used';
const STATUS_EXPIRED = 'Expired';
const STATUS_CANCELLED = 'Cancelled';

// Transaction type constants
const TR_TYPE_EARNED = 'Earned';
const TR_TYPE_USED = 'Used';
const TR_TYPE_EXPIRED = 'Expired';
const TR_TYPE_ADJUSTED = 'Adjusted';

// -----------------------------------------------------------------------------
// Migration Functions
// -----------------------------------------------------------------------------

/**
 * MIGRATION: Populates MONTH_YEAR column for all COC_Records that have it empty.
 * This is needed because older records were created without the MONTH_YEAR column,
 * and the new apiListCOCRecordsForMonth function filters by this column.
 *
 * Run this function once to migrate old data.
 */
function migrateCOCRecordsMonthYear() {
  try {
    const db = SpreadsheetApp.openById(DATABASE_ID);
    const recordsSheet = db.getSheetByName('COC_Records');

    if (!recordsSheet) {
      Logger.log('COC_Records sheet not found');
      return { success: false, message: 'COC_Records sheet not found' };
    }

    const data = recordsSheet.getDataRange().getValues();
    let updatedCount = 0;
    let fixedCount = 0;

    Logger.log(`Starting migration. Total rows: ${data.length}`);

    // Start from row 2 (index 1) to skip header
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const monthYearValue = row[RECORD_COLS.MONTH_YEAR];
      const dateRendered = row[RECORD_COLS.DATE_RENDERED];

      let needsUpdate = false;
      let newMonthYear = null;

      // Case 1: MONTH_YEAR is empty
      if (!monthYearValue && dateRendered) {
        const date = new Date(dateRendered);
        if (!isNaN(date.getTime())) {
          const year = date.getFullYear();
          const month = String(date.getMonth() + 1).padStart(2, '0');
          newMonthYear = `${year}-${month}`;
          needsUpdate = true;
          Logger.log(`Row ${i + 1}: EMPTY - Set to ${newMonthYear} from date ${date}`);
        }
      }
      // Case 2: MONTH_YEAR is a Date object (should be a string)
      else if (monthYearValue instanceof Date) {
        const date = new Date(monthYearValue);
        if (!isNaN(date.getTime())) {
          const year = date.getFullYear();
          const month = String(date.getMonth() + 1).padStart(2, '0');
          newMonthYear = `${year}-${month}`;
          needsUpdate = true;
          fixedCount++;
          Logger.log(`Row ${i + 1}: DATE OBJECT - Convert to ${newMonthYear}`);
        }
      }
      // Case 3: MONTH_YEAR is in wrong format (MM-YYYY instead of YYYY-MM)
      else if (typeof monthYearValue === 'string') {
        const trimmed = monthYearValue.trim();
        // Check if it matches MM-YYYY format (e.g., "10-2025")
        const mmYyyyMatch = trimmed.match(/^(\d{2})-(\d{4})$/);
        if (mmYyyyMatch) {
          const month = mmYyyyMatch[1];
          const year = mmYyyyMatch[2];
          newMonthYear = `${year}-${month}`;
          needsUpdate = true;
          fixedCount++;
          Logger.log(`Row ${i + 1}: WRONG FORMAT "${trimmed}" - Fix to ${newMonthYear}`);
        }
        // If it's already in YYYY-MM format, skip
        else if (trimmed.match(/^\d{4}-\d{2}$/)) {
          Logger.log(`Row ${i + 1}: Already correct format: ${trimmed}`);
        }
        // Unknown format - try to fix from date
        else if (dateRendered) {
          const date = new Date(dateRendered);
          if (!isNaN(date.getTime())) {
            const year = date.getFullYear();
            const month = String(date.getMonth() + 1).padStart(2, '0');
            newMonthYear = `${year}-${month}`;
            needsUpdate = true;
            fixedCount++;
            Logger.log(`Row ${i + 1}: UNKNOWN FORMAT "${trimmed}" - Fix to ${newMonthYear} from date`);
          }
        }
      }

      // Update the cell if needed
      if (needsUpdate && newMonthYear) {
        recordsSheet.getRange(i + 1, RECORD_COLS.MONTH_YEAR + 1).setValue(newMonthYear);
        updatedCount++;
      }
    }

    Logger.log(`Migration completed. Updated ${updatedCount} records (${fixedCount} were format fixes).`);
    return {
      success: true,
      updatedCount: updatedCount,
      fixedCount: fixedCount
    };

  } catch (e) {
    Logger.log(`Error in migrateCOCRecordsMonthYear: ${e}`);
    return { success: false, message: e.message };
  }
}

/**
 * DEBUG FUNCTION: Check what data exists for a specific employee and month
 * This helps diagnose why records aren't showing
 */
function debugCOCRecords(employeeId, month, year) {
  try {
    const db = SpreadsheetApp.openById(DATABASE_ID);
    const recordsSheet = db.getSheetByName('COC_Records');
    const data = recordsSheet.getDataRange().getValues();

    const monthYear = `${year}-${String(month).padStart(2, '0')}`;

    Logger.log('=== DEBUG COC RECORDS ===');
    Logger.log(`Looking for: Employee ID = "${employeeId}", Month/Year = "${monthYear}"`);
    Logger.log(`Total rows in sheet: ${data.length}`);
    Logger.log('');

    // Check header
    Logger.log('Headers:');
    Logger.log(data[0]);
    Logger.log('');

    let matchingRecords = 0;
    let recordsForEmployee = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmpId = row[RECORD_COLS.EMPLOYEE_ID];
      const rowMonthYear = row[RECORD_COLS.MONTH_YEAR];
      const rowStatus = row[RECORD_COLS.STATUS];
      const rowDate = row[RECORD_COLS.DATE_RENDERED];

      // Count all records for this employee
      if (rowEmpId === employeeId) {
        recordsForEmployee++;
        Logger.log(`Row ${i + 1}: EmpID="${rowEmpId}", MonthYear="${rowMonthYear}", Status="${rowStatus}", Date="${rowDate}"`);

        // Check if it matches our filter
        if (rowMonthYear === monthYear) {
          matchingRecords++;
          Logger.log(`  ^^^ MATCHES! ^^^`);
        }
      }
    }

    Logger.log('');
    Logger.log(`Summary: Found ${recordsForEmployee} total records for employee ${employeeId}`);
    Logger.log(`Found ${matchingRecords} records matching month/year ${monthYear}`);

    return {
      totalRecordsForEmployee: recordsForEmployee,
      matchingRecords: matchingRecords
    };

  } catch (e) {
    Logger.log(`Error in debugCOCRecords: ${e}`);
    return { error: e.message };
  }
}

/**
 * CORRECTED: Calculates expiration date based on CERTIFICATE ISSUE DATE
 * Formula: Certificate Issue Date + 1 Year - 1 Day
 *
 * Example:
 *   Certificate issued: Nov 5, 2025
 *   Expiration: Nov 4, 2026 (NOT Nov 5, 2026)
 */
function calculateCertificateExpiration(issueDate) {
  const expiration = new Date(issueDate);
  expiration.setFullYear(expiration.getFullYear() + 1);
  expiration.setDate(expiration.getDate() - 1);
  expiration.setHours(0, 0, 0, 0);
  return expiration;
}

/**
 * CORRECTED: Generate monthly COC certificate for an employee
 * This consolidates all COC earned in a specific month
 * 
 * Example: October 2025
 *   - Oct 4: 4 hrs
 *   - Oct 6: 2 hrs  
 *   - Oct 25: 8 hrs
 *   Total: 14 hrs
 *   Certificate issued: Nov 5, 2025
 *   All 3 records expire: Nov 4, 2026
 */
function generateMonthlyCOCCertificate(employeeId, monthYear) {
  const db = getDatabase();
  const recSheet = db.getSheetByName('COC_Records');
  const detailSheet = ensureCOCBalanceDetailSheet();
  const TIME_ZONE = getScriptTimeZone();
  
  if (!recSheet) throw new Error('COC_Records sheet not found');
  
  // Parse month-year (format: "2025-10" or "10-2025")
  const [part1, part2] = monthYear.split('-');
  const year = part1.length === 4 ? parseInt(part1) : parseInt(part2);
  const month = part1.length === 4 ? parseInt(part2) : parseInt(part1);
  
  // Get all approved COC records for this employee and month
  const recData = recSheet.getDataRange().getValues();
  const monthRecords = [];
  
  for (let i = 1; i < recData.length; i++) {
    const row = recData[i];
    const empId = row[1]; // Employee ID
    const dateRendered = new Date(row[4]); // Date Rendered
    const status = row[15]; // Status
    
    if (empId === employeeId && status === 'Active') {
      const recordMonth = dateRendered.getMonth() + 1;
      const recordYear = dateRendered.getFullYear();
      
      if (recordMonth === month && recordYear === year) {
        monthRecords.push({
          recordId: row[0],
          dateRendered: dateRendered,
          hoursWorked: parseFloat(row[10]) || 0,
          cocEarned: parseFloat(row[12]) || 0,
          rowIndex: i + 1
        });
      }
    }
  }
  
  if (monthRecords.length === 0) {
    throw new Error('No COC records found for this employee and month');
  }
  
  // Get employee name
  const employeeName = recData[monthRecords[0].rowIndex - 1][2];
  
  // Calculate total COC earned
  const totalCOC = monthRecords.reduce((sum, rec) => sum + rec.cocEarned, 0);
  
  // Create certificate ID based on current timestamp
  const issueDate = new Date();
  const certificateId = 'CERT-' + Utilities.formatDate(issueDate, TIME_ZONE, 'yyyyMMddHHmmssSSS');
  
  // Calculate expiration: Issue Date + 1 Year - 1 Day
  const expirationDate = calculateCertificateExpiration(issueDate);
  
  // Create Google Doc certificate
  const docName = `COC Certificate - ${employeeName} - ${monthYear}`;
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();
  
  // Certificate header
  body.appendParagraph('Republic of the Philippines')
    .setBold(true).setFontSize(12)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Department of Public Works and Highways')
    .setBold(true).setFontSize(12)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Certificate of Compensatory Overtime Credit (COC)')
    .setBold(true).setFontSize(14)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('\n');
  
  // Certificate body
  body.appendParagraph(`Certificate ID: ${certificateId}`)
    .setBold(true);
  body.appendParagraph(`Employee: ${employeeName}`);
  body.appendParagraph(`Period: ${monthYear}`);
  body.appendParagraph('\n');
  
  // List all dates
  body.appendParagraph('This certifies that the above-named employee has rendered overtime services during the following dates:')
    .setBold(true);
  
  monthRecords.forEach(rec => {
    body.appendParagraph(`  • ${Utilities.formatDate(rec.dateRendered, TIME_ZONE, 'MMMM dd, yyyy')}: ${rec.hoursWorked.toFixed(2)} hours worked = ${rec.cocEarned.toFixed(2)} COC hours`);
  });
  
  body.appendParagraph('\n');
  body.appendParagraph(`Total COC Earned: ${totalCOC.toFixed(2)} hours`)
    .setBold(true).setFontSize(13);
  body.appendParagraph('\n');
  
  // Issue and expiration info
  body.appendParagraph(`Date Issued: ${Utilities.formatDate(issueDate, TIME_ZONE, 'MMMM dd, yyyy')}`);
  body.appendParagraph(`Expiration Date: ${Utilities.formatDate(expirationDate, TIME_ZONE, 'MMMM dd, yyyy')}`)
    .setBold(true).setForegroundColor('#DC2626'); // Red color for emphasis
  body.appendParagraph('\n');
  body.appendParagraph(`Processed by: ${Session.getActiveUser().getEmail()}`);
  
  doc.saveAndClose();
  const docId = doc.getId();
  const docUrl = doc.getUrl();
  const pdfUrl = `https://docs.google.com/document/d/${docId}/export?format=pdf`;
  
  // Update COC_Certificates sheet
  let certSheet = db.getSheetByName('COC_Certificates');
  if (!certSheet) {
    certSheet = db.insertSheet('COC_Certificates');
    certSheet.getRange(1, 1, 1, 12).setValues([[
      'Certificate ID', 'Employee ID', 'Employee Name', 'Month-Year', 
      'Total COC Earned', 'Issue Date', 'Expiration Date', 
      'Certificate URL', 'PDF URL', 'Status', 'Created Date', 'Created By'
    ]]);
  }
  
  certSheet.appendRow([
    certificateId,
    employeeId,
    employeeName,
    monthYear,
    totalCOC,
    issueDate,
    expirationDate,
    docUrl,
    pdfUrl,
    'Active',
    issueDate,
    Session.getActiveUser().getEmail()
  ]);
  
  // UPDATE ALL RELATED RECORDS in COC_Balance_Detail with certificate info and expiration
  const detailData = detailSheet.getDataRange().getValues();
  
  for (let i = 1; i < detailData.length; i++) {
    const recordId = detailData[i][DETAIL_COLS.RECORD_ID];
    
    // Check if this detail record matches any of our month records
    if (monthRecords.some(rec => rec.recordId === recordId)) {
      // Update Certificate ID
      detailSheet.getRange(i + 1, DETAIL_COLS.CERTIFICATE_ID + 1).setValue(certificateId);
      
      // Update Certificate Issue Date
      detailSheet.getRange(i + 1, DETAIL_COLS.CERTIFICATE_ISSUE_DATE + 1).setValue(issueDate);
      
      // Update Expiration Date (CRITICAL: Based on certificate issue date, NOT date earned!)
      detailSheet.getRange(i + 1, DETAIL_COLS.EXPIRATION_DATE + 1).setValue(expirationDate);
      
      // Update notes
      const existingNotes = detailData[i][DETAIL_COLS.NOTES] || '';
      const newNote = `[${Utilities.formatDate(issueDate, TIME_ZONE, 'yyyy-MM-dd')}] Certificate issued: ${certificateId}. Expires: ${Utilities.formatDate(expirationDate, TIME_ZONE, 'yyyy-MM-dd')}`;
      detailSheet.getRange(i + 1, DETAIL_COLS.NOTES + 1).setValue(
        existingNotes ? existingNotes + '\n' + newNote : newNote
      );
    }
  }
  
  return {
    certificateId: certificateId,
    docUrl: docUrl,
    pdfUrl: pdfUrl,
    totalCOC: totalCOC,
    issueDate: issueDate,
    expirationDate: expirationDate,
    recordsUpdated: monthRecords.length
  };
}

/**
 * Parses the stored month-year string from COC_Balance_Detail and returns a
 * Date representing the first day of that month in the script's time zone.
 * Supports "yyyy-MM" and "MM-yyyy" formats.
 *
 * @param {string} monthYear The stored month-year value.
 * @return {Date|null} Parsed date or null if parsing failed.
 */
function parseMonthYear(monthYear) {
  if (!monthYear) return null;
  const normalized = String(monthYear).trim();
  let year, month;
  if (/^\d{4}-\d{2}$/.test(normalized)) {
    const parts = normalized.split('-');
    year = parseInt(parts[0], 10);
    month = parseInt(parts[1], 10) - 1;
  } else if (/^\d{2}-\d{4}$/.test(normalized)) {
    const parts = normalized.split('-');
    month = parseInt(parts[0], 10) - 1;
    year = parseInt(parts[1], 10);
  } else {
    const parsed = new Date(normalized);
    if (!isNaN(parsed.getTime())) {
      parsed.setDate(1);
      parsed.setHours(0, 0, 0, 0);
      return parsed;
    }
    return null;
  }

  if (isNaN(year) || isNaN(month)) return null;
  const date = new Date(year, month, 1);
  date.setHours(0, 0, 0, 0);
  return date;
}

// -----------------------------------------------------------------------------
// Helpers
// -----------------------------------------------------------------------------

/**
 * Returns the script's time zone. Must be called inside a function.
 */
function getScriptTimeZone() {
  return Session.getScriptTimeZone();
}

/**
 * Formats a JavaScript Date into a human friendly string (e.g. “May 24, 2025”).
 * All user‑visible dates should pass through this helper to enforce a
 * consistent display across the application. Internally we leverage
 * `Utilities.formatDate` because it allows specifying both the time zone and
 * the desired output pattern. See the official documentation for examples【640255559486048†L203-L214】.
 *
 * @param {Date} date The date to format.
 * @return {string} The formatted date.
 */
function formatDate(date) {
  const TIME_ZONE = getScriptTimeZone();
  return Utilities.formatDate(new Date(date), TIME_ZONE, 'MMMM dd, yyyy');
}

/**
 * Returns the database spreadsheet object. All data reads/writes should go
 * through this helper so that future changes (like migrating to another file)
 * require editing only one location.
 *
 * @return {Spreadsheet} The database spreadsheet.
 */
function getDatabase() {
  try {
    if (DATABASE_ID && typeof DATABASE_ID === 'string' && DATABASE_ID !== 'REPLACE_WITH_DATABASE_SHEET_ID') {
      // --- MODIFICATION: Added error logging ---
      const db = SpreadsheetApp.openById(DATABASE_ID);
      db.getName(); // This simple call will trigger auth/permission errors
      return db;
      // --- END MODIFICATION ---
    }
  } catch (err) {
    // --- MODIFICATION: Added a visible log for debugging ---
    // This will write the *actual* error to the Apps Script console.
    Logger.log('WARNING: Could not open database by ID (' + DATABASE_ID + '). Error: ' + err.message + '. Falling back to active spreadsheet.');
    // --- END MODIFICATION ---
  }
  // Fall back to the active spreadsheet if openById fails
  return SpreadsheetApp.getActive();
}

/**
 * Gets all data from a sheet excluding the header row.
 * Returns an array of arrays representing the data rows.
 *
 * @param {string} sheetName The name of the sheet to retrieve data from
 * @return {Array<Array>} Array of row arrays (excluding header)
 */
function getSheetDataNoHeader(sheetName) {
  const db = getDatabase();
  const sheet = db.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found in database.`);
  }

  const data = sheet.getDataRange().getValues();

  // Return empty array if sheet only has header or is empty
  if (data.length <= 1) {
    return [];
  }

  // Return all rows except the first one (header)
  return data.slice(1);
}

/**
 * Gets the current user's email address.
 *
 * @return {string} The email address of the current user
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Generates a unique ID with a given prefix.
 * Uses timestamp and random component for uniqueness.
 *
 * @param {string} prefix The prefix for the ID (e.g., "COC-", "CERT-")
 * @return {string} A unique ID string
 */
function generateUniqueId(prefix) {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return `${prefix}${timestamp}${random}`;
}

// ============================================================================
// --- NEW HELPER FUNCTION ---
// This function contains the logic to determine the day type.
// It is now called by apiGetDayType and calculateOvertimeForDate.
// ============================================================================
/**
 * Determines the type of day (Weekday, Weekend, Holiday)
 * @param {Date} date The date to check.
 * @return {string} The day type.
 */
function getDayType(date) {
  const TIME_ZONE = getScriptTimeZone();
  const dow = date.getDay(); // 0=Sun
  let dayType = 'Weekday';
  if (dow === 0 || dow === 6) {
    dayType = 'Weekend';
  }

  // Check holidays sheet for overrides
  const db = getDatabase();
  const holidaysSheet = db.getSheetByName('Holidays');
  if (holidaysSheet) {
    const holData = holidaysSheet.getDataRange().getValues();
    const target = Utilities.formatDate(date, TIME_ZONE, 'yyyy-MM-dd');
    for (let i = 1; i < holData.length; i++) {
      const holDate = holData[i][0];
      const holType = holData[i][1];
      if (holDate && Utilities.formatDate(new Date(holDate), TIME_ZONE, 'yyyy-MM-dd') === target) {
        if (holType === 'Regular') {
          dayType = 'Regular Holiday';
        } else {
          // Consolidate 'Special Non-Working' and other types
          dayType = 'Special Non-Working';
        }
        break;
      }
    }
  }
  return dayType;
}

/**
 * Reads the Settings sheet into a plain object. Settings provide system
 * defaults like maximum COC per month or the validity period. Adding new
 * settings to the sheet automatically exposes them here without additional
 * code changes.
 *
 * @return {Object<string,string>} A mapping of setting keys to their values.
 */
function getSettings() {
  const db = getDatabase();
  const sheet = db.getSheetByName('Settings');
  if (!sheet) return {};
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const settings = {};
  values.forEach(row => {
    settings[row[0]] = row[1];
  });
  return settings;
}

/**
 * Reads lists from the 'Library' sheet for use in dropdowns.
 * Assumes Positions are in Column A and Offices are in Column B (starting from row 2).
 *
 * @return {Object} An object { positions: [...], offices: [...] }.
 */
function getDropdownOptions() {
  const db = getDatabase();
  const options = {
    positions: [],
    offices: []
  };

  // Get data from 'Library' sheet
  const libSheet = db.getSheetByName('Library');
  if (libSheet && libSheet.getLastRow() > 1) {
    // Read both columns A and B in one go for efficiency
    const lastRow = libSheet.getLastRow();
    const data = libSheet.getRange(2, 1, lastRow - 1, 2).getValues();

    const positions = [];
    const offices = [];

    data.forEach(row => {
      if (row[0] && row[0].trim() !== '') { // If there's a value in Column A (Position)
        positions.push(row[0].trim());
      }
      if (row[1] && row[1].trim() !== '') { // If there's a value in Column B (Office)
        offices.push(row[1].trim());
      }
    });

    options.positions = positions;
    options.offices = offices;
  }

  return options;
}


/** Pads a number with leading zeroes. */
function padNumber(num, length) {
  return num.toString().padStart(length, '0');
}

/**
 * Generates the next sequential employee ID based off the existing entries in
 * the Employees sheet. IDs are prefaced with “EMP” and padded to three
 * digits (e.g. EMP001, EMP010, EMP100). When adding new employees the
 * generated ID will always be unique.
 *
 * @return {string} The newly generated employee ID.
 */
function generateEmployeeId() {
  const db = getDatabase();
  const sheet = db.getSheetByName('Employees');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'EMP001';
  const lastId = sheet.getRange(lastRow, 1).getValue();
  const num = parseInt(String(lastId).replace('EMP', ''), 10) + 1;
  return 'EMP' + padNumber(num, 3);
}

/**
 * Generates unique IDs for COC, CTO and Ledger records. The IDs encode the
 * current date/time down to seconds to ensure uniqueness. If called more
 * scenario is extremely unlikely in this use case.
 */
function generateRecordId() {
  const now = new Date();
  const TIME_ZONE = getScriptTimeZone();
  return 'COC-' + Utilities.formatDate(now, TIME_ZONE, 'yyyyMMddHHmmssSSS');
}
function generateCTOId() {
  const now = new Date();
  const TIME_ZONE = getScriptTimeZone();
  return 'CTO-' + Utilities.formatDate(now, TIME_ZONE, 'yyyyMMddHHmmssSSS');
}
function generateLedgerId() {
  const now = new Date();
  const TIME_ZONE = getScriptTimeZone();
  return 'LED-' + Utilities.formatDate(now, TIME_ZONE, 'yyyyMMddHHmmssSSS');
}

// -----------------------------------------------------------------------------
// Employee management
// -----------------------------------------------------------------------------

/**
 * Returns a list of employees. Optionally exclude inactive employees. Each
 * employee object includes the employee ID, full name, position, office,
 * email, status, *total* initial COC balance, and the earliest "as-of" date
 * for that balance.
 *
 * @param {boolean} includeInactive When true inactive employees are also
 * included in the list. Default false.
 * @return {Array<Object>} An array of employee objects.
 */
function listEmployees(includeInactive) {
  // --- NEW DEBUG LOG ---
  Logger.log('apiListEmployees called. includeInactive = ' + includeInactive);
  // ---
  const db = getDatabase();
  // --- NEW DEBUG LOG ---
  // This will tell us if we are reading the correct spreadsheet.
  Logger.log('getDatabase() returned spreadsheet: ' + db.getName());
  // ---
  const sheet = db.getSheetByName('Employees');

  // --- NEW DEBUG LOG ---
  if (!sheet) {
    Logger.log('ERROR: "Employees" sheet was NOT FOUND in ' + db.getName());
    return []; // Return early if sheet not found
  }
  Logger.log('"Employees" sheet was found.');
  // ---

  // Read up to column K (InitialBalanceDate)
  // --- NEW DEBUG LOG ---
  const lastRow = sheet.getLastRow();
  Logger.log('Last row in "Employees" sheet is: ' + lastRow);
  if (lastRow < 2) {
      Logger.log('No data rows found (lastRow < 2). Returning empty list.');
      return [];
  }
  // ---
  const rows = sheet.getRange(1, 1, lastRow, 11).getValues();
  const employees = [];
  const TIME_ZONE = getScriptTimeZone(); // Added this for the formatter
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const id = row[0];
    const lastName = row[1];
    const firstName = row[2];
    const middleInitial = row[3];
    const suffix = row[4];
    const position = row[5];
    const office = row[6];
    const email = row[7];
    const status = row[8];
    const initialBalance = parseFloat(row[9]) || 0; // This is the TOTAL initial balance
    const initialBalanceDate = row[10] ? Utilities.formatDate(new Date(row[10]), TIME_ZONE, 'yyyy-MM-dd') : '';

    if (!includeInactive && status !== 'Active') continue;
    const fullName = firstName + (middleInitial ? ' ' + middleInitial + ' ' : ' ') + lastName + (suffix ? ' ' + suffix : '');

    employees.push({
      id: id,
      fullName: fullName.trim(),
      position: position,
      office: office,
      email: email,
      status: status,
      initialBalance: initialBalance,
      initialBalanceDate: initialBalanceDate, // This is the *earliest* date for reference
      lastName: lastName,
      firstName: firstName,
      middleInitial: middleInitial,
      suffix: suffix
    });
  }
  return employees;
}

/**
 * Retrieves an employee record by ID. The returned object contains the row
 * number to facilitate updates along with all employee fields. Returns null
 * if the employee does not exist.
 *
 * @param {string} employeeId The employee ID (e.g. EMP001).
 * @return {Object|null} The employee record or null if not found.
 */
function getEmployeeById(employeeId) {
  Logger.log('getEmployeeById called for: ' + employeeId);
  const db = getDatabase();
  const sheet = db.getSheetByName('Employees');
  const TIME_ZONE = getScriptTimeZone(); // Added this for the formatter
  // Read up to column K
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 11).getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === employeeId) {
      const row = data[i];
      const fullName = row[2] + (row[3] ? ' ' + row[3] + ' ' : ' ') + row[1] + (row[4] ? ' ' + row[4] : '');
      return {
        row: i + 1,
        id: row[0],
        lastName: row[1],
        firstName: row[2],
        middleInitial: row[3],
        suffix: row[4],
        position: row[5],
        office: row[6],
        email: row[7],
        status: row[8],
        initialBalance: parseFloat(row[9]) || 0,
        initialBalanceDate: row[10] ? Utilities.formatDate(new Date(row[10]), TIME_ZONE, 'yyyy-MM-dd') : '',
        fullName: fullName.trim()
      };
    }
  }
  return null;
}

/**
 * Adds a new employee to the database. (Original, non-FIFO version)
 *
 * --- MODIFIED (Option C) ---
 * - Now accepts `data.initialBalanceEntries` which is an array of objects:
 * `[{amount: 5.0, date: '2024-10-01'}, ...]`.
 * - It calculates the *total* initial balance and the *earliest* date to
 * store in the 'Employees' sheet (Cols J & K).
 * - It then loops through the `initialBalanceEntries` array and creates a
 * *separate ledger entry for each one*, correctly calculating expirations
 * and running balances.
 *
 * @param {Object} data An object with employee details and
 * `initialBalanceEntries` array.
 * @return {string} The generated employee ID.
 */
function addEmployee(data) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees sheet not found in database');

  const id = generateEmployeeId();
  const fullName = data.firstName + (data.middleInitial ? ' ' + data.middleInitial + ' ' : ' ') + data.lastName + (data.suffix ? ' ' + data.suffix : '');

  // --- NEW LOGIC for Option C ---
  const entries = data.initialBalanceEntries || [];

  // 1. Calculate Total Balance and Earliest Date
  let totalInitialBalance = 0;
  let earliestDate = null;
  const validDates = [];

  entries.forEach(entry => {
      const amount = parseFloat(entry.amount) || 0;
      if (amount > 0 && entry.date) {
          totalInitialBalance += amount;
          validDates.push(new Date(entry.date));
      }
  });

  if (validDates.length > 0) {
      // Find the minimum (earliest) date
      earliestDate = new Date(Math.min.apply(null, validDates));
  }
  // --- END NEW LOGIC ---

  // 2. Append to 'Employees' sheet
  sheet.appendRow([
    id,
    data.lastName,
    data.firstName,
    data.middleInitial || '',
    data.suffix || '',
    data.position,
    data.office,
    data.email,
    data.status || 'Active',
    totalInitialBalance, // Store the calculated total
    earliestDate // Store the earliest date for reference
  ]);

  // 3. Create multiple ledger entries if balance > 0
  if (totalInitialBalance > 0) {
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');
    const settings = getSettings();
    const validityMonths = parseInt(settings['COC_VALIDITY_MONTHS']) || 12;
    const TIME_ZONE = getScriptTimeZone(); // Added this

    let runningBalance = 0; // New employee's balance starts at 0
    const ledgerRows = []; // To batch-write for efficiency

    // Sort entries by date to ensure correct running balance
    const sortedEntries = entries
      .filter(e => (parseFloat(e.amount) || 0) > 0 && e.date)
      .sort((a, b) => new Date(a.date) - new Date(b.date));

    sortedEntries.forEach(entry => {
      const amount = parseFloat(entry.amount);
      const earnedDate = new Date(entry.date);
      const ledgerId = generateLedgerId(); // Generates unique ID with milliseconds
      const monthYear = Utilities.formatDate(earnedDate, TIME_ZONE, 'MM-yyyy');

      const expiration = new Date(earnedDate);
      expiration.setMonth(expiration.getMonth() + validityMonths);

      runningBalance += amount; // Increment running balance

      const remarks = 'Initial COC balance (Earned: ' + formatDate(earnedDate) + ')';

      ledgerRows.push([
        ledgerId,
        id,
        fullName.trim(),
        new Date(), // transactionDate (today)
        'Initial Balance',
        'INIT-BALANCE-' + ledgerId.substring(4, 12), // Reference
        amount, // cocEarned
        0, // ctoUsed
        runningBalance, // cocBalance (cumulative)
        monthYear, // monthYearEarned
        expiration, // expirationDate
        Session.getActiveUser().getEmail(),
        remarks
      ]);
    });

    // 4. Batch-write all ledger rows
    if (ledgerRows.length > 0) {
       ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, ledgerRows.length, ledgerRows[0].length).setValues(ledgerRows);
    }
  }
  return id;
}

/*
 * DPWH COC/CTO Recording System – Code.gs (UPDATED with FIFO)
 *
 * ENHANCEMENTS:
 * 1. Added COC_Balance_Detail sheet for FIFO tracking
 * 2. Initial Balance now creates detailed entries with expiration dates
 * 3. CTO usage implements FIFO (First In, First Out)
 * 4. Automatic expiration checking
 */

// ============================================================================
// INSERT THIS CODE AFTER LINE 441 (after addEmployee function)
// ============================================================================

/**
 * Creates or ensures the COC_Balance_Detail sheet exists
 * This sheet tracks individual COC entries for FIFO consumption
 */
function ensureCOCBalanceDetailSheet() {
  const db = getDatabase();
  let detailSheet = db.getSheetByName('COC_Balance_Detail');

  if (!detailSheet) {
    detailSheet = db.insertSheet('COC_Balance_Detail');

    const headers = [
      'Record ID', 'Employee ID', 'Employee Name', 'Month-Year Earned',
      'Certificate Date', 'Hours Earned', 'Hours Used', 'Hours Remaining',
      'Status', 'Expiration Date', 'Certificate ID', 'Date Created',
      'Created By', 'Remarks'
    ];

    detailSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    detailSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    detailSheet.setFrozenRows(1);
  }

  return detailSheet;
}

/**
 * API: Get day type only (Weekday, Weekend, Holiday)
 * Used for immediate day type detection without needing time entries
 *
 * @param {number} year - The year (e.g., 2025)
 * @param {number} month - The month (1-12)
 * @param {number} day - The day of the month (1-31)
 * @return {string} The day type (Weekday, Weekend, Regular Holiday, Special Non-Working Holiday)
 */
function apiGetDayType(year, month, day) {
  const date = new Date(year, month - 1, day);
  return getDayType(date); // This will now work
}

/**
 * API wrapper for calculateOvertimeForDate
 * This function is called from the frontend to automatically detect day type
 * and calculate COC when times are entered
 *
 * @param {number} year - The year (e.g., 2025)
 * @param {number} month - The month (1-12)
 * @param {number} day - The day of the month (1-31)
 * @param {string} amIn - AM in time (HH:mm format)
 * @param {string} amOut - AM out time (HH:mm format)
 * @param {string} pmIn - PM in time (HH:mm format)
 * @param {string} pmOut - PM out time (HH:mm format)
 * @return {Object} Result with dayType, hoursWorked, multiplier, cocEarned
 */
function apiCalculateOvertimeForDate(year, month, day, amIn, amOut, pmIn, pmOut) {
  const date = new Date(year, month - 1, day);
  return calculateOvertimeForDate(date, amIn, amOut, pmIn, pmOut);
}

/**
 * Generates unique Entry ID for COC_Balance_Detail
 */
function generateCOCDetailEntryId() {
  const now = new Date();
  const TIME_ZONE = getScriptTimeZone();
  return 'COCD-' + Utilities.formatDate(now, TIME_ZONE, 'yyyyMMddHHmmssSSS');
}

/**
 * MODIFIED: Enhanced addEmployee to create COC_Balance_Detail entries
 * This replaces the existing addEmployee function starting at line 345
 */
function addEmployeeWithFIFO(data) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees sheet not found in database');

  const id = generateEmployeeId();
  const fullName = data.firstName + (data.middleInitial ? ' ' + data.middleInitial + ' ' : ' ') +
                     data.lastName + (data.suffix ? ' ' + data.suffix : '');

  const entries = data.initialBalanceEntries || [];

  // Calculate Total Balance and Earliest Date
  let totalInitialBalance = 0;
  let earliestDate = null;
  const validDates = [];

  entries.forEach(entry => {
    const amount = parseFloat(entry.amount) || 0;
    if (amount > 0 && entry.date) {
      totalInitialBalance += amount;
      validDates.push(new Date(entry.date));
    }
  });

  if (validDates.length > 0) {
    earliestDate = new Date(Math.min.apply(null, validDates));
  }

  // Append to 'Employees' sheet
  sheet.appendRow([
    id,
    data.lastName,
    data.firstName,
    data.middleInitial || '',
    data.suffix || '',
    data.position,
    data.office,
    data.email,
    data.status || 'Active',
    totalInitialBalance,
    earliestDate
  ]);

  // Create ledger and detail entries
  if (totalInitialBalance > 0) {
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    const detailSheet = ensureCOCBalanceDetailSheet();

    if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');

    const settings = getSettings();
    const validityMonths = parseInt(settings['COC_VALIDITY_MONTHS']) || 12;
    const TIME_ZONE = getScriptTimeZone();

    let runningBalance = 0;
    const ledgerRows = [];
    const detailRows = [];

    // Sort entries by date
    const sortedEntries = entries
      .filter(e => (parseFloat(e.amount) || 0) > 0 && e.date)
      .sort((a, b) => new Date(a.date) - new Date(b.date));

    sortedEntries.forEach(entry => {
      const amount = parseFloat(entry.amount);
      const earnedDate = new Date(entry.date);
      const ledgerId = generateLedgerId();
      const monthYear = Utilities.formatDate(earnedDate, TIME_ZONE, 'MM-yyyy');

      const certificateDate = new Date(earnedDate);
      const expiration = calculateCertificateExpiration(certificateDate);

      runningBalance += amount;

      const remarks = 'Initial COC balance (Earned: ' + formatDate(earnedDate) + ')';
      const recordId = 'INITIAL-' + id + '-' + Utilities.formatDate(earnedDate, TIME_ZONE, 'yyyyMM');

      // Ledger entry
      ledgerRows.push([
        ledgerId,
        id,
        fullName.trim(),
        new Date(),
        'Initial Balance',
        recordId,
        amount,
        0,
        runningBalance,
        monthYear,
        expiration,
        Session.getActiveUser().getEmail(),
        remarks
      ]);

      // Detail entry (for FIFO tracking)
      detailRows.push([
        recordId,
        id,
        fullName.trim(),
        monthYear,
        certificateDate,
        amount,
        0,
        amount,
        'Active',
        expiration,
        'CERT-' + Utilities.formatDate(certificateDate, TIME_ZONE, 'yyyyMMddHHmmssSSS'),
        new Date(),
        Session.getActiveUser().getEmail(),
        remarks
      ]);
    });

    // Batch write
    if (ledgerRows.length > 0) {
      ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, ledgerRows.length, ledgerRows[0].length)
        .setValues(ledgerRows);
    }

    if (detailRows.length > 0) {
      detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length)
        .setValues(detailRows);
    }
  }

  return id;
}

/**
 * Adds a COC entry to COC_Balance_Detail when COC is earned
 * Call this from recordCOCEntries after creating COC_Records entry
 */
function addCOCToBalanceDetail(employeeId, employeeName, recordId, dateEarned, hoursEarned) {
  const detailSheet = ensureCOCBalanceDetailSheet();
  const TIME_ZONE = getScriptTimeZone();

  const earnedDate = new Date(dateEarned);
  const monthYearEarned = Utilities.formatDate(earnedDate, TIME_ZONE, 'MM-yyyy');
  const now = new Date();
  const createdBy = Session.getActiveUser().getEmail();

  detailSheet.appendRow([
    recordId,                            // Record ID (matches COC_Records ID)
    employeeId,
    employeeName,
    monthYearEarned,
    '',                                  // Certificate Date (set once certificate is issued)
    hoursEarned,
    0,                                   // Hours Used
    hoursEarned,                         // Hours Remaining
    'Active',
    '',                                  // Expiration Date (computed after certificate issuance)
    '',                                  // Certificate ID
    now,
    createdBy,
    'COC earned from ' + recordId
  ]);
}

/**
 * FIFO COC Consumption
 * Returns array of deductions showing which entries were consumed
 */
function consumeCOCWithFIFO(employeeId, hoursToConsume, reference) {
  const detailSheet = ensureCOCBalanceDetailSheet();
  const data = detailSheet.getDataRange().getValues();
  const TIME_ZONE = getScriptTimeZone();

  // Get all active entries for this employee
  const availableEntries = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[DETAIL_COLS.STATUS] || '').trim();
    const remaining = parseFloat(row[DETAIL_COLS.HOURS_REMAINING]) || 0;
    if (row[DETAIL_COLS.EMPLOYEE_ID] === employeeId && status === 'Active' && remaining > 0) {
      const certificateDateVal = row[DETAIL_COLS.CERTIFICATE_DATE];
      const orderingDate = (certificateDateVal instanceof Date && !isNaN(certificateDateVal.getTime()))
        ? new Date(certificateDateVal)
        : (parseMonthYear(row[DETAIL_COLS.MONTH_YEAR]) || new Date(row[DETAIL_COLS.DATE_CREATED]));

      const expirationVal = row[DETAIL_COLS.EXPIRATION_DATE];
      const expirationDate = (expirationVal instanceof Date && !isNaN(expirationVal.getTime()))
        ? new Date(expirationVal)
        : null;

      availableEntries.push({
        rowIndex: i + 1,
        recordId: row[DETAIL_COLS.RECORD_ID],
        dateReference: orderingDate,
        hoursEarned: parseFloat(row[DETAIL_COLS.HOURS_EARNED]) || 0,
        hoursUsed: parseFloat(row[DETAIL_COLS.HOURS_USED]) || 0,
        hoursRemaining: remaining,
        expirationDate: expirationDate
      });
    }
  }

  // Sort by certificate date (FIFO - oldest first)
  availableEntries.sort((a, b) => a.dateReference - b.dateReference);

  // Calculate total available
  const totalAvailable = availableEntries.reduce((sum, entry) => sum + entry.hoursRemaining, 0);

  if (hoursToConsume > totalAvailable) {
    throw new Error('Insufficient COC balance. Available: ' + totalAvailable.toFixed(2) +
                      ' hours, Requested: ' + hoursToConsume.toFixed(2) + ' hours');
  }

  // Apply FIFO deduction
  let remainingToConsume = hoursToConsume;
  const deductions = [];

  for (const entry of availableEntries) {
    if (remainingToConsume <= 0) break;

    const deductFromThis = Math.min(remainingToConsume, entry.hoursRemaining);
    const newRemaining = entry.hoursRemaining - deductFromThis;
    const newUsed = entry.hoursUsed + deductFromThis;
    const timestamp = new Date();
    const note = `[${Utilities.formatDate(timestamp, TIME_ZONE, 'yyyy-MM-dd HH:mm')}] Consumed ${deductFromThis.toFixed(2)} hrs for ${reference}`;

    detailSheet.getRange(entry.rowIndex, DETAIL_COLS.HOURS_USED + 1).setValue(newUsed);
    detailSheet.getRange(entry.rowIndex, DETAIL_COLS.HOURS_REMAINING + 1).setValue(newRemaining);
    if (newRemaining === 0) {
      detailSheet.getRange(entry.rowIndex, DETAIL_COLS.STATUS + 1).setValue('Depleted');
    }

    const existingNote = detailSheet.getRange(entry.rowIndex, DETAIL_COLS.REMARKS + 1).getValue() || '';
    const combinedNote = existingNote ? existingNote + '\n' + note : note;
    detailSheet.getRange(entry.rowIndex, DETAIL_COLS.REMARKS + 1).setValue(combinedNote);

    deductions.push({
      recordId: entry.recordId,
      hoursDeducted: deductFromThis,
      remainingHours: newRemaining
    });

    remainingToConsume -= deductFromThis;
  }

  return deductions;
}

/**
 * MODIFIED: Enhanced recordCTOApplication with FIFO
 * This should replace or enhance the existing recordCTOApplication function
*/
function recordCTOApplicationWithFIFO(employeeId, hours, startDate, endDate, remarks) {
  // Ensure startDate and endDate are Date objects
startDate = new Date(startDate);
endDate   = new Date(endDate);
const db = getDatabase();
  const employee = getEmployeeById(employeeId);
  if (!employee) throw new Error('Employee not found');

  const balance = getCurrentCOCBalance(employeeId);
  if (hours > balance) {
    throw new Error('Insufficient COC balance. Available: ' + balance.toFixed(2) +
                      ' hours, Requested: ' + hours.toFixed(2) + ' hours');
  }

  const ctoId = generateCTOId();
  const ledgerId = generateLedgerId();
  const inclusiveDates = formatDate(startDate) + ' to ' + formatDate(endDate);
  const TIME_ZONE = getScriptTimeZone();

  // Apply FIFO consumption
  const deductions = consumeCOCWithFIFO(employeeId, hours, ctoId);

  // Create detailed remarks showing FIFO consumption
  const fifoDetails = deductions.map(d =>
    '  • ' + d.hoursDeducted.toFixed(2) + ' hrs from ' + d.recordId +
    ' (earned ' + d.dateEarned + ')'
  ).join('\n');

  const detailedRemarks = (remarks || 'CTO application') + '\n\nFIFO Consumption:\n' + fifoDetails;

  // Record in CTO_Applications
  const ctoSheet = db.getSheetByName('CTO_Applications');
  if (ctoSheet) {
    ctoSheet.appendRow([
      ctoId,
      employeeId,
      employee.fullName,
      employee.office,
      hours,
      startDate,
      endDate,
      inclusiveDates,
      balance,
      new Date(),
      'Approved',
      new Date(),
      detailedRemarks
    ]);
  }

  // Record in COC_Ledger (one entry per FIFO deduction for audit trail)
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  if (ledgerSheet) {
    deductions.forEach(deduction => {
      const deductionLedgerId = generateLedgerId();
      const monthYear = deduction.dateEarned.substring(0, 7).split('-').reverse().join('-'); // Convert to MM-yyyy

      ledgerSheet.appendRow([
        deductionLedgerId,
        employeeId,
        employee.fullName,
        new Date(),
        'CTO Used',
        ctoId,
        0, // COC Earned
        deduction.hoursDeducted, // CTO Used
        balance - hours, // New balance
        monthYear,
        '', // Expiration (not applicable for usage)
        Session.getActiveUser().getEmail(),
        'FIFO: Used ' + deduction.hoursDeducted.toFixed(2) + ' hrs from ' +
        deduction.recordId + ' (earned ' + deduction.dateEarned + ') for ' + ctoId
      ]);
    });
  }

  return {
    success: true,
    ctoId: ctoId,
    hoursUsed: hours,
    newBalance: balance - hours,
    fifoDeductions: deductions
  };
}

/**
 * Check and mark expired COC entries
 * Run this daily or on-demand
 */
function checkAndExpireCOC() {
  const detailSheet = ensureCOCBalanceDetailSheet();
  const data = detailSheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let expiredCount = 0;
  let totalHoursForfeited = 0;
  const TIME_ZONE = getScriptTimeZone();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[DETAIL_COLS.STATUS] || '').trim();
    const hoursRemaining = parseFloat(row[DETAIL_COLS.HOURS_REMAINING]) || 0;
    const expirationVal = row[DETAIL_COLS.EXPIRATION_DATE];
    if (!expirationVal) continue;

    const expirationDate = new Date(expirationVal);
    if (isNaN(expirationDate.getTime())) continue;
    expirationDate.setHours(0, 0, 0, 0);

    if (status === 'Active' && hoursRemaining > 0 && expirationDate < today) {
      detailSheet.getRange(i + 1, DETAIL_COLS.STATUS + 1).setValue('Expired');

      const note = `[${Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm')}] EXPIRED - ${hoursRemaining.toFixed(2)} hours forfeited`;
      const currentNote = detailSheet.getRange(i + 1, DETAIL_COLS.REMARKS + 1).getValue() || '';
      const updatedNote = currentNote ? currentNote + '\n' + note : note;
      detailSheet.getRange(i + 1, DETAIL_COLS.REMARKS + 1).setValue(updatedNote);

      expiredCount++;
      totalHoursForfeited += hoursRemaining;

      const ledgerSheet = getDatabase().getSheetByName('COC_Ledger');
      if (ledgerSheet) {
        ledgerSheet.appendRow([
          generateLedgerId(),
          row[DETAIL_COLS.EMPLOYEE_ID],
          row[DETAIL_COLS.EMPLOYEE_NAME],
          new Date(),
          'COC Expired',
          row[DETAIL_COLS.RECORD_ID],
          0,
          hoursRemaining,
          '',
          row[DETAIL_COLS.MONTH_YEAR],
          expirationDate,
          Session.getActiveUser().getEmail(),
          'Auto-expired: ' + hoursRemaining.toFixed(2) + ' hours forfeited from ' + row[DETAIL_COLS.RECORD_ID]
        ]);
      }
    }
  }

  return {
    expiredCount: expiredCount,
    totalHoursForfeited: totalHoursForfeited.toFixed(2)
  };
}

/**
 * FIXED: Get COC Balance breakdown with robust date handling
 */
function getCOCBalanceBreakdown(employeeId) {
  const detailSheet = ensureCOCBalanceDetailSheet();
  const data = detailSheet.getDataRange().getValues();
  const TIME_ZONE = getScriptTimeZone();
  
  /**
   * Helper to safely parse dates
   */
  function safeDateParse(value) {
    if (!value) return null;
    try {
      if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
      if (typeof value === 'string') {
        const date = new Date(value);
        if (!isNaN(date.getTime())) return date;
        const usMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (usMatch) {
          const [, m, d, y] = usMatch;
          return new Date(y, parseInt(m) - 1, d);
        }
      }
      return new Date(value);
    } catch (e) {
      return null;
    }
  }
  
  function safeFormatDate(value) {
    const date = safeDateParse(value);
    if (!date) return 'N/A';
    try {
      return Utilities.formatDate(date, TIME_ZONE, 'yyyy-MM-dd');
    } catch (e) {
      return date.toISOString().split('T')[0];
    }
  }

  const breakdown = [];
  
  // CRITICAL FIX: Normalize employee ID for comparison
  const searchId = String(employeeId).trim();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // CRITICAL FIX: Normalize row employee ID before comparison
    const rowEmpId = String(row[1] || '').trim();
    const status = row[8];
    const hoursRemaining = parseFloat(row[6]) || 0;
    
    if (rowEmpId === searchId && status === 'Active' && hoursRemaining > 0) {
      const expDate = safeDateParse(row[7]);
      const daysUntil = expDate ? Math.ceil((expDate - new Date()) / (1000 * 60 * 60 * 24)) : -999;
      
      breakdown.push({
        entryId: row[0],
        recordId: row[3],
        dateEarned: safeFormatDate(row[4]),
        hoursEarned: parseFloat(row[5]) || 0,
        hoursRemaining: hoursRemaining,
        expirationDate: safeFormatDate(row[7]),
        daysUntilExpiration: daysUntil
      });
    }
  }

  breakdown.sort((a, b) => new Date(a.dateEarned) - new Date(b.dateEarned));
  return breakdown;
}

// ============================================================================
// API WRAPPERS - Add these to your existing API functions section
// ============================================================================

/**
 * API: Add employee with FIFO tracking
 * This replaces apiAddEmployee
 */
function apiAddEmployee(data) {
  return addEmployeeWithFIFO(data);
}

/**
 * API: Record CTO with FIFO
 */
function apiRecordCTOWithFIFO(employeeId, hours, startDate, endDate, remarks) {
  return recordCTOApplicationWithFIFO(employeeId, hours, startDate, endDate, remarks);
}

/**
 * API: Get COC balance breakdown
 */
function apiGetCOCBalanceBreakdown(employeeId) {
  try {
    const breakdown = getCOCBalanceBreakdown(employeeId);
    
    // Convert to serializable format
    const safeBreakdown = [];
    
    if (breakdown && Array.isArray(breakdown)) {
      breakdown.forEach(function(item) {
        safeBreakdown.push({
          entryId: String(item.entryId || ''),
          recordId: String(item.recordId || ''),
          dateEarned: String(item.dateEarned || ''),
          hoursEarned: parseFloat(item.hoursEarned) || 0,
          hoursRemaining: parseFloat(item.hoursRemaining) || 0,
          expirationDate: String(item.expirationDate || ''),
          daysUntilExpiration: parseInt(item.daysUntilExpiration) || 0
        });
      });
    }
    
    return safeBreakdown;
    
  } catch (error) {
    Logger.log('ERROR in apiGetCOCBalanceBreakdown: ' + error.message);
    return [];
  }
}

/**
 * API: Check expired COC
 */
function apiCheckAndExpireCOC() {
  return checkAndExpireCOC();
}

/**
 * API: Initialize COC_Balance_Detail sheet
 */
function apiInitializeCOCBalanceDetail() {
  ensureCOCBalanceDetailSheet();
  return { success: true, message: 'COC_Balance_Detail sheet initialized' };
}

// ============================================================================
// ONE-TIME MIGRATION: Migrate existing data to COC_Balance_Detail
// ============================================================================

/**
 * ONE-TIME MIGRATION: Convert existing Initial Balance entries to detail entries
 * Run this ONCE after deploying the new code
 */
function migrateExistingInitialBalances() {
  const db = getDatabase();
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  const detailSheet = ensureCOCBalanceDetailSheet();

  if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');

  const ledgerData = ledgerSheet.getDataRange().getValues();
  const detailRows = [];
  const TIME_ZONE = getScriptTimeZone();

  let migratedCount = 0;

  for (let i = 1; i < ledgerData.length; i++) {
    const row = ledgerData[i];
    const transactionType = row[4];

    if (transactionType === 'Initial Balance') {
      const employeeId = row[1];
      const employeeName = row[2];
      const cocEarned = row[6];
      const monthYearEarned = row[9];
      const expirationDate = row[10];
      const referenceId = row[5];

      // Skip if no amount or already migrated
      if (!cocEarned || cocEarned <= 0) continue;

      // Check if already exists in detail sheet
      const detailData = detailSheet.getDataRange().getValues();
      let alreadyExists = false;
      for (let j = 1; j < detailData.length; j++) {
        if (detailData[j][DETAIL_COLS.RECORD_ID] === referenceId) {
          alreadyExists = true;
          break;
        }
      }

      if (alreadyExists) continue;

      // Convert MM-yyyy to date
      const earnedDate = monthYearEarned ? parseMonthYear(monthYearEarned) : new Date(row[3]);
      const certDate = earnedDate ? new Date(earnedDate) : new Date(row[3]);
      const normalizedMonthYear = Utilities.formatDate(certDate, TIME_ZONE, 'MM-yyyy');
      const calculatedExpiration = expirationDate ? new Date(expirationDate) : calculateCertificateExpiration(certDate);

      detailRows.push([
        referenceId,
        employeeId,
        employeeName,
        normalizedMonthYear,
        certDate,
        cocEarned,
        0,
        cocEarned,
        'Active',
        calculatedExpiration,
        'CERT-' + Utilities.formatDate(certDate, TIME_ZONE, 'yyyyMMddHHmmssSSS'),
        new Date(),
        Session.getActiveUser().getEmail(),
        'Migrated from ledger: ' + row[12]
      ]);

      migratedCount++;
    }
  }

  // Batch write
  if (detailRows.length > 0) {
    detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length)
      .setValues(detailRows);
  }

  return {
    success: true,
    migratedCount: migratedCount,
    message: 'Migrated ' + migratedCount + ' initial balance entries to COC_Balance_Detail'
  };
}

/**
 * API: Run migration
 */
function apiMigrateExistingInitialBalances() {
  return migrateExistingInitialBalances();
}
/**
 * Updates an existing employee’s details.
 *
 * --- IMPORTANT ---
 * This function intentionally DOES NOT update the initial balance (Cols J & K).
 * Initial balances are considered historical data tied to the ledger and
 * should not be changed after employee creation.
 *
 * @param {string} employeeId The ID of the employee to update.
 * @param {Object} data The fields to update (e.g., lastName, firstName, etc.).
 * @return {boolean} True on success.
 */
function updateEmployee(employeeId, data) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Employees');
  if (!sheet) throw new Error('Employees sheet not found');
  const employee = getEmployeeById(employeeId);
  if (!employee) throw new Error('Employee not found');

  // Build updated row values (B to I)
  const updated = [];
  updated.push(data.lastName !== undefined ? data.lastName : employee.lastName);
  updated.push(data.firstName !== undefined ? data.firstName : employee.firstName);
  updated.push(data.middleInitial !== undefined ? data.middleInitial : employee.middleInitial);
  updated.push(data.suffix !== undefined ? data.suffix : employee.suffix);
  updated.push(data.position !== undefined ? data.position : employee.position);
  updated.push(data.office !== undefined ? data.office : employee.office);
  updated.push(data.email !== undefined ? data.email : employee.email);
  updated.push(data.status !== undefined ? data.status : employee.status);

  // This updates 8 columns, starting from column 2 (B to I)
  sheet.getRange(employee.row, 2, 1, updated.length).setValues([updated]);
  return true;
}

// -----------------------------------------------------------------------------
// Balance calculation and ledger
// -----------------------------------------------------------------------------

/**
 * Returns the current COC balance for a given employee. First tries to read
 * the most recent balance from the ledger sheet (Column I). If no ledger
 * records exist for the employee, it falls back to the 'initialBalance'
 * value in the 'Employees' sheet (Column J).
 *
 * @param {string} employeeId The employee ID.
 * @return {number} The current COC balance (in hours).
 */
function getCurrentCOCBalance(employeeId) {
  try {
    const db = getDatabase();
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    
    if (!ledgerSheet) {
      return 0;
    }

    if (ledgerSheet.getLastRow() <= 1) {
      return calculateBalanceFallback(employeeId);
    }
    
    const ledgerData = ledgerSheet.getDataRange().getValues();
    
    // CRITICAL FIX: Normalize employee ID for comparison
    const searchId = String(employeeId).trim();
    
    // Search from bottom for the latest transaction
    for (let i = ledgerData.length - 1; i >= 1; i--) {
      const row = ledgerData[i];
      
      // Skip empty rows
      if (!row || row.length === 0 || !row[0]) {
        continue;
      }
      
      // CRITICAL FIX: Normalize row employee ID before comparison
      const rowEmpId = String(row[1] || '').trim();
      
      if (rowEmpId === searchId) {
        const balance = parseFloat(row[8]) || 0;
        return balance < 0 ? 0 : balance;
      }
    }

    // No ledger entries found, use fallback
    return calculateBalanceFallback(employeeId);
    
  } catch (error) {
    Logger.log('ERROR in getCurrentCOCBalance: ' + error.message);
    return 0;
  }
}

/**
 * Fallback balance calculation when no ledger entries exist
 */
function calculateBalanceFallback(employeeId) {
  try {
    const db = getDatabase();
    const employee = getEmployeeById(employeeId);
    
    if (!employee) return 0;

    let balance = parseFloat(employee.initialBalance) || 0;

    // Sum all active COC earned
    const recordsSheet = db.getSheetByName('COC_Records');
    if (recordsSheet && recordsSheet.getLastRow() > 1) {
      const recData = recordsSheet.getDataRange().getValues();
      const searchId = String(employeeId).trim();
      
      for (let i = 1; i < recData.length; i++) {
        const row = recData[i];
        const rowEmpId = String(row[1] || '').trim();
        const status = row[15];
        
        if (rowEmpId === searchId && status === 'Active') {
          const isInitial = (row[0] || '').startsWith('INIT-');
          if (!isInitial) {
            balance += parseFloat(row[12]) || 0;
          }
        }
      }
    }

    // Subtract all approved CTO used
    const ctoSheet = db.getSheetByName('CTO_Applications');
    if (ctoSheet && ctoSheet.getLastRow() > 1) {
      const ctoData = ctoSheet.getDataRange().getValues();
      const searchId = String(employeeId).trim();
      
      for (let i = 1; i < ctoData.length; i++) {
        const row = ctoData[i];
        const rowEmpId = String(row[1] || '').trim();
        const status = row[10];
        
        if (rowEmpId === searchId && status === 'Approved') {
          balance -= parseFloat(row[4]) || 0;
        }
      }
    }
    
    return balance < 0 ? 0 : balance;
    
  } catch (error) {
    Logger.log('ERROR in calculateBalanceFallback: ' + error.message);
    return 0;
  }
}

// -----------------------------------------------------------------------------
// Overtime and COC earning calculations
// -----------------------------------------------------------------------------

/**
 * Calculates overtime hours, the appropriate multiplier, and the resulting COC
 * earned for a single date. The day type is automatically determined by
 * checking the day of week and any matching entry in the Holidays sheet.
 * Weekdays allow overtime only between 5:00 PM and 7:00 PM (maximum 2 hours).
 * Weekends and holidays allow morning (8:00–12:00) and afternoon (1:00–5:00)
 * blocks, excluding the lunch break from 12:01–12:59 PM. For weekends and
 * holidays the multiplier is 1.5. All calculations are capped as per the
 * business rules.
 *
 * @param {Date} date The calendar date being processed.
 * @param {string} amIn  Optional AM start time in HH:MM (24‑hour) format.
 * @param {string} amOut Optional AM end time in HH:MM (24‑hour) format.
 * @param {string} pmIn  Optional PM start time in HH:MM (24‑hour) format.
 * @param {string} pmOut Optional PM end time in HH:MM (24‑hour) format.
 * @return {Object} An object describing the day type, hours worked, multiplier
 * and COC earned.
*/
function calculateOvertimeForDate(date, amIn, amOut, pmIn, pmOut) {
  const settings = getSettings();
  const TIME_ZONE = getScriptTimeZone(); // Added this

  // --- MODIFICATION ---
  // Call the new helper function to get the day type
  const dayType = getDayType(date);
  // --- END MODIFICATION ---

  let hoursWorked = 0;
  let multiplier = 1.0;
  // Helper to convert time string to a Date on same day
  function parseTime(timeStr) {
    if (!timeStr) return null;
    const parts = timeStr.split(':');
    const h = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const d = new Date(date);
    d.setHours(h, m, 0, 0);
    return d;
  }
  const amStart = parseTime(amIn);
  const amEnd = parseTime(amOut);
  const pmStart = parseTime(pmIn);
  const pmEnd = parseTime(pmOut);
  if (dayType === 'Weekday') {
    // Weekday overtime from 5:00 PM to 7:00 PM
    const otStart = new Date(date);
    otStart.setHours(17, 0, 0, 0);
    const otEnd = new Date(date);
    otEnd.setHours(19, 0, 0, 0);
    const outTime = pmEnd;
    if (outTime && outTime > otStart) {
      let endTime = outTime > otEnd ? otEnd : outTime;
      const ms = endTime.getTime() - otStart.getTime();
      const hours = ms / (1000 * 60 * 60);
      hoursWorked = Math.min(Math.max(hours, 0), 2);
    }
    multiplier = 1.0;
  } else {
    // Weekend/Holiday schedule
    const morningStart = new Date(date);
    morningStart.setHours(8, 0, 0, 0);
    const morningEnd = new Date(date);
    morningEnd.setHours(12, 0, 0, 0);
    const afternoonStart = new Date(date);
    afternoonStart.setHours(13, 0, 0, 0);
    const afternoonEnd = new Date(date);
    afternoonEnd.setHours(17, 0, 0, 0);
    // Morning block
    if (amStart && amEnd) {
      let startTime = amStart < morningStart ? morningStart : amStart;
      let endTime = amEnd > morningEnd ? morningEnd : amEnd;
      const ms = endTime.getTime() - startTime.getTime();
      if (ms > 0) {
        hoursWorked += ms / (1000 * 60 * 60);
      }
    }
    // Afternoon block
    if (pmStart && pmEnd) {
      let startTime = pmStart < afternoonStart ? afternoonStart : pmStart;
      let endTime = pmEnd > afternoonEnd ? afternoonEnd : pmEnd;
      const ms = endTime.getTime() - startTime.getTime();
      if (ms > 0) {
        hoursWorked += ms / (1000 * 60 * 60);
      }
    }
    multiplier = 1.5;
  }
  const cocEarned = hoursWorked * multiplier;
  return {
    dayType: dayType,
    hoursWorked: hoursWorked,
    multiplier: multiplier,
    cocEarned: cocEarned,
    // Added these to pass them to the next step
    amIn: amIn || '',
    amOut: amOut || '',
    pmIn: pmIn || '',
    pmOut: pmOut || ''
  };
}

// -----------------------------------------------------------------------------
// COC recording
// -----------------------------------------------------------------------------

/**
 * Records one or more overtime entries for a single employee and month. Entries
 * are validated against monthly limits, duplicate dates and maximum overall
 * balance constraints. If any entry fails validation the entire operation
 * aborts and no rows are appended.
 *
 * @param {string} employeeId The employee ID.
 * @param {number} month The month number (1–12).
 * @param {number} year The year (four digits).
 * @param {Array<Object>} entries List of objects {day, amIn, amOut, pmIn, pmOut}.
 * @return {Object} Summary of the operation including how many rows were added,
 * total COC earned and the resulting balance.
 */
function recordCOCEntries(employeeId, month, year, entries) {
  const db = getDatabase();
  const recordsSheet = db.getSheetByName('COC_Records');
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  if (!recordsSheet || !ledgerSheet) {
    throw new Error('Required sheets missing');
  }
  const settings = getSettings();
  const employee = getEmployeeById(employeeId);
  if (!employee) {
    throw new Error('Employee not found');
  }
  const TIME_ZONE = getScriptTimeZone(); // Added this
  // Calculate existing hours for this month
  let monthlyTotal = 0;
  // Use MM-yyyy format consistently
  const monthYear = String(month).padStart(2, '0') + '-' + String(year);
  const recData = recordsSheet.getDataRange().getValues();
  for (let i = 1; i < recData.length; i++) {
    const row = recData[i];
    if (row[1] === employeeId && String(row[3]).trim() === monthYear) {
      monthlyTotal += parseFloat(row[10]) || 0;
    }
  }
  // Prepare new rows and ledger entries but do not write yet
  let totalNewHours = 0;
  let totalNewCOC = 0;
  let runningBalance = getCurrentCOCBalance(employeeId);
  const newRows = [];
  const newLedgerRows = [];
  // Duplicate detection: collect existing date strings for this employee.
  // Skip rows that have been cancelled so that users may re‑enter a date
  // after cancelling a previous record. Status is in column 16 (index 15).

  const existingDates = {};
  const existingRecordDetails = {};

  for (let i = 1; i < recData.length; i++) {
    const row = recData[i];
    if (row[1] === employeeId) {
      const status = row[15];
      if (status && String(status).toLowerCase() === 'cancelled') continue;
      const d = new Date(row[4]);
      const key = Utilities.formatDate(d, TIME_ZONE, 'yyyy-MM-dd');
      existingDates[key] = true;
      existingRecordDetails[key] = {
        recordId: row[0],
        date: formatDate(d),
        dayType: row[5],
        hoursWorked: row[10],
        cocEarned: row[11]
      };
    }
  }

  // Check for duplicates and collect them
  const duplicates = [];
  entries.forEach(entry => {
    const date = new Date(year, month - 1, entry.day);
    const key = Utilities.formatDate(date, TIME_ZONE, 'yyyy-MM-dd');
    if (existingDates[key]) {
      duplicates.push({
        date: formatDate(date),
        existing: existingRecordDetails[key]
      });
    }
    const result = calculateOvertimeForDate(date, entry.amIn, entry.amOut, entry.pmIn, entry.pmOut);
    totalNewHours += result.hoursWorked;
    totalNewCOC += result.cocEarned;
  });

  // If duplicates found, throw detailed error
  if (duplicates.length > 0) {
    let errorMsg = 'The following date(s) already have COC records:\n\n';
    duplicates.forEach(dup => {
      errorMsg += '• ' + dup.date + ' (' + dup.existing.dayType + ', ' +
                    dup.existing.hoursWorked + ' hours, ' +
                    dup.existing.cocEarned + ' COC)\n';
    });
    errorMsg += '\nTo update these records:\n';
    errorMsg += '1. Remove these entries from your current form, OR\n';
    errorMsg += '2. Delete the existing records from the database first, then resubmit\n\n';
    errorMsg += 'Note: You cannot have multiple entries for the same date.';
    throw new Error(errorMsg);
  }
  // Enforce monthly hours limit
  const maxPerMonth = parseFloat(settings['MAX_COC_PER_MONTH']) || 40;
  if (monthlyTotal + totalNewHours > maxPerMonth) {
    throw new Error('Monthly COC limit of ' + maxPerMonth + ' hours exceeded');
  }
  // Enforce total balance limit
  const maxBalance = parseFloat(settings['MAX_COC_BALANCE']) || 120;
  if (runningBalance + totalNewCOC > maxBalance) {
    throw new Error('Total COC balance cannot exceed ' + maxBalance + ' hours');
  }
  // Now build rows
  entries.forEach(entry => {
    const date = new Date(year, month - 1, entry.day);
    const result = calculateOvertimeForDate(date, entry.amIn, entry.amOut, entry.pmIn, entry.pmOut);
    if(result.cocEarned <= 0) return; // Don't log entries with 0 hours

    runningBalance += result.cocEarned;
    const recordId = generateRecordId();
    newRows.push([
      recordId,
      employeeId,
      employee.fullName,
      monthYear,
      date,
      result.dayType,
      entry.amIn || '',
      entry.amOut || '',
      entry.pmIn || '',
      entry.pmOut || '',
      result.hoursWorked,
      result.multiplier,
      result.cocEarned,
      new Date(),
      '',
      'Active'
    ]);
    // Mirror the earned entry in the FIFO detail sheet so that CTO consumption
    // and expiration tracking have a single source of truth. Certificate details
    // are filled in once the consolidated certificate is issued.
    addCOCToBalanceDetail(
      employeeId,
      employee.fullName,
      recordId,
      date,
      result.cocEarned
    );
    const ledgerId = generateLedgerId();
    newLedgerRows.push([
      ledgerId,
      employeeId,
      employee.fullName,
      new Date(),
      'COC Earned',
      recordId,
      result.cocEarned,
      0,
      runningBalance,
      monthYear,
      '',
      Session.getActiveUser().getEmail(),
      'COC entry for ' + formatDate(date)
    ]);
  });
  // Append rows to both sheets
  if (newRows.length > 0) {
    recordsSheet.getRange(recordsSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, newLedgerRows.length, newLedgerRows[0].length).setValues(newLedgerRows);
  }
  return {
    added: newRows.length,
    totalNewCOC: totalNewCOC,
    balanceAfter: runningBalance
  };
}

// -----------------------------------------------------------------------------
// CTO applications
// -----------------------------------------------------------------------------

/**
 * Records a CTO application for an employee. Validation ensures hours are a
 * multiple of four, requested days do not exceed five consecutive days,
 * sufficient balance exists, and then writes both the application and ledger
 * entries.
 *
 * @param {string} employeeId The employee ID.
 * @param {number} hours The number of CTO hours requested (must be multiple of 4).
 * @param {Date} startDate The first day of CTO.
 * @param {Date} endDate The last day of CTO.
 * @param {string} remarks Optional remarks.
 * @return {Object} Contains the generated application ID and the new balance.
*/
function recordCTOApplication(employeeId, hours, startDate, endDate, remarks) {
  const db = getDatabase();
  const ctoSheet = db.getSheetByName('CTO_Applications');
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  if (!ctoSheet || !ledgerSheet) throw new Error('Required sheets missing');
  const employee = getEmployeeById(employeeId);
  if (!employee) throw new Error('Employee not found');
  hours = parseFloat(hours);
  if (hours <= 0 || hours % 4 !== 0) throw new Error('CTO hours must be a positive multiple of 4');
  // If the request is for a half‑day (4hrs) or a full day (8hrs) the start and end
  // dates must be the same. This prevents the user from spanning multiple days
  // with only a partial day of CTO, which is prohibited by the business rules.
  if ((hours === 4 || hours === 8) && startDate.getTime() !== endDate.getTime()) {
    throw new Error('For 4 or 8 hour CTO applications the start and end date must be the same');
  }
  // Check consecutive days (inclusive)
  const daysRequested = Math.floor((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  if (daysRequested > 5) throw new Error('Maximum consecutive CTO days is 5');
  const currentBalance = getCurrentCOCBalance(employeeId);
  if (currentBalance < hours) throw new Error('Insufficient COC balance');
  const inclusiveDates = formatDate(startDate) + ' - ' + formatDate(endDate);
  const ctoId = generateCTOId();
  // Record CTO application
  ctoSheet.appendRow([
    ctoId,
    employeeId,
    employee.fullName,
    employee.office,
    hours,
    startDate,
    endDate,
    inclusiveDates,
    currentBalance,
    new Date(),
    'Approved',
    new Date(),
    remarks || ''
  ]);
  // Update ledger
  const ledgerId = generateLedgerId();
  const newBalance = currentBalance - hours;
  ledgerSheet.appendRow([
    ledgerId,
    employeeId,
    employee.fullName,
    new Date(),
    'CTO Used',
    ctoId,
    0,
    hours,
    newBalance < 0 ? 0 : newBalance,
    '',
    '',
    Session.getActiveUser().getEmail(),
    remarks || 'CTO application'
  ]);
  return {
    applicationId: ctoId,
    balanceAfter: newBalance
  };
}

// -----------------------------------------------------------------------------
// Ledger retrieval and reporting
// -----------------------------------------------------------------------------

/**
 * Retrieves the ledger for a specific employee ordered newest first. Each
 * ledger entry is returned as an object with formatted dates. The current
 * balance is computed using getCurrentCOCBalance so that the UI always shows
 * up‑to‑date information.
 *
 * @param {string} employeeId The employee ID.
 * @return {Object} Contains the current balance and an array of ledger entries.
 */
/**
 * FIXED: Retrieves the ledger for a specific employee with robust date handling
 * Handles multiple date formats: ISO dates, US dates, and mixed formats
 *
 * @param {string} employeeId The employee ID.
 * @return {Object} Contains the current balance and an array of ledger entries.
 */
function getLedgerForEmployee(employeeId) {
  try {
    const db = getDatabase();
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    
    if (!ledgerSheet) {
      return { balance: 0, entries: [] };
    }
    
    if (ledgerSheet.getLastRow() <= 1) {
      const balance = getCurrentCOCBalance(employeeId);
      return { balance: balance, entries: [] };
    }
    
    const data = ledgerSheet.getDataRange().getValues();
    const TIME_ZONE = getScriptTimeZone();
    const entries = [];
    
    // CRITICAL FIX: Normalize employee ID for comparison
    const searchId = String(employeeId).trim();
    
    /**
     * Helper to safely parse dates
     */
    function safeDateParse(value) {
      if (!value) return null;
      
      try {
        if (value instanceof Date) {
          return isNaN(value.getTime()) ? null : value;
        }
        
        if (typeof value === 'string') {
          value = value.trim();
          let date = new Date(value);
          if (!isNaN(date.getTime())) {
            return date;
          }
          
          const usDateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
          if (usDateMatch) {
            const [, month, day, year] = usDateMatch;
            date = new Date(year, parseInt(month) - 1, day);
            if (!isNaN(date.getTime())) {
              return date;
            }
          }
        }
        
        const date = new Date(value);
        return isNaN(date.getTime()) ? null : date;
        
      } catch (e) {
        return null;
      }
    }
    
    /**
     * Helper to format dates safely
     */
    function safeFormatDate(value, format) {
      const date = safeDateParse(value);
      if (!date) return 'N/A';
      
      try {
        return Utilities.formatDate(date, TIME_ZONE, format || 'MMMM dd, yyyy');
      } catch (e) {
        return date.toLocaleDateString();
      }
    }
    
    // Process ledger rows in reverse order (newest first)
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      
      // Skip empty rows
      if (!row || row.length === 0 || !row[0]) {
        continue;
      }
      
      // CRITICAL FIX: Normalize row employee ID before comparison
      const rowEmpId = String(row[1] || '').trim();
      
      if (rowEmpId !== searchId) {
        continue;
      }
      
      try {
        const entry = {
          ledgerId: row[0] || '',
          employeeId: row[1] || '',
          employeeName: row[2] || '',
          transactionDate: safeFormatDate(row[3], 'MMMM dd, yyyy'),
          transactionType: row[4] || '',
          referenceId: row[5] || '',
          cocEarned: parseFloat(row[6]) || 0,
          ctoUsed: parseFloat(row[7]) || 0,
          cocBalance: parseFloat(row[8]) || 0,
          monthYearEarned: row[9] || '',
          expirationDate: row[10] ? safeFormatDate(row[10], 'MMMM dd, yyyy') : '',
          processedBy: row[11] || '',
          remarks: row[12] || ''
        };
        
        entries.push(entry);
        
      } catch (rowError) {
        Logger.log('Error processing row ' + i + ': ' + rowError.message);
        continue;
      }
    }
    
    const balance = getCurrentCOCBalance(employeeId);
    
    return { 
      balance: balance, 
      entries: entries 
    };
    
  } catch (error) {
    Logger.log('CRITICAL ERROR in getLedgerForEmployee: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    
    return { 
      balance: 0, 
      entries: [],
      error: error.message 
    };
  }
}

/**
s* Generates report data for different report types. Currently supports
 * “monthly” which summarizes COC earned and CTO used by employee over a
 * specified date range. Additional report types can be added by extending
 * this function.
 *
 * @param {string} type The report type (e.g. "monthly").
 * @param {Date} startDate Inclusive start date of the report.
 * @param {Date} endDate Inclusive end date of the report.
 * @return {Array<Object>} An array of summary objects keyed by employee.
 */
function getReportData(type, startDate, endDate) {
  const db = getDatabase();
  const employees = listEmployees(false);
  if (type === 'monthly') {
    const recSheet = db.getSheetByName('COC_Records');
    const ctoSheet = db.getSheetByName('CTO_Applications');
    const recData = recSheet ? recSheet.getDataRange().getValues() : [];
    const ctoData = ctoSheet ? ctoSheet.getDataRange().getValues() : [];
    const summary = [];
    employees.forEach(emp => {
      let cocTotal = 0;
      let ctoTotal = 0;
      // Sum COC earned within range
      for (let i = 1; i < recData.length; i++) {
        const row = recData[i];
        if (row[1] === emp.id) {
          const d = new Date(row[4]);
          if (d >= startDate && d <= endDate) {
            cocTotal += parseFloat(row[12]) || 0;
          }
        }
      }
      // Sum CTO used within range
      for (let i = 1; i < ctoData.length; i++) {
        const row = ctoData[i];
        if (row[1] === emp.id && row[10] === 'Approved') {
          const d = new Date(row[9]);
          if (d >= startDate && d <= endDate) {
            ctoTotal += parseFloat(row[4]) || 0;
          }
        }
      }
      summary.push({
        employeeId: emp.id,
        employeeName: emp.fullName,
        cocEarned: cocTotal,
        ctoUsed: ctoTotal,
        net: cocTotal - ctoTotal
      });
    });
    return summary;
  }
  return [];
}

// -----------------------------------------------------------------------------
// Holiday management
// -----------------------------------------------------------------------------

/**
* Retrieves a list of holidays from the Holidays sheet. Each entry contains
 * the row number (for editing/deletion), the date, type and description.
 * If the Holidays sheet does not exist an empty array is returned. Dates
 * are returned as JavaScript Date objects to allow the client to format
 * appropriately.
 *
 * @return {Array<Object>} A list of holiday entries.
 */
function listHolidays() {
  const db = getDatabase();
  const sheet = db.getSheetByName('Holidays');
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const dateValue = row[0];
    // Only process rows with a valid date
    if (!dateValue) continue;
    result.push({
      rowNumber: i + 1, // 1‑based index including header
      date: new Date(dateValue),
      type: row[1],
      description: row[2] || ''
    });
  }
  return result;
}

/**
 * Adds a new holiday to the Holidays sheet. If the sheet does not exist
 * it will be created with appropriate headers.
 *
 * @param {Date} date The holiday date.
 * @param {string} type The holiday type (e.g. "Regular", "Special Non‑Working", "Local", "No Work").
 * @param {string} description Optional description.
 * @return {boolean} True on success.
 */
function addHoliday(date, type, description) {
  const db = getDatabase();
  let sheet = db.getSheetByName('Holidays');
  if (!sheet) {
    sheet = db.insertSheet('Holidays');
    // Create header row
    sheet.getRange(1, 1, 1, 3).setValues([
      ['Date', 'Type', 'Description']
    ]);
  }
  sheet.appendRow([date, type, description || '']);
  return true;
}

/**
 * Updates an existing holiday entry. The caller must supply the row number
 * (including the header row) along with the new values. If the row does
 * not exist an error is thrown.
 *
 * @param {number} rowNumber The row number to update (1‑based, includes header).
 * @param {Date} date The new date.
 * @param {string} type The new type.
 * @param {string} description The new description.
 * @return {boolean} True on success.
*/
function updateHoliday(rowNumber, date, type, description) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Holidays');
  if (!sheet) throw new Error('Holidays sheet not found');
  const lastRow = sheet.getLastRow();
  if (rowNumber < 2 || rowNumber > lastRow) throw new Error('Invalid row number');
  sheet.getRange(rowNumber, 1, 1, 3).setValues([[date, type, description || '']]);
  return true;
}

/**
 * Deletes a holiday entry by row number. Throws an error if the row is
 * invalid. Deleting the header row is not permitted.
 *
 * @param {number} rowNumber The row number to delete.
 * @return {boolean} True on success.
 */
function deleteHoliday(rowNumber) {
  const db = getDatabase();
  const sheet = db.getSheetByName('Holidays');
  if (!sheet) throw new Error('Holidays sheet not found');
  const lastRow = sheet.getLastRow();
  if (rowNumber < 2 || rowNumber > lastRow) throw new Error('Invalid row number');
  sheet.deleteRow(rowNumber);
  return true;
}

// -----------------------------------------------------------------------------
// Cancellation and update of COC/CTO entries
// -----------------------------------------------------------------------------

/**
 * Cancels a CTO application and credits the hours back to the employee’s
 * balance. Only applications with status "Approved" can be cancelled.
 * A new ledger entry is created to record the reversal. The CTO record
 * status is set to "Cancelled" and the approval date is preserved.
 *
 * @param {string} ctoId The application ID (e.g. CTO‑20250101123000).
 * @param {string} remarks Optional remarks explaining the cancellation.
 * @return {Object} Contains the new balance after cancellation.
 */
function cancelCTOApplication(ctoId, remarks) {
  const db = getDatabase();
  const ctoSheet = db.getSheetByName('CTO_Applications');
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  if (!ctoSheet || !ledgerSheet) throw new Error('Required sheets missing');
  const data = ctoSheet.getDataRange().getValues();
  let rowIndex = -1;
  let record = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === ctoId) {
      rowIndex = i + 1;
      record = data[i];
      break;
    }
  }
  if (rowIndex === -1) throw new Error('CTO application not found');
  // record fields: [0]=ID, [1]=employeeId, [2]=employeeName, [3]=office,
  // [4]=hours, [5]=startDate, [6]=endDate, [7]=inclusiveDates,
  // [8]=balanceBefore, [9]=applicationDate, [10]=status, [11]=approvalDate, [12]=remarks
  const status = record[10];
  if (status !== 'Approved') throw new Error('Only approved CTO applications can be cancelled');
  const employeeId = record[1];
  const hours = parseFloat(record[4]) || 0;
  // Update CTO status
  ctoSheet.getRange(rowIndex, 11).setValue('Cancelled');
  ctoSheet.getRange(rowIndex, 13).setValue(remarks || 'CTO cancelled');
  // Credit back hours: new balance = current + hours
  const currentBalance = getCurrentCOCBalance(employeeId);
  const newBalance = currentBalance + hours;
  // Create ledger reversal entry
  const ledgerId = generateLedgerId();
  const employee = getEmployeeById(employeeId);
  ledgerSheet.appendRow([
    ledgerId,
    employeeId,
    employee.fullName,
    new Date(),
    'CTO Cancelled',
    ctoId,
    0,
    -hours,
    newBalance,
    '',
    '',
    Session.getActiveUser().getEmail(),
    remarks || 'CTO cancellation'
  ]);
  return { balanceAfter: newBalance };
}

/**
 * Cancels a COC record. The record status is set to "Cancelled" and a
 * reversal ledger entry is created. Users may then re‑record a new COC
 * entry for the same date. The original COC earned hours are deducted
 * from the employee’s balance.
 *
 * @param {string} recordId The COC record ID.
 * @param {string} remarks Optional remarks explaining the cancellation.
 * @return {Object} Contains the new balance after cancellation.
 */
function cancelCOCRecord(recordId, remarks) {
  const db = getDatabase();
  const recSheet = db.getSheetByName('COC_Records');
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  if (!recSheet || !ledgerSheet) throw new Error('Required sheets missing');
  const data = recSheet.getDataRange().getValues();
  let rowIndex = -1;
  let record = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === recordId) {
      rowIndex = i + 1;
      record = data[i];
      break;
    }
  }
  if (rowIndex === -1) throw new Error('COC record not found');
  const status = record[15];
  // Only active records can be cancelled
  if (status === 'Cancelled') throw new Error('COC record is already cancelled');
  // Mark record as cancelled
  recSheet.getRange(rowIndex, 16).setValue('Cancelled');
  // Deduct COC earned
  const employeeId = record[1];
  const cocEarned = parseFloat(record[12]) || 0;
  const currentBalance = getCurrentCOCBalance(employeeId);
  const newBalance = currentBalance - cocEarned;
  const employee = getEmployeeById(employeeId);
  // Append reversal in ledger
  const ledgerId = generateLedgerId();
  ledgerSheet.appendRow([
    ledgerId,
    employeeId,
    employee.fullName,
    new Date(),
    'COC Cancelled',
    recordId,
    -cocEarned,
    0,
    newBalance,
    record[3], // Month‑Year Earned
    record[14], // Expiration Date
    Session.getActiveUser().getEmail(),
    remarks || 'COC cancellation'
  ]);
  return { balanceAfter: newBalance };
}

// -----------------------------------------------------------------------------
// COC certificate generation
// -----------------------------------------------------------------------------

/**
 * Generates a certificate document for a specific COC record. If a certificate
 * has already been generated for this record it will simply return the
 * existing certificate URL. Certificates are tracked in a dedicated sheet
 * named "COC_Certificates" which stores the record ID, employee ID, hours
 * earned, date rendered and the certificate document URL. A new Google
 * Document is created with a simple layout summarising the overtime and
 * signed by the processing officer.
 *
 * @param {string} recordId The COC record ID.
 * @return {Object} An object containing the certificate URL and document ID.
 */
function generateCOCCertificate(recordId) {
  const db = getDatabase();
  const recSheet = db.getSheetByName('COC_Records');
  if (!recSheet) throw new Error('COC_Records sheet not found');
  const data = recSheet.getDataRange().getValues();
  let record = null;
  let recordRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === recordId) {
      record = data[i];
      recordRowIndex = i + 1;
      break;
    }
  }
  if (!record || recordRowIndex === -1) throw new Error('COC record not found');
  // Check certificate sheet for existing certificate
  let certSheet = db.getSheetByName('COC_Certificates');
  if (!certSheet) {
    certSheet = db.insertSheet('COC_Certificates');
    certSheet.getRange(1, 1, 1, 9).setValues([
      ['Record ID','Employee ID','Employee Name','Date Rendered','Hours Worked','COC Earned','Certificate URL','PDF URL','Issued Date']
    ]);
  }
  const certData = certSheet.getDataRange().getValues();
  for (let i = 1; i < certData.length; i++) {
    if (certData[i][0] === recordId) {
      // Already exists - return both doc and PDF URLs
      return { docUrl: certData[i][6], pdfUrl: certData[i][7] || certData[i][6], docId: '' };
    }
  }
  const employeeId = record[1];
  const employeeName = record[2];
  const dateRendered = new Date(record[4]);
  const hoursWorked = parseFloat(record[10]) || 0;
  const cocEarned = parseFloat(record[12]) || 0;
  // Create document
  const docName = 'COC Certificate - ' + employeeName + ' - ' + formatDate(dateRendered);
  const doc = DocumentApp.create(docName);
  const body = doc.getBody();
  body.appendParagraph('Republic of the Philippines').setBold(true).setFontSize(12).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Department of Public Works and Highways').setBold(true).setFontSize(12).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Certificate of Compensatory Overtime Credit (COC)').setBold(true).setFontSize(14).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('\n');
  body.appendParagraph('This certifies that ' + employeeName + ' has rendered overtime services on ' + formatDate(dateRendered) + ' totaling ' + hoursWorked.toFixed(2) + ' hour(s).');
  body.appendParagraph('In accordance with Civil Service Commission policies, the services rendered have been credited as ' + cocEarned.toFixed(2) + ' hour(s) of Compensatory Overtime Credit.');
  body.appendParagraph('\n');
  const issueDate = new Date();
  body.appendParagraph('Processed by: ' + Session.getActiveUser().getEmail());
  body.appendParagraph('Date Issued: ' + formatDate(issueDate));
  doc.saveAndClose();
  const docId = doc.getId();
  const docUrl = doc.getUrl();

  // Create PDF export URL
  const pdfUrl = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf';

  // Calculate standardized expiration and persist back to the record sheet
  const expirationDate = calculateCertificateExpiration(issueDate);
  recSheet.getRange(recordRowIndex, 15).setValue(expirationDate);

  // Synchronise with the FIFO detail sheet so expiration logic stays correct
  const detailSheet = ensureCOCBalanceDetailSheet();
  const detailData = detailSheet.getDataRange().getValues();
  const certificateId = 'CERT-' + Utilities.formatDate(issueDate, getScriptTimeZone(), 'yyyyMMddHHmmssSSS');
  for (let i = 1; i < detailData.length; i++) {
    if (detailData[i][DETAIL_COLS.RECORD_ID] === recordId) {
      detailSheet.getRange(i + 1, DETAIL_COLS.CERTIFICATE_DATE + 1).setValue(issueDate);
      detailSheet.getRange(i + 1, DETAIL_COLS.EXPIRATION_DATE + 1).setValue(expirationDate);
      detailSheet.getRange(i + 1, DETAIL_COLS.CERTIFICATE_ID + 1).setValue(certificateId);
      const existingRemarks = detailSheet.getRange(i + 1, DETAIL_COLS.REMARKS + 1).getValue() || '';
      const remarkNote = `[${Utilities.formatDate(issueDate, getScriptTimeZone(), 'yyyy-MM-dd')}] Certificate issued (${certificateId}).`;
      detailSheet.getRange(i + 1, DETAIL_COLS.REMARKS + 1).setValue(existingRemarks ? existingRemarks + '\n' + remarkNote : remarkNote);
      break;
    }
  }

  const ledgerSheet = db.getSheetByName('COC_Ledger');
  if (ledgerSheet) {
    const ledgerData = ledgerSheet.getDataRange().getValues();
    for (let i = 1; i < ledgerData.length; i++) {
      if (ledgerData[i][5] === recordId) {
        ledgerSheet.getRange(i + 1, 11).setValue(expirationDate);
      }
    }
  }

  // Record certificate
  certSheet.appendRow([
    recordId,
    employeeId,
    employeeName,
    dateRendered,
    hoursWorked,
    cocEarned,
    docUrl,
    pdfUrl,
    issueDate
  ]);
  return { docUrl: docUrl, pdfUrl: pdfUrl, docId: docId, expirationDate: expirationDate, certificateId: certificateId };
}

// -----------------------------------------------------------------------------
// API wrappers for holiday management, cancellation and certificates
// -----------------------------------------------------------------------------

function apiListHolidays() {
  return listHolidays();
}

function apiAddHoliday(date, type, description) {
  return addHoliday(date, type, description);
}

function apiUpdateHoliday(rowNumber, date, type, description) {
  return updateHoliday(rowNumber, date, type, description);
}

function apiDeleteHoliday(rowNumber) {
  return deleteHoliday(rowNumber);
}

function apiCancelCTO(ctoId, remarks) {
  return cancelCTOApplication(ctoId, remarks);
}

function apiCancelCOC(recordId, remarks) {
  return cancelCOCRecord(recordId, remarks);
}

function apiGenerateCOCCertificate(recordId) {
  return generateCOCCertificate(recordId);
}

/**
 * OLD VERSION - DEPRECATED
 * This function has been replaced by the newer version at line 6383+.
 * Kept here for reference but should not be used.
 * The newer version uses MONTH_YEAR column filtering instead of parsing dates.
 */
/*
function apiListCOCRecordsForMonth_OLD(employeeId, month, year) {
  const db = getDatabase();
  const recSheet = db.getSheetByName('COC_Records');
  if (!recSheet) throw new Error('COC_Records sheet not found');

  const certSheet = db.getSheetByName('COC_Certificates');
  const TIME_ZONE = getScriptTimeZone(); // Get timezone

  // Get all records
  const recData = recSheet.getDataRange().getValues();

  // Get all certificates (if sheet exists)
  const certData = certSheet ? certSheet.getDataRange().getValues() : [];

  // Build certificate lookup by recordId
  const certMap = {};
  for (let i = 1; i < certData.length; i++) {
    const recordId = certData[i][0];
    certMap[recordId] = {
      certificateUrl: certData[i][6],
      pdfUrl: certData[i][7],
      issuedDate: certData[i][8]
    };
  }

  const targetMonth = month - 1; // JS months are 0-indexed
  const targetYear = year;
  const results = [];

  Logger.log('apiListCOCRecordsForMonth: Looking for employeeId=' + employeeId + ', month=' + month + ', year=' + year);

  for (let i = 1; i < recData.length; i++) {
    const row = recData[i];
    const recEmployeeId = String(row[1] || '').trim(); // Convert to string and trim
    const status = String(row[15] || '').trim();

    // Skip if not matching employee or if cancelled
    if (recEmployeeId !== employeeId) continue;
    if (status.toLowerCase() === 'cancelled') continue;

    const recordId = row[0];

    // Parse date carefully
    let dateRendered;
    if (row[4] instanceof Date) {
      dateRendered = new Date(row[4]);
    } else if (row[4]) {
      dateRendered = new Date(row[4]);
      if (isNaN(dateRendered.getTime())) {
        Logger.log(`WARNING: Invalid date at row ${i}: ${row[4]}`);
        continue;
      }
    } else {
      Logger.log(`WARNING: Missing date at row ${i}`);
      continue;
    }

    // Check if date matches target month/year
    if (dateRendered.getMonth() !== targetMonth || dateRendered.getFullYear() !== targetYear) {
      continue;
    }

    const dayType = row[5] || '';
    const hoursWorked = parseFloat(row[10]) || 0;
    const cocEarned = parseFloat(row[12]) || 0;

    // Get certificate info if exists
    const cert = certMap[recordId];

    // FIX: Use Utilities.formatDate() instead of formatDate()
    let displayDate;
    try {
      displayDate = Utilities.formatDate(dateRendered, TIME_ZONE, 'MMM dd, yyyy');
    } catch (e) {
      // Fallback if formatting fails
      displayDate = dateRendered.toLocaleDateString();
      Logger.log(`WARNING: Date formatting failed for ${dateRendered}, using fallback`);
    }

    // Format certificate date if exists
    let certificateDisplayDate = null;
    if (cert && cert.issuedDate) {
      try {
        const certDate = new Date(cert.issuedDate);
        if (!isNaN(certDate.getTime())) {
          certificateDisplayDate = Utilities.formatDate(certDate, TIME_ZONE, 'MMM dd, yyyy');
        }
      } catch (e) {
        Logger.log(`WARNING: Certificate date formatting failed for ${cert.issuedDate}`);
      }
    }

    results.push({
      recordId: recordId,
      displayDate: displayDate, // NOW PROPERLY FORMATTED!
      dayType: dayType,
      hoursWorked: hoursWorked,
      cocEarned: cocEarned,
      certificateUrl: cert ? cert.certificateUrl : null,
      pdfUrl: cert ? cert.pdfUrl : null,
      certificateDisplayDate: certificateDisplayDate
    });
  }

  // Sort by date
  results.sort((a, b) => {
    // Extract dates from displayDate strings for sorting
    const dateA = new Date(a.displayDate);
    const dateB = new Date(b.displayDate);
    return dateA - dateB;
  });

  Logger.log(`Found ${results.length} records for employee ${employeeId} in ${month}/${year}`);
  return results;
}
*/

/**
 * Iterates through all COC records and marks entries as expired when their
 * expiration date has passed. This function should be scheduled as a
 * time‑driven trigger (e.g. daily at midnight) so that expired entries are
 * automatically updated without user intervention.
 *
 * @return {number} The number of entries marked as expired.
 */
function checkExpiredCOC() {
  const db = getDatabase();
  const recSheet = db.getSheetByName('COC_Records');
  if (!recSheet) return 0;
  const data = recSheet.getDataRange().getValues();
  const today = new Date();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const status = data[i][15];
    const expDate = data[i][14];
    if (status === 'Active' && expDate) {
      const exp = new Date(expDate);
      if (exp.getTime() <= today.getTime()) {
        recSheet.getRange(i + 1, 16).setValue('Expired');
        count++;
      }
    }
  }
  return count;
}

/**
A* Computes summary statistics for the dashboard: number of active employees,
 * count of active COC records, count of CTO applications and the number of
 * active COC entries expiring within the next 30 days. The calculation is
 * performed server‑side to avoid transferring large datasets to the client.
 *
 * @return {Object} Dashboard statistics.
 */
function getDashboardStats() {
  const db = getDatabase();
  // Active employees
  const employees = listEmployees(false);
  const employeeCount = employees.length;
  // COC records
  let activeCOCCount = 0;
  let expiringSoonCount = 0;
  const recSheet = db.getSheetByName('COC_Records');
  if (recSheet) {
    const recData = recSheet.getDataRange().getValues();
    const today = new Date();
    const thirty = 30 * 24 * 60 * 60 * 1000;
    for (let i = 1; i < recData.length; i++) {
      const status = recData[i][15];
      if (status === 'Active') {
        activeCOCCount++;
        const exp = recData[i][14];
        if (exp) {
          const expDate = new Date(exp);
          const diff = expDate.getTime() - today.getTime();
          if (diff >= 0 && diff <= thirty) {
            expiringSoonCount++;
          }
        }
      }
    }
  }
  // CTO applications
  let ctoCount = 0;
  const ctoSheet = db.getSheetByName('CTO_Applications');
  if (ctoSheet) {
    const ctoData = ctoSheet.getDataRange().getValues();
    ctoCount = Math.max(0, ctoData.length - 1);
  }
  return {
    employees: employeeCount,
    activeCOC: activeCOCCount,
    ctoApplications: ctoCount,
    expiringSoon: expiringSoonCount
  };
}

/**
 * Server API wrapper for dashboard stats.
 *
 * @return {Object} Stats for dashboard display.
*/
function apiGetDashboardStats() {
  return getDashboardStats();
}

/**
 * Retrieves the latest ledger transactions across all employees. This is used
 * by the dashboard to show a recap of recent activity. The results are
 * returned in reverse chronological order (newest first). Each entry includes
 * the transaction date, employee name, type, reference and remarks.
 *
 * @param {number} limit The maximum number of entries to return.
 * @return {Array<Object>} An array of recent ledger entries.
 */
function getRecentActivities(limit) {
  const db = getDatabase();
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  if (!ledgerSheet) return [];
  const data = ledgerSheet.getDataRange().getValues();
  const entries = [];
  // Iterate backwards to gather newest first
  for (let i = data.length - 1; i >= 1 && entries.length < (limit || 10); i--) {
    const row = data[i];
    entries.push({
      transactionDate: formatDate(row[3]),
      employeeName: row[2],
      transactionType: row[4],
      referenceId: row[5],
      remarks: row[12]
    });
  }
  return entries;
}

/**
 * API wrapper for recent ledger activity. Accepts an optional limit.
 *
 * @param {number} limit Maximum number of records to return.
 * @return {Array<Object>} Recent activities.
 */
function apiGetRecentActivities(limit) {
  return getRecentActivities(limit);
}

// -----------------------------------------------------------------------------
// UI entry points – these functions are called from the HTML forms via
// google.script.run. Each returns plain objects suitable for JSON serialisation.
// -----------------------------------------------------------------------------

/**
 * API wrapper for getting dropdown options for Positions and Offices.
 * @return {Object} An object { positions: [...], offices: [...] }.
 */
function apiGetDropdownOptions() {
  return getDropdownOptions();
}

/**
* API wrapper for listing employees. Accepts an optional boolean flag to
 * include inactive employees. Without parameters the list defaults to
 * active employees only. This change allows HTML clients to request
 * either view while remaining backwards compatible with existing calls.
 *
 * @param {boolean} includeInactive Whether to include inactive employees.
 * @return {Array<Object>} List of employee objects.
 */
function apiListEmployees(includeInactive) {
  // Explicitly coerce undefined/null to false. When a parameter is passed
  // through google.script.run it will arrive as the first argument.
  return listEmployees(Boolean(includeInactive));
}

function apiGetEmployee(employeeId) {
  return getEmployeeById(employeeId);
}

/**
 * Get employee's current COC balance
 * This reads from the COC_Ledger to get the most recent balance
 * 
 * @param {string} empId - Employee ID (e.g., "EMP008")
 * @returns {number} Current COC balance in hours
 */
function apiGetBalance(empId) {
  try {
    // Use the correct spreadsheet ID
    let ss;
    if (typeof CONFIG !== 'undefined' && CONFIG.SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } else if (typeof SPREADSHEET_ID !== 'undefined') {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      const SPREADSHEET_ID = '1LIQRnQb7lL-6hdSpsOwq-XVwRM6Go6u4q6RkSB2ZCXo';
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    
    // Read from COC_Ledger sheet (NOT COC_Balance_Detail)
    const ledgerSheet = ss.getSheetByName('COC_Ledger');
    
    if (!ledgerSheet) {
      Logger.log('COC_Ledger sheet not found');
      return 0;
    }
    
    const data = ledgerSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const empIdCol = headers.indexOf('Employee ID');
    const transDateCol = headers.indexOf('Transaction Date');
    const balanceCol = headers.indexOf('Balance After'); // Use the correct header name
    
    if (empIdCol === -1 || balanceCol === -1 || transDateCol === -1) {
      Logger.log('Required columns not found in COC_Ledger');
      return 0;
    }
    
    // Find the most recent transaction for this employee
    let latestBalance = 0;
    let latestDate = null;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[empIdCol] === empId) {
        const transDate = new Date(row[transDateCol]);
        const balance = parseFloat(row[balanceCol]) || 0;
        
        if (!latestDate || transDate > latestDate) {
          latestDate = transDate;
          latestBalance = balance;
        }
      }
    }
    
    Logger.log('Balance for ' + empId + ': ' + latestBalance);
    return latestBalance;
    
  } catch (error) {
    Logger.log('Error in apiGetBalance: ' + error.toString());
    return 0;
  }
}


/**
 * Alternative: Get balance from COC_Balance_Detail (FIFO method)
 * Use this if you want to calculate from remaining hours in detail sheet
 */
function apiGetBalanceFromDetail(empId) {
  try {
    let ss;
    if (typeof CONFIG !== 'undefined' && CONFIG.SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } else if (typeof SPREADSHEET_ID !== 'undefined') {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } else {
      const SPREADSHEET_ID = '1LIQRnQb7lL-6hdSpsOwq-XVwRM6Go6u4q6RkSB2ZCXo';
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    
    const detailSheet = ss.getSheetByName('COC_Balance_Detail');
    
    if (!detailSheet) {
      Logger.log('COC_Balance_Detail sheet not found');
      return 0;
    }
    
    const data = detailSheet.getDataRange().getValues();
    const headers = data[0];
    
    const empIdCol = headers.indexOf('Employee ID');
    const hoursRemainingCol = headers.indexOf('Hours Remaining');
    const statusCol = headers.indexOf('Status');
    
    if (empIdCol === -1 || hoursRemainingCol === -1) {
      Logger.log('Required columns not found in COC_Balance_Detail');
      return 0;
    }
    
    let totalBalance = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (row[empIdCol] === empId) {
        const status = row[statusCol];
        // Only count Active entries
        if (status === 'Active') {
          const hoursRemaining = parseFloat(row[hoursRemainingCol]) || 0;
          totalBalance += hoursRemaining;
        }
      }
    }
    
    Logger.log('Balance from detail for ' + empId + ': ' + totalBalance);
    return totalBalance;
    
  } catch (error) {
    Logger.log('Error in apiGetBalanceFromDetail: ' + error.toString());
    return 0;
  }
}


/**
 * Test function
 */
function testGetBalance() {
  const empId = 'EMP008'; // Patrick E Tan
  
  const balanceFromLedger = apiGetBalance(empId);
  Logger.log('Balance from Ledger: ' + balanceFromLedger);
  
  const balanceFromDetail = apiGetBalanceFromDetail(empId);
  Logger.log('Balance from Detail: ' + balanceFromDetail);
}

// --- END MODIFICATION ---

function apiCalculateOvertime(year, month, day, amIn, amOut, pmIn, pmOut) {
  const date = new Date(year, month - 1, day);
  return calculateOvertimeForDate(date, amIn, amOut, pmIn, pmOut);
}

function apiRecordCTO(employeeId, hours, startDate, endDate, remarks) {
  return recordCTOApplication(employeeId, hours, startDate, endDate, remarks);
}

/**
 * FINAL FIX - Add this to Code.gs
 * 
 * Replace the apiGetLedger function (around line 2486) with this version
 * that has extensive logging
 */

function apiGetLedger(employeeId) {
  try {
    const result = getLedgerForEmployee(employeeId);
    
    // Convert to serializable format
    const safeResult = {
      balance: parseFloat(result.balance) || 0,
      entries: []
    };
    
    if (result.entries && Array.isArray(result.entries)) {
      result.entries.forEach(function(entry) {
        safeResult.entries.push({
          ledgerId: String(entry.ledgerId || ''),
          employeeId: String(entry.employeeId || ''),
          employeeName: String(entry.employeeName || ''),
          transactionDate: String(entry.transactionDate || ''),
          transactionType: String(entry.transactionType || ''),
          referenceId: String(entry.referenceId || ''),
          cocEarned: parseFloat(entry.cocEarned) || 0,
          ctoUsed: parseFloat(entry.ctoUsed) || 0,
          cocBalance: parseFloat(entry.cocBalance) || 0,
          monthYearEarned: String(entry.monthYearEarned || ''),
          expirationDate: String(entry.expirationDate || ''),
          processedBy: String(entry.processedBy || ''),
          remarks: String(entry.remarks || '')
        });
      });
    }
    
    return safeResult;
    
  } catch (error) {
    Logger.log('ERROR in apiGetLedger: ' + error.message);
    return { balance: 0, entries: [] };
  }
}


/**
 * TEST THIS NEW VERSION
 * Run this after replacing apiGetLedger
 */
function testApiGetLedgerDirect() {
  const empId = 'EMP002';
  Logger.log('Testing apiGetLedger for ' + empId);
  
  const result = apiGetLedger(empId);
  
  Logger.log('\n=== RESULT ===');
  Logger.log('Type: ' + typeof result);
  Logger.log('Balance: ' + result.balance);
  Logger.log('Entries: ' + result.entries.length);
  
  if (result.entries.length > 0) {
    Logger.log('\nFirst entry:');
    Logger.log(JSON.stringify(result.entries[0], null, 2));
  }
  
  if (result.error) {
    Logger.log('\nERROR: ' + result.error);
  }
}

// --- Using FIFO version ---
// function apiAddEmployee(data) {
//   return addEmployee(data);
// }

function apiUpdateEmployee(employeeId, data) {
  return updateEmployee(employeeId, data);
}

function apiGetReport(type, startDate, endDate) {
  return getReportData(type, startDate, endDate);
}

// --- This is an old version, replaced by the validation-enabled one below ---
// function apiRecordCOC(employeeId, month, year, entries) {
//   return recordCOCEntries(employeeId, month, year, entries);
// }

// --- This is an old version, replaced by apiCheckAndExpireCOC ---
// function apiCheckExpiredCOC() {
//   return checkExpiredCOC();
// }

// -----------------------------------------------------------------------------
// Spreadsheet UI menu
// -----------------------------------------------------------------------------

/**
 * Adds a custom menu to the spreadsheet upon opening. The menu items open
 * each of the HTML forms as a modal dialog. Users can navigate to all
 * system functions from the “COC/CTO System” menu.
 */
// ============================================================================
// ADD TO MENU (update your onOpen function)
// ============================================================================

/**
 * Update your existing onOpen function to include the new menu item
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('COC/CTO System')
    .addItem('Dashboard', 'showDashboard')
    .addSeparator()
    .addItem('Record COC Earned', 'showMonthlyCOCEntry')
    .addItem('Record CTO Application', 'showCTORecordForm')
    .addItem('CTO Applications Manager', 'showCTOApplicationsManager')
    .addSeparator()
    .addItem('Employee Ledger', 'showEmployeeLedger')
    .addItem('Employee Manager', 'showEmployeeManager')
    .addItem('Holiday Manager', 'showHolidayManager')
    .addItem('Import', 'showHistoricalimport')
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addItem('Reports', 'showReports')
    .addSeparator()
    .addSubMenu(ui.createMenu('Admin Tools')
      .addItem('Run Data Migration (MONTH_YEAR)', 'runCOCRecordsMigration')
      .addItem('Debug COC Records', 'debugMariaOctober2025'))
    .addToUi();
}

/**
 * Menu function to run the COC records MONTH_YEAR migration
 */
function runCOCRecordsMigration() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Run Data Migration',
    'This will populate the MONTH_YEAR column for all existing COC records that don\'t have it set.\n\nThis is safe to run multiple times.\n\nProceed?',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    try {
      const result = migrateCOCRecordsMonthYear();
      if (result.success) {
        const message = `Successfully updated ${result.updatedCount} records.\n` +
          `Format fixes: ${result.fixedCount || 0}\n\n` +
          `Please refresh the page and try viewing records again.\n\n` +
          `Check the Execution log (View > Execution log) for details.`;
        ui.alert('Migration Complete', message, ui.ButtonSet.OK);
      } else {
        ui.alert('Migration Failed', result.message || 'An unknown error occurred.', ui.ButtonSet.OK);
      }
    } catch (e) {
      ui.alert('Migration Error', e.message || String(e), ui.ButtonSet.OK);
    }
  }
}

/**
 * Menu function to debug COC records for Maria L Garcia in October 2025
 * This helps diagnose the "No entries recorded" issue
 */
function debugMariaOctober2025() {
  const ui = SpreadsheetApp.getUi();

  // Get Maria's employee ID - you may need to adjust this
  const employeeId = ui.prompt(
    'Enter Employee ID',
    'Enter the Employee ID to debug (e.g., EMP001):',
    ui.ButtonSet.OK_CANCEL
  );

  if (employeeId.getSelectedButton() === ui.Button.OK) {
    const empId = employeeId.getResponseText();

    try {
      // Run debug function
      const result = debugCOCRecords(empId, 10, 2025);

      const message = `Debug Results:\n\n` +
        `Total records for ${empId}: ${result.totalRecordsForEmployee}\n` +
        `Matching October 2025: ${result.matchingRecords}\n\n` +
        `Check the Execution log (View > Execution log) for detailed information.`;

      ui.alert('Debug Complete', message, ui.ButtonSet.OK);
    } catch (e) {
      ui.alert('Debug Error', e.message || String(e), ui.ButtonSet.OK);
    }
  }
}

/**
 * Opens the CTO Applications Manager
 */
function showCTOApplicationsManager() {
  const html = HtmlService.createHtmlOutputFromFile('CTOApplicationsManager')
    .setWidth(1400)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'CTO Applications Manager');
}

function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'System Settings');
}

/**
 * The following functions open their corresponding HTML files as modal
 * dialogs. The widths/heights can be adjusted as needed to fit the contents.
 */
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard.html')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard');
}

function showMonthlyCOCEntry() {
  const html = HtmlService.createHtmlOutputFromFile('MonthlyCOCEntry')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Monthly COC Entry');
}

function showCTORecordForm() {
  const html = HtmlService.createHtmlOutputFromFile('CTORecordForm')
    .setWidth(1000)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'CTO Application');
}

function showEmployeeLedger() {
  const html = HtmlService.createHtmlOutputFromFile('EmployeeLedger')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Employee Ledger');
}

function showEmployeeManager() {
  const html = HtmlService.createHtmlOutputFromFile('EmployeeManager')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Employee Manager');
}

function showHistoricalimport() {
  const html = HtmlService.createHtmlOutputFromFile('Historicalimport')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Historical Import');
}

function showReports() {
  const html = HtmlService.createHtmlOutputFromFile('Reports')
    .setWidth(1200)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Reports');
}

/**
 * Opens the Holiday Manager modal dialog. This interface allows HR staff
 * to add, edit and delete holidays and no‑work days. Holidays are stored
 * in a dedicated sheet and used to classify overtime dates correctly.
 */
function showHolidayManager() {
  const html = HtmlService.createHtmlOutputFromFile('HolidayManager')
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Holiday Manager');
}

function navigateToPage(pageName) {
  // We use the 'show' functions that already exist in your Code.gs
  try {
    if (pageName === 'MonthlyCOCEntry') {
      showMonthlyCOCEntry();
    } else if (pageName === 'CTORecordForm') {
      showCTORecordForm();
    } else if (pageName === 'EmployeeLedger') {
      showEmployeeLedger();
    } else if (pageName === 'EmployeeManager') {
      showEmployeeManager();
    } else if (pageName === 'Reports') {
      showReports();
    } else if (pageName === 'HolidayManager') {
      showHolidayManager();
    } else {
      // Fallback in case of an unknown pageName
      showDashboard();
    }
  } catch (e) {
    Logger.log('Error navigating to page: ' + pageName + '. Error: ' + e.message);
    // If it fails, just show the dashboard again
    showDashboard();
  }
}

// ============================================================================
// ADD THESE FUNCTIONS TO YOUR Code.gs
// These support the MonthlyCOCEntry.html with COC limits validation
// ============================================================================

/**
 * Gets the total COC earned for the specified month.
 * @param {string} employeeId The ID of the employee.
 * @param {number} month The month (1-12) from the dropdown.
 * @param {number} year The year (e.g., 2025) from the dropdown.
 * @returns {object} An object containing { currentMonthTotal, monthlyRemaining, etc. }.
 */
function apiGetEmployeeCOCStats(employeeId, month, year) {
  const db = getDatabase(); // Assuming you have this function
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  const TIME_ZONE = getScriptTimeZone(); // Assuming you have this function

  // Get current balance (This is the TOTAL balance, which is fine)
  const currentBalance = getCurrentCOCBalanceFromDetail(employeeId); // Assuming you have this function

  // Get current month's COC total
  const now = new Date(); // We use this *only* as a fallback

  // --- START OF FIX ---
  
  // 1. Use the passed-in month and year. If they are missing, use the current month/year.
  const targetMonth = month || (now.getMonth() + 1);
  const targetYear = year || now.getFullYear();

  // 2. Create the target Month-Year string in YYYY-MM format
  //    This format MUST match your 'COC_Balance_Detail' sheet's 'Month Year' column (Column F)
  //    Based on your CSV, it looks like it's "YYYY-MM" (e.g., "2025-10").
  const targetMonthYear = `${targetYear}-${String(targetMonth).padStart(2, '0')}`;
  
  // --- END OF FIX ---

  let currentMonthTotal = 0;

  if (detailSheet && detailSheet.getLastRow() > 1) {
    const data = detailSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const empId = row[1]; // Employee ID (Column B)
      
      // --- START OF FIX 2 ---
      // We read the 'Month Year' column directly. It's much faster and more reliable.
      // This assumes 'Month Year' is in Column F (index 5)
      const rowMonthYear = row[5]; 
      const hoursEarned = parseFloat(row[8]) || 0; // 'Hours Earned' (Column I)

      // 3. Compare against the targetMonthYear string
      if (empId === employeeId && rowMonthYear === targetMonthYear) {
        currentMonthTotal += hoursEarned;
      }
      // --- END OF FIX 2 ---
    }
  }

  // This return object matches your original function
  return {
    employeeId: employeeId,
    currentBalance: currentBalance,
    balanceLimit: 120,
    balanceRemaining: 120 - currentBalance,
    currentMonthTotal: currentMonthTotal, // This will now be for the SELECTED month
    monthlyLimit: 40,
    monthlyRemaining: 40 - currentMonthTotal,
    canAddThisMonth: 40 - currentMonthTotal,
    canAddTotal: Math.min(40 - currentMonthTotal, 120 - currentBalance)
  };
}

/**
 * Get current COC balance from COC_Balance_Detail (active entries only)
 * More accurate than ledger-based calculation
 */
function getCurrentCOCBalanceFromDetail(employeeId) {
  const db = getDatabase();
  const detailSheet = db.getSheetByName('COC_Balance_Detail');

  if (!detailSheet || detailSheet.getLastRow() < 2) {
    // Fallback to original method if detail sheet doesn't exist
    return getCurrentCOCBalance(employeeId);
  }

  const data = detailSheet.getDataRange().getValues();
  let totalRemaining = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const empId = row[1]; // Employee ID
    const hoursRemaining = row[6]; // Hours Remaining
    const status = row[8]; // Status

    if (empId === employeeId && status === 'Active' && hoursRemaining > 0) {
      totalRemaining += hoursRemaining;
    }
  }

  return totalRemaining;
}

/**
 * Check if adding COC hours would exceed the monthly 40-hour limit
 */
function checkMonthlyLimitForMonth(employeeId, monthYear, hoursToAdd) {
  const db = getDatabase();
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  const TIME_ZONE = getScriptTimeZone();
  const MONTHLY_LIMIT = 40;

  if (!detailSheet || detailSheet.getLastRow() < 2) {
    return {
      valid: true,
      currentMonthTotal: 0,
      message: 'OK'
    };
  }

  const data = detailSheet.getDataRange().getValues();
  let monthTotal = 0;

  // Sum all COC earned in the same month for this employee
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const empId = row[1]; // Employee ID
    const dateEarned = new Date(row[4]); // Date Earned
    const hoursEarned = row[5]; // Hours Earned

    if (empId === employeeId) {
      // Use MM-yyyy format to match the monthYear parameter
      const earnedMonthYear = Utilities.formatDate(dateEarned, TIME_ZONE, 'MM-yyyy');
      if (earnedMonthYear === monthYear) {
        monthTotal += hoursEarned;
      }
    }
  }

  const newTotal = monthTotal + hoursToAdd;

  if (newTotal > MONTHLY_LIMIT) {
    return {
      valid: false,
      currentMonthTotal: monthTotal,
      message: `Monthly limit exceeded! Current: ${monthTotal.toFixed(2)} hrs + New: ${hoursToAdd.toFixed(2)} hrs = ${newTotal.toFixed(2)} hrs. Maximum: ${MONTHLY_LIMIT} hrs/month`
    };
  }

  return {
    valid: true,
    currentMonthTotal: monthTotal,
    message: `OK. Monthly total: ${newTotal.toFixed(2)} hrs / ${MONTHLY_LIMIT} hrs`
  };
}

/**
 * Check if adding COC hours would exceed the 120-hour total balance limit
 */
function checkTotalBalanceLimitForEmployee(employeeId, hoursToAdd) {
  const BALANCE_LIMIT = 120;
  const currentBalance = getCurrentCOCBalanceFromDetail(employeeId);
  const newBalance = currentBalance + hoursToAdd;

  if (newBalance > BALANCE_LIMIT) {
    return {
      valid: false,
      currentBalance: currentBalance,
      message: `Total balance limit exceeded! Current: ${currentBalance.toFixed(2)} hrs + New: ${hoursToAdd.toFixed(2)} hrs = ${newBalance.toFixed(2)} hrs. Maximum: ${BALANCE_LIMIT} hrs`
    };
  }

  return {
    valid: true,
    currentBalance: currentBalance,
    message: `OK. Total balance: ${newBalance.toFixed(2)} hrs / ${BALANCE_LIMIT} hrs`
  };
}

/**
 * OLD VERSION - DEPRECATED
 * This function has been replaced by the newer version at line 6436+.
 * The newer version uses the updated schema with MONTH_YEAR column and improved logic.
 */
/*
function apiRecordCOCWithValidation_OLD(employeeId, month, year, entries) {
  const db = getDatabase();
  const TIME_ZONE = getScriptTimeZone();

  // Get employee
  const employee = getEmployeeById(employeeId);
  if (!employee) {
    throw new Error('Employee not found');
  }

  // Get settings
  const settings = getSettings();

  // Validate inputs
  if (!month || !year || !entries || entries.length === 0) {
    throw new Error('Invalid parameters: month, year, and entries are required');
  }

  // Calculate total COC that will be added
  let totalNewCOC = 0;
  const calculatedEntries = [];

  for (const entry of entries) {
    // Calculate overtime for each entry
    const date = new Date(year, month - 1, entry.day);
    const result = calculateOvertimeForDate(date, entry.amIn, entry.amOut, entry.pmIn, entry.pmOut);
    // --- REMOVED STRAY 'G' ---
    totalNewCOC += result.cocEarned;
    calculatedEntries.push({
      day: entry.day,
      ...result
    });
  }

  // === START OF NEWLY INSERTED CODE ===
  // === CHECK FOR DUPLICATE DATES ===
  const cocRecordsSheet = db.getSheetByName('COC_Records');
  if (!cocRecordsSheet) throw new Error('COC_Records sheet not found');

  const recordsData = cocRecordsSheet.getDataRange().getValues();
  const existingDates = {};
  const duplicateDetails = [];

  // Build map of existing dates for this employee (exclude cancelled)
  for (let i = 1; i < recordsData.length; i++) {
    const row = recordsData[i];
    const empId = row[1];
    const status = row[15];

    if (empId === employeeId && status !== 'Cancelled') {
      const dateRendered = new Date(row[4]);
      const dateKey = Utilities.formatDate(dateRendered, TIME_ZONE, 'yyyy-MM-dd');

      existingDates[dateKey] = {
        date: dateRendered,
        dayType: row[5],
        hoursWorked: parseFloat(row[10]) || 0,
        cocEarned: parseFloat(row[12]) || 0
      };
    }
  }

  // Check each entry for duplicates
  calculatedEntries.forEach(calc => {
    const dateToCheck = new Date(year, month - 1, calc.day);
    const dateKey = Utilities.formatDate(dateToCheck, TIME_ZONE, 'yyyy-MM-dd');

    if (existingDates[dateKey]) {
      const existing = existingDates[dateKey];
      duplicateDetails.push({
        date: Utilities.formatDate(existing.date, TIME_ZONE, 'MMMM dd, yyyy'),
        dayType: existing.dayType,
        hoursWorked: existing.hoursWorked,
        cocEarned: existing.cocEarned
      });
    }
  });

  // If duplicates found, throw enhanced error message
  if (duplicateDetails.length > 0) {
    let errorMsg = 'The following date(s) already have COC records:\n\n';

    duplicateDetails.forEach(dup => {
      errorMsg += `• <strong>${dup.date}</strong> (${dup.dayType}, `;
      errorMsg += `<span style="color: #dc2626; font-weight: bold;">${dup.hoursWorked.toFixed(1)} hours</span>, `;
      errorMsg += `<span style="color: #16a34a; font-weight: bold;">${dup.cocEarned.toFixed(1)} COC</span>)\n`;
    });

    errorMsg += '\n<strong>To update these records:</strong>\n';
    errorMsg += '<ol style="margin: 8px 0; padding-left: 20px;">';
    errorMsg += '<li>Remove these entries from your current form, OR</li>';
    errorMsg += '<li>Delete the existing records from the database first, then resubmit</li>';
    errorMsg += '</ol>\n';
    errorMsg += '<div style="background-color: #fef2f2; border-left: 4px solid #dc2626; padding: 12px; margin-top: 12px;">';
    errorMsg += '<strong>Note:</strong> You cannot have multiple entries for the same date.';
    errorMsg += '</div>';

    throw new Error(errorMsg);
  }
  // === END OF NEWLY INSERTED CODE ===


  // === VALIDATION: Check COC limits ===
  // Use MM-yyyy format consistently
  const monthYear = String(month).padStart(2, '0') + '-' + String(year);

  // Check monthly limit
  const monthlyCheck = checkMonthlyLimitForMonth(employeeId, monthYear, totalNewCOC);
  if (!monthlyCheck.valid) {
    throw new Error(monthlyCheck.message);
  }

  // Check total balance limit
  const balanceCheck = checkTotalBalanceLimitForEmployee(employeeId, totalNewCOC);
  if (!balanceCheck.valid) {
    throw new Error(balanceCheck.message);
  }

  // === Validations passed, proceed with recording ===

  // const cocRecordsSheet = db.getSheetByName('COC_Records'); // <-- This is now defined above in the new code
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  const detailSheet = ensureCOCBalanceDetailSheet();

  // if (!cocRecordsSheet) throw new Error('COC_Records sheet not found'); // <-- This is now defined above
  if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');

  let runningBalance = getCurrentCOCBalanceFromDetail(employeeId);
  const recordRows = [];
  const ledgerRows = [];
  const detailRows = [];

  // Process each entry
  calculatedEntries.forEach(calc => {
    // Skip entries that didn't earn any COC
    if (calc.cocEarned <= 0) return;

    const recordId = generateRecordId();
    const dateRendered = new Date(year, month - 1, calc.day);
    const monthYearEarned = Utilities.formatDate(dateRendered, TIME_ZONE, 'MM-yyyy');

    runningBalance += calc.cocEarned;

    // COC_Records row
    recordRows.push([
      recordId,
      employeeId,
      employee.fullName,
      monthYearEarned,
      dateRendered,
      calc.dayType,
      calc.amIn || '',
      calc.amOut || '',
      calc.pmIn || '',
      calc.pmOut || '',
      calc.hoursWorked,
      calc.multiplier,
      calc.cocEarned,
      new Date(),
      '',
      'Active'
    ]);

    // COC_Ledger row
    const ledgerId = generateLedgerId();
    ledgerRows.push([
      ledgerId,
      employeeId,
      employee.fullName,
      new Date(),
      'COC Earned',
      recordId,
      calc.cocEarned,
      0,
      runningBalance,
      monthYearEarned,
      '',
      Session.getActiveUser().getEmail(),
      `COC earned on ${Utilities.formatDate(dateRendered, TIME_ZONE, 'MMMM dd, yyyy')} (${calc.dayType}, ${calc.hoursWorked.toFixed(2)} hrs × ${calc.multiplier})`
    ]);

    // COC_Balance_Detail row (for FIFO tracking)
    detailRows.push([
      recordId,
      employeeId,
      employee.fullName,
      monthYearEarned,
      '',
      calc.cocEarned,
      0,
      calc.cocEarned,
      'Active',
      '',
      '',
      new Date(),
      Session.getActiveUser().getEmail(),
      `COC earned from ${calc.dayType} overtime. ${monthlyCheck.message}`
    ]);
  });

  // Batch write to sheets
  if (recordRows.length > 0) {
    cocRecordsSheet.getRange(cocRecordsSheet.getLastRow() + 1, 1, recordRows.length, recordRows[0].length)
      .setValues(recordRows);
  }

  if (ledgerRows.length > 0) {
    ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, ledgerRows.length, ledgerRows[0].length)
      .setValues(ledgerRows);
  }

  if (detailRows.length > 0) {
    detailSheet.getRange(detailSheet.getLastRow() + 1, 1, detailRows.length, detailRows[0].length)
      .setValues(detailRows);
  }

  return {
    success: true,
    added: calculatedEntries.length,
    totalNewCOC: totalNewCOC,
    balanceAfter: runningBalance
  };
}

function apiRecordCOC_OLD_WRAPPER(employeeId, month, year, entries) {
  return apiRecordCOCWithValidation_OLD(employeeId, month, year, entries);
}
*/

/**
 * Get all CTO applications for a specific employee
 * This should be added to your Code.gs file
 * 
 * @param {string} empId - Employee ID (e.g., "EMP008")
 * @returns {Array} Array of CTO application objects
 */
/**
 * FIXED: Get all CTO applications for a specific employee
 * Enhanced with robust date handling for inconsistent formats
 * 
 * @param {string} empId - Employee ID (e.g., "EMP008")
 * @returns {Array} Array of CTO application objects
 */
function apiGetEmployeeCTOApplications(employeeId) {
  try {
    // Get the original result (you already have this function somewhere)
    const db = getDatabase();
    const ctoSheet = db.getSheetByName('CTO_Applications');
    
    if (!ctoSheet || ctoSheet.getLastRow() < 2) {
      return [];
    }
    
    const data = ctoSheet.getDataRange().getValues();
    const applications = [];
    
    // Normalize search ID
    const searchId = String(employeeId).trim();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmpId = String(row[1] || '').trim();
      
      if (rowEmpId === searchId) {
        applications.push({
          appId: String(row[0] || ''),
          employeeId: String(row[1] || ''),
          employeeName: String(row[2] || ''),
          office: String(row[3] || ''),
          hours: parseFloat(row[4]) || 0,
          startDate: String(row[5] || ''),
          endDate: String(row[6] || ''),
          appliedDate: String(row[7] || ''),
          status: String(row[10] || ''),
          remarks: String(row[11] || '')
        });
      }
    }
    
    return applications;
    
  } catch (error) {
    Logger.log('ERROR in apiGetEmployeeCTOApplications: ' + error.message);
    return [];
  }
}


/**
 * Test function - use this to debug CTO applications
 */
function testGetEmployeeCTOApplications() {
  const empId = 'EMP002'; // Juan A Dela Cruz Jr.
  const apps = apiGetEmployeeCTOApplications(empId);
  Logger.log('Applications found: ' + apps.length);
  Logger.log(JSON.stringify(apps, null, 2));
}


/**
 * Test function - use this to debug
 */
function testGetEmployeeCTOApplications() {
  const empId = 'EMP008'; // Patrick E Tan
  const apps = apiGetEmployeeCTOApplications(empId);
  Logger.log('Applications found: ' + apps.length);
  Logger.log(JSON.stringify(apps, null, 2));
}

function debugAllThreeFunctions() {
  const empId = 'EMP002'; // Juan A Dela Cruz Jr.
  
  try {
    Logger.log('=== Testing apiGetLedger ===');
    const ledger = apiGetLedger(empId);
    Logger.log('✓ Ledger: ' + ledger.entries.length + ' entries, balance: ' + ledger.balance);
  } catch (e) {
    Logger.log('✗ apiGetLedger failed: ' + e.message);
  }
  
  try {
    Logger.log('=== Testing apiGetCOCBalanceBreakdown ===');
    const breakdown = apiGetCOCBalanceBreakdown(empId);
    Logger.log('✓ Breakdown: ' + breakdown.length + ' active entries');
  } catch (e) {
    Logger.log('✗ apiGetCOCBalanceBreakdown failed: ' + e.message);
  }
  
  try {
    Logger.log('=== Testing apiGetEmployeeCTOApplications ===');
    const ctos = apiGetEmployeeCTOApplications(empId);
    Logger.log('✓ CTOs: ' + ctos.length + ' applications');
  } catch (e) {
    Logger.log('✗ apiGetEmployeeCTOApplications failed: ' + e.message);
  }
  
  Logger.log('=== Test Complete ===');
}

/**
 * DIAGNOSTIC FUNCTION
 * Run this manually to check the COC_Ledger structure and find EMP002's data
 */
function diagnosticCheckLedgerForEMP002() {
  const empId = 'EMP002';
  const db = getDatabase();
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  
  Logger.log('=== DIAGNOSTIC: COC_Ledger Structure ===');
  Logger.log('Sheet name: ' + ledgerSheet.getName());
  Logger.log('Last row: ' + ledgerSheet.getLastRow());
  Logger.log('Last column: ' + ledgerSheet.getLastColumn());
  
  // Get first few rows to check structure
  const data = ledgerSheet.getDataRange().getValues();
  
  Logger.log('\n--- HEADER ROW (Row 1) ---');
  Logger.log('Headers: ' + JSON.stringify(data[0]));
  
  Logger.log('\n--- FIRST DATA ROW (Row 2) ---');
  if (data.length > 1) {
    Logger.log('Row 2 data: ' + JSON.stringify(data[1]));
    Logger.log('Row 2, Column B (index 1): "' + data[1][1] + '"');
    Logger.log('Row 2, Column B type: ' + typeof data[1][1]);
  }
  
  Logger.log('\n--- SEARCHING FOR EMP002 ---');
  let foundCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const cellValue = row[1]; // Column B (Employee ID)
    
    // Check if this row contains EMP002
    if (cellValue === empId) {
      foundCount++;
      Logger.log('✓ Found exact match at row ' + (i + 1));
      Logger.log('  Full row: ' + JSON.stringify(row));
      
      if (foundCount <= 3) { // Show first 3 matches in detail
        Logger.log('  Column A (Ledger ID): "' + row[0] + '"');
        Logger.log('  Column B (Employee ID): "' + row[1] + '"');
        Logger.log('  Column C (Employee Name): "' + row[2] + '"');
        Logger.log('  Column D (Transaction Date): "' + row[3] + '"');
        Logger.log('  Column E (Transaction Type): "' + row[4] + '"');
      }
    } else if (String(cellValue).indexOf(empId) !== -1) {
      Logger.log('⚠ Found partial match at row ' + (i + 1) + ': "' + cellValue + '"');
    }
  }
  
  Logger.log('\n--- SUMMARY ---');
  Logger.log('Total exact matches for "' + empId + '": ' + foundCount);
  
  if (foundCount === 0) {
    Logger.log('\n⚠ WARNING: No exact matches found!');
    Logger.log('Checking for similar values in column B...');
    
    const uniqueIds = new Set();
    for (let i = 1; i < Math.min(data.length, 20); i++) {
      uniqueIds.add('"' + data[i][1] + '"');
    }
    Logger.log('Sample Employee IDs from first 20 rows:');
    Array.from(uniqueIds).forEach(id => Logger.log('  ' + id));
  }
  
  Logger.log('\n=== END DIAGNOSTIC ===');
}

/**
 * DIAGNOSTIC: Check what getCurrentCOCBalance returns for EMP002
 */
function diagnosticCheckBalanceForEMP002() {
  const empId = 'EMP002';
  
  Logger.log('=== DIAGNOSTIC: getCurrentCOCBalance for ' + empId + ' ===');
  
  try {
    const balance = getCurrentCOCBalance(empId);
    Logger.log('✓ Balance returned: ' + balance);
  } catch (e) {
    Logger.log('✗ Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
  
  Logger.log('=== END DIAGNOSTIC ===');
}

/**
 * DIAGNOSTIC: Check what apiGetLedger returns for EMP002
 */
function diagnosticCheckApiGetLedger() {
  const empId = 'EMP002';
  
  Logger.log('=== DIAGNOSTIC: apiGetLedger for ' + empId + ' ===');
  
  try {
    const result = apiGetLedger(empId);
    Logger.log('✓ Result returned:');
    Logger.log('  Balance: ' + result.balance);
    Logger.log('  Entries count: ' + result.entries.length);
    
    if (result.entries.length > 0) {
      Logger.log('  First entry: ' + JSON.stringify(result.entries[0]));
    } else {
      Logger.log('  ⚠ No entries found!');
    }
    
    if (result.error) {
      Logger.log('  ⚠ Error property: ' + result.error);
    }
  } catch (e) {
    Logger.log('✗ Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
  }
  
  Logger.log('=== END DIAGNOSTIC ===');
}

/**
 * RUN ALL DIAGNOSTICS
 */
function runAllDiagnostics() {
  diagnosticCheckLedgerForEMP002();
  Logger.log('\n\n');
  diagnosticCheckBalanceForEMP002();
  Logger.log('\n\n');
  diagnosticCheckApiGetLedger();
}

function testAllSerializations() {
  const empId = 'EMP002';
  
  Logger.log('========================================');
  Logger.log('Testing all API serializations for ' + empId);
  Logger.log('========================================\n');
  
  // Test 1: Ledger
  Logger.log('TEST 1: apiGetLedger');
  try {
    const ledger = apiGetLedger(empId);
    const json1 = JSON.stringify(ledger);
    Logger.log('✅ Ledger serialization OK (' + json1.length + ' chars)');
    Logger.log('   Balance: ' + ledger.balance);
    Logger.log('   Entries: ' + ledger.entries.length);
  } catch (e) {
    Logger.log('❌ Ledger serialization FAILED: ' + e.message);
  }
  
  // Test 2: Breakdown
  Logger.log('\nTEST 2: apiGetCOCBalanceBreakdown');
  try {
    const breakdown = apiGetCOCBalanceBreakdown(empId);
    const json2 = JSON.stringify(breakdown);
    Logger.log('✅ Breakdown serialization OK (' + json2.length + ' chars)');
    Logger.log('   Entries: ' + breakdown.length);
  } catch (e) {
    Logger.log('❌ Breakdown serialization FAILED: ' + e.message);
  }
  
  // Test 3: CTOs
  Logger.log('\nTEST 3: apiGetEmployeeCTOApplications');
  try {
    const ctos = apiGetEmployeeCTOApplications(empId);
    const json3 = JSON.stringify(ctos);
    Logger.log('✅ CTOs serialization OK (' + json3.length + ' chars)');
    Logger.log('   Entries: ' + ctos.length);
  } catch (e) {
    Logger.log('❌ CTOs serialization FAILED: ' + e.message);
  }
  
  Logger.log('\n========================================');
  Logger.log('All tests complete!');
  Logger.log('========================================');
}

/**
 * BACKEND FUNCTIONS FOR CTO APPLICATIONS MANAGER
 * 
 * Add these functions to your Code.gs file
 */

// ============================================================================
// API: Get All CTO Applications
// ============================================================================

/**
 * Get all CTO applications from the database with proper serialization
 */
function apiGetAllCTOApplications() {
  try {
    const db = getDatabase();
    const ctoSheet = db.getSheetByName('CTO_Applications');
    
    if (!ctoSheet || ctoSheet.getLastRow() < 2) {
      Logger.log('No CTO applications found');
      return [];
    }
    
    const data = ctoSheet.getDataRange().getValues();
    const applications = [];
    const TIME_ZONE = getScriptTimeZone();
    
    // Helper to format dates safely
    function safeFormatDate(value) {
      if (!value) return '';
      try {
        if (value instanceof Date) {
          return Utilities.formatDate(value, TIME_ZONE, 'yyyy-MM-dd');
        }
        return String(value);
      } catch (e) {
        return String(value);
      }
    }
    
    // Process each row (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row[0]) continue;
      
      try {
        // Ensure all data is serializable
        const app = {
          appId: String(row[0] || ''),
          employeeId: String(row[1] || ''),
          employeeName: String(row[2] || ''),
          office: String(row[3] || ''),
          hours: parseFloat(row[4]) || 0,
          startDate: safeFormatDate(row[5]),
          endDate: safeFormatDate(row[6]),
          appliedDate: safeFormatDate(row[7]),
          approvedBy: String(row[8] || ''),
          approvedDate: safeFormatDate(row[9]),
          status: String(row[10] || 'Pending'),
          remarks: String(row[11] || '')
        };
        
        applications.push(app);
        
      } catch (rowError) {
        Logger.log('Error processing CTO application row ' + (i + 1) + ': ' + rowError.message);
        continue;
      }
    }
    
    Logger.log('Loaded ' + applications.length + ' CTO applications');
    return applications;
    
  } catch (error) {
    Logger.log('ERROR in apiGetAllCTOApplications: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    return [];
  }
}

// ============================================================================
// API: Cancel CTO Application
// ============================================================================

/**
 * Cancel a CTO application and restore the COC hours using FIFO
 */
function apiCancelCTOApplication(applicationId) {
  try {
    Logger.log('=== Cancelling CTO Application: ' + applicationId + ' ===');
    
    const db = getDatabase();
    const ctoSheet = db.getSheetByName('CTO_Applications');
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    
    if (!ctoSheet) throw new Error('CTO_Applications sheet not found');
    if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');
    if (!detailSheet) throw new Error('COC_Balance_Detail sheet not found');
    
    // Find the application
    const ctoData = ctoSheet.getDataRange().getValues();
    let appRow = -1;
    let application = null;
    
    for (let i = 1; i < ctoData.length; i++) {
      if (ctoData[i][0] === applicationId) {
        appRow = i + 1; // Sheet row (1-indexed)
        application = {
          id: ctoData[i][0],
          employeeId: ctoData[i][1],
          employeeName: ctoData[i][2],
          hours: parseFloat(ctoData[i][4]) || 0,
          status: ctoData[i][10]
        };
        break;
      }
    }
    
    if (!application) {
      throw new Error('CTO application not found: ' + applicationId);
    }
    
    // Check if already cancelled
    if (application.status === 'Cancelled') {
      return {
        success: false,
        message: 'This application is already cancelled'
      };
    }
    
    // --- START MODIFICATION ---
    // Allow cancelling 'Pending' or 'Approved' applications
    if (application.status !== 'Pending' && application.status !== 'Approved') {
      return {
        success: false,
        // Updated error message
        message: 'Only Pending or Approved applications can be cancelled. This application is ' + application.status
      };
    }
    // --- END MODIFICATION ---
    
    Logger.log('Application found: ' + application.employeeName + ', Status: ' + application.status + ', Hours: ' + application.hours);
    
    // Update CTO application status
    ctoSheet.getRange(appRow, 11).setValue('Cancelled'); // Column K (Status)
    ctoSheet.getRange(appRow, 12).setValue('Cancelled by user on ' + new Date()); // Column L (Remarks)
    
    Logger.log('✓ CTO application status updated to Cancelled');
    
    // --- START MODIFICATION: Restore hours ONLY if it was Approved ---
    // If it was Pending, no hours were deducted yet, so no need to restore.
    let restoredHours = 0; // Initialize restoredHours
    if (application.status === 'Approved') {
        Logger.log('Restoring hours because status was Approved.');
        // Find and reverse the ledger entry
        const ledgerData = ledgerSheet.getDataRange().getValues();
        let ledgerRow = -1;

        for (let i = ledgerData.length - 1; i >= 1; i--) {
            if (ledgerData[i][1] === application.employeeId &&
                ledgerData[i][5] === applicationId &&
                ledgerData[i][4] === 'CTO Used') {
                ledgerRow = i + 1;
                break;
            }
        }

        if (ledgerRow > 0) {
            // Mark ledger entry as cancelled
            ledgerSheet.getRange(ledgerRow, 5).setValue('CTO Cancelled (Approved)'); // Clarify transaction type
            ledgerSheet.getRange(ledgerRow, 13).setValue('Cancelled on ' + new Date()); // Remarks
            Logger.log('✓ Original CTO Used ledger entry marked as cancelled');
        } else {
             Logger.log('⚠ Original CTO Used ledger entry not found, proceeding with balance restoration.');
        }

        // Restore the COC hours in detail sheet (reverse FIFO deduction)
        try {
          // --- Call the separate restore function ---
          restoredHours = restoreCOCHoursFIFO(application.employeeId, application.hours, applicationId);
          if (restoredHours > 0) {
             Logger.log('✓ Restored ' + restoredHours + ' hours to COC balance via FIFO.');
          } else {
             Logger.log('⚠ No hours were restored via FIFO. This might be okay if the deduction logic changed or was manual.');
          }
        } catch (restoreError) {
           Logger.log('ERROR during FIFO restore: ' + restoreError.message);
           // Even if restore fails, continue with ledger update to reflect cancellation attempt
        }

        // Add a reversal entry to the ledger ONLY if hours were deducted (status was Approved)
        const newBalance = getCurrentCOCBalance(application.employeeId); // Recalculate balance AFTER potential restore
        const ledgerEntry = [
          generateLedgerEntryId(),
          application.employeeId,
          application.employeeName,
          new Date(),
          'CTO Cancelled (Restore)', // Clearer transaction type
          applicationId,
          application.hours, // COC Earned (restored)
          0, // CTO Used
          newBalance,
          '', // Month-Year Earned
          '', // Expiration Date
          Session.getActiveUser().getEmail(),
          'CTO application (Approved) cancelled, ' + application.hours + ' hrs restored'
        ];
        ledgerSheet.appendRow(ledgerEntry);
        Logger.log('✓ Added restoration entry to ledger');

    } else {
       // If status was 'Pending', just log that no restoration needed
       Logger.log('Status was Pending, no hours to restore.');
       // Optionally add a simpler ledger entry just marking cancellation
        const newBalance = getCurrentCOCBalance(application.employeeId); // Get current balance
        const ledgerEntry = [
          generateLedgerEntryId(),
          application.employeeId,
          application.employeeName,
          new Date(),
          'CTO Cancelled (Pending)', // Clearer transaction type
          applicationId,
          0, // COC Earned
          0, // CTO Used
          newBalance, // Balance remains unchanged
          '', // Month-Year Earned
          '', // Expiration Date
          Session.getActiveUser().getEmail(),
          'CTO application (Pending) cancelled before approval.'
        ];
        ledgerSheet.appendRow(ledgerEntry);
        Logger.log('✓ Added ledger entry for Pending cancellation.');
    }
    // --- END MODIFICATION ---
    
    Logger.log('=== CTO Cancellation Complete ===');
    
    return {
      success: true,
      // --- MODIFICATION: Updated success message based on status ---
      message: application.status === 'Approved' 
         ? 'CTO application cancelled successfully. ' + application.hours + ' hours have been restored to the employee\'s COC balance.'
         : 'Pending CTO application cancelled successfully.'
      // --- END MODIFICATION ---
    };
    
  } catch (error) {
    Logger.log('ERROR in apiCancelCTOApplication: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    
    return {
      success: false,
      message: 'Failed to cancel CTO application: ' + error.message
    };
  }
}

// ============================================================================
// HELPER: Generate Ledger Entry ID
// ============================================================================

/**
 * Generates a unique ledger entry ID
 * Format: LED-YYYYMMDDHHMMSSmmm
 */
function generateLedgerEntryId() {
  const now = new Date();
  const TIME_ZONE = getScriptTimeZone();
  const timestamp = Utilities.formatDate(now, TIME_ZONE, 'yyyyMMddHHmmss');
  const millis = String(now.getMilliseconds()).padStart(3, '0');
  return 'LED-' + timestamp + millis;
}



// ============================================================================
// TEST FUNCTIONS
// ============================================================================

/**
 * Test getting all CTO applications
 */
function testGetAllCTOApplications() {
  Logger.log('=== Testing apiGetAllCTOApplications ===');
  
  const apps = apiGetAllCTOApplications();
  
  Logger.log('Total applications: ' + apps.length);
  
  if (apps.length > 0) {
    Logger.log('\nFirst application:');
    Logger.log(JSON.stringify(apps[0], null, 2));
    
    // Count by status
    const stats = {};
    apps.forEach(app => {
      stats[app.status] = (stats[app.status] || 0) + 1;
    });
    
    Logger.log('\nStatistics:');
    Object.keys(stats).forEach(status => {
      Logger.log('  ' + status + ': ' + stats[status]);
    });
  }
  
  Logger.log('\n=== Test Complete ===');
}

/**
 * Test cancelling a CTO application
 */
function testCancelCTOApplication() {
  // CHANGE THIS to a real application ID from your sheet
  const testAppId = 'CTO-20251025001613300'; 
  
  Logger.log('=== Testing apiCancelCTOApplication ===');
  Logger.log('Application ID: ' + testAppId);
  
  const result = apiCancelCTOApplication(testAppId);
  
  Logger.log('\nResult:');
  Logger.log('Success: ' + result.success);
  Logger.log('Message: ' + result.message);
  
  Logger.log('\n=== Test Complete ===');
}

/**
 * ADDITIONAL BACKEND FUNCTION FOR CTO UPDATE
 * 
 * Add this to your Code.gs file (along with the other CTO functions)
 */

// ============================================================================
// API: Update CTO Application
// ============================================================================


function apiUpdateCTOApplication(applicationId, newHours, newStartDate, newEndDate, newRemarks) {
  try {
    Logger.log('=== Updating CTO Application: ' + applicationId + ' ===');
    Logger.log('New hours: ' + newHours);
    Logger.log('New dates: ' + newStartDate + ' to ' + newEndDate);

    const db = getDatabase();
    const ctoSheet = db.getSheetByName('CTO_Applications');
    const ledgerSheet = db.getSheetByName('COC_Ledger'); // Ensure ledger sheet is available

    if (!ctoSheet) throw new Error('CTO_Applications sheet not found');
    if (!ledgerSheet) throw new Error('COC_Ledger sheet not found'); // Added check

    // Find the application
    const ctoData = ctoSheet.getDataRange().getValues();
    let appRow = -1;
    let application = null;

    for (let i = 1; i < ctoData.length; i++) {
      if (ctoData[i][0] === applicationId) {
        appRow = i + 1; // Sheet row (1-indexed)
        application = {
          id: ctoData[i][0],
          employeeId: ctoData[i][1],
          employeeName: ctoData[i][2],
          oldHours: parseFloat(ctoData[i][4]) || 0,
          oldStartDate: ctoData[i][5], // Keep original format for comparison/logging
          oldEndDate: ctoData[i][6],   // Keep original format for comparison/logging
          oldRemarks: ctoData[i][11],
          status: ctoData[i][10]
        };
        break;
      }
    }

    if (!application) {
      throw new Error('CTO application not found: ' + applicationId);
    }

    // Check if already cancelled
    if (application.status === 'Cancelled') {
      return {
        success: false,
        message: 'Cannot update a cancelled application'
      };
    }

    // Allow updating Approved or Pending
     if (application.status !== 'Pending' && application.status !== 'Approved') {
        return {
            success: false,
            message: 'Only Pending or Approved applications can be updated. This application is ' + application.status
        };
    }


    Logger.log('Found application: ' + application.employeeName);
    Logger.log('Old hours: ' + application.oldHours + ', New hours: ' + newHours);

    // --- START VALIDATION ---
    const validationError = validateCTOUpdate(application, newHours, newStartDate, newEndDate);
     if (validationError) {
        return { success: false, message: validationError };
    }
    // --- END VALIDATION ---


    // --- Apply Updates ---
    const TIME_ZONE = getScriptTimeZone();
    const updateTimestamp = new Date(); // Use consistent timestamp for updates

    // Update CTO_Applications sheet
    ctoSheet.getRange(appRow, 5).setValue(newHours); // Column E (Hours)
    ctoSheet.getRange(appRow, 6).setValue(new Date(newStartDate)); // Column F (Start Date) - Ensure it's a Date object
    ctoSheet.getRange(appRow, 7).setValue(new Date(newEndDate));   // Column G (End Date) - Ensure it's a Date object
    // Append update info to remarks
    const updateRemark = `\n[Updated on ${Utilities.formatDate(updateTimestamp, TIME_ZONE, 'yyyy-MM-dd HH:mm')} by ${Session.getActiveUser().getEmail()}] ${newRemarks || '(No additional remarks)'}`;
    ctoSheet.getRange(appRow, 12).setValue((application.oldRemarks || '') + updateRemark); // Column L (Remarks)

    Logger.log('✓ CTO application updated in database');

    // --- Ledger and FIFO Adjustments (Only if status is Approved and hours changed) ---
     let hoursDifference = 0;
     if (application.status === 'Approved') {
        hoursDifference = newHours - application.oldHours;
        Logger.log(`Status is Approved. Hours difference: ${hoursDifference}`);

        if (hoursDifference !== 0) {
             // If hours increased, deduct using FIFO
            if (hoursDifference > 0) {
                try {
                    deductCOCHoursFIFO(application.employeeId, hoursDifference, applicationId + '-UPDATE'); // Use a distinct reference for update adjustment
                    Logger.log('✓ Deducted additional ' + hoursDifference + ' hours using FIFO for update.');
                } catch (deductError) {
                     Logger.log('ERROR during FIFO deduction for update: ' + deductError.message);
                     // Rollback changes? Or just return error? For now, return error.
                     // It might be better to revert the CTO sheet change here.
                    return { success: false, message: 'Failed to apply update: ' + deductError.message };
                }
            }
            // If hours decreased, restore using FIFO
            else if (hoursDifference < 0) {
                try {
                    restoreCOCHoursFIFO(application.employeeId, Math.abs(hoursDifference), applicationId); // Restore based on original ID
                    Logger.log('✓ Restored ' + Math.abs(hoursDifference) + ' hours to FIFO due to update.');
                } catch (restoreError) {
                     Logger.log('ERROR during FIFO restoration for update: ' + restoreError.message);
                     // Rollback changes? Or just return error? For now, return error.
                    return { success: false, message: 'Failed to apply update: ' + restoreError.message };
                }
            }
        }
    } else {
         Logger.log('Status is Pending. No ledger/FIFO adjustments needed yet.');
    }


    // --- Add a specific Ledger Entry for the UPDATE action itself ---
    const currentBalance = getCurrentCOCBalance(application.employeeId); // Get balance AFTER potential adjustments
    const updateLedgerRemark = `CTO Updated: ${applicationId}. ` +
                             (hoursDifference !== 0 ? `Hours changed ${application.oldHours} -> ${newHours} (${hoursDifference > 0 ? '+' : ''}${hoursDifference.toFixed(2)}). ` : '') +
                             `Dates changed. ${newRemarks || ''}`;

    const updateLedgerEntry = [
      generateLedgerEntryId(),
      application.employeeId,
      application.employeeName,
      updateTimestamp, // Use the consistent timestamp
      'CTO Updated', // New Transaction Type
      applicationId, // Reference the original CTO ID
      (hoursDifference < 0) ? Math.abs(hoursDifference) : 0, // Restore is like earning COC back
      (hoursDifference > 0) ? hoursDifference : 0, // Increase is like using more CTO
      currentBalance,
      '', // Month-Year Earned (N/A for update)
      '', // Expiration Date (N/A for update)
      Session.getActiveUser().getEmail(),
      updateLedgerRemark.trim() // Trim potential extra space at the end
    ];
    ledgerSheet.appendRow(updateLedgerEntry);
    Logger.log('✓ Added "CTO Updated" entry to ledger');
    // --- End Ledger Update ---


    Logger.log('=== CTO Update Complete ===');

    return {
      success: true,
      message: 'CTO application updated successfully'
    };

  } catch (error) {
    Logger.log('ERROR in apiUpdateCTOApplication: ' + error.message);
    Logger.log('Stack: ' + error.stack);

    return {
      success: false,
      message: 'Failed to update CTO application: ' + error.message
    };
  }
}

// ============================================================================
// HELPER: Validate CTO Update Data (Separated Logic)
// ============================================================================
/**
 * Validates the data for updating a CTO application.
 * Returns null if valid, or an error message string if invalid.
 */
function validateCTOUpdate(application, newHours, newStartDateStr, newEndDateStr) {
     // 1. Check for valid hours (positive number, multiple of 4)
    if (isNaN(newHours) || newHours <= 0 || newHours % 4 !== 0) {
        return 'CTO hours must be a positive multiple of 4 (e.g., 4, 8, 12).';
    }

    // 2. Check for dates
    if (!newStartDateStr || !newEndDateStr) {
        return 'Start and End dates are required.';
    }

    const startDate = new Date(newStartDateStr + 'T00:00:00');
    const endDate = new Date(newEndDateStr + 'T00:00:00');

    // Check if dates are valid
     if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        return 'Invalid date format provided.';
    }

    // 3. Check chronological order
    if (endDate < startDate) {
        return 'End date must not be before the start date.';
    }

     // 4. Check 5-day limit (inclusive)
    const diffDays = Math.floor((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)) + 1;
    if (diffDays > 5) {
        return 'Date range must not exceed 5 days.';
    }

    // 5. Check 4/8 hour rule
    if ((newHours === 4 || newHours === 8) && startDate.getTime() !== endDate.getTime()) {
        return 'For 4 or 8-hour CTO applications, the start and end dates must be the same.';
    }

     // 6. Check balance IF hours are increasing AND status is Approved
     if (application.status === 'Approved' && newHours > application.oldHours) {
        const currentBalance = getCurrentCOCBalance(application.employeeId);
        const additionalHoursNeeded = newHours - application.oldHours;
        if (currentBalance < additionalHoursNeeded) {
            return `Insufficient COC balance. Employee has ${currentBalance.toFixed(2)} hrs available, but needs ${additionalHoursNeeded.toFixed(2)} additional hours for this update.`;
        }
    }

    // All checks passed
    return null;
}


/**
 * Cancel a CTO application and restore the COC hours using FIFO
 */
function apiCancelCTOApplication(applicationId) {
  try {
    Logger.log('=== Cancelling CTO Application: ' + applicationId + ' ===');
    
    const db = getDatabase();
    const ctoSheet = db.getSheetByName('CTO_Applications');
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    
    if (!ctoSheet) throw new Error('CTO_Applications sheet not found');
    if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');
    if (!detailSheet) throw new Error('COC_Balance_Detail sheet not found');
    
    // Find the application
    const ctoData = ctoSheet.getDataRange().getValues();
    let appRow = -1;
    let application = null;
    
    for (let i = 1; i < ctoData.length; i++) {
      if (ctoData[i][0] === applicationId) {
        appRow = i + 1; // Sheet row (1-indexed)
        application = {
          id: ctoData[i][0],
          employeeId: ctoData[i][1],
          employeeName: ctoData[i][2],
          hours: parseFloat(ctoData[i][4]) || 0,
          status: ctoData[i][10]
        };
        break;
      }
    }
    
    if (!application) {
      throw new Error('CTO application not found: ' + applicationId);
    }
    
    // Check if already cancelled
    if (application.status === 'Cancelled') {
      return {
        success: false,
        message: 'This application is already cancelled'
      };
    }
    
    // --- START MODIFICATION ---
    // Allow cancelling 'Pending' or 'Approved' applications
    if (application.status !== 'Pending' && application.status !== 'Approved') {
      return {
        success: false,
        // Updated error message
        message: 'Only Pending or Approved applications can be cancelled. This application is ' + application.status
      };
    }
    // --- END MODIFICATION ---
    
    Logger.log('Application found: ' + application.employeeName + ', Status: ' + application.status + ', Hours: ' + application.hours);
    
    // Update CTO application status
    ctoSheet.getRange(appRow, 11).setValue('Cancelled'); // Column K (Status)
    ctoSheet.getRange(appRow, 12).setValue('Cancelled by user on ' + new Date()); // Column L (Remarks)
    
    Logger.log('✓ CTO application status updated to Cancelled');
    
    // --- START MODIFICATION: Restore hours ONLY if it was Approved ---
    // If it was Pending, no hours were deducted yet, so no need to restore.
    let restoredHours = 0; // Initialize restoredHours
    if (application.status === 'Approved') {
        Logger.log('Restoring hours because status was Approved.');
        // Find and reverse the ledger entry
        const ledgerData = ledgerSheet.getDataRange().getValues();
        let ledgerRow = -1;

        for (let i = ledgerData.length - 1; i >= 1; i--) {
            if (ledgerData[i][1] === application.employeeId &&
                ledgerData[i][5] === applicationId &&
                ledgerData[i][4] === 'CTO Used') {
                ledgerRow = i + 1;
                break;
            }
        }

        if (ledgerRow > 0) {
            // Mark ledger entry as cancelled
            ledgerSheet.getRange(ledgerRow, 5).setValue('CTO Cancelled (Approved)'); // Clarify transaction type
            ledgerSheet.getRange(ledgerRow, 13).setValue('Cancelled on ' + new Date()); // Remarks
            Logger.log('✓ Original CTO Used ledger entry marked as cancelled');
        } else {
             Logger.log('⚠ Original CTO Used ledger entry not found, proceeding with balance restoration.');
        }

        // Restore the COC hours in detail sheet (reverse FIFO deduction)
        try {
          // --- Call the separate restore function ---
          restoredHours = restoreCOCHoursFIFO(application.employeeId, application.hours, applicationId);
          if (restoredHours > 0) {
             Logger.log('✓ Restored ' + restoredHours + ' hours to COC balance via FIFO.');
          } else {
             Logger.log('⚠ No hours were restored via FIFO. This might be okay if the deduction logic changed or was manual.');
          }
        } catch (restoreError) {
           Logger.log('ERROR during FIFO restore: ' + restoreError.message);
           // Even if restore fails, continue with ledger update to reflect cancellation attempt
        }

        // Add a reversal entry to the ledger ONLY if hours were deducted (status was Approved)
        const newBalance = getCurrentCOCBalance(application.employeeId); // Recalculate balance AFTER potential restore
        const ledgerEntry = [
          generateLedgerEntryId(),
          application.employeeId,
          application.employeeName,
          new Date(),
          'CTO Cancelled (Restore)', // Clearer transaction type
          applicationId,
          application.hours, // COC Earned (restored)
          0, // CTO Used
          newBalance,
          '', // Month-Year Earned
          '', // Expiration Date
          Session.getActiveUser().getEmail(),
          'CTO application (Approved) cancelled, ' + application.hours + ' hrs restored'
        ];
        ledgerSheet.appendRow(ledgerEntry);
        Logger.log('✓ Added restoration entry to ledger');

    } else {
       // If status was 'Pending', just log that no restoration needed
       Logger.log('Status was Pending, no hours to restore.');
       // Optionally add a simpler ledger entry just marking cancellation
        const newBalance = getCurrentCOCBalance(application.employeeId); // Get current balance
        const ledgerEntry = [
          generateLedgerEntryId(),
          application.employeeId,
          application.employeeName,
          new Date(),
          'CTO Cancelled (Pending)', // Clearer transaction type
          applicationId,
          0, // COC Earned
          0, // CTO Used
          newBalance, // Balance remains unchanged
          '', // Month-Year Earned
          '', // Expiration Date
          Session.getActiveUser().getEmail(),
          'CTO application (Pending) cancelled before approval.'
        ];
        ledgerSheet.appendRow(ledgerEntry);
        Logger.log('✓ Added ledger entry for Pending cancellation.');
    }
    // --- END MODIFICATION ---
    
    Logger.log('=== CTO Cancellation Complete ===');
    
    return {
      success: true,
      // --- MODIFICATION: Updated success message based on status ---
      message: application.status === 'Approved' 
         ? 'CTO application cancelled successfully. ' + application.hours + ' hours have been restored to the employee\'s COC balance.'
         : 'Pending CTO application cancelled successfully.'
      // --- END MODIFICATION ---
    };
    
  } catch (error) {
    Logger.log('ERROR in apiCancelCTOApplication: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    
    return {
      success: false,
      message: 'Failed to cancel CTO application: ' + error.message
    };
  }
}

// ============================================================================
// HELPER: Deduct COC Hours using FIFO
// ============================================================================

/**
 * Deduct COC hours from employee balance using FIFO method
 * Used when increasing CTO hours
 */
function deductCOCHoursFIFO(employeeId, hoursToDeduct, referenceId) {
  const db = getDatabase();
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  
  if (!detailSheet) throw new Error('COC_Balance_Detail sheet not found');
  
  const data = detailSheet.getDataRange().getValues();
  let remainingToDeduct = hoursToDeduct;
  
  // Get active entries for this employee, sorted by date (FIFO)
  const activeEntries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === employeeId && row[8] === 'Active' && row[6] > 0) {
      activeEntries.push({
        rowIndex: i + 1,
        entryId: row[0],
        recordId: row[3],
        dateEarned: new Date(row[4]),
        hoursRemaining: parseFloat(row[6]) || 0
      });
    }
  }
  
  // Sort by date earned (oldest first - FIFO)
  activeEntries.sort((a, b) => a.dateEarned - b.dateEarned);
  
  // Deduct from entries
  activeEntries.forEach(entry => {
    if (remainingToDeduct <= 0) return;
    
    const hoursFromThisEntry = Math.min(entry.hoursRemaining, remainingToDeduct);
    const newRemaining = entry.hoursRemaining - hoursFromThisEntry;
    
    // Update detail sheet
    detailSheet.getRange(entry.rowIndex, 7).setValue(newRemaining); // Hours Remaining
    
    if (newRemaining <= 0) {
      detailSheet.getRange(entry.rowIndex, 9).setValue('Used'); // Status
    }
    
    // Add ledger entry for tracking
    const ledgerEntry = [
      generateLedgerEntryId(),
      employeeId,
      '', // Employee name (will be filled by caller)
      new Date(),
      'FIFO Adjustment',
      referenceId,
      0,
      hoursFromThisEntry,
      getCurrentCOCBalance(employeeId),
      '',
      '',
      Session.getActiveUser().getEmail(),
      `FIFO: Used ${hoursFromThisEntry.toFixed(2)} hrs from ${entry.recordId} for ${referenceId}`
    ];
    
    remainingToDeduct -= hoursFromThisEntry;
  });
  
  if (remainingToDeduct > 0) {
    throw new Error('Insufficient COC balance for deduction');
  }
}

// ============================================================================
// HELPER: Restore COC Hours using FIFO (in reverse)
// ============================================================================

/**
 * Restore COC hours to employee balance
 * Used when decreasing CTO hours
 */
function restoreCOCHoursFIFO(employeeId, hoursToRestore, referenceId) {
  const db = getDatabase();
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  const ledgerSheet = db.getSheetByName('COC_Ledger'); // Needed to find original deductions
  
  if (!detailSheet) throw new Error('COC_Balance_Detail sheet not found');
  if (!ledgerSheet) throw new Error('COC_Ledger sheet not found'); // Added check

  Logger.log(`Attempting to restore ${hoursToRestore} hours for ${employeeId} related to ${referenceId}`);
  
  // Find the ledger entries representing the *original* FIFO deductions for this referenceId
  const ledgerData = ledgerSheet.getDataRange().getValues();
  const originalDeductions = [];
  
  // Search from newest to oldest for CTO Used or FIFO Adjustment entries
  for (let i = ledgerData.length - 1; i >= 1; i--) {
    const row = ledgerData[i];
    // Check Employee ID, Reference ID, and Transaction Type or Remarks
    if (row[1] === employeeId && row[5] === referenceId) {
       // Look for "CTO Used" or "FIFO Adjustment" or remarks containing "FIFO: Used"
       const type = row[4];
       const remarks = row[12] || '';
       if (type === 'CTO Used' || type === 'FIFO Adjustment' || remarks.startsWith('FIFO: Used')) {
          // Extract hours and source record ID from remarks if possible
          const remarksMatch = remarks.match(/Used\s+([\d.]+)\s+hrs\s+from\s+(\S+)/);
          if (remarksMatch) {
            originalDeductions.push({
              hours: parseFloat(remarksMatch[1]),
              recordId: remarksMatch[2] // This is the Record ID from COC_Balance_Detail
            });
             Logger.log(`Found original deduction: ${remarksMatch[1]} hrs from ${remarksMatch[2]}`);
          } else if (type === 'CTO Used' && (parseFloat(row[7]) || 0) > 0) {
            // Fallback for simple CTO Used entries without detailed FIFO remarks (less accurate)
            // We need to guess which detail record it came from - this part is tricky without proper remarks
            // For now, let's log a warning if we can't find specific FIFO details
            Logger.log(`Warning: Found CTO Used entry for ${referenceId} but couldn't parse specific FIFO details from remarks: "${remarks}"`);
            // We could try to add the total CTO Used hours here, but reversing it accurately is hard.
            // It's better if the original deduction logged the source record ID in remarks.
          }
       }
    }
  }

  // --- MODIFICATION: Reverse the deductions array so we restore in the reverse order they were taken ---
  originalDeductions.reverse();
  // --- END MODIFICATION ---

  let totalRestored = 0;
  let remainingToRestore = hoursToRestore;
  const detailData = detailSheet.getDataRange().getValues();
  const updatesToDetailSheet = []; // Batch updates

  // Iterate through the identified original deductions
  for (const deduction of originalDeductions) {
    if (remainingToRestore <= 0) break;

    const hoursToRestoreHere = Math.min(deduction.hours, remainingToRestore);
    Logger.log(`Processing deduction: Need to restore ${hoursToRestoreHere} hrs originally from ${deduction.recordId}`);

    // Find the corresponding detail entry using the Record ID
    let detailRowIndex = -1;
    for (let i = 1; i < detailData.length; i++) {
        // --- MODIFICATION: Match using Record ID (column D, index 3) ---
      if (detailData[i][3] === deduction.recordId && detailData[i][1] === employeeId) {
        detailRowIndex = i + 1; // 1-based index for getRange
         Logger.log(`Found matching detail entry at row ${detailRowIndex}`);
        break;
      }
       // --- END MODIFICATION ---
    }

    if (detailRowIndex !== -1) {
      const currentRemaining = parseFloat(detailSheet.getRange(detailRowIndex, 7).getValue()) || 0; // Column G
      const newRemaining = currentRemaining + hoursToRestoreHere;
      const currentStatus = detailSheet.getRange(detailRowIndex, 9).getValue(); // Column I

      updatesToDetailSheet.push({ row: detailRowIndex, col: 7, value: newRemaining }); // Update Hours Remaining
       Logger.log(`Updating row ${detailRowIndex}, col 7 (Hours Remaining) from ${currentRemaining} to ${newRemaining}`);

      // If the entry was Depleted or Used, mark it back to Active
      if (currentStatus === 'Depleted' || currentStatus === 'Used') {
        updatesToDetailSheet.push({ row: detailRowIndex, col: 9, value: 'Active' }); // Update Status
         Logger.log(`Updating row ${detailRowIndex}, col 9 (Status) to Active`);
      }
      
      // Add a note about the restoration
       const currentNote = detailSheet.getRange(detailRowIndex, 11).getValue(); // Column K
       const TIME_ZONE = getScriptTimeZone();
       const newNote = currentNote + '\n[' + Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm') +
                      '] Restored ' + hoursToRestoreHere.toFixed(2) + ' hrs due to cancellation of ' + referenceId;
       updatesToDetailSheet.push({ row: detailRowIndex, col: 11, value: newNote});
       Logger.log(`Updating row ${detailRowIndex}, col 11 (Notes)`);


      totalRestored += hoursToRestoreHere;
      remainingToRestore -= hoursToRestoreHere;
    } else {
        Logger.log(`Warning: Could not find detail entry for Record ID ${deduction.recordId} to restore hours.`);
    }
  }

  // Apply batch updates to the detail sheet
  updatesToDetailSheet.forEach(update => {
    detailSheet.getRange(update.row, update.col).setValue(update.value);
  });
   Logger.log(`Applied ${updatesToDetailSheet.length} updates to COC_Balance_Detail.`);


  if (remainingToRestore > 0.01) { // Allow for small floating point inaccuracies
     Logger.log(`Warning: Could only restore ${totalRestored.toFixed(2)} out of ${hoursToRestore.toFixed(2)} hours requested. There might be discrepancies in ledger remarks or data.`);
     // Optionally throw an error here if exact restoration is critical
     // throw new Error(`Could only restore ${totalRestored.toFixed(2)} out of ${hoursToRestore.toFixed(2)} hours.`);
  }

  Logger.log(`Total hours restored: ${totalRestored.toFixed(2)}`);
  return totalRestored; // Return the actual amount restored
}

// ============================================================================
// TEST FUNCTION
// ============================================================================

function deductCOCHoursFIFO(employeeId, hoursToDeduct, referenceId) {
  const db = getDatabase();
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  // --- MODIFICATION: Added Ledger Sheet ---
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  // --- END MODIFICATION ---

  if (!detailSheet) throw new Error('COC_Balance_Detail sheet not found');
  // --- MODIFICATION: Added Ledger Sheet Check ---
  if (!ledgerSheet) throw new Error('COC_Ledger sheet not found');
  // --- END MODIFICATION ---


  const data = detailSheet.getDataRange().getValues();
  let remainingToDeduct = hoursToDeduct;
  const TIME_ZONE = getScriptTimeZone(); // Get timezone
  const updatesToDetailSheet = []; // For batch updates
  const ledgerEntries = []; // For batch ledger updates

  Logger.log(`FIFO Deduction: Attempting to deduct ${hoursToDeduct} hrs for ${employeeId}, reference: ${referenceId}`);


  // Get active entries for this employee, sorted by date (FIFO)
  const activeEntries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Check Employee ID, Status 'Active', and Hours Remaining > 0
    if (row[1] === employeeId && row[8] === 'Active' && (parseFloat(row[6]) || 0) > 0) {
      activeEntries.push({
        rowIndex: i + 1, // 1-based index
        entryId: row[0],
        recordId: row[3],
        dateEarned: new Date(row[4]),
        hoursRemaining: parseFloat(row[6]) || 0
      });
    }
  }

  // Sort by date earned (oldest first - FIFO)
  activeEntries.sort((a, b) => a.dateEarned - b.dateEarned);

   // Check if enough balance exists before starting deductions
    const totalAvailable = activeEntries.reduce((sum, entry) => sum + entry.hoursRemaining, 0);
    if (hoursToDeduct > totalAvailable) {
        Logger.log(`FIFO Deduction Error: Insufficient balance. Available: ${totalAvailable.toFixed(2)}, Needed: ${hoursToDeduct.toFixed(2)}`);
        throw new Error(`Insufficient COC balance for deduction. Available: ${totalAvailable.toFixed(2)}, Requested: ${hoursToDeduct.toFixed(2)}`);
    }


  // Deduct from entries
  activeEntries.forEach(entry => {
    if (remainingToDeduct <= 0) return;

    const hoursFromThisEntry = Math.min(entry.hoursRemaining, remainingToDeduct);
    const newRemaining = entry.hoursRemaining - hoursFromThisEntry;

    // Prepare detail sheet updates
    updatesToDetailSheet.push({ row: entry.rowIndex, col: 7, value: newRemaining }); // Hours Remaining (Col G)
     Logger.log(`FIFO Deduction: Using ${hoursFromThisEntry.toFixed(2)} hrs from ${entry.recordId} (Row ${entry.rowIndex}). New remaining: ${newRemaining.toFixed(2)}`);


    if (newRemaining <= 0) {
      updatesToDetailSheet.push({ row: entry.rowIndex, col: 9, value: 'Used' }); // Status (Col I) - Mark as 'Used' instead of 'Depleted' for clarity
       Logger.log(`   Marking ${entry.recordId} (Row ${entry.rowIndex}) as Used.`);

    }
     // Add note about consumption (Column K = index 11)
    const currentNote = detailSheet.getRange(entry.rowIndex, 11).getValue(); // Get current note value
    const newNote = currentNote + '\n[' + Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm') +
                      '] Consumed ' + hoursFromThisEntry.toFixed(2) + ' hrs for ' + referenceId;
    updatesToDetailSheet.push({ row: entry.rowIndex, col: 11, value: newNote }); // Notes (Col K)



    // --- MODIFICATION: Prepare ledger entry INSTEAD of writing directly ---
    // We add this later AFTER calculating the final balance
    const ledgerRemark = `FIFO: Used ${hoursFromThisEntry.toFixed(2)} hrs from ${entry.recordId} (earned ${Utilities.formatDate(entry.dateEarned, TIME_ZONE, 'yyyy-MM-dd')}) for ${referenceId}`;
    ledgerEntries.push({
        employeeId: employeeId,
        // employeeName will be filled later if needed, or retrieved via getEmployeeById
        transactionDate: new Date(),
        transactionType: 'CTO Used (FIFO)', // More specific type
        referenceId: referenceId, // Link to the CTO or update action
        cocEarned: 0,
        ctoUsed: hoursFromThisEntry, // Log the specific amount deducted
        // cocBalance will be calculated and added later
        monthYearEarned: '', // Not applicable for usage
        expirationDate: '', // Not applicable for usage
        processedBy: Session.getActiveUser().getEmail(),
        remarks: ledgerRemark
    });
    // --- END MODIFICATION ---

    remainingToDeduct -= hoursFromThisEntry;
  });

   // Apply batch updates to the detail sheet
    updatesToDetailSheet.forEach(update => {
        detailSheet.getRange(update.row, update.col).setValue(update.value);
    });
    Logger.log(`FIFO Deduction: Applied ${updatesToDetailSheet.length} updates to COC_Balance_Detail.`);


  // --- MODIFICATION: Add Ledger Entries with correct running balance ---
   if (ledgerEntries.length > 0) {
        const employeeData = getEmployeeById(employeeId); // Get employee name
        const employeeName = employeeData ? employeeData.fullName : 'Unknown Employee';
        const finalBalance = getCurrentCOCBalance(employeeId); // Get final balance AFTER deductions are applied to detail sheet

        const rowsToAdd = ledgerEntries.map((entry, index) => {
             // Calculate the running balance for each entry (might be slightly off if multiple entries, but good for audit)
             // A more accurate way is to just put the final balance on all entries for this transaction.
            return [
                generateLedgerEntryId(),
                entry.employeeId,
                employeeName, // Add employee name
                entry.transactionDate,
                entry.transactionType,
                entry.referenceId,
                entry.cocEarned,
                entry.ctoUsed,
                finalBalance, // Use the final balance for all related entries
                entry.monthYearEarned,
                entry.expirationDate,
                entry.processedBy,
                entry.remarks
            ];
        });

        // Batch write ledger entries
        ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
        Logger.log(`FIFO Deduction: Added ${rowsToAdd.length} ledger entries.`);
    }
  // --- END MODIFICATION ---

  // This check should ideally not be needed if balance check is done before calling
  if (remainingToDeduct > 0.01) { // Allow tiny floating point differences
    Logger.log(`FIFO Deduction Error: Could not deduct all required hours. ${remainingToDeduct.toFixed(2)} hrs remaining.`);
    throw new Error('Insufficient COC balance detected during FIFO deduction.');
  }
   Logger.log(`FIFO Deduction: Successfully deducted ${hoursToDeduct.toFixed(2)} hrs.`);
}


// ... rest of the code including restoreCOCHoursFIFO ...

// ============================================================================
// TEST FUNCTION
// ============================================================================

/**
 * Test updating a CTO application
 */
function testUpdateCTOApplication() {
  // CHANGE THESE to real values from your sheet
  const testAppId = 'CTO-20251025001613300'; // Example ID - CHANGE THIS
  const newHours = 8.0; // Change from current value
  const newStartDate = '2025-10-25'; // Example Date - CHANGE THIS
  const newEndDate = '2025-10-25';   // Example Date - CHANGE THIS
  const newRemarks = 'Updated hours for testing - Now 8 hours';

  Logger.log('=== Testing apiUpdateCTOApplication ===');
  Logger.log('Application ID: ' + testAppId);

  const result = apiUpdateCTOApplication(testAppId, newHours, newStartDate, newEndDate, newRemarks);

  Logger.log('\nResult:');
  Logger.log('Success: ' + result.success);
  Logger.log('Message: ' + result.message);

  Logger.log('\n=== Test Complete ===');
}

/**
 * ============================================================================
 * API: Get Employees with Expiring COC (within 30 days)
 * ============================================================================
 * Retrieves a list of employees who have active COC entries expiring within
 * the next 30 days, along with the total hours expiring for each employee
 * and the date of the earliest expiration.
 *
 * @return {Array<Object>} An array of objects: { employeeId, employeeName, totalHoursExpiring, earliestExpiryDate }
 */
function apiGetEmployeesWithExpiringCOC() {
  try {
    const db = getDatabase();
    const detailSheet = ensureCOCBalanceDetailSheet(); // Ensure the sheet exists
    const TIME_ZONE = getScriptTimeZone();

    if (!detailSheet || detailSheet.getLastRow() < 2) {
      Logger.log('COC_Balance_Detail sheet is empty or not found.');
      return [];
    }

    const data = detailSheet.getDataRange().getValues();
    const expiringMap = {}; // Use a map to group by employeeId

    const today = new Date();
    today.setHours(0, 0, 0, 0); // Start of today
    const thirtyDaysFromNow = new Date(today);
    thirtyDaysFromNow.setDate(today.getDate() + 30); // End of the 30-day window

    Logger.log(`Checking for expirations between ${today.toISOString()} and ${thirtyDaysFromNow.toISOString()}`);

    // Iterate through detail entries (skip header row)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const employeeId = String(row[DETAIL_COLS.EMPLOYEE_ID] || '').trim();
      const employeeName = String(row[DETAIL_COLS.EMPLOYEE_NAME] || '').trim();
      const hoursRemaining = parseFloat(row[DETAIL_COLS.HOURS_REMAINING]) || 0;
      const expirationDateVal = row[DETAIL_COLS.EXPIRATION_DATE];
      const status = String(row[DETAIL_COLS.STATUS] || '').trim();

      // Skip if no employee ID, status is not Active, or no hours remaining
      if (!employeeId || status !== 'Active' || hoursRemaining <= 0) {
        continue;
      }

      // --- Robust Date Parsing for Expiration Date ---
      let expirationDate = null;
      if (expirationDateVal instanceof Date && !isNaN(expirationDateVal.getTime())) {
          expirationDate = new Date(expirationDateVal); // Clone to avoid modifying original
      } else if (expirationDateVal) {
          try {
              const parsed = new Date(expirationDateVal);
              if (!isNaN(parsed.getTime())) {
                  expirationDate = parsed;
              } else {
                 Logger.log(`Skipping row ${i+1} due to invalid expiration date format: ${expirationDateVal}`);
                 continue; // Skip if date is invalid
              }
          } catch (e) {
               Logger.log(`Error parsing expiration date at row ${i+1}: ${expirationDateVal}`);
               continue; // Skip on parsing error
          }
      } else {
           Logger.log(`Skipping row ${i+1} due to missing expiration date.`);
           continue; // Skip if no expiration date
      }
      expirationDate.setHours(0,0,0,0); // Use start of the expiration day for comparison
      // --- End Date Parsing ---


      // Check if the expiration date is within the next 30 days (inclusive of today)
      if (expirationDate >= today && expirationDate <= thirtyDaysFromNow) {
        // If employee not yet in map, initialize
        if (!expiringMap[employeeId]) {
          expiringMap[employeeId] = {
            employeeId: employeeId,
            employeeName: employeeName,
            totalHoursExpiring: 0,
            earliestExpiryDate: expirationDate // Initialize with the first one found
          };
        }

        // Add hours to the total for this employee
        expiringMap[employeeId].totalHoursExpiring += hoursRemaining;

        // Update the earliest expiration date if this one is earlier
        if (expirationDate < expiringMap[employeeId].earliestExpiryDate) {
          expiringMap[employeeId].earliestExpiryDate = expirationDate;
        }
         Logger.log(`Found expiring entry for ${employeeId}: ${hoursRemaining} hrs on ${Utilities.formatDate(expirationDate, TIME_ZONE, 'yyyy-MM-dd')}`);
      }
    }

    // Convert map to array
    const expiringList = Object.values(expiringMap);

    // Sort the list (e.g., by earliest expiry date, then by name)
    expiringList.sort((a, b) => {
      if (a.earliestExpiryDate.getTime() !== b.earliestExpiryDate.getTime()) {
        return a.earliestExpiryDate - b.earliestExpiryDate; // Sort by date first
      }
      return a.employeeName.localeCompare(b.employeeName); // Then by name
    });

     // Format the date string for the final output
     const formattedList = expiringList.map(item => ({
       ...item,
       earliestExpiryDate: Utilities.formatDate(item.earliestExpiryDate, TIME_ZONE, 'yyyy-MM-dd') // Format date as YYYY-MM-DD string
     }));


    Logger.log(`Found ${formattedList.length} employees with COC expiring within 30 days.`);
    return formattedList; // Return the array

  } catch (error) {
    Logger.log('ERROR in apiGetEmployeesWithExpiringCOC: ' + error.message + '\nStack: ' + error.stack);
    return []; // Return empty array on error
  }
}

// ============================================================================
// HOLIDAY MANAGER API FUNCTIONS
// ============================================================================

/**
 * API: List all holidays
 * @return {Array<Object>} Array of holiday objects
 */
function apiListHolidays() {
  try {
    const db = getDatabase();
    const sheet = db.getSheetByName('Holidays');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const results = [];
    const TIME_ZONE = getScriptTimeZone();

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[0]; // Column A: Date
      
      if (!date) continue; // Skip empty rows

      results.push({
        rowNumber: i + 1,
        date: date,
        dateISO: Utilities.formatDate(new Date(date), TIME_ZONE, 'yyyy-MM-dd'),
        type: row[1] || '',           // Column B: Type
        description: row[2] || '',     // Column C: Description
        halfdayTime: row[3] || '',     // Column D: Half-day Start Time
        suspensionTime: row[4] || '',  // Column E: Work Suspension Time
        remarks: row[5] || ''          // Column F: Remarks
      });
    }

    // Sort by date (most recent first for better UX)
    results.sort((a, b) => new Date(b.date) - new Date(a.date));

    return results;

  } catch (error) {
    Logger.log('ERROR in apiListHolidays: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to load holidays: ' + error.message);
  }
}

/**
 * API: Add a new holiday
 * @param {Date} date The date of the holiday
 * @param {string} type The type of holiday
 * @param {string} description Optional description
 * @param {string} halfdayTime Optional time for half-day holidays
 * @param {string} suspensionTime Optional time for work suspension
 * @param {string} remarks Optional additional remarks
 * @return {Object} Success result
 */
function apiAddHoliday(date, type, description, halfdayTime, suspensionTime, remarks) {
  try {
    const db = getDatabase();
    let sheet = db.getSheetByName('Holidays');

    if (!sheet) {
      sheet = db.insertSheet('Holidays');
      sheet.getRange(1, 1, 1, 6).setValues([[
        'Date', 'Type', 'Description', 'Half-day Start Time', 'Suspension Time', 'Remarks'
      ]]);
      sheet.setFrozenRows(1);
    }

    // Validate inputs
    if (!date || !(date instanceof Date)) {
      throw new Error('Invalid date provided');
    }

    if (!type) {
      throw new Error('Holiday type is required');
    }

    // Check for duplicate date
    const existingData = sheet.getDataRange().getValues();
    const TIME_ZONE = getScriptTimeZone();
    const dateStr = Utilities.formatDate(new Date(date), TIME_ZONE, 'yyyy-MM-dd');
    
    for (let i = 1; i < existingData.length; i++) {
      const existingDate = existingData[i][0];
      if (existingDate && Utilities.formatDate(new Date(existingDate), TIME_ZONE, 'yyyy-MM-dd') === dateStr) {
        throw new Error('A holiday already exists for this date. Please edit the existing entry instead.');
      }
    }

    // Add new row
    const newRow = [
      new Date(date),
      type,
      description || '',
      halfdayTime || '',
      suspensionTime || '',
      remarks || ''
    ];

    sheet.appendRow(newRow);

    Logger.log('Holiday added successfully: ' + dateStr + ' - ' + type);
    return { success: true, message: 'Holiday added successfully' };

  } catch (error) {
    Logger.log('ERROR in apiAddHoliday: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to add holiday: ' + error.message);
  }
}

/**
 * API: Update an existing holiday
 * @param {number} rowNumber The row number to update
 * @param {Date} date The date of the holiday
 * @param {string} type The type of holiday
 * @param {string} description Optional description
 * @param {string} halfdayTime Optional time for half-day holidays
 * @param {string} suspensionTime Optional time for work suspension
 * @param {string} remarks Optional additional remarks
 * @return {Object} Success result
 */
function apiUpdateHoliday(rowNumber, date, type, description, halfdayTime, suspensionTime, remarks) {
  try {
    const db = getDatabase();
    const sheet = db.getSheetByName('Holidays');
    
    if (!sheet) {
      throw new Error('Holidays sheet not found');
    }

    // Validate inputs
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number');
    }

    if (!date || !(date instanceof Date)) {
      throw new Error('Invalid date provided');
    }

    if (!type) {
      throw new Error('Holiday type is required');
    }

    // Check for duplicate date (excluding current row)
    const existingData = sheet.getDataRange().getValues();
    const TIME_ZONE = getScriptTimeZone();
    const dateStr = Utilities.formatDate(new Date(date), TIME_ZONE, 'yyyy-MM-dd');
    
    for (let i = 1; i < existingData.length; i++) {
      if (i + 1 === rowNumber) continue; // Skip current row
      const existingDate = existingData[i][0];
      if (existingDate && Utilities.formatDate(new Date(existingDate), TIME_ZONE, 'yyyy-MM-dd') === dateStr) {
        throw new Error('A holiday already exists for this date. Please choose a different date.');
      }
    }

    // Update the row
    const range = sheet.getRange(rowNumber, 1, 1, 6);
    range.setValues([[
      new Date(date),
      type,
      description || '',
      halfdayTime || '',
      suspensionTime || '',
      remarks || ''
    ]]);

    Logger.log('Holiday updated successfully: ' + dateStr + ' - ' + type);
    return { success: true, message: 'Holiday updated successfully' };

  } catch (error) {
    Logger.log('ERROR in apiUpdateHoliday: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to update holiday: ' + error.message);
  }
}

/**
 * API: Delete a holiday
 * @param {number} rowNumber The row number to delete
 * @return {Object} Success result
 */
function apiDeleteHoliday(rowNumber) {
  try {
    const db = getDatabase();
    const sheet = db.getSheetByName('Holidays');
    
    if (!sheet) {
      throw new Error('Holidays sheet not found');
    }

    // Validate row number
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number');
    }

    if (rowNumber > sheet.getLastRow()) {
      throw new Error('Row number does not exist');
    }

    // Delete the row
    sheet.deleteRow(rowNumber);

    Logger.log('Holiday deleted successfully from row: ' + rowNumber);
    return { success: true, message: 'Holiday deleted successfully' };

  } catch (error) {
    Logger.log('ERROR in apiDeleteHoliday: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to delete holiday: ' + error.message);
  }
}

/**
 * Enhanced getDayType function with support for new holiday types
 * This replaces the existing getDayType function in the main code
 * 
 * @param {Date} date The date to check
 * @return {Object} Object with dayType, multiplier, and additional info
 */
function getDayTypeEnhanced(date) {
  const TIME_ZONE = getScriptTimeZone();
  const dow = date.getDay(); // 0=Sun
  let dayType = 'Weekday';
  let multiplier = 1.0;
  let additionalInfo = {};
  
  if (dow === 0 || dow === 6) {
    dayType = 'Weekend';
    multiplier = 1.5;
  }

  // Check holidays sheet for overrides
  const db = getDatabase();
  const holidaysSheet = db.getSheetByName('Holidays');
  if (holidaysSheet) {
    const holData = holidaysSheet.getDataRange().getValues();
    const target = Utilities.formatDate(date, TIME_ZONE, 'yyyy-MM-dd');
    
    for (let i = 1; i < holData.length; i++) {
      const holDate = holData[i][0];
      const holType = holData[i][1];
      const halfdayTime = holData[i][3];
      const suspensionTime = holData[i][4];
      
      if (holDate && Utilities.formatDate(new Date(holDate), TIME_ZONE, 'yyyy-MM-dd') === target) {
        if (holType === 'Regular') {
          dayType = 'Regular Holiday';
          multiplier = 1.5;
        } else if (holType === 'Special Non-Working') {
          dayType = 'Special Non-Working';
          multiplier = 1.5;
        } else if (holType === 'Local') {
          dayType = 'Local Holiday';
          multiplier = 1.5;
        } else if (holType === 'No Work') {
          dayType = 'No Work / Typhoon';
          multiplier = 1.0;
        } else if (holType === 'Half-day') {
          dayType = 'Half-day Holiday';
          multiplier = 1.5;
          additionalInfo.halfdayTime = halfdayTime;
        } else if (holType === 'Work Suspended') {
          dayType = 'Work Suspended';
          multiplier = 1.0;
          additionalInfo.suspensionTime = suspensionTime;
        }
        break;
      }
    }
  }
  
  return {
    dayType: dayType,
    multiplier: multiplier,
    additionalInfo: additionalInfo
  };
}

/**
 * Calculate COC for a specific date with enhanced holiday logic
 * @param {Date} date The date to calculate COC for
 * @param {number} hoursWorked The hours worked
 * @param {string} timeIn Optional time in (for half-day calculations)
 * @param {string} timeOut Optional time out (for half-day calculations)
 * @return {Object} Object with COC hours and day type info
 */
function calculateCOCForDate(date, hoursWorked, timeIn, timeOut) {
  const dayInfo = getDayTypeEnhanced(date);
  let cocHours = 0;
  
  // For half-day holidays, we need to calculate based on time worked
  if (dayInfo.dayType === 'Half-day Holiday' && dayInfo.additionalInfo.halfdayTime && timeIn && timeOut) {
    // Parse times
    const halfdayStartMins = timeToMinutes(dayInfo.additionalInfo.halfdayTime);
    const timeInMins = timeToMinutes(timeIn);
    const timeOutMins = timeToMinutes(timeOut);
    
    // Calculate hours before and after half-day start
    if (timeOutMins <= halfdayStartMins) {
      // All work was before half-day, use normal weekday rate
      cocHours = hoursWorked * 1.0;
    } else if (timeInMins >= halfdayStartMins) {
      // All work was during half-day, use holiday rate
      cocHours = hoursWorked * 1.5;
    } else {
      // Work spanned both periods, calculate proportionally
      const hoursBeforeHalfday = (halfdayStartMins - timeInMins) / 60;
      const hoursAfterHalfday = (timeOutMins - halfdayStartMins) / 60;
      cocHours = (hoursBeforeHalfday * 1.0) + (hoursAfterHalfday * 1.5);
    }
  } else {
    // Standard calculation using multiplier
    cocHours = hoursWorked * dayInfo.multiplier;
  }
  
  return {
    cocHours: cocHours,
    dayType: dayInfo.dayType,
    multiplier: dayInfo.multiplier,
    additionalInfo: dayInfo.additionalInfo
  };
}

/**
 * Helper function to convert time string to minutes
 * @param {string} timeStr Time in HH:mm format
 * @return {number} Minutes since midnight
 */
function timeToMinutes(timeStr) {
  if (!timeStr) return 0;
  const parts = timeStr.split(':');
  return parseInt(parts[0]) * 60 + parseInt(parts[1]);
}

// ============================================================================
// HISTORICAL COC/CTO IMPORT API FUNCTIONS
// ============================================================================

/**
 * API: Import a single historical COC entry
 * @param {Object} data Import data object
 * @return {Object} Success result
 */
function apiImportHistoricalCOC(data) {
  try {
    const db = getDatabase();
    const detailSheet = ensureCOCBalanceDetailSheet();
    const ledgerSheet = ensureLedgerSheet();
    
    if (!detailSheet) {
      throw new Error('COC_Balance_Detail sheet not found');
    }

    // Validate employee
    const employee = getEmployeeById(data.employeeId);
    if (!employee) {
      throw new Error('Employee not found: ' + data.employeeId);
    }

    // Validate data
    if (!data.monthYear || !data.certificateDate || !data.hoursEarned) {
      throw new Error('Missing required fields');
    }

    if (data.hoursUsed > data.hoursEarned) {
      throw new Error('Hours Used cannot exceed Hours Earned');
    }

    const hoursRemaining = data.hoursEarned - data.hoursUsed;

    // Generate IDs
    const recordId = generateRecordId();
    const certificateId = 'CERT-' + Utilities.formatDate(new Date(), getScriptTimeZone(), 'yyyyMMddHHmmssSSS');
    
    // Parse dates
    const certDate = new Date(data.certificateDate);
    const expDate = new Date(data.expirationDate);
    const importDate = new Date();

    // Create the detail entry
    const detailRow = [
      recordId,                    // A: Record ID
      data.employeeId,             // B: Employee ID
      employee.fullName,           // C: Employee Name
      data.monthYear,              // D: Month-Year Earned
      certDate,                    // E: Certificate Date
      data.hoursEarned,            // F: Hours Earned
      data.hoursUsed,              // G: Hours Used
      hoursRemaining,              // H: Hours Remaining
      data.status,                 // I: Status
      expDate,                     // J: Expiration Date
      certificateId,               // K: Certificate ID
      importDate,                  // L: Date Created
      Session.getActiveUser().getEmail(), // M: Created By
      data.remarks || 'Historical Import' // N: Remarks
    ];

    detailSheet.appendRow(detailRow);
    Logger.log('Historical COC entry imported: ' + recordId);

    // Add ledger entry if there are hours
    if (data.hoursEarned > 0) {
      const ledgerRow = [
        generateLedgerEntryId(),
        data.employeeId,
        employee.fullName,
        importDate,
        'Historical Import',
        recordId,
        data.hoursEarned,
        0, // No CTO used in import
        getCurrentCOCBalance(data.employeeId),
        data.monthYear,
        expDate,
        Session.getActiveUser().getEmail(),
        data.remarks || 'Historical data import'
      ];
      
      ledgerSheet.appendRow(ledgerRow);
      Logger.log('Ledger entry created for historical import');
    }

    return { 
      success: true, 
      message: 'Historical COC entry imported successfully',
      recordId: recordId
    };

  } catch (error) {
    Logger.log('ERROR in apiImportHistoricalCOC: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to import historical COC: ' + error.message);
  }
}

/**
 * API: Import multiple historical COC entries from CSV
 * @param {Array} csvData Array of CSV rows (first row is headers)
 * @return {Object} Result with success and error counts
 */
function apiImportHistoricalCOCBatch(csvData) {
  try {
    if (!csvData || csvData.length < 2) {
      throw new Error('CSV data is empty or invalid');
    }

    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    // Expected headers: Employee ID, Month-Year, Certificate Date, Hours Earned, Hours Used, Status, Remarks
    const headers = csvData[0];
    
    // Skip header row, process data rows
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      
      // Skip empty rows
      if (!row[0] || row[0].trim() === '') {
        continue;
      }

      try {
        // Parse row data
        const employeeId = row[0].trim();
        const monthYear = row[1].trim();
        const certificateDate = row[2].trim();
        const hoursEarned = parseFloat(row[3]) || 0;
        const hoursUsed = parseFloat(row[4]) || 0;
        const status = row[5].trim() || 'Active';
        const remarks = row[6] || 'CSV Import';

        // Calculate expiration date (Certificate Date + 1 year - 1 day)
        const certDate = new Date(certificateDate);
        const expDate = new Date(certDate);
        expDate.setFullYear(expDate.getFullYear() + 1);
        expDate.setDate(expDate.getDate() - 1);

        // Import the entry
        const data = {
          employeeId: employeeId,
          monthYear: monthYear,
          certificateDate: certificateDate,
          expirationDate: expDate.toISOString().split('T')[0],
          hoursEarned: hoursEarned,
          hoursUsed: hoursUsed,
          status: status,
          remarks: remarks
        };

        apiImportHistoricalCOC(data);
        successCount++;

      } catch (rowError) {
        errorCount++;
        errors.push(`Row ${i + 1}: ${rowError.message}`);
        Logger.log(`Error importing row ${i + 1}: ${rowError.message}`);
      }
    }

    const result = {
      success: true,
      successCount: successCount,
      errorCount: errorCount,
      errors: errors
    };

    if (errorCount > 0) {
      Logger.log(`Batch import completed with ${errorCount} errors: ${errors.join('; ')}`);
    }

    return result;

  } catch (error) {
    Logger.log('ERROR in apiImportHistoricalCOCBatch: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to import CSV: ' + error.message);
  }
}

/**
 * API: Get historical import records (last 50)
 * @return {Array<Object>} Array of import history objects
 */
function apiGetHistoricalImports() {
  try {
    const db = getDatabase();
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    
    if (!detailSheet || detailSheet.getLastRow() < 2) {
      return [];
    }

    const data = detailSheet.getDataRange().getValues();
    const results = [];
    const TIME_ZONE = getScriptTimeZone();

    // Look for entries with "Historical Import" in remarks
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const remarks = String(row[13] || '').toLowerCase();
      
      if (remarks.includes('historical') || remarks.includes('import') || remarks.includes('migrat')) {
        results.push({
          employeeId: row[1],
          employeeName: row[2],
          monthYear: row[3],
          certificateDate: Utilities.formatDate(new Date(row[4]), TIME_ZONE, 'MMM dd, yyyy'),
          hoursEarned: Number(row[5]).toFixed(2),
          hoursUsed: Number(row[6]).toFixed(2),
          hoursRemaining: Number(row[7]).toFixed(2),
          status: row[8],
          importedDate: Utilities.formatDate(new Date(row[11]), TIME_ZONE, 'MMM dd, yyyy HH:mm')
        });
      }
    }

    // Sort by import date (most recent first)
    results.sort((a, b) => {
      const dateA = new Date(a.importedDate);
      const dateB = new Date(b.importedDate);
      return dateB - dateA;
    });

    // Return last 50 entries
    return results.slice(0, 50);

  } catch (error) {
    Logger.log('ERROR in apiGetHistoricalImports: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to load import history: ' + error.message);
  }
}

/**
 * API: List employees for dropdown (simplified)
 * @return {Array<Object>} Array of employee objects with id and fullName
 */
function apiListEmployeesForDropdown() {
  try {
    const db = getDatabase();
    const empSheet = db.getSheetByName('Employees');
    
    if (!empSheet || empSheet.getLastRow() < 2) {
      return [];
    }

    const data = empSheet.getDataRange().getValues();
    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[6]; // Assuming status is in column G
      
      // Only include active employees
      if (status !== 'Active') continue;
      
      results.push({
        id: row[0],        // Employee ID
        fullName: row[1]   // Full Name
      });
    }

    // Sort by name
    results.sort((a, b) => a.fullName.localeCompare(b.fullName));

    return results;

  } catch (error) {
    Logger.log('ERROR in apiListEmployeesForDropdown: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to load employees: ' + error.message);
  }
}

/**
 * Ensure Ledger sheet exists
 * @return {Sheet} The ledger sheet
 */
function ensureLedgerSheet() {
  const db = getDatabase();
  let sheet = db.getSheetByName('Ledger');
  
  if (!sheet) {
    sheet = db.insertSheet('Ledger');
    // Add headers
    const headers = [
      'Ledger ID', 'Employee ID', 'Employee Name', 'Transaction Date', 
      'Transaction Type', 'Reference ID', 'COC Earned', 'CTO Used', 
      'Running Balance', 'Month-Year Earned', 'Expiration Date', 
      'Processed By', 'Remarks'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('Ledger sheet created');
  }
  
  return sheet;
}

/**
 * Generate a unique ledger entry ID
 * @return {string} Ledger entry ID
 */
function generateLedgerEntryId() {
  const now = new Date();
  const TIME_ZONE = getScriptTimeZone();
  return 'LED-' + Utilities.formatDate(now, TIME_ZONE, 'yyyyMMddHHmmssSSS');
}

// ============================================================================
// FIFO INTEGRITY CHECK TOOL
// ============================================================================

/**
 * FIFO Integrity Check and Reconciliation Tool
 * 
 * This tool verifies that all CTO deductions follow First-In, First-Out logic
 * and provides utilities to fix any discrepancies found.
 */

/**
 * API: Run FIFO integrity check for all employees or specific employee
 * @param {string} employeeId Optional - specific employee to check, or null for all
 * @return {Object} Report with findings and discrepancies
 */
function apiFIFOIntegrityCheck(employeeId) {
  try {
    const db = getDatabase();
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    const ctoSheet = db.getSheetByName('CTO_Applications');
    
    if (!detailSheet) {
      throw new Error('COC_Balance_Detail sheet not found');
    }
    
    if (!ctoSheet) {
      throw new Error('CTO_Applications sheet not found');
    }

    Logger.log('=== Starting FIFO Integrity Check ===');
    Logger.log('Employee filter: ' + (employeeId || 'ALL'));

    const detailData = detailSheet.getDataRange().getValues();
    const ctoData = ctoSheet.getDataRange().getValues();
    
    const report = {
      checkDate: new Date(),
      employeeId: employeeId || 'ALL',
      totalEmployeesChecked: 0,
      totalDiscrepancies: 0,
      discrepancyDetails: [],
      integrityIssues: [],
      summary: ''
    };

    // Get unique employees to check
    const employeesToCheck = new Set();
    for (let i = 1; i < detailData.length; i++) {
      const empId = detailData[i][DETAIL_COLS.EMPLOYEE_ID];
      if (!employeeId || empId === employeeId) {
        employeesToCheck.add(empId);
      }
    }

    // Check each employee
    for (const empId of employeesToCheck) {
      Logger.log(`Checking employee: ${empId}`);
      const employeeIssues = checkEmployeeFIFO(empId, detailData, ctoData);
      
      if (employeeIssues.length > 0) {
        report.totalDiscrepancies++;
        report.integrityIssues.push({
          employeeId: empId,
          issues: employeeIssues
        });
      }
      
      report.totalEmployeesChecked++;
    }

    // Generate summary
    if (report.totalDiscrepancies === 0) {
      report.summary = `✓ All ${report.totalEmployeesChecked} employee(s) passed FIFO integrity check. No discrepancies found.`;
    } else {
      report.summary = `⚠ Found FIFO discrepancies for ${report.totalDiscrepancies} out of ${report.totalEmployeesChecked} employee(s).`;
    }

    Logger.log('=== FIFO Integrity Check Complete ===');
    Logger.log(report.summary);

    return report;

  } catch (error) {
    Logger.log('ERROR in apiFIFOIntegrityCheck: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('FIFO integrity check failed: ' + error.message);
  }
}

/**
 * Check FIFO integrity for a specific employee
 * @param {string} employeeId Employee ID to check
 * @param {Array} detailData COC_Balance_Detail data
 * @param {Array} ctoData CTO_Applications data
 * @return {Array} Array of issues found
 */
function checkEmployeeFIFO(employeeId, detailData, ctoData) {
  const issues = [];
  
  // Get all COC entries for this employee, sorted by certificate date (FIFO order)
  const cocEntries = [];
  for (let i = 1; i < detailData.length; i++) {
    const row = detailData[i];
    if (row[DETAIL_COLS.EMPLOYEE_ID] === employeeId && String(row[DETAIL_COLS.STATUS] || '').trim() === 'Active') {
      let certificateDate = row[DETAIL_COLS.CERTIFICATE_DATE];
      if (!(certificateDate instanceof Date) || isNaN(certificateDate.getTime())) {
        certificateDate = parseMonthYear(row[DETAIL_COLS.MONTH_YEAR]) || new Date(row[DETAIL_COLS.DATE_CREATED]);
      }

      let expirationDate = row[DETAIL_COLS.EXPIRATION_DATE];
      if (!(expirationDate instanceof Date) || isNaN(expirationDate.getTime())) {
        expirationDate = null;
      }

      cocEntries.push({
        rowNumber: i + 1,
        recordId: row[DETAIL_COLS.RECORD_ID],
        certificateDate: certificateDate,
        hoursEarned: parseFloat(row[DETAIL_COLS.HOURS_EARNED]) || 0,
        hoursUsed: parseFloat(row[DETAIL_COLS.HOURS_USED]) || 0,
        hoursRemaining: parseFloat(row[DETAIL_COLS.HOURS_REMAINING]) || 0,
        expirationDate: expirationDate
      });
    }
  }

  // Sort by certificate date (oldest first - FIFO)
  cocEntries.sort((a, b) => a.certificateDate - b.certificateDate);

  // Get all approved CTO applications for this employee
  const ctoApplications = [];
  for (let i = 1; i < ctoData.length; i++) {
    const row = ctoData[i];
    if (row[1] === employeeId && row[8] === 'Approved') { // Column B: Employee ID, Column I: Status
      ctoApplications.push({
        applicationId: row[0],
        applicationDate: new Date(row[2]),
        hoursRequested: parseFloat(row[3]) || 0
      });
    }
  }

  // Sort CTO applications by date
  ctoApplications.sort((a, b) => a.applicationDate - b.applicationDate);

  // Simulate FIFO deduction and compare with actual
  let simulatedDeductions = cocEntries.map(entry => ({
    recordId: entry.recordId,
    expectedUsed: 0,
    expectedRemaining: entry.hoursEarned
  }));

  for (const cto of ctoApplications) {
    let remainingToDeduct = cto.hoursRequested;
    
    for (let i = 0; i < simulatedDeductions.length && remainingToDeduct > 0.01; i++) {
      const availableHours = simulatedDeductions[i].expectedRemaining;
      
      if (availableHours > 0) {
        const hoursToDeduct = Math.min(availableHours, remainingToDeduct);
        simulatedDeductions[i].expectedUsed += hoursToDeduct;
        simulatedDeductions[i].expectedRemaining -= hoursToDeduct;
        remainingToDeduct -= hoursToDeduct;
      }
    }
    
    if (remainingToDeduct > 0.01) {
      issues.push({
        type: 'INSUFFICIENT_BALANCE',
        ctoApplicationId: cto.applicationId,
        message: `CTO application ${cto.applicationId} deducted hours when insufficient balance existed`,
        shortfall: remainingToDeduct.toFixed(2)
      });
    }
  }

  // Compare simulated vs actual
  for (let i = 0; i < cocEntries.length; i++) {
    const actual = cocEntries[i];
    const simulated = simulatedDeductions.find(s => s.recordId === actual.recordId);
    
    if (simulated) {
      const usedDiff = Math.abs(actual.hoursUsed - simulated.expectedUsed);
      const remainingDiff = Math.abs(actual.hoursRemaining - simulated.expectedRemaining);
      
      if (usedDiff > 0.01 || remainingDiff > 0.01) {
        issues.push({
          type: 'FIFO_VIOLATION',
          recordId: actual.recordId,
          certificateDate: actual.certificateDate,
          actualUsed: actual.hoursUsed.toFixed(2),
          expectedUsed: simulated.expectedUsed.toFixed(2),
          actualRemaining: actual.hoursRemaining.toFixed(2),
          expectedRemaining: simulated.expectedRemaining.toFixed(2),
          message: 'Hours used/remaining do not match FIFO order'
        });
      }
    }
  }

  // Check for negative balances
  for (const entry of cocEntries) {
    if (entry.hoursRemaining < -0.01) {
      issues.push({
        type: 'NEGATIVE_BALANCE',
        recordId: entry.recordId,
        hoursRemaining: entry.hoursRemaining.toFixed(2),
        message: 'COC entry has negative remaining hours'
      });
    }
  }

  return issues;
}

/**
 * API: Fix FIFO discrepancies for an employee
 * This will recalculate all COC balances using proper FIFO logic
 * @param {string} employeeId Employee ID to fix
 * @param {boolean} dryRun If true, only simulate the fix without applying changes
 * @return {Object} Result of the fix operation
 */
function apiFIFOFix(employeeId, dryRun) {
  try {
    if (!employeeId) {
      throw new Error('Employee ID is required');
    }

    const db = getDatabase();
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    const ctoSheet = db.getSheetByName('CTO_Applications');
    
    Logger.log(`=== ${dryRun ? 'Simulating' : 'Executing'} FIFO Fix for ${employeeId} ===`);

    const detailData = detailSheet.getDataRange().getValues();
    const ctoData = ctoSheet.getDataRange().getValues();

    // Get all COC entries for this employee
    const cocEntries = [];
    for (let i = 1; i < detailData.length; i++) {
      const row = detailData[i];
      if (row[DETAIL_COLS.EMPLOYEE_ID] === employeeId) {
        let certificateDate = row[DETAIL_COLS.CERTIFICATE_DATE];
        if (!(certificateDate instanceof Date) || isNaN(certificateDate.getTime())) {
          certificateDate = parseMonthYear(row[DETAIL_COLS.MONTH_YEAR]) || new Date(row[DETAIL_COLS.DATE_CREATED]);
        }

        cocEntries.push({
          rowNumber: i + 1,
          recordId: row[DETAIL_COLS.RECORD_ID],
          certificateDate: certificateDate,
          hoursEarned: parseFloat(row[DETAIL_COLS.HOURS_EARNED]) || 0,
          currentUsed: parseFloat(row[DETAIL_COLS.HOURS_USED]) || 0,
          currentRemaining: parseFloat(row[DETAIL_COLS.HOURS_REMAINING]) || 0,
          status: row[DETAIL_COLS.STATUS]
        });
      }
    }

    // Sort by certificate date (FIFO)
    cocEntries.sort((a, b) => a.certificateDate - b.certificateDate);

    // Reset all to unused state (for recalculation)
    cocEntries.forEach(entry => {
      entry.newUsed = 0;
      entry.newRemaining = entry.hoursEarned;
      entry.newStatus = 'Active';
    });

    // Get all approved CTO applications
    const ctoApplications = [];
    for (let i = 1; i < ctoData.length; i++) {
      const row = ctoData[i];
      if (row[1] === employeeId && row[8] === 'Approved') {
        ctoApplications.push({
          applicationId: row[0],
          applicationDate: new Date(row[2]),
          hoursRequested: parseFloat(row[3]) || 0
        });
      }
    }

    // Sort by application date
    ctoApplications.sort((a, b) => a.applicationDate - b.applicationDate);

    // Apply FIFO deduction
    for (const cto of ctoApplications) {
      let remainingToDeduct = cto.hoursRequested;
      
      for (let i = 0; i < cocEntries.length && remainingToDeduct > 0.01; i++) {
        if (cocEntries[i].newRemaining > 0) {
          const hoursToDeduct = Math.min(cocEntries[i].newRemaining, remainingToDeduct);
          cocEntries[i].newUsed += hoursToDeduct;
          cocEntries[i].newRemaining -= hoursToDeduct;
          
          // Update status
          if (cocEntries[i].newRemaining < 0.01) {
            cocEntries[i].newStatus = 'Fully Used';
          } else if (cocEntries[i].newUsed > 0.01) {
            cocEntries[i].newStatus = 'Partially Used';
          }
          
          remainingToDeduct -= hoursToDeduct;
        }
      }
    }

    // Prepare result
    const result = {
      employeeId: employeeId,
      dryRun: dryRun,
      changes: [],
      summary: ''
    };

    // Compare and apply changes
    for (const entry of cocEntries) {
      const usedChanged = Math.abs(entry.currentUsed - entry.newUsed) > 0.01;
      const remainingChanged = Math.abs(entry.currentRemaining - entry.newRemaining) > 0.01;
      const statusChanged = entry.status !== entry.newStatus;
      
      if (usedChanged || remainingChanged || statusChanged) {
        result.changes.push({
          recordId: entry.recordId,
          certificateDate: entry.certificateDate,
          before: {
            used: entry.currentUsed.toFixed(2),
            remaining: entry.currentRemaining.toFixed(2),
            status: entry.status
          },
          after: {
            used: entry.newUsed.toFixed(2),
            remaining: entry.newRemaining.toFixed(2),
            status: entry.newStatus
          }
        });

        // Apply changes if not dry run
        if (!dryRun) {
          detailSheet.getRange(entry.rowNumber, DETAIL_COLS.HOURS_USED + 1).setValue(entry.newUsed);
          detailSheet.getRange(entry.rowNumber, DETAIL_COLS.HOURS_REMAINING + 1).setValue(entry.newRemaining);
          detailSheet.getRange(entry.rowNumber, DETAIL_COLS.STATUS + 1).setValue(entry.newStatus);
        }
      }
    }

    if (result.changes.length === 0) {
      result.summary = '✓ No FIFO discrepancies found. No changes needed.';
    } else {
      result.summary = `${dryRun ? 'Would update' : 'Updated'} ${result.changes.length} COC record(s) to match FIFO order.`;
    }

    Logger.log(result.summary);
    return result;

  } catch (error) {
    Logger.log('ERROR in apiFIFOFix: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('FIFO fix failed: ' + error.message);
  }
}

/**
 * API: Get detailed FIFO report for a specific employee
 * @param {string} employeeId Employee ID
 * @return {Object} Detailed report
 */
function apiFIFOEmployeeReport(employeeId) {
  try {
    if (!employeeId) {
      throw new Error('Employee ID is required');
    }

    const db = getDatabase();
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    const ctoSheet = db.getSheetByName('CTO_Applications');
    const TIME_ZONE = getScriptTimeZone();

    const detailData = detailSheet.getDataRange().getValues();
    const ctoData = ctoSheet.getDataRange().getValues();

    // Get employee info
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      throw new Error('Employee not found');
    }

    // Get COC entries
    const cocEntries = [];
    for (let i = 1; i < detailData.length; i++) {
      const row = detailData[i];
      if (row[DETAIL_COLS.EMPLOYEE_ID] === employeeId) {
        const certificateDateVal = row[DETAIL_COLS.CERTIFICATE_DATE];
        const expirationVal = row[DETAIL_COLS.EXPIRATION_DATE];
        const certificateDate = certificateDateVal instanceof Date && !isNaN(certificateDateVal.getTime())
          ? certificateDateVal
          : parseMonthYear(row[DETAIL_COLS.MONTH_YEAR]) || new Date(row[DETAIL_COLS.DATE_CREATED]);
        const expirationDate = expirationVal instanceof Date && !isNaN(expirationVal.getTime())
          ? expirationVal
          : null;

        cocEntries.push({
          recordId: row[DETAIL_COLS.RECORD_ID],
          monthYear: row[DETAIL_COLS.MONTH_YEAR],
          certificateDate: Utilities.formatDate(certificateDate, TIME_ZONE, 'MMM dd, yyyy'),
          hoursEarned: Number(row[DETAIL_COLS.HOURS_EARNED]).toFixed(2),
          hoursUsed: Number(row[DETAIL_COLS.HOURS_USED]).toFixed(2),
          hoursRemaining: Number(row[DETAIL_COLS.HOURS_REMAINING]).toFixed(2),
          status: row[DETAIL_COLS.STATUS],
          expirationDate: expirationDate ? Utilities.formatDate(expirationDate, TIME_ZONE, 'MMM dd, yyyy') : '—'
        });
      }
    }

    // Sort by certificate date
    cocEntries.sort((a, b) => new Date(a.certificateDate) - new Date(b.certificateDate));

    // Get CTO applications
    const ctoApps = [];
    for (let i = 1; i < ctoData.length; i++) {
      const row = ctoData[i];
      if (row[1] === employeeId && row[8] === 'Approved') {
        ctoApps.push({
          applicationId: row[0],
          applicationDate: Utilities.formatDate(new Date(row[2]), TIME_ZONE, 'MMM dd, yyyy'),
          hoursUsed: Number(row[3]).toFixed(2),
          status: row[8]
        });
      }
    }

    const report = {
      employeeId: employeeId,
      employeeName: employee.fullName,
      totalCOCEarned: cocEntries.reduce((sum, e) => sum + parseFloat(e.hoursEarned), 0).toFixed(2),
      totalCOCUsed: cocEntries.reduce((sum, e) => sum + parseFloat(e.hoursUsed), 0).toFixed(2),
      totalCOCRemaining: cocEntries.reduce((sum, e) => sum + parseFloat(e.hoursRemaining), 0).toFixed(2),
      cocEntries: cocEntries,
      ctoApplications: ctoApps,
      integrityCheck: checkEmployeeFIFO(employeeId, detailData, ctoData)
    };

    return report;

  } catch (error) {
    Logger.log('ERROR in apiFIFOEmployeeReport: ' + error.message + '\nStack: ' + error.stack);
    throw new Error('Failed to generate FIFO report: ' + error.message);
  }
}

/**
 * TEST FUNCTION: Run FIFO integrity check
 */
function testFIFOIntegrityCheck() {
  Logger.log('=== Testing FIFO Integrity Check ===');
  
  // Test on a specific employee (replace with actual employee ID)
  const testEmployeeId = 'EMP001';
  
  try {
    const report = apiFIFOIntegrityCheck(testEmployeeId);
    Logger.log('Check completed: ' + report.summary);
    Logger.log('Total discrepancies: ' + report.totalDiscrepancies);
    
    if (report.integrityIssues.length > 0) {
      Logger.log('Issues found:');
      report.integrityIssues.forEach(issue => {
        Logger.log(`Employee ${issue.employeeId}: ${issue.issues.length} issue(s)`);
        issue.issues.forEach(i => {
          Logger.log(`  - ${i.type}: ${i.message}`);
        });
      });
    }
  } catch (error) {
    Logger.log('Test failed: ' + error.message);
  }
}



/**
 * =================================================================
 * NEW FUNCTIONS - ADD THIS ENTIRE BLOCK TO THE END OF YOUR CODE.GS
 * =================================================================
 */

// !!! IMPORTANT: REPLACE 'YOUR_TEMPLATE_ID_HERE' WITH YOUR ACTUAL GOOGLE DOC TEMPLATE ID
const COC_CERTIFICATE_TEMPLATE_ID = 'YOUR_TEMPLATE_ID_HERE'; 

// -----------------------------------------------------------------------------
// API: List & Record COC Entries
// -----------------------------------------------------------------------------

/**
 * Lists all recorded COC entries for a given employee and month.
 * This is called by loadExistingRecords() in the HTML.
 */
function apiListCOCRecordsForMonth(employeeId, month, year) {
  if (!employeeId || !month || !year) throw new Error("Employee ID, month, and year are required.");

  try {
    const monthYear = `${year}-${String(month).padStart(2, '0')}`;
    const recordsData = getSheetDataNoHeader('COC_Records');
    const certsData = getSheetDataNoHeader('COC_Certificates');

    Logger.log(`apiListCOCRecordsForMonth: Searching for empId="${employeeId}", month=${month}, year=${year}, monthYear="${monthYear}"`);
    Logger.log(`Total records in sheet: ${recordsData.length}`);

    // Create a map of Certificate IDs to their URLs for easy lookup
    const certMap = new Map();
    certsData.forEach(cert => {
      certMap.set(cert[CERT_COLS.CERTIFICATE_ID], {
        url: cert[CERT_COLS.CERTIFICATE_URL],
        pdf: cert[CERT_COLS.PDF_URL]
      });
    });

    // Filter records - EXCLUDE CANCELLED status
    const employeeMonthRecords = recordsData.filter(r => {
      const rowEmpId = String(r[RECORD_COLS.EMPLOYEE_ID] || '').trim();
      const rowMonthYear = String(r[RECORD_COLS.MONTH_YEAR] || '').trim();
      const rowStatus = String(r[RECORD_COLS.STATUS] || '').trim();

      const empMatch = rowEmpId === employeeId;
      const monthMatch = rowMonthYear === monthYear;
      const notCancelled = rowStatus !== STATUS_CANCELLED;

      if (empMatch) {
        Logger.log(`  Row: empId="${rowEmpId}", monthYear="${rowMonthYear}", status="${rowStatus}", matches=${empMatch && monthMatch && notCancelled}`);
      }

      return empMatch && monthMatch && notCancelled;
    });

    Logger.log(`Found ${employeeMonthRecords.length} matching records`);

    // Format the records for the client
    const formattedRecords = employeeMonthRecords.map(r => {
      const recordId = r[RECORD_COLS.RECORD_ID];
      const certificateId = r[RECORD_COLS.CERTIFICATE_ID];
      const certInfo = certMap.get(certificateId);

      return {
        recordId: recordId,
        displayDate: Utilities.formatDate(new Date(r[RECORD_COLS.DATE_RENDERED]), "GMT+8", "MMM dd, yyyy"),
        dayType: r[RECORD_COLS.DAY_TYPE],
        hoursWorked: parseFloat(r[RECORD_COLS.HOURS_WORKED] || 0),
        cocEarned: parseFloat(r[RECORD_COLS.COC_EARNED] || 0),
        certificateId: certificateId,
        certificateUrl: certInfo ? certInfo.url : null,
        pdfUrl: certInfo ? (certInfo.pdf || certInfo.url) : null // Fallback pdf to url
      };
    });

    return formattedRecords;

  } catch (e) {
    Logger.log(`Error in apiListCOCRecordsForMonth: ${e}`);
    throw new Error(`Failed to list COC records: ${e.message}`);
  }
}

/**
 * Records new COC entries from the form submission.
 * This just saves the records as "Pending".
 * The "Generate Certificate" step will finalize them.
 */
function apiRecordCOC(employeeId, month, year, entries) {
  if (!employeeId || !month || !year) throw new Error("Employee ID, month, and year are required.");
  if (!entries || entries.length === 0) throw new Error("At least one entry is required.");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Wait 30 seconds

  try {
    const monthYear = `${year}-${String(month).padStart(2, '0')}`;
    const currentUser = getCurrentUserEmail();
    const now = new Date();
    
    // 1. Get Employee Details
    const empDetails = getEmployeeDetails(employeeId);
    if (!empDetails) throw new Error("Employee not found.");

    let totalNewCOC = 0;
    let addedCount = 0;
    const newRecordRows = [];

    // 2. Process each entry
    for (const entry of entries) {
      // Server-side validation and calculation
      const result = apiCalculateOvertimeForDate(year, month, entry.day, entry.amIn, entry.amOut, entry.pmIn, entry.pmOut);
      
      if (result.cocEarned > 0) {
        const recordId = generateUniqueId("COC-");
        const dateRendered = new Date(year, month - 1, entry.day);

        const newRow = new Array(22).fill(''); // Initialize empty row
        newRow[RECORD_COLS.RECORD_ID] = recordId;
        newRow[RECORD_COLS.EMPLOYEE_ID] = employeeId;
        newRow[RECORD_COLS.EMPLOYEE_NAME] = empDetails.fullName;
        newRow[RECORD_COLS.MONTH_YEAR] = monthYear;
        newRow[RECORD_COLS.DATE_RENDERED] = dateRendered;
        newRow[RECORD_COLS.DAY_TYPE] = result.dayType;
        newRow[RECORD_COLS.AM_IN] = entry.amIn;
        newRow[RECORD_COLS.AM_OUT] = entry.amOut;
        newRow[RECORD_COLS.PM_IN] = entry.pmIn;
        newRow[RECORD_COLS.PM_OUT] = entry.pmOut;
        newRow[RECORD_COLS.HOURS_WORKED] = result.hoursWorked;
        newRow[RECORD_COLS.MULTIPLIER] = result.multiplier;
        newRow[RECORD_COLS.COC_EARNED] = result.cocEarned;
        newRow[RECORD_COLS.CERTIFICATE_ID] = ''; // Will be set upon generation
        newRow[RECORD_COLS.DATE_RECORDED] = now;
        newRow[RECORD_COLS.EXPIRATION_DATE] = ''; // Will be set upon generation
        newRow[RECORD_COLS.STATUS] = STATUS_PENDING; // Set to Pending
        newRow[RECORD_COLS.APPROVED_BY] = ''; // Assumes an approval step, or can be auto-approved
        newRow[RECORD_COLS.APPROVED_DATE] = '';
        newRow[RECORD_COLS.CREATED_BY] = currentUser;
        newRow[RECORD_COLS.LAST_MODIFIED] = now;
        newRow[RECORD_COLS.MODIFIED_BY] = currentUser;

        newRecordRows.push(newRow);
        totalNewCOC += result.cocEarned;
        addedCount++;
      }
    }

    if (newRecordRows.length > 0) {
      const recordsSheet = SpreadsheetApp.openById(DATABASE_ID).getSheetByName('COC_Records');
      recordsSheet.getRange(recordsSheet.getLastRow() + 1, 1, newRecordRows.length, newRecordRows[0].length).setValues(newRecordRows);
    } else {
      throw new Error("No valid COC entries to record.");
    }

    return {
      added: addedCount,
      totalNewCOC: totalNewCOC
    };

  } catch (e) {
    Logger.log(`Error in apiRecordCOC: ${e}`);
    throw new Error(`Failed to record COC entries: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------------------------------------------------------
// API: Certificate Generation (NEW)
// -----------------------------------------------------------------------------

/**
 * Generates a single monthly certificate for all uncertificated records.
 * This is the new function called by the "Generate" button.
 */
function apiGenerateMonthlyCOCCertificate(employeeId, month, year, issueDateString) {
  if (!employeeId || !month || !year || !issueDateString) {
    throw new Error("Employee ID, month, year, and issue date are required.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const db = SpreadsheetApp.openById(DATABASE_ID);
    const recordsSheet = db.getSheetByName('COC_Records');
    const recordsData = recordsSheet.getDataRange().getValues();
    const certSheet = db.getSheetByName('COC_Certificates');
    const detailSheet = db.getSheetByName('COC_Balance_Detail');
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    
    const settings = getSheetDataNoHeader('Settings');
    const validityMonths = parseInt(settings.find(r => r[0] === 'COC_VALIDITY_MONTHS')[1] || 12);

    const currentUser = getCurrentUserEmail();
    const now = new Date();
    const monthYear = `${year}-${String(month).padStart(2, '0')}`;
    
    // 1. Get Employee Details
    const empDetails = getEmployeeDetails(employeeId);
    if (!empDetails) throw new Error("Employee not found.");

    // 2. Find records to certify
    const recordsToCertifyIndices = []; // Store row indices (1-based)
    const recordsToCertifyData = [];
    
    // Start from row 2 (index 1) to skip header
    for (let i = 1; i < recordsData.length; i++) {
      const row = recordsData[i];
      if (row[RECORD_COLS.EMPLOYEE_ID] === employeeId &&
          row[RECORD_COLS.MONTH_YEAR] === monthYear &&
          (row[RECORD_COLS.STATUS] === STATUS_PENDING || row[RECORD_COLS.CERTIFICATE_ID] === '')) {
        recordsToCertifyIndices.push(i + 1); // Store 1-based index
        recordsToCertifyData.push(row);
      }
    }

    if (recordsToCertifyData.length === 0) {
      throw new Error("No pending COC records found for this employee and month to certify.");
    }

    // 3. Calculate totals and dates
    const totalNewCOC = recordsToCertifyData.reduce((sum, r) => sum + parseFloat(r[RECORD_COLS.COC_EARNED] || 0), 0);
    const numRecords = recordsToCertifyData.length;
    
    const issueDate = new Date(issueDateString);
    const expirationDate = new Date(issueDate);
    expirationDate.setFullYear(expirationDate.getFullYear() + validityMonths);
    expirationDate.setDate(expirationDate.getDate() - 1);

    const certificateId = generateUniqueId("CERT-");

    // 4. Generate the Google Doc
    const doc = generateCertificateDocument(certificateId, empDetails, recordsToCertifyData, issueDate, expirationDate);

    // 5. Create new rows for other sheets
    const newCertRow = new Array(13).fill('');
    newCertRow[CERT_COLS.CERTIFICATE_ID] = certificateId;
    newCertRow[CERT_COLS.EMPLOYEE_ID] = employeeId;
    newCertRow[CERT_COLS.EMPLOYEE_NAME] = empDetails.fullName;
    newCertRow[CERT_COLS.MONTH_YEAR] = monthYear;
    newCertRow[CERT_COLS.TOTAL_COC_EARNED] = totalNewCOC;
    newCertRow[CERT_COLS.NUMBER_OF_RECORDS] = numRecords;
    newCertRow[CERT_COLS.ISSUE_DATE] = issueDate;
    newCertRow[CERT_COLS.EXPIRATION_DATE] = expirationDate;
    newCertRow[CERT_COLS.CERTIFICATE_URL] = doc.url;
    newCertRow[CERT_COLS.PDF_URL] = doc.pdfUrl;
    newCertRow[CERT_COLS.STATUS] = STATUS_ACTIVE;
    newCertRow[CERT_COLS.CREATED_DATE] = now;
    newCertRow[CERT_COLS.CREATED_BY] = currentUser;
    certSheet.appendRow(newCertRow);

    const newDetailRows = [];
    recordsToCertifyData.forEach(record => {
      const newDetailRow = new Array(19).fill('');
      newDetailRow[DETAIL_COLS.ENTRY_ID] = generateUniqueId("COCD-");
      newDetailRow[DETAIL_COLS.EMPLOYEE_ID] = employeeId;
      newDetailRow[DETAIL_COLS.EMPLOYEE_NAME] = empDetails.fullName;
      newDetailRow[DETAIL_COLS.CERTIFICATE_ID] = certificateId;
      newDetailRow[DETAIL_COLS.RECORD_ID] = record[RECORD_COLS.RECORD_ID];
      newDetailRow[DETAIL_COLS.MONTH_YEAR] = monthYear;
      newDetailRow[DETAIL_COLS.DATE_EARNED] = record[RECORD_COLS.DATE_RENDERED];
      newDetailRow[DETAIL_COLS.DAY_TYPE] = record[RECORD_COLS.DAY_TYPE];
      newDetailRow[DETAIL_COLS.HOURS_EARNED] = record[RECORD_COLS.COC_EARNED];
      newDetailRow[DETAIL_COLS.HOURS_USED] = 0;
      newDetailRow[DETAIL_COLS.HOURS_REMAINING] = record[RECORD_COLS.COC_EARNED];
      newDetailRow[DETAIL_COLS.CERTIFICATE_ISSUE_DATE] = issueDate;
      newDetailRow[DETAIL_COLS.EXPIRATION_DATE] = expirationDate;
      newDetailRow[DETAIL_COLS.STATUS] = STATUS_ACTIVE;
      newDetailRow[DETAIL_COLS.DATE_CREATED] = record[RECORD_COLS.DATE_RECORDED];
      newDetailRow[DETAIL_COLS.CREATED_BY] = record[RECORD_COLS.CREATED_BY];
      newDetailRow[DETAIL_COLS.LAST_UPDATED] = now;
      newDetailRow[DETAIL_COLS.NOTES] = `Part of ${monthYear} certificate. Expires ${Utilities.formatDate(expirationDate, "GMT+8", "yyyy-MM-dd")}.`;
      
      newDetailRows.push(newDetailRow);
    });
    if (newDetailRows.length > 0) {
      detailSheet.getRange(detailSheet.getLastRow() + 1, 1, newDetailRows.length, newDetailRows[0].length).setValues(newDetailRows);
    }

    // 6. Create Ledger Entry
    const currentBalance = apiGetBalance(employeeId);
    const newBalance = currentBalance + totalNewCOC;
    const newLedgerRow = new Array(19).fill('');
    newLedgerRow[LEDGER_COLS.LEDGER_ID] = generateUniqueId("LED-");
    newLedgerRow[LEDGER_COLS.EMPLOYEE_ID] = employeeId;
    newLedgerRow[LEDGER_COLS.EMPLOYEE_NAME] = empDetails.fullName;
    newLedgerRow[LEDGER_COLS.TRANSACTION_DATE] = issueDate;
    newLedgerRow[LEDGER_COLS.TRANSACTION_TYPE] = TR_TYPE_EARNED;
    newLedgerRow[LEDGER_COLS.REFERENCE_ID] = certificateId;
    newLedgerRow[LEDGER_COLS.BALANCE_BEFORE] = currentBalance;
    newLedgerRow[LEDGER_COLS.COC_EARNED] = totalNewCOC;
    newLedgerRow[LEDGER_COLS.CTO_USED] = 0;
    newLedgerRow[LEDGER_COLS.COC_EXPIRED] = 0;
    newLedgerRow[LEDGER_COLS.BALANCE_ADJUSTMENT] = 0;
    newLedgerRow[LEDGER_COLS.BALANCE_AFTER] = newBalance;
    newLedgerRow[LEDGER_COLS.MONTH_YEAR_EARNED] = monthYear;
    newLedgerRow[LEDGER_COLS.EXPIRATION_DATE] = expirationDate;
    newLedgerRow[LEDGER_COLS.PROCESSED_BY] = currentUser;
    newLedgerRow[LEDGER_COLS.PROCESSED_DATE] = now;
    newLedgerRow[LEDGER_COLS.REMARKS] = `Certificate ${certificateId} issued for ${totalNewCOC} hrs.`;
    ledgerSheet.appendRow(newLedgerRow);

    // 7. Update original COC_Records
    recordsToCertifyIndices.forEach(rowIndex => {
      recordsSheet.getRange(rowIndex, RECORD_COLS.CERTIFICATE_ID + 1).setValue(certificateId);
      recordsSheet.getRange(rowIndex, RECORD_COLS.EXPIRATION_DATE + 1).setValue(expirationDate);
      recordsSheet.getRange(rowIndex, RECORD_COLS.STATUS + 1).setValue(STATUS_ACTIVE);
      recordsSheet.getRange(rowIndex, RECORD_COLS.LAST_MODIFIED + 1).setValue(now);
      recordsSheet.getRange(rowIndex, RECORD_COLS.MODIFIED_BY + 1).setValue(currentUser);
    });

    // 8. Return success
    return {
      certificateId: certificateId,
      numRecords: numRecords,
      totalHours: totalNewCOC
    };

  } catch (e) {
    Logger.log(`Error in apiGenerateMonthlyCOCCertificate: ${e}\nStack: ${e.stack}`);
    throw new Error(`Failed to generate certificate: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper function to create the Google Doc certificate.
 */
function generateCertificateDocument(certificateId, empDetails, records, issueDate, expirationDate) {
  try {
    const templateId = COC_CERTIFICATE_TEMPLATE_ID;
    if (templateId === '114b0thKcB-TCO6Zc53e5OgNpWNI5WGWbf2aNDmZjlNE') {
      throw new Error("Invalid Certificate Template ID. Please update COC_CERTIFICATE_TEMPLATE_ID in Code.gs.");
    }
    const templateFile = DriveApp.getFileById(templateId);
    
    // Create a copy
    const newFileName = `COC Certificate - ${empDetails.fullName} - ${certificateId}`;
    const newFile = templateFile.makeCopy(newFileName);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();

    // Prepare data
    const totalCOC = records.reduce((sum, r) => sum + parseFloat(r[RECORD_COLS.COC_EARNED] || 0), 0);
    const inclusiveDates = formatInclusiveDates(records.map(r => new Date(r[RECORD_COLS.DATE_RENDERED])));
    const issueDateFormatted = formatLongDate(issueDate);
    const expiryDateFormatted = formatLongDate(expirationDate);

    // *** MODIFIED: Get signatories dynamically from "Signatories" sheet ***
    const settings = getSheetDataNoHeader('Signatories'); // Changed from 'Settings'
    const getSetting = (key) => {
      const row = settings.find(r => r[0] === key);
      return row ? row[1] : `[${key} not found]`;
    };

    const issuedByName = getSetting("SIGNATORY_ISSUED_BY_NAME");
    const issuedByPosition = getSetting("SIGNATORY_ISSUED_BY_POSITION");
    // REMOVED: notedByName and notedByPosition
    // *** END OF MODIFICATION ***

    // Replace placeholders
    body.replaceText("{{EMPLOYEE_NAME}}", empDetails.fullName.toUpperCase());
    body.replaceText("{{POSITION}}", empDetails.position);
    body.replaceText("{{OFFICE}}", empDetails.office);
    body.replaceText("{{TOTAL_COC}}", totalCOC.toFixed(1)); // Use 1 decimal place
    body.replaceText("{{INCLUSIVE_DATES}}", inclusiveDates);
    body.replaceText("{{ISSUE_DATE_FORMATTED}}", issueDateFormatted);
    body.replaceText("{{EXPIRY_DATE_FORMATTED}}", expiryDateFormatted);
    body.replaceText("{{ISSUED_BY_NAME}}", issuedByName.toUpperCase());
    body.replaceText("{{ISSUED_BY_POSITION}}", issuedByPosition);
    // REMOVED: notedBy replaceText calls
    body.replaceText("{{CERTIFICATE_ID}}", certificateId); // Add this if you have a placeholder for it

    doc.saveAndClose();

    // Set permissions to "anyone with link can view"
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const pdfBlob = newFile.getAs('application/pdf');
    const pdfFile = DriveApp.createFile(pdfBlob).setName(newFileName + ".pdf");
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);


    return {
      url: newFile.getUrl(),
      pdfUrl: pdfFile.getUrl()
    };

  } catch (e) {
    Logger.log(`Error in generateCertificateDocument: ${e}`);
    throw new Error(`Failed to create document: ${e.message}`);
  }
}

// -----------------------------------------------------------------------------
// HELPER FUNCTIONS (Add these if they don't exist)
// -----------------------------------------------------------------------------

/**
 * Gets full employee details from the 'Employees' sheet.
 * @param {string} employeeId The Employee ID.
 * @returns {object} An object with fullName, position, and office.
 */
function getEmployeeDetails(employeeId) {
  const data = getSheetDataNoHeader('Employees');
  const empRow = data.find(r => r[EMP_COLS.EMPLOYEE_ID] === employeeId);
  if (!empRow) return null;
  
  const middle = empRow[EMP_COLS.MIDDLE_INITIAL] ? ` ${empRow[EMP_COLS.MIDDLE_INITIAL]}.` : '';
  const suffix = empRow[EMP_COLS.SUFFIX] ? ` ${empRow[EMP_COLS.SUFFIX]}` : '';
  
  return {
    fullName: `${empRow[EMP_COLS.FIRST_NAME]}${middle} ${empRow[EMP_COLS.LAST_NAME]}${suffix}`,
    position: empRow[EMP_COLS.POSITION],
    office: empRow[EMP_COLS.OFFICE]
  };
}

/**
 * Formats a Date object into "Month Day, Year" (e.g., "October 27, 2025").
 * @param {Date} date The date object.
 * @returns {string} The formatted date string.
 */
function formatLongDate(date) {
  return Utilities.formatDate(date, "GMT+8", "MMMM dd, yyyy");
}

/**
 * Formats an array of dates into a string like "Month day1, day2, and day3, Year".
 * @param {Array<Date>} dateArray An array of Date objects.
 * @returns {string} The formatted inclusive dates string.
 */
function formatInclusiveDates(dateArray) {
  if (!dateArray || dateArray.length === 0) return "";

  // Sort dates
  dateArray.sort((a, b) => a.getTime() - b.getTime());

  // Get unique days
  const days = [...new Set(dateArray.map(d => d.getDate()))];
  
  const month = Utilities.formatDate(dateArray[0], "GMT+8", "MMMM");
  const year = Utilities.formatDate(dateArray[0], "GMT+8", "yyyy");

  let dayString = "";
  if (days.length === 1) {
    dayString = days[0];
  } else if (days.length === 2) {
    dayString = `${days[0]} and ${days[1]}`;
  } else {
    dayString = `${days.slice(0, -1).join(", ")}, and ${days[days.length - 1]}`;
  }

  return `${month} ${dayString}, ${year}`;
}


/**
 * =================================================================
 * NEW SETTINGS FUNCTIONS - ADD THESE TO YOUR CODE.GS
 * =================================================================
 */

/**
 * Gets the current signatories from the Settings sheet.
 */
function apiGetSignatories() {
  try {
    const settings = getSheetDataNoHeader('Signatories'); // Changed from 'Settings'
    const getSetting = (key, defaultValue = "") => {
      const row = settings.find(r => r[0] === key);
      return row ? row[1] : defaultValue;
    };

    return {
      issuedByName: getSetting("SIGNATORY_ISSUED_BY_NAME", "NIDA O. TRINIDAD"),
      issuedByPosition: getSetting("SIGNATORY_ISSUED_BY_POSITION", "Administrative Officer V")
      // REMOVED: notedByName and notedByPosition
    };
  } catch (e) {
    Logger.log(`Error in apiGetSignatories: ${e}`);
    throw new Error(`Failed to get signatories: ${e.message}`);
  }
}

/**
 * Saves the signatories to the Settings sheet.
 */
function apiSaveSignatories(signatories) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const db = SpreadsheetApp.openById(DATABASE_ID);
    const sheet = db.getSheetByName('Signatories'); // Changed from 'Settings'
    const data = sheet.getDataRange().getValues();

    const settingsToUpdate = {
      "SIGNATORY_ISSUED_BY_NAME": signatories.issuedByName,
      "SIGNATORY_ISSUED_BY_POSITION": signatories.issuedByPosition
      // REMOVED: notedByName and notedByPosition
    };

    const keysToUpdate = Object.keys(settingsToUpdate);
    let updatedCount = 0;

    // Update existing keys
    for (let i = 1; i < data.length; i++) { // Start from row 2 (index 1)
      const key = data[i][0];
      if (keysToUpdate.includes(key)) {
        sheet.getRange(i + 1, 2).setValue(settingsToUpdate[key]); // Column 2 (B) is the value
        sheet.getRange(i + 1, 4).setValue(new Date()); // Column 4 (D) is Last Updated
        sheet.getRange(i + 1, 5).setValue(getCurrentUserEmail()); // Column 5 (E) is Updated By
        
        // Remove key from list once updated
        keysToUpdate.splice(keysToUpdate.indexOf(key), 1);
        updatedCount++;
      }
    }

    // Add new keys if they didn't exist
    const newRows = [];
    for (const key of keysToUpdate) {
      newRows.push([
        key,
        settingsToUpdate[key],
        `Signatory for COC Certificates`, // Description
        new Date(), // Last Updated
        getCurrentUserEmail() // Updated By
      ]);
    }

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
    
    return { success: true, updated: updatedCount, added: newRows.length };

  } catch (e) {
    Logger.log(`Error in apiSaveSignatories: ${e}`);
    throw new Error(`Failed to save signatories: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------------------------------------------------------
// API: Delete COC Record
// -----------------------------------------------------------------------------

/**
 * Deletes (marks as Cancelled) a COC record.
 * This is used for pending records that haven't been certificated yet.
 *
 * @param {string} recordId - The record ID to delete
 * @param {string} reason - Reason for deletion
 * @returns {Object} Result object
 */
function apiDeleteCOCRecord(recordId, reason) {
  if (!recordId) throw new Error("Record ID is required.");
  if (!reason) throw new Error("Reason for deletion is required.");

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const db = SpreadsheetApp.openById(DATABASE_ID);
    const recordsSheet = db.getSheetByName('COC_Records');
    const data = recordsSheet.getDataRange().getValues();

    const currentUser = getCurrentUserEmail();
    const now = new Date();

    // Find the record
    let rowIndex = -1;
    let recordData = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][RECORD_COLS.RECORD_ID] === recordId) {
        rowIndex = i;
        recordData = data[i];
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error("Record not found.");
    }

    // Check if record is already certificated
    const certificateId = recordData[RECORD_COLS.CERTIFICATE_ID];
    if (certificateId) {
      throw new Error("Cannot delete a record that has already been certificated. Please contact administrator if you need to cancel a certificated record.");
    }

    // Check current status
    const currentStatus = recordData[RECORD_COLS.STATUS];
    if (currentStatus === STATUS_CANCELLED) {
      throw new Error("This record has already been cancelled.");
    }

    // Update the record to mark it as Cancelled
    // We don't actually delete it to maintain audit trail
    const row = rowIndex + 1; // Convert to 1-based index
    recordsSheet.getRange(row, RECORD_COLS.STATUS + 1).setValue(STATUS_CANCELLED);
    recordsSheet.getRange(row, RECORD_COLS.LAST_MODIFIED + 1).setValue(now);
    recordsSheet.getRange(row, RECORD_COLS.MODIFIED_BY + 1).setValue(currentUser);

    // Add a note/remark in a comments column if available, or we can add it to the APPROVED_BY column as a workaround
    // Since we don't have a dedicated "Remarks" column in COC_Records, we'll store it in the ledger

    // Create a ledger entry for this cancellation
    const ledgerSheet = db.getSheetByName('COC_Ledger');
    if (ledgerSheet) {
      const employeeId = recordData[RECORD_COLS.EMPLOYEE_ID];
      const employeeName = recordData[RECORD_COLS.EMPLOYEE_NAME];
      const cocEarned = parseFloat(recordData[RECORD_COLS.COC_EARNED] || 0);
      const monthYear = recordData[RECORD_COLS.MONTH_YEAR];

      const ledgerId = generateUniqueId("LDG-");
      const ledgerRow = new Array(15).fill('');
      ledgerRow[LEDGER_COLS.LEDGER_ID] = ledgerId;
      ledgerRow[LEDGER_COLS.EMPLOYEE_ID] = employeeId;
      ledgerRow[LEDGER_COLS.EMPLOYEE_NAME] = employeeName;
      ledgerRow[LEDGER_COLS.TRANSACTION_DATE] = now;
      ledgerRow[LEDGER_COLS.TRANSACTION_TYPE] = 'COC Cancelled';
      ledgerRow[LEDGER_COLS.REFERENCE_ID] = recordId;
      ledgerRow[LEDGER_COLS.BALANCE_BEFORE] = 0; // Since it was pending, it never affected balance
      ledgerRow[LEDGER_COLS.COC_EARNED] = 0;
      ledgerRow[LEDGER_COLS.CTO_USED] = 0;
      ledgerRow[LEDGER_COLS.COC_EXPIRED] = 0;
      ledgerRow[LEDGER_COLS.BALANCE_ADJUSTMENT] = 0;
      ledgerRow[LEDGER_COLS.BALANCE_AFTER] = 0;
      ledgerRow[LEDGER_COLS.MONTH_YEAR_EARNED] = monthYear;
      ledgerRow[LEDGER_COLS.PROCESSED_BY] = currentUser;
      ledgerRow[14] = `Cancelled pending COC record (${cocEarned.toFixed(2)} hrs). Reason: ${reason}`; // Remarks column

      ledgerSheet.getRange(ledgerSheet.getLastRow() + 1, 1, 1, ledgerRow.length).setValues([ledgerRow]);
    }

    Logger.log(`COC Record ${recordId} cancelled by ${currentUser}. Reason: ${reason}`);

    return {
      success: true,
      message: "Record cancelled successfully."
    };

  } catch (e) {
    Logger.log(`Error in apiDeleteCOCRecord: ${e}`);
    throw new Error(`Failed to delete record: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}





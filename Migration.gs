/**
 * MIGRATION SCRIPT - Add this to your Code.gs file
 * 
 * This script migrates COC data from COC_Records and COC_Ledger 
 * into the COC_Balance_Detail sheet for FIFO tracking.
 * 
 * RUN THESE FUNCTIONS IN ORDER:
 * 1. runMigrationStep1_Initialize()
 * 2. runMigrationStep2_MigrateInitialBalances()
 * 3. runMigrationStep3_MigrateCOCRecords()
 */

/**
 * STEP 1: Initialize COC_Balance_Detail sheet
 * This ensures the sheet exists with proper headers
 */
function runMigrationStep1_Initialize() {
  Logger.log('=== STEP 1: Initializing COC_Balance_Detail ===');
  
  try {
    const result = apiInitializeCOCBalanceDetail();
    Logger.log('✓ Success: ' + result.message);
    return result;
  } catch (error) {
    Logger.log('✗ Error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * STEP 2: Migrate Initial Balances from COC_Ledger
 * This migrates all "Initial Balance" entries
 */
function runMigrationStep2_MigrateInitialBalances() {
  Logger.log('=== STEP 2: Migrating Initial Balances ===');
  
  try {
    const result = apiMigrateExistingInitialBalances();
    Logger.log('✓ Success: Migrated ' + result.migratedCount + ' initial balance entries');
    return result;
  } catch (error) {
    Logger.log('✗ Error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * STEP 3: Migrate Active COC_Records to COC_Balance_Detail
 * This is the main migration that will fix Juan's issue
 */
function runMigrationStep3_MigrateCOCRecords() {
  Logger.log('=== STEP 3: Migrating COC_Records ===');
  
  const db = getDatabase();
  const recordsSheet = db.getSheetByName('COC_Records');
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  
  if (!recordsSheet) {
    Logger.log('✗ ERROR: COC_Records sheet not found');
    return { success: false, error: 'COC_Records sheet not found' };
  }
  
  if (!detailSheet) {
    Logger.log('✗ ERROR: COC_Balance_Detail sheet not found. Run Step 1 first.');
    return { success: false, error: 'COC_Balance_Detail sheet not found' };
  }
  
  const recordsData = recordsSheet.getDataRange().getValues();
  const detailRows = [];
  const TIME_ZONE = getScriptTimeZone();
  
  let migratedCount = 0;
  let skippedCount = 0;
  
  // Check existing entries to avoid duplicates
  const existingDetailData = detailSheet.getDataRange().getValues();
  const existingRecordIds = new Set();
  
  Logger.log('Checking existing COC_Balance_Detail entries...');
  for (let i = 1; i < existingDetailData.length; i++) {
    const recordId = existingDetailData[i][3]; // Record ID column (index 3)
    if (recordId) {
      existingRecordIds.add(recordId);
    }
  }
  Logger.log('Found ' + existingRecordIds.size + ' existing entries in COC_Balance_Detail');
  
  // Process each COC_Records entry
  Logger.log('Processing COC_Records...');
  for (let i = 1; i < recordsData.length; i++) {
    const row = recordsData[i];
    
    // Column mapping based on COC_Records structure:
    // 0:Record ID, 1:Employee ID, 2:Employee Name, 3:Month-Year, 4:Date Rendered
    // 5:Day Type, 6:AM In, 7:AM Out, 8:PM In, 9:PM Out, 10:Hours Worked
    // 11:COC Multiplier, 12:COC Earned, 13:Date Recorded, 14:Expiration Date, 15:Status
    
    const recordId = row[0];
    const employeeId = row[1];
    const employeeName = row[2];
    const dateRendered = new Date(row[4]);
    const cocEarned = parseFloat(row[12]) || 0;
    const expirationDate = row[14];
    const status = row[15];
    
    // Skip conditions
    if (!recordId) {
      skippedCount++;
      continue;
    }
    
    if (existingRecordIds.has(recordId)) {
      Logger.log('Skipping ' + recordId + ' - already exists in COC_Balance_Detail');
      skippedCount++;
      continue;
    }
    
    if (cocEarned <= 0) {
      Logger.log('Skipping ' + recordId + ' - no COC hours earned');
      skippedCount++;
      continue;
    }
    
    if (status !== 'Active') {
      Logger.log('Skipping ' + recordId + ' - status is ' + status);
      skippedCount++;
      continue;
    }
    
    // Calculate expiration date if not set (default: 1 year from date earned)
    let expDate = expirationDate;
    if (!expDate || !(expDate instanceof Date)) {
      expDate = new Date(dateRendered.getTime() + 365 * 24 * 60 * 60 * 1000);
    }
    
    // Create detail entry
    // COC_Balance_Detail columns:
    // 0:Entry ID, 1:Employee ID, 2:Employee Name, 3:Record ID, 4:Date Earned
    // 5:Hours Earned, 6:Hours Remaining, 7:Expiration Date, 8:Status
    // 9:Last Updated, 10:Notes
    
    detailRows.push([
      generateCOCDetailEntryId(),      // Entry ID
      employeeId,                       // Employee ID
      employeeName,                     // Employee Name
      recordId,                         // Record ID
      dateRendered,                     // Date Earned
      cocEarned,                        // Hours Earned
      cocEarned,                        // Hours Remaining (full amount for migration)
      expDate,                          // Expiration Date
      'Active',                         // Status
      new Date(),                       // Last Updated
      'Migrated from COC_Records on ' + Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss')
    ]);
    
    Logger.log('✓ Queued for migration: ' + recordId + ' (' + employeeId + ') - ' + cocEarned + ' hours');
    migratedCount++;
  }
  
  // Batch write all entries
  if (detailRows.length > 0) {
    Logger.log('Writing ' + detailRows.length + ' entries to COC_Balance_Detail...');
    const startRow = detailSheet.getLastRow() + 1;
    detailSheet.getRange(startRow, 1, detailRows.length, detailRows[0].length)
      .setValues(detailRows);
    Logger.log('✓ Successfully wrote all entries');
  } else {
    Logger.log('No new entries to migrate');
  }
  
  const summary = {
    success: true,
    migratedCount: migratedCount,
    skippedCount: skippedCount,
    message: 'Migrated ' + migratedCount + ' active COC records to COC_Balance_Detail. Skipped ' + skippedCount + ' records.'
  };
  
  Logger.log('=== MIGRATION COMPLETE ===');
  Logger.log('Migrated: ' + migratedCount);
  Logger.log('Skipped: ' + skippedCount);
  
  return summary;
}

/**
 * VERIFICATION: Check balance in both systems
 * Run this after migration to verify data integrity
 */
function verifyMigrationForEmployee(employeeId) {
  Logger.log('=== VERIFICATION FOR ' + employeeId + ' ===');
  
  const db = getDatabase();
  
  // Get balance from COC_Ledger (old system)
  const ledgerBalance = getCurrentCOCBalance(employeeId);
  Logger.log('COC_Ledger Balance: ' + ledgerBalance.toFixed(2) + ' hours');
  
  // Get balance from COC_Balance_Detail (new FIFO system)
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  let detailBalance = 0;
  
  if (detailSheet) {
    const data = detailSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === employeeId && row[8] === 'Active') {
        detailBalance += parseFloat(row[6]) || 0; // Hours Remaining
      }
    }
  }
  
  Logger.log('COC_Balance_Detail Balance: ' + detailBalance.toFixed(2) + ' hours');
  
  const difference = Math.abs(ledgerBalance - detailBalance);
  
  if (difference < 0.01) {
    Logger.log('✓ BALANCES MATCH - Migration successful!');
    return { success: true, balanced: true, ledgerBalance, detailBalance };
  } else {
    Logger.log('✗ WARNING: Balances do not match! Difference: ' + difference.toFixed(2) + ' hours');
    Logger.log('This may be due to expired COC not being migrated. Check manually.');
    return { success: true, balanced: false, ledgerBalance, detailBalance, difference };
  }
}

/**
 * CONVENIENCE FUNCTION: Run all migration steps
 * Use this to run all 3 steps in sequence
 */
function runCompleteMigration() {
  Logger.log('========================================');
  Logger.log('STARTING COMPLETE MIGRATION PROCESS');
  Logger.log('========================================\n');
  
  const step1 = runMigrationStep1_Initialize();
  if (!step1.success) {
    Logger.log('Migration aborted at Step 1');
    return step1;
  }
  
  Logger.log('\n');
  
  const step2 = runMigrationStep2_MigrateInitialBalances();
  if (!step2.success) {
    Logger.log('Migration aborted at Step 2');
    return step2;
  }
  
  Logger.log('\n');
  
  const step3 = runMigrationStep3_MigrateCOCRecords();
  if (!step3.success) {
    Logger.log('Migration aborted at Step 3');
    return step3;
  }
  
  Logger.log('\n========================================');
  Logger.log('MIGRATION COMPLETE!');
  Logger.log('Initial Balances Migrated: ' + step2.migratedCount);
  Logger.log('COC Records Migrated: ' + step3.migratedCount);
  Logger.log('========================================\n');
  
  // Verify Juan's balance
  Logger.log('Verifying Juan A Dela Cruz Jr. (EMP002)...\n');
  const verification = verifyMigrationForEmployee('EMP002');
  
  return {
    success: true,
    step1: step1,
    step2: step2,
    step3: step3,
    verification: verification
  };
}

/**
 * HELPER: Generate unique Entry ID for COC_Balance_Detail
 * This should already exist in your Code.gs, but included here for completeness
 */
function generateCOCDetailEntryId() {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 1000);
  return 'COCD-' + timestamp + random;
}

// ============================================================================
// INSTRUCTIONS FOR USE:
// ============================================================================
// 
// OPTION 1 - Run all at once:
//   1. Open Apps Script Editor (Extensions > Apps Script)
//   2. Paste this entire file into Code.gs (or create new file Migration.gs)
//   3. Click on "runCompleteMigration" in the function dropdown
//   4. Click "Run"
//   5. Check execution log (View > Logs)
//
// OPTION 2 - Run step by step:
//   1. Run "runMigrationStep1_Initialize"
//   2. Check log, then run "runMigrationStep2_MigrateInitialBalances"
//   3. Check log, then run "runMigrationStep3_MigrateCOCRecords"
//   4. Run "verifyMigrationForEmployee('EMP002')" to verify Juan's data
//
// AFTER MIGRATION:
//   - Test CTO application with Juan A Dela Cruz Jr.
//   - Verify balance shows 37.0 hours
//   - Submit 4-hour CTO application
//   - Should work without errors!
//
// ============================================================================



/**
 * BULK MIGRATION - Set Initial Balances with Expiration
 * Copy this to your Code.gs file
 */
function bulkMigrateInitialBalances() {
  // Prepare your migration data here
  const migrationData = [
    {
      employeeId: 'EMP001',
      balances: [
        { monthYear: '2024-12', amount: 15.0, note: '2024 COC total' },
        { monthYear: '2025-09', amount: 10.0, note: '2025 COC total' }
      ]
    },
    {
      employeeId: 'EMP002',
      balances: [
        { monthYear: '2024-12', amount: 14.0, note: '2024 COC total' },
        { monthYear: '2025-10', amount: 23.0, note: '2025 COC total' }
      ]
    }
    // Add more employees...
  ];
  
  const db = getDatabase();
  const employeesSheet = db.getSheetByName('Employees');
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  const recordsSheet = db.getSheetByName('COC_Records');
  
  if (!employeesSheet || !ledgerSheet || !detailSheet || !recordsSheet) {
    throw new Error('Required sheets not found');
  }
  
  const TIME_ZONE = getScriptTimeZone();
  let successCount = 0;
  let errorCount = 0;
  const errors = [];
  
  for (const empData of migrationData) {
    try {
      const employee = getEmployeeById(empData.employeeId);
      if (!employee) {
        errors.push(`Employee ${empData.employeeId} not found`);
        errorCount++;
        continue;
      }
      
      let totalInitialBalance = 0;
      let earliestDate = null;
      
      // Process each balance entry
      for (const balance of empData.balances) {
        // Parse month-year
        const [year, month] = balance.monthYear.split('-').map(Number);
        const earnedDate = new Date(year, month - 1, 1);
        
        // Calculate expiration: end of succeeding year
        const expirationDate = new Date(year + 1, 11, 31, 23, 59, 59);
        
        // Track earliest date for employee record
        if (!earliestDate || earnedDate < earliestDate) {
          earliestDate = earnedDate;
        }
        
        totalInitialBalance += balance.amount;
        
        // Generate IDs
        const recordId = 'INIT-' + empData.employeeId + '-' + balance.monthYear;
        const ledgerId = generateLedgerId();
        const detailId = generateCOCDetailEntryId();
        
        // 1. Add to COC_Records
        recordsSheet.appendRow([
          recordId,                               // Record ID
          empData.employeeId,                     // Employee ID
          employee.fullName,                      // Employee Name
          balance.monthYear,                      // Month-Year
          earnedDate,                             // Date Rendered
          'Initial Balance',                      // Day Type
          '',                                     // AM In
          '',                                     // AM Out
          '',                                     // PM In
          '',                                     // PM Out
          balance.amount,                         // Hours Worked
          1.0,                                    // COC Multiplier
          balance.amount,                         // COC Earned
          new Date(),                             // Date Recorded
          expirationDate,                         // Expiration Date
          'Active'                                // Status
        ]);
        
        // 2. Add to COC_Ledger
        ledgerSheet.appendRow([
          ledgerId,                               // Ledger ID
          empData.employeeId,                     // Employee ID
          employee.fullName,                      // Employee Name
          new Date(),                             // Transaction Date
          'Initial Balance',                      // Transaction Type
          recordId,                               // Reference ID
          balance.amount,                         // COC Earned
          0,                                      // CTO Used
          totalInitialBalance,                    // Running Balance (cumulative)
          balance.monthYear,                      // Month-Year Earned
          Utilities.formatDate(expirationDate, TIME_ZONE, 'yyyy-MM-dd'), // Expiration
          Session.getActiveUser().getEmail(),    // Processed By
          'Initial balance migration: ' + (balance.note || '')  // Remarks
        ]);
        
        // 3. Add to COC_Balance_Detail (for FIFO)
        detailSheet.appendRow([
          detailId,                               // Entry ID
          empData.employeeId,                     // Employee ID
          employee.fullName,                      // Employee Name
          recordId,                               // Record ID
          earnedDate,                             // Date Earned
          balance.amount,                         // Hours Earned
          balance.amount,                         // Hours Remaining
          expirationDate,                         // Expiration Date
          'Active',                               // Status
          new Date(),                             // Last Updated
          'Initial balance migration: ' + (balance.note || '')  // Notes
        ]);
      }
      
      // 4. Update Employees sheet with total initial balance and earliest date
      const empData2Update = employeesSheet.getDataRange().getValues();
      for (let i = 1; i < empData2Update.length; i++) {
        if (empData2Update[i][0] === empData.employeeId) {
          employeesSheet.getRange(i + 1, 10).setValue(totalInitialBalance); // Column J
          employeesSheet.getRange(i + 1, 11).setValue(earliestDate); // Column K (if exists)
          break;
        }
      }
      
      Logger.log(`✓ Migrated ${empData.employeeId}: ${totalInitialBalance} hours`);
      successCount++;
      
    } catch (error) {
      Logger.log(`✗ Error migrating ${empData.employeeId}: ${error.message}`);
      errors.push(`${empData.employeeId}: ${error.message}`);
      errorCount++;
    }
  }
  
  // Summary
  const summary = {
    success: successCount,
    errors: errorCount,
    errorDetails: errors,
    message: `Migration complete. Success: ${successCount}, Errors: ${errorCount}`
  };
  
  Logger.log('=== MIGRATION SUMMARY ===');
  Logger.log(JSON.stringify(summary, null, 2));
  
  return summary;
}

/**
 * Migrate existing CTO applications
 */
function migrateExistingCTOs() {
  const existingCTOs = [
    {
      employeeId: 'EMP001',
      ctoId: 'CTO-MANUAL-001',
      hoursUsed: 8.0,
      startDate: '2024-03-20',
      endDate: '2024-03-20',
      dateApplied: '2024-03-15',
      status: 'Approved',
      remarks: 'Migrated from manual records'
    }
    // Add more...
  ];
  
  const db = getDatabase();
  const ctoSheet = db.getSheetByName('CTO_Applications');
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  
  for (const cto of existingCTOs) {
    const employee = getEmployeeById(cto.employeeId);
    if (!employee) continue;
    
    // 1. Add to CTO_Applications
    ctoSheet.appendRow([
      cto.ctoId,
      cto.employeeId,
      employee.fullName,
      employee.office,
      cto.hoursUsed,
      new Date(cto.startDate),
      new Date(cto.endDate),
      formatDate(new Date(cto.startDate)) + ' to ' + formatDate(new Date(cto.endDate)),
      0, // Balance before (unknown for migration)
      new Date(cto.dateApplied),
      cto.status,
      new Date(cto.dateApplied),
      cto.remarks
    ]);
    
    // 2. Deduct from COC_Balance_Detail using FIFO
    // This will automatically deduct from oldest entries
    const deductions = consumeCOCWithFIFO(cto.employeeId, cto.hoursUsed, cto.ctoId);
    
    // 3. Add ledger entry
    const ledgerId = generateLedgerId();
    const currentBalance = getCurrentCOCBalance(cto.employeeId);
    
    ledgerSheet.appendRow([
      ledgerId,
      cto.employeeId,
      employee.fullName,
      new Date(cto.dateApplied),
      'CTO Used',
      cto.ctoId,
      0,
      cto.hoursUsed,
      currentBalance,
      '',
      '',
      Session.getActiveUser().getEmail(),
      'Migrated CTO: ' + cto.remarks
    ]);
    
    Logger.log(`✓ Migrated CTO ${cto.ctoId} for ${cto.employeeId}`);
  }
  
  return { success: true, message: 'CTO migration complete' };
}

function checkWhoNeedsMigration() {
  const db = getDatabase();
  const employeesSheet = db.getSheetByName('Employees');
  const detailSheet = db.getSheetByName('COC_Balance_Detail');
  const ledgerSheet = db.getSheetByName('COC_Ledger');
  
  const employees = employeesSheet.getDataRange().getValues();
  const detailData = detailSheet.getDataRange().getValues();
  const ledgerData = ledgerSheet.getDataRange().getValues();
  
  // Build a set of employees who have entries in Detail
  const employeesInDetail = new Set();
  for (let i = 1; i < detailData.length; i++) {
    employeesInDetail.add(detailData[i][1]); // Employee ID
  }
  
  Logger.log('=== MIGRATION STATUS CHECK ===\n');
  
  const needsMigration = [];
  const alreadyMigrated = [];
  
  for (let i = 1; i < employees.length; i++) {
    const empId = employees[i][0];
    const empName = employees[i][2] + ' ' + employees[i][1]; // First + Last
    const status = employees[i][8];
    
    if (status !== 'Active') continue; // Skip inactive
    
    if (employeesInDetail.has(empId)) {
      alreadyMigrated.push(empId + ': ' + empName);
    } else {
      // Check if they have COC balance in ledger
      let hasBalance = false;
      for (let j = 1; j < ledgerData.length; j++) {
        if (ledgerData[j][1] === empId) {
          hasBalance = true;
          break;
        }
      }
      
      if (hasBalance) {
        needsMigration.push(empId + ': ' + empName);
      }
    }
  }
  
  Logger.log('✅ ALREADY MIGRATED (' + alreadyMigrated.length + '):');
  alreadyMigrated.forEach(emp => Logger.log('  ' + emp));
  
  Logger.log('\n❌ NEEDS MIGRATION (' + needsMigration.length + '):');
  needsMigration.forEach(emp => Logger.log('  ' + emp));
  
  return {
    alreadyMigrated: alreadyMigrated.length,
    needsMigration: needsMigration.length,
    needsMigrationList: needsMigration
  };
}

/**
 * MIGRATION: Add PDF URLs to existing COC_Certificates
 * This migration adds PDF export URLs for certificates that don't have them yet.
 * Run this after updating the certificate generation code to support PDFs.
 */
function runMigrationStep4_AddPDFUrlsToCertificates() {
  Logger.log('=== STEP 4: Adding PDF URLs to COC_Certificates ===');

  const db = getDatabase();
  const certSheet = db.getSheetByName('COC_Certificates');

  if (!certSheet) {
    Logger.log('✗ ERROR: COC_Certificates sheet not found');
    return { success: false, error: 'COC_Certificates sheet not found' };
  }

  const data = certSheet.getDataRange().getValues();
  const headers = data[0];

  // Check if PDF URL column already exists
  let pdfUrlColIndex = headers.indexOf('PDF URL');
  let issuedDateColIndex = headers.indexOf('Issued Date');

  if (pdfUrlColIndex === -1) {
    // Add PDF URL column after Certificate URL
    Logger.log('Adding PDF URL column to COC_Certificates sheet...');
    certSheet.insertColumnAfter(6); // Insert after column 6 (Certificate URL)
    certSheet.getRange(1, 8).setValue('PDF URL');
    pdfUrlColIndex = 7;
    issuedDateColIndex = 8; // Issued Date shifted one column
    Logger.log('✓ Added PDF URL column');
  }

  // Process each certificate
  let updatedCount = 0;
  let skippedCount = 0;

  Logger.log('Processing existing certificates...');
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const recordId = row[0];
    const certUrl = row[6]; // Certificate URL
    const existingPdfUrl = row[pdfUrlColIndex];

    // Skip if PDF URL already exists
    if (existingPdfUrl && existingPdfUrl !== '') {
      skippedCount++;
      continue;
    }

    // Extract document ID from certificate URL
    if (certUrl && certUrl.indexOf('/document/d/') !== -1) {
      const urlParts = certUrl.split('/document/d/');
      if (urlParts.length > 1) {
        const docId = urlParts[1].split('/')[0];
        const pdfUrl = 'https://docs.google.com/document/d/' + docId + '/export?format=pdf';

        // Update the cell with PDF URL
        certSheet.getRange(i + 1, pdfUrlColIndex + 1).setValue(pdfUrl);
        Logger.log('✓ Updated ' + recordId + ' with PDF URL');
        updatedCount++;
      }
    }
  }

  const summary = {
    success: true,
    updatedCount: updatedCount,
    skippedCount: skippedCount,
    message: 'Added PDF URLs to ' + updatedCount + ' certificates. Skipped ' + skippedCount + ' certificates.'
  };

  Logger.log('=== MIGRATION COMPLETE ===');
  Logger.log('Updated: ' + updatedCount);
  Logger.log('Skipped: ' + skippedCount);

  return summary;
}

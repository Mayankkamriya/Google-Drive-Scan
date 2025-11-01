/*************** MENU SETUP ***************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Soluvery')
    .addItem('Open Drive Audit', 'openSidebar')
    .addToUi();
}

/*************** SIDEBAR ***************/
function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Drive Audit');
  SpreadsheetApp.getUi().showSidebar(html);
}

/*************** PROGRESS TRACKING ***************/
function resetProgress() {
  PropertiesService.getScriptProperties().setProperty('progress', '0');
}
function getProgress() {
  return PropertiesService.getScriptProperties().getProperty('progress') || '0';
}

/*************** SCAN DRIVE (SAFE + CHUNKED) ***************/
function scanDriveStep(startIndex) {
  const MAX_FILES = 30; // 👈 Only scan first 30 files
  const progress = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;

  try {
    sheet = ss.getSheetByName('Drive Audit');
    if (!sheet) {
      sheet = ss.insertSheet('Drive Audit');
      setupAuditSheet(sheet);
    }
  } catch (e) {
    progress.setProperty('progress', '100');
    return { done: true, message: '❌ Unable to access sheet: ' + e.message };
  }

  try {
    const batchSize = 10;
    const files = DriveApp.searchFiles("mimeType != 'application/vnd.google-apps.folder'");
    const totalFiles = Math.min(countAllFiles(), MAX_FILES); // only count up to 30

    // Skip already processed files
    for (let i = 0; i < startIndex && files.hasNext(); i++) files.next();

    let processed = 0;
    const rows = [];

    while (files.hasNext() && processed < batchSize && startIndex + processed < MAX_FILES) {
      const file = files.next();
      let owner = 'Unknown';
      let sharedStatus = 'Private';
      let permissionCount = 0;
      let modifiedDate = 'N/A';
      let mimeType = 'Unknown';

      try {
        owner = file.getOwner() ? file.getOwner().getEmail() : 'Unknown';
        mimeType = file.getMimeType();
        modifiedDate = file.getLastUpdated();
        sharedStatus = file.isShared() ? 'Shared' : 'Private';
        const editors = file.getEditors().length;
        const viewers = file.getViewers().length;
        permissionCount = editors + viewers + 1;
      } catch (metaErr) {
        console.log(`⚠️ Metadata read failed for ${file.getName()}:`, metaErr.message);
      }

      rows.push([
        file.getId(),
        file.getName(),
        mimeType,
        owner,
        modifiedDate,
        sharedStatus,
        permissionCount,
      ]);
      processed++;
    }

    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      sheet.autoResizeColumns(1, 7);
      const minWidths = [150, 200, 180, 200, 150, 120, 120];
      minWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
      sheet.getRange(2, 2, sheet.getLastRow(), 1).setWrap(true);
    }

    const newProgress = totalFiles
      ? Math.min(Math.round(((startIndex + processed) / totalFiles) * 100), 100)
      : 0;
    progress.setProperty('progress', newProgress.toString());

    // 👇 stop once 30 files scanned
    if (startIndex + processed >= MAX_FILES || !files.hasNext()) {
      progress.setProperty('progress', '100');
      addSummaryRow(sheet);
      formatAuditSheet(sheet);
      return { done: true, message: `✅ Scan complete: scanned ${startIndex + processed} of ${MAX_FILES} files.` };
    }

    return { done: false, nextIndex: startIndex + processed };
  } catch (err) {
    progress.setProperty('progress', '100');
    return { done: true, message: '❌ Unexpected error: ' + err.message };
  }
}


/*************** COUNT HELPER ***************/
function countAllFiles() {
  try {
    const files = DriveApp.searchFiles("mimeType != 'application/vnd.google-apps.folder'");
    let count = 0;
    while (files.hasNext()) {
      files.next();
      count++;
    }
    return count;
  } catch (err) {
    console.log('Error counting files:', err);
    return 0;
  }
}

/*************** SETUP AUDIT SHEET ***************/
function setupAuditSheet(sheet) {
  const headers = [
    'File ID',
    'File Name',
    'MIME Type',
    'Owner',
    'Last Modified',
    'Shared Status',
    'Permission Count',
  ];
  sheet.clear();
  sheet.appendRow(headers);
  formatAuditSheet(sheet);
}

/*************** FORMAT HEADERS AND SUMMARY ***************/
function formatAuditSheet(sheet) {
  const range = sheet.getRange('A1:G1');
  range.setFontWeight('bold');
  range.setBackground('#000000');
  range.setFontColor('#FFFFFF');
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setWrap(true);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 7);
}

/*************** SUMMARY ROW ***************/
function addSummaryRow(sheet) {
  const lastRow = sheet.getLastRow();
  const summaryRow = lastRow + 2;

  sheet.getRange(summaryRow, 1).setValue('Summary');
  sheet.getRange(summaryRow, 1).setFontWeight('bold');

  // Total file count
  sheet.getRange(summaryRow, 2).setFormula(`=COUNTA(B2:B${lastRow})`);
  sheet.getRange(summaryRow, 3).setValue('Files Audited');
  sheet.getRange(summaryRow, 1, 1, 3).setBackground('#E0E0E0');
}

let SCAN_STATE = PropertiesService.getScriptProperties();
const MAX_FILES = 100; // Increased batch size for better throughput
const PROP_KEY = 'TEMP_AUDIT_DATA';
const PAGE_TOKEN_KEY = 'NEXT_PAGE_TOKEN';
const TOTAL_COUNT_KEY = 'TOTAL_FILE_COUNT';

/*************** MENU SETUP & SIDEBAR ***************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Soluvery')
    .addItem('Open Drive Audit', 'openSidebar')
    .addToUi();
}

function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Drive Audit');
  SpreadsheetApp.getUi().showSidebar(html);
}

/*************** SHEET HELPERS ***************/
function setupAuditSheet(sheet) {
  const headers = [
    'File ID', 'File Name', 'MIME Type', 'Owner',
    'Last Modified', 'Shared Status', 'Permission Count'
  ];
  sheet.clear();
  sheet.appendRow(headers);
  formatAuditSheet(sheet);
}

function formatAuditSheet(sheet) {
  const range = sheet.getRange('A1:G1');
  range.setFontWeight('bold');
  range.setBackground('#000000');
  range.setFontColor('#FFFFFF');
  range.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 7);
}

function addSummaryRow(sheet) {
  const lastRow = sheet.getLastRow();
  const summaryRow = lastRow + 2;
  sheet.getRange(summaryRow, 1).setValue('Summary');
  sheet.getRange(summaryRow, 1).setFontWeight('bold');
  sheet.getRange(summaryRow, 2).setFormula(`=COUNTA(B2:B${lastRow})`);
  sheet.getRange(summaryRow, 3).setValue('Files Audited');
  sheet.getRange(summaryRow, 1, 1, 3).setBackground('#E0E0E0');
  sheet.getRange(summaryRow, 2).setHorizontalAlignment('center');
}

/*************** RESET & PROGRESS (UPDATED) ***************/
function resetProgress() {
  SCAN_STATE.deleteProperty(PROP_KEY);
  SCAN_STATE.deleteProperty(PAGE_TOKEN_KEY);
  SCAN_STATE.deleteProperty(TOTAL_COUNT_KEY);
  SCAN_STATE.deleteProperty('PROCESSED_COUNT');
  SCAN_STATE.setProperty('progress', '0');
  return true;
}

function updateProgress(scannedCount, totalFiles, nextPageToken) {
  const progress = Math.min((scannedCount / totalFiles) * 100, 100).toFixed(2);
  SCAN_STATE.setProperty('PROCESSED_COUNT', scannedCount.toString());
  SCAN_STATE.setProperty('progress', progress);
  if (nextPageToken) {
    SCAN_STATE.setProperty(PAGE_TOKEN_KEY, nextPageToken);
  } else {
    SCAN_STATE.deleteProperty(PAGE_TOKEN_KEY);
  }
  return progress;
}

function getProgress() {
  const total = parseInt(SCAN_STATE.getProperty(TOTAL_COUNT_KEY) || '0', 10);
  const done = parseInt(SCAN_STATE.getProperty('PROCESSED_COUNT') || '0', 10);
  const progress = SCAN_STATE.getProperty('progress') || '0';
  const nextPageToken = SCAN_STATE.getProperty(PAGE_TOKEN_KEY) || null;

  return {
    total,
    done,
    progress,
    nextPageToken,
  };
}

/*************** COUNT FILES ***************/
function getTotalFileCount() {
  let count = SCAN_STATE.getProperty(TOTAL_COUNT_KEY);
  if (count) return parseInt(count, 10);

  let files = DriveApp.searchFiles("mimeType != 'application/vnd.google-apps.folder'");
  let total = 0;
  while (files.hasNext()) {
    files.next();
    total++;
  }
  SCAN_STATE.setProperty(TOTAL_COUNT_KEY, total.toString());
  return total;
}

/*************** FULL DRIVE SCAN (Optimized) ***************/
function scanEntireDrive() {
  resetProgress();
  const totalFiles = getTotalFileCount();
  let data = [];
  let pageToken = null;
  let scannedCount = 0;

  do {
    const response = Drive.Files.list({
      q: "mimeType != 'application/vnd.google-apps.folder'",
      pageSize: MAX_FILES,
      pageToken: pageToken,
      fields: "nextPageToken, files(id, name, mimeType, owners(emailAddress), modifiedTime, shared)"
    });

    const files = response.files || [];
    pageToken = response.nextPageToken || null;

    // Batch permission requests using UrlFetchApp for parallelism
    const batch = files.map(file => ({
      id: file.id,
      name: file.name || 'Untitled',
      mimeType: file.mimeType || 'Unknown',
      owner: (file.owners && file.owners.length > 0)
        ? file.owners[0].emailAddress
        : 'Unknown',
      modified: file.modifiedTime || 'Unknown',
      shared: file.shared ? 'Shared' : 'Private'
    }));

    // Get permissions efficiently in smaller chunks
    const permissionData = getPermissionsInBatch(batch);

    permissionData.forEach(item => {
      data.push([
        item.id,
        item.name,
        item.mimeType,
        item.owner,
        item.modified,
        item.shared,
        item.permissionCount
      ]);
    });

   scannedCount += batch.length;
const progress = updateProgress(scannedCount, totalFiles, pageToken);
console.log(`Scanned ${scannedCount}/${totalFiles} (${progress}%)`);

    // Light delay to avoid rate limit issues
    Utilities.sleep(200);

  } while (pageToken);

  SCAN_STATE.setProperty(PROP_KEY, JSON.stringify(data));
  SCAN_STATE.setProperty('progress', '100');
  console.log(`✅ Completed scanning ${data.length} files.`);
  return `✅ Completed scanning ${data.length} files. Click "Write to Sheet" to write data in the sheet..`;
}

/*************** PERMISSIONS BATCH (Parallel Requests) ***************/
function getPermissionsInBatch(files) {
  const BATCH_LIMIT = 10; // max parallel requests at a time
  const results = [];

  for (let i = 0; i < files.length; i += BATCH_LIMIT) {
    const chunk = files.slice(i, i + BATCH_LIMIT);
    const requests = chunk.map(f => ({
      url: `https://www.googleapis.com/drive/v3/files/${f.id}/permissions`,
      method: 'get',
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: true
    }));

    const responses = UrlFetchApp.fetchAll(requests);
    responses.forEach((res, idx) => {
      let count = 0;
      try {
        const json = JSON.parse(res.getContentText());
        count = json.permissions ? json.permissions.length : 0;
      } catch (e) {
        count = 0;
      }
      results.push({ ...chunk[idx], permissionCount: count });
    });
  }

  return results;
}

/*************** WRITE TO SHEET ***************/
function writeToSheet() {
  const dataString = SCAN_STATE.getProperty(PROP_KEY);
  if (!dataString || dataString === '[]') {
    return '⚠️ No scan data found. Please scan first.';
  }

  const data = JSON.parse(dataString);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Drive Audit');
  if (!sheet) sheet = ss.insertSheet('Drive Audit');

  setupAuditSheet(sheet);
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  addSummaryRow(sheet);

  // Auto resize, but limit to a max width (e.g., 300 pixels)
  sheet.autoResizeColumns(1, data[0].length);
  const maxWidth = 300;
  for (let i = 1; i <= data[0].length; i++) {
    const width = sheet.getColumnWidth(i);
    if (width > maxWidth) sheet.setColumnWidth(i, maxWidth);
  }

  resetProgress();
  return `✅ ${data.length} files written to "Drive Audit" sheet.`;
}


/*************** BACKGROUND RESCAN ***************/
function reScanAsync() {
  ScriptApp.newTrigger('startBackgroundScanTrigger')
    .timeBased()
    .after(5 * 1000)
    .create();
  return '🔄 Background re-scan scheduled.';
}

function startBackgroundScanTrigger() {
  resetProgress();
  scanEntireDrive();
}
/**
 * Green Team Asset Tracker
 * Google Sheets + Drive + Apps Script backend
 *
 * Setup:
 * 1) Create a Google Sheet and import Assets_import_for_Google_Sheets.csv into a tab named Assets.
 * 2) Add a tab named Users with columns: Email, Role, Active.
 * 3) Create a Drive folder for asset photos.
 * 4) In Apps Script, Project Settings > Script Properties:
 *    ASSET_PHOTO_FOLDER_ID = your Drive folder ID
 * 5) Deploy > New deployment > Web app.
 */

const SHEET_ASSETS = 'Assets';
const SHEET_DELETED = 'Deleted Assets';
const SHEET_USERS = 'Users';

const REQUIRED_ASSET_HEADERS = [
  'Record ID',
  'Asset Number',
  'Brand',
  'Tool Description',
  'Model Number',
  'Serial Number',
  'Asset Tag Y/N',
  'Status',
  'Location',
  'Assigned To',
  'Photo Folder URL',
  'Photo File URLs',
  'Notes',
  'Last Updated',
  'Updated By'
];

function doGet() {
  ensureSetup_();
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle("McCoy's Green Team Asset Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getBootstrap() {
  ensureSetup_();
  const user = getCurrentUser_();
  return {
    appName: "McCoy's Green Team Asset Tracker",
    user,
    canEdit: ['Admin', 'Editor'].includes(user.role),
    canDelete: user.role === 'Admin',
    headers: REQUIRED_ASSET_HEADERS
  };
}

function listAssets() {
  ensureSetup_();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ASSETS);
  const rows = getSheetObjects_(sheet);
  return rows.filter(r => String(r.Status || '').toLowerCase() !== 'deleted');
}

function saveAsset(asset, photoPayloads) {
  ensureSetup_();
  const user = getCurrentUser_();
  if (!['Admin', 'Editor'].includes(user.role)) {
    throw new Error('You do not have permission to add or edit assets.');
  }

  asset = sanitizeAsset_(asset || {});
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ASSETS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const recordIdCol = headers.indexOf('Record ID');
  const assetNumberCol = headers.indexOf('Asset Number');

  if (!asset['Asset Number']) {
    throw new Error('Asset Number is required.');
  }

  asset['Record ID'] = asset['Record ID'] || Utilities.getUuid();
  asset['Last Updated'] = new Date();
  asset['Updated By'] = user.email;

  let rowIndex = -1;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][recordIdCol]) === String(asset['Record ID'])) {
      rowIndex = r + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    // Warn on duplicate asset numbers but allow it because the uploaded file already has one.
    // The internal Record ID keeps each row unique.
    const nextRow = headers.map(h => asset[h] || '');
    sheet.appendRow(nextRow);
    rowIndex = sheet.getLastRow();
  } else {
    const nextRow = headers.map(h => asset[h] !== undefined ? asset[h] : values[rowIndex - 1][headers.indexOf(h)]);
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([nextRow]);
  }

  if (photoPayloads && photoPayloads.length) {
    const uploadResult = uploadPhotosForAsset_(asset['Asset Number'], photoPayloads);
    setCellByHeader_(sheet, rowIndex, 'Photo Folder URL', uploadResult.folderUrl);
    setCellByHeader_(sheet, rowIndex, 'Photo File URLs', uploadResult.fileUrls.join('\n'));
  }

  return getAssetByRecordId_(asset['Record ID']);
}

function deleteAsset(recordId) {
  ensureSetup_();
  const user = getCurrentUser_();
  if (user.role !== 'Admin') {
    throw new Error('Only Admin users can delete assets.');
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_ASSETS);
  const deletedSheet = ss.getSheetByName(SHEET_DELETED);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const recordIdCol = headers.indexOf('Record ID');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][recordIdCol]) === String(recordId)) {
      const row = values[r].slice();
      const statusCol = headers.indexOf('Status');
      const lastUpdatedCol = headers.indexOf('Last Updated');
      const updatedByCol = headers.indexOf('Updated By');

      if (statusCol >= 0) row[statusCol] = 'Deleted';
      if (lastUpdatedCol >= 0) row[lastUpdatedCol] = new Date();
      if (updatedByCol >= 0) row[updatedByCol] = user.email;

      deletedSheet.appendRow(row);
      sheet.deleteRow(r + 1);
      return { ok: true };
    }
  }

  throw new Error('Asset was not found.');
}

function getAssetByRecordId_(recordId) {
  const rows = listAssets();
  return rows.find(r => String(r['Record ID']) === String(recordId)) || null;
}

function uploadPhotosForAsset_(assetNumber, photoPayloads) {
  const root = getRootPhotoFolder_();
  const folder = getOrCreateChildFolder_(root, cleanFolderName_(assetNumber));
  const fileUrls = [];

  photoPayloads.forEach((p, index) => {
    if (!p || !p.dataUrl) return;
    const blob = dataUrlToBlob_(p.dataUrl, p.name || `asset-${assetNumber}-${Date.now()}-${index}.jpg`);
    const file = folder.createFile(blob);
    file.setDescription(`Asset ${assetNumber}`);
    fileUrls.push(file.getUrl());
  });

  return {
    folderUrl: folder.getUrl(),
    fileUrls
  };
}

function ensureSetup_() {
  const ss = SpreadsheetApp.getActive();
  let assets = ss.getSheetByName(SHEET_ASSETS);
  if (!assets) assets = ss.insertSheet(SHEET_ASSETS);
  ensureHeaders_(assets, REQUIRED_ASSET_HEADERS);

  let deleted = ss.getSheetByName(SHEET_DELETED);
  if (!deleted) deleted = ss.insertSheet(SHEET_DELETED);
  ensureHeaders_(deleted, REQUIRED_ASSET_HEADERS);

  let users = ss.getSheetByName(SHEET_USERS);
  if (!users) {
    users = ss.insertSheet(SHEET_USERS);
    users.getRange(1, 1, 1, 3).setValues([['Email', 'Role', 'Active']]);
    users.getRange(2, 1, 1, 3).setValues([[Session.getActiveUser().getEmail(), 'Admin', true]]);
  }
}

function ensureHeaders_(sheet, requiredHeaders) {
  const lastCol = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (sheet.getLastRow() === 0 || current.every(h => !h.trim())) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    return;
  }

  const missing = requiredHeaders.filter(h => !current.includes(h));
  if (missing.length) {
    sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
  }
  sheet.setFrozenRows(1);
}

function getSheetObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(String);
  return data.slice(1)
    .filter(row => row.some(cell => cell !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        if (!h) return;
        obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i];
      });
      return obj;
    });
}

function sanitizeAsset_(asset) {
  const clean = {};
  REQUIRED_ASSET_HEADERS.forEach(h => clean[h] = asset[h] || '');
  clean['Asset Number'] = String(clean['Asset Number']).trim();
  clean['Brand'] = String(clean['Brand']).trim();
  clean['Tool Description'] = String(clean['Tool Description']).trim();
  clean['Model Number'] = String(clean['Model Number']).trim();
  clean['Serial Number'] = String(clean['Serial Number']).trim();
  clean['Asset Tag Y/N'] = String(clean['Asset Tag Y/N']).trim().toUpperCase();
  clean['Status'] = clean['Status'] || 'Active';
  return clean;
}

function setCellByHeader_(sheet, rowIndex, header, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const colIndex = headers.indexOf(header) + 1;
  if (colIndex > 0) sheet.getRange(rowIndex, colIndex).setValue(value);
}

function getCurrentUser_() {
  const email = Session.getActiveUser().getEmail() || '';
  const role = getUserRole_(email);
  return { email, role };
}

function getUserRole_(email) {
  if (!email) return 'Viewer';
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_USERS);
  if (!sheet) return 'Viewer';

  const rows = getSheetObjects_(sheet);
  const match = rows.find(r =>
    String(r.Email || '').toLowerCase() === email.toLowerCase() &&
    String(r.Active || '').toLowerCase() !== 'false'
  );

  // If Users sheet is blank, make current signed-in owner admin.
  if (!match && rows.length === 0) return 'Admin';

  return match ? String(match.Role || 'Viewer') : 'Viewer';
}

function getRootPhotoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('ASSET_PHOTO_FOLDER_ID');
  if (folderId) return DriveApp.getFolderById(folderId);

  // Fallback: create a folder if no script property was set.
  const folder = getOrCreateRootFolderByName_('Green Team Asset Photos');
  props.setProperty('ASSET_PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

function getOrCreateRootFolderByName_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getOrCreateChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function cleanFolderName_(assetNumber) {
  return String(assetNumber || 'NO-ASSET-NUMBER')
    .replace(/[\\/:*?"<>|#%{}~&]/g, '-')
    .trim();
}

function dataUrlToBlob_(dataUrl, filename) {
  const parts = String(dataUrl).split(',');
  const metadata = parts[0];
  const base64 = parts[1];
  const contentTypeMatch = metadata.match(/data:(.*?);base64/);
  const contentType = contentTypeMatch ? contentTypeMatch[1] : 'image/jpeg';
  const bytes = Utilities.base64Decode(base64);
  return Utilities.newBlob(bytes, contentType, filename);
}

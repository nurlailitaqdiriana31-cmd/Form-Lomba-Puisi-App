/**
 * Backend for the Poetry Team Registration web app.
 * Upload the supplied HTML as an Apps Script HTML file named "Index".
 */
const APP_CONFIG = Object.freeze({
  title: 'Pendaftaran Lomba Membaca Puisi Beregu 2026',
  spreadsheetName: 'Data Pendaftaran Lomba Membaca Puisi Beregu 2026',
  sheetName: 'Pendaftaran',
  spreadsheetProperty: 'REGISTRATION_SPREADSHEET_ID',
  adminEmailProperty: 'REGISTRATION_ADMIN_EMAIL'
});

const COLUMNS = Object.freeze([
  'id', 'namaSekolah', 'alamatSekolah', 'namaKepalaSekolah', 'nipKepalaSekolah',
  'nomorHP', 'peserta1_nama', 'peserta1_ttl', 'peserta1_jk',
  'peserta2_nama', 'peserta2_ttl', 'peserta2_jk',
  'peserta3_nama', 'peserta3_ttl', 'peserta3_jk',
  'judulPuisi', 'musikAda', 'tanggalSurat', 'updatedAt'
]);

/** Serves Index.html as the Apps Script web app. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_CONFIG.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Returns all saved registrations as plain objects for the existing frontend. */
function getAll() {
  const sheet = getRegistrationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getDisplayValues();
  return values
    .filter(row => row[0])
    .map(row => {
      const item = {};
      COLUMNS.forEach((key, index) => {
        if (key !== 'updatedAt') item[key] = row[index];
      });
      return item;
    });
}

/** Inserts or updates one registration. Called by google.script.run. */
function saveRegistration(input) {
  const record = normalizeRecord_(input);
  const sheet = getRegistrationSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const rowNumber = findRowById_(sheet, record.id);
    record.updatedAt = new Date().toISOString();
    const row = COLUMNS.map(key => safeCellText_(record[key]));
    if (rowNumber) sheet.getRange(rowNumber, 1, 1, COLUMNS.length).setValues([row]);
    else sheet.appendRow(row);
    SpreadsheetApp.flush();
    return { ok: true, id: record.id };
  } finally {
    lock.releaseLock();
  }
}
/**
 * Deletes only for the configured administrator. Set REGISTRATION_ADMIN_EMAIL
 * in Project Settings > Script properties before deploying. If identity is
 * unavailable for a deployment, deletion fails closed; rows can be removed
 * directly in the private spreadsheet by its owner.
 */
function deleteRegistration(id) {
  const expected = PropertiesService.getScriptProperties().getProperty(APP_CONFIG.adminEmailProperty);
  const activeEmail = Session.getActiveUser().getEmail();
  if (!expected || !activeEmail || activeEmail.toLowerCase() !== expected.toLowerCase()) {
    throw new Error('Aksi hapus hanya tersedia untuk akun administrator yang dikonfigurasi.');
  }
  const cleanId = String(id || '').trim();
  if (!cleanId) throw new Error('ID pendaftaran tidak valid.');
  const sheet = getRegistrationSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const rowNumber = findRowById_(sheet, cleanId);
    if (rowNumber) sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
/** Run this from the editor to find the spreadsheet created for this app. */
function getRegistrationSpreadsheetUrl() {
  return getRegistrationSheet_().getParent().getUrl();
}

function getRegistrationSheet_() {
  const properties = PropertiesService.getScriptProperties();
  let id = properties.getProperty(APP_CONFIG.spreadsheetProperty);
  let spreadsheet;

  if (id) {
    try {
      spreadsheet = SpreadsheetApp.openById(id);
    } catch (error) {
      throw new Error('Spreadsheet penyimpanan tidak bisa dibuka. Periksa akses akun pemilik script.');
    }
  } else {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      id = properties.getProperty(APP_CONFIG.spreadsheetProperty);
      if (id) {
        spreadsheet = SpreadsheetApp.openById(id);
      } else {
        spreadsheet = SpreadsheetApp.create(APP_CONFIG.spreadsheetName);
        properties.setProperty(APP_CONFIG.spreadsheetProperty, spreadsheet.getId());
      }
    } finally {
      lock.releaseLock();
    }
  }

  let sheet = spreadsheet.getSheetByName(APP_CONFIG.sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(APP_CONFIG.sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, COLUMNS.length)
      .setBackground('#0E3B3C').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setColumnWidths(1, COLUMNS.length, 170);
    sheet.setColumnWidth(1, 190);
    sheet.setColumnWidth(3, 260);
  }
  return sheet;
}

function normalizeRecord_(input) {
  if (!input || typeof input !== 'object') throw new Error('Data pendaftaran tidak valid.');
  const record = {};
  COLUMNS.forEach(key => {
    if (key === 'updatedAt') return;
    const value = input[key] == null ? '' : String(input[key]);
    if (value.length > 2000) throw new Error('Ada kolom yang melebihi batas panjang.');
    record[key] = value.trim();
  });
  if (!record.id) throw new Error('ID pendaftaran tidak valid.');
  return record;
}

function findRowById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const found = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(String(id)).matchEntireCell(true).findNext();
  return found ? found.getRow() : 0;
}

/** Prefix formula-like user input so Google Sheets stores it as plain text. */
function safeCellText_(value) {
  const text = String(value == null ? '' : value);
  return /^[\s]*[=+@-]/.test(text) ? "'" + text : text;
}
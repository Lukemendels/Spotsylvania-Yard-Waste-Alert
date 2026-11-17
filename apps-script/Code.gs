const SITE = 'chancellor';
const SERVICE = 'yard-waste';
const STATUSES = ['accepting', 'filling', 'full'];
const LOG_HEADERS = ['Update ID', 'Site ID', 'Service', 'Status', 'Confirmed At (UTC)', 'Publisher Source'];
const PUBLIC_HEADERS = ['Status', 'Confirmed At (UTC)', 'Expires At (UTC)', 'Version'];
const FRESHNESS_MS = 2 * 60 * 60 * 1000;

function doGet(event) {
  const query = (event && event.parameter) || {};
  if (query.feed !== '1') {
    return HtmlService.createHtmlOutputFromFile('Staff')
      .setTitle('Chancellor Yard Waste · Staff')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  let record;
  try {
    record = publicStatus_();
  } catch (error) {
    console.error('Unable to read the published yard-waste status.');
    record = { status: 'unavailable', version: 0 };
  }

  const body = JSON.stringify(record);
  if (query.callback && /^yardWasteStatus_[a-f0-9]{32}$/.test(query.callback)) {
    return ContentService.createTextOutput(query.callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function openOffice(key) {
  authorize_(key);
  return latestCommitted_(spreadsheet_().getSheetByName('Status Log'));
}

function saveStatus(key, status, requestId) {
  authorize_(key);
  if (!STATUSES.includes(status) || typeof requestId !== 'string' ||
      !/^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/.test(requestId)) {
    throw new Error('Invalid status request.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const log = spreadsheet_().getSheetByName('Status Log');
    const lastRow = log.getLastRow();
    const duplicate = lastRow > 1 && log.getRange(2, 1, lastRow - 1, 1)
      .createTextFinder(requestId).matchEntireCell(true).findNext();
    if (duplicate) {
      const version = duplicate.getRow();
      const row = log.getRange(version, 1, 1, LOG_HEADERS.length).getValues()[0];
      if (row[3] !== status) throw new Error('Update ID is already in use.');
      return {
        saved: committedRow_(row, version),
        latest: latestCommitted_(log)
      };
    }

    const confirmedAt = new Date().toISOString();
    log.appendRow([requestId, SITE, SERVICE, status, confirmedAt, 'Chancellor office']);
    SpreadsheetApp.flush();
    const saved = { status, confirmedAt, version: lastRow + 1 };
    return { saved, latest: saved };
  } finally {
    lock.releaseLock();
  }
}

function authorize_(candidate) {
  const expected = PropertiesService.getScriptProperties().getProperty('OFFICE_KEY');
  if (!expected || expected.length < 32 || typeof candidate !== 'string' ||
      candidate.length !== expected.length) {
    throw new Error('Office key not accepted.');
  }
  let difference = 0;
  for (let index = 0; index < expected.length; index++) {
    difference |= expected.charCodeAt(index) ^ candidate.charCodeAt(index);
  }
  if (difference !== 0) throw new Error('Office key not accepted.');
}

function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Demo setup is incomplete.');
  return SpreadsheetApp.openById(id);
}

function committedRow_(row, version) {
  const confirmedAt = timestamp_(row[4]);
  if (row[1] !== SITE || row[2] !== SERVICE || !STATUSES.includes(row[3]) || !confirmedAt) {
    throw new Error('The latest status row is invalid.');
  }
  return { status: row[3], confirmedAt, version };
}

function latestCommitted_(log) {
  const lastRow = log.getLastRow();
  if (lastRow < 2) return null;
  return committedRow_(log.getRange(lastRow, 1, 1, LOG_HEADERS.length).getValues()[0], lastRow);
}

function timestamp_(value) {
  const time = value instanceof Date ? value.getTime() :
    typeof value === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/.test(value)
      ? Date.parse(value)
      : NaN;
  return Number.isFinite(time) ? new Date(time).toISOString() : null;
}

function publishPending_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = spreadsheet_();
    const latest = latestCommitted_(spreadsheet.getSheetByName('Status Log'));
    if (!latest) return;

    const output = spreadsheet.getSheetByName('Published Status').getRange(2, 1, 1, PUBLIC_HEADERS.length);
    if (Number(output.getValues()[0][3]) >= latest.version) return;

    const expiresAt = new Date(Date.parse(latest.confirmedAt) + FRESHNESS_MS).toISOString();
    output.setValues([[latest.status, latest.confirmedAt, expiresAt, latest.version]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function publicStatus_() {
  const sheet = spreadsheet_().getSheetByName('Published Status');
  const [status, rawConfirmed, rawExpiry, rawVersion] = sheet.getRange(2, 1, 1, PUBLIC_HEADERS.length).getValues()[0];
  const confirmedAt = timestamp_(rawConfirmed);
  const expiresAt = timestamp_(rawExpiry);
  const version = Number(rawVersion) || 0;
  const now = Date.now();
  if (!STATUSES.includes(status) || !confirmedAt || !expiresAt || !Number.isInteger(version) ||
      version < 2 || Date.parse(confirmedAt) > now || Date.parse(expiresAt) <= now ||
      Date.parse(expiresAt) - Date.parse(confirmedAt) !== FRESHNESS_MS) {
    return { status: 'unavailable', version };
  }
  return { status, confirmedAt, expiresAt, version };
}

// Run this function once from the editor attached to the private demo Sheet.
function setupDemo_() {
  const properties = PropertiesService.getScriptProperties();
  const key = properties.getProperty('OFFICE_KEY');
  if (!key || key.length < 32) throw new Error('Set OFFICE_KEY in Script Properties first.');
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Run setup from the Apps Script project attached to a Sheet.');
  properties.setProperty('SPREADSHEET_ID', spreadsheet.getId());
  prepareTab_(spreadsheet, 'Status Log', LOG_HEADERS);
  prepareTab_(spreadsheet, 'Published Status', PUBLIC_HEADERS);
  const hasPublisher = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === 'publishPending_');
  if (!hasPublisher) ScriptApp.newTrigger('publishPending_').timeBased().everyMinutes(1).create();
}

function prepareTab_(spreadsheet, name, headers) {
  const tab = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  const firstRow = tab.getRange(1, 1, 1, headers.length);
  if (tab.getLastRow() === 0) {
    firstRow.setValues([headers]).setFontWeight('bold');
    tab.setFrozenRows(1);
  } else if (firstRow.getValues()[0].some((value, index) => value !== headers[index])) {
    throw new Error('Unexpected columns in ' + name + '.');
  }
}

const SITE = 'chancellor';
const SERVICE = 'yard-waste';
const STATUSES = ['accepting', 'filling', 'full'];
const LOG_HEADERS = ['Update ID', 'Site ID', 'Service', 'Status', 'Confirmed At (UTC)', 'Publisher Source'];
const PUBLIC_HEADERS = ['Status', 'Confirmed At (UTC)', 'Expires At (UTC)', 'Version'];
const QUEUE_LIMIT = 30;
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
  const properties = PropertiesService.getScriptProperties().getProperties();
  const queue = readQueue_(properties);
  return queue[queue.length - 1] || readPublished_(properties);
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
    const store = PropertiesService.getScriptProperties();
    const properties = store.getProperties();
    const queue = readQueue_(properties);
    const published = readPublished_(properties);
    const receipts = JSON.parse(properties.RECEIPTS || '[]');
    const duplicate = queue.concat(receipts).find(record => record.id === requestId);
    if (duplicate) {
      if (duplicate.status !== status) throw new Error('Update ID is already in use.');
      return queue[queue.length - 1] || published;
    }

    if (queue.length >= QUEUE_LIMIT) throw new Error('Publisher queue is full.');
    const record = { id: requestId, status, confirmedAt: new Date().toISOString(), queued: true };
    queue.push(record);
    store.setProperty('STATUS_QUEUE', JSON.stringify(queue));
    return record;
  } finally {
    lock.releaseLock();
  }
}

function readQueue_(properties) {
  return JSON.parse(properties.STATUS_QUEUE || '[]');
}

function readPublished_(properties) {
  return JSON.parse(properties.PUBLISHED_STATUS || 'null');
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
    const store = PropertiesService.getScriptProperties();
    const properties = store.getProperties();
    const queue = readQueue_(properties);
    if (!queue.length) return;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) throw new Error('Publisher requires the bound demo Sheet.');
    const log = spreadsheet.getSheetByName('Status Log');
    let latest;
    for (const record of queue) {
      const lastRow = log.getLastRow();
      const duplicate = lastRow > 1 && log.getRange(2, 1, lastRow - 1, 1)
        .createTextFinder(record.id).matchEntireCell(true).findNext();
      if (!duplicate) {
        log.appendRow([record.id, SITE, SERVICE, record.status, record.confirmedAt, 'Chancellor office']);
      }
      latest = { status: record.status, confirmedAt: record.confirmedAt,
        expiresAt: new Date(Date.parse(record.confirmedAt) + FRESHNESS_MS).toISOString(),
        version: duplicate ? duplicate.getRow() : lastRow + 1 };
    }

    const published = spreadsheet.getSheetByName('Published Status');
    published.getRange(2, 1, 1, PUBLIC_HEADERS.length)
      .setValues([[latest.status, latest.confirmedAt, latest.expiresAt, latest.version]]);
    SpreadsheetApp.flush();

    const receipts = JSON.parse(properties.RECEIPTS || '[]');
    const newReceipts = queue.map(({ id, status }) => ({ id, status }));
    store.setProperties({
      STATUS_QUEUE: '[]',
      PUBLISHED_STATUS: JSON.stringify(latest),
      RECEIPTS: JSON.stringify(receipts.concat(newReceipts).slice(-QUEUE_LIMIT))
    });
  } finally {
    lock.releaseLock();
  }
}

function publicStatus_() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const {
    status, confirmedAt: rawConfirmed, expiresAt: rawExpiry, version = 0
  } = readPublished_(properties) || {};
  const confirmedAt = timestamp_(rawConfirmed);
  const expiresAt = timestamp_(rawExpiry);
  const now = Date.now();
  if (!STATUSES.includes(status) || !confirmedAt || !expiresAt || !Number.isInteger(version) ||
      version < 2 || Date.parse(confirmedAt) > now || Date.parse(expiresAt) <= now ||
      Date.parse(expiresAt) - Date.parse(confirmedAt) !== FRESHNESS_MS) {
    return { status: 'unavailable', version };
  }
  return { status, confirmedAt, expiresAt, version };
}

// Run this function once from the editor attached to the private demo Sheet.
function setupDemo() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Run setup from the Apps Script project attached to a Sheet.');
  const properties = PropertiesService.getScriptProperties();
  const key = properties.getProperty('OFFICE_KEY');
  if (!key || key.length < 32) throw new Error('Set OFFICE_KEY in Script Properties first.');
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

/**
 * Neon CRM -> Google Sheets weekly prospect pull, with notes synced back to Neon.
 *
 * Every Monday: pulls 10 NEW accounts that
 *   - have opened our emails,
 *   - have never donated,
 *   - have a Windfall ID and a calculated Net Worth,
 * and appends them to the "Prospects" sheet.
 *
 * When someone types in the "Notes" column, the note is posted to that account
 * in Neon and the row is stamped in "Synced to Neon".
 *
 * Setup: see README.md in this folder.
 */

// ---------------- CONFIG ----------------
const CONFIG = {
  API_BASE: 'https://api.neoncrm.com/v2',
  API_VERSION: '2.10',
  SHEET_NAME: 'Prospects',
  HISTORY_SHEET: '_pulled_ids',       // hidden; stops the same person being pulled twice
  RECORDS_PER_WEEK: 10,

  // Custom field IDs from Neon (Settings > Custom Fields, or run logSearchFields()).
  WINDFALL_ID_FIELD: 'REPLACE_WITH_WINDFALL_ID_CUSTOM_FIELD_ID',
  NET_WORTH_FIELD: 'REPLACE_WITH_NET_WORTH_CUSTOM_FIELD_ID',

  // Standard search fields. Names can differ slightly per Neon instance —
  // run logSearchFields() once and adjust if a search errors.
  EMAIL_OPENED_FIELD: 'Email Opened',        // e.g. "Email Opened" / "Last Email Open Date"
  EMAIL_OPENED_OPERATOR: 'NOT_BLANK',
  DONATION_COUNT_FIELD: 'Lifetime Donation Count',

  // How notes are written back: Neon's API stores account notes as Activities.
  NOTE_SUBJECT: 'Prospect note (from Google Sheet)',
};

const HEADERS = ['Pulled On', 'Account ID', 'Full Name', 'Phone 1', 'Email 1', 'Net Worth', 'Notes', 'Synced to Neon'];
const COL = { ACCOUNT_ID: 2, NOTES: 7, SYNCED: 8 };

// ---------------- SETUP ----------------
/** Run once: stores nothing secret in code, creates sheet + triggers. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('NEON_ORG_ID') || !props.getProperty('NEON_API_KEY')) {
    throw new Error('Set NEON_ORG_ID and NEON_API_KEY in Project Settings > Script Properties first.');
  }
  getSheet_();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('weeklyPull').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  ScriptApp.newTrigger('onNoteEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  SpreadsheetApp.getActive().toast('Neon sync set up: Mondays 8am + notes sync.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Neon')
    .addItem('Pull 10 prospects now', 'weeklyPull')
    .addItem('Sync unsynced notes', 'syncAllNotes')
    .addToUi();
}

// ---------------- WEEKLY PULL ----------------
function weeklyPull() {
  const sheet = getSheet_();
  const already = getPulledIds_();
  const picked = [];
  let page = 0, totalPages = 1;

  while (picked.length < CONFIG.RECORDS_PER_WEEK && page < totalPages) {
    const res = neon_('post', '/accounts/search', {
      searchFields: [
        { field: CONFIG.EMAIL_OPENED_FIELD, operator: CONFIG.EMAIL_OPENED_OPERATOR },
        { field: CONFIG.DONATION_COUNT_FIELD, operator: 'EQUAL', value: '0' },
        { field: CONFIG.WINDFALL_ID_FIELD, operator: 'NOT_BLANK' },
        { field: CONFIG.NET_WORTH_FIELD, operator: 'NOT_BLANK' },
      ],
      outputFields: ['Account ID', 'Full Name (F)', 'Phone 1 Full Number (F)', 'Email 1', CONFIG.NET_WORTH_FIELD],
      pagination: { currentPage: page, pageSize: 100 },
    });
    totalPages = (res.pagination && res.pagination.totalPages) || 0;
    (res.searchResults || []).forEach(r => {
      const id = String(r['Account ID']);
      if (picked.length < CONFIG.RECORDS_PER_WEEK && !already.has(id)) {
        picked.push([new Date(), id, r['Full Name (F)'] || '', r['Phone 1 Full Number (F)'] || '',
          r['Email 1'] || '', findNetWorth_(r), '', '']);
        already.add(id);
      }
    });
    page++;
  }

  if (!picked.length) { Logger.log('No new matching accounts.'); return; }
  sheet.getRange(sheet.getLastRow() + 1, 1, picked.length, HEADERS.length).setValues(picked);
  addPulledIds_(picked.map(r => r[1]));
  Logger.log('Added ' + picked.length + ' prospects.');
}

// Custom field output keys come back under the field's display name, not its ID.
function findNetWorth_(row) {
  if (row[CONFIG.NET_WORTH_FIELD] !== undefined) return row[CONFIG.NET_WORTH_FIELD];
  const key = Object.keys(row).find(k => /net\s*worth/i.test(k));
  return key ? row[key] : '';
}

// ---------------- NOTES -> NEON ----------------
/** Installable onEdit trigger (simple onEdit can't call external APIs). */
function onNoteEdit(e) {
  const range = e.range;
  if (range.getSheet().getName() !== CONFIG.SHEET_NAME) return;
  if (range.getColumn() > COL.NOTES || range.getLastColumn() < COL.NOTES) return;
  for (let r = range.getRow(); r <= range.getLastRow(); r++) {
    if (r > 1) syncRow_(range.getSheet(), r);
  }
}

/** Menu fallback: push any notes that aren't marked synced yet. */
function syncAllNotes() {
  const sheet = getSheet_();
  for (let r = 2; r <= sheet.getLastRow(); r++) {
    if (!sheet.getRange(r, COL.SYNCED).getValue()) syncRow_(sheet, r);
  }
}

function syncRow_(sheet, row) {
  const accountId = sheet.getRange(row, COL.ACCOUNT_ID).getValue();
  const note = String(sheet.getRange(row, COL.NOTES).getValue()).trim();
  if (!accountId || !note) return;
  const synced = sheet.getRange(row, COL.SYNCED);
  try {
    neon_('post', '/activities', {
      clientId: String(accountId),
      subject: CONFIG.NOTE_SUBJECT,
      note: note + '\n\n— ' + (Session.getActiveUser().getEmail() || 'Google Sheet'),
      activityStatus: { name: 'Completed' },
      startDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    });
    synced.setValue('✅ ' + new Date().toLocaleString());
  } catch (err) {
    synced.setValue('❌ ' + err.message.slice(0, 200));
  }
}

// ---------------- HELPERS ----------------
function neon_(method, path, body) {
  const p = PropertiesService.getScriptProperties();
  const auth = Utilities.base64Encode(p.getProperty('NEON_ORG_ID') + ':' + p.getProperty('NEON_API_KEY'));
  const res = UrlFetchApp.fetch(CONFIG.API_BASE + path, {
    method: method,
    contentType: 'application/json',
    headers: { Authorization: 'Basic ' + auth, 'NEON-API-VERSION': CONFIG.API_VERSION },
    payload: body ? JSON.stringify(body) : undefined,
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code >= 300) throw new Error('Neon ' + code + ': ' + res.getContentText());
  const text = res.getContentText();
  return text ? JSON.parse(text) : {};
}

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(COL.NOTES, 350);
  }
  return sheet;
}

function getPulledIds_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.HISTORY_SHEET);
  if (!sh || sh.getLastRow() === 0) return new Set();
  return new Set(sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(v => String(v[0])));
}

function addPulledIds_(ids) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!sh) { sh = ss.insertSheet(CONFIG.HISTORY_SHEET); sh.hideSheet(); }
  sh.getRange(sh.getLastRow() + 1, 1, ids.length, 1).setValues(ids.map(id => [id]));
}

/** Run once to see the exact search/output field names + custom field IDs in YOUR Neon. */
function logSearchFields() {
  Logger.log(JSON.stringify(neon_('get', '/accounts/search/searchFields'), null, 1));
  Logger.log(JSON.stringify(neon_('get', '/customFields?category=Account'), null, 1));
}

/**
 * Troubleshooting: tests each filter on its own and logs how many accounts match,
 * so you can see which one is returning 0 (usually a wrong field name/ID).
 */
function debugFilters() {
  const filters = {
    'Email opened': { field: CONFIG.EMAIL_OPENED_FIELD, operator: CONFIG.EMAIL_OPENED_OPERATOR },
    'Never donated': { field: CONFIG.DONATION_COUNT_FIELD, operator: 'EQUAL', value: '0' },
    'Has Windfall ID': { field: CONFIG.WINDFALL_ID_FIELD, operator: 'NOT_BLANK' },
    'Has Net Worth': { field: CONFIG.NET_WORTH_FIELD, operator: 'NOT_BLANK' },
  };
  const all = [];
  Object.keys(filters).forEach(name => {
    all.push(filters[name]);
    [[name + ' (alone)', [filters[name]]], ['...combined up to ' + name, all.slice()]].forEach(([label, fields]) => {
      try {
        const res = neon_('post', '/accounts/search', {
          searchFields: fields, outputFields: ['Account ID'], pagination: { currentPage: 0, pageSize: 1 },
        });
        Logger.log(label + ': ' + ((res.pagination && res.pagination.totalResults) || 0) + ' accounts');
      } catch (err) {
        Logger.log(label + ': ERROR -> ' + err.message);
      }
    });
  });
  Logger.log('Already pulled (skipped): ' + getPulledIds_().size);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);
  return sheet;
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim()] = i + 1; });
  return { map: map, headers: headers };
}

function ensureManagedColumns() {
  const sheet = getSheet_();
  MANAGED_COLUMN_HEADERS.forEach(header => {
    const { map } = getHeaderMap_(sheet);
    let col = map[header];
    if (!col) {
      col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header);
    }
    sheet.hideColumns(col);
  });
}

function readRows() {
  const sheet = getSheet_();
  const { map, headers } = getHeaderMap_(sheet);
  if (!map[DATE_COLUMN_HEADER]) throw new Error('Missing column: ' + DATE_COLUMN_HEADER);
  if (!map[WEIGHT_COLUMN_HEADER]) throw new Error('Missing column: ' + WEIGHT_COLUMN_HEADER);

  const syncedAtCol = map[SYNCED_AT_COLUMN_HEADER] || null;
  const healthIdsCol = map[HEALTH_IDS_COLUMN_HEADER] || null;
  const dateCol = map[DATE_COLUMN_HEADER];
  const weightCol = map[WEIGHT_COLUMN_HEADER];

  const exerciseCols = [];
  headers.forEach((h, i) => {
    const name = String(h).trim();
    if (!name) return;
    if (name === DATE_COLUMN_HEADER || name === WEIGHT_COLUMN_HEADER) return;
    if (MANAGED_COLUMN_HEADERS.indexOf(name) !== -1) return;
    exerciseCols.push({ name: name, col: i + 1 });
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], syncedAtCol: syncedAtCol, healthIdsCol: healthIdsCol };

  const width = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();

  const rows = [];
  values.forEach((row, idx) => {
    const rowNum = idx + 2;
    const dateVal = row[dateCol - 1];
    if (!dateVal) return;
    const date = toDate_(dateVal);
    if (!date) return;
    const exercises = [];
    exerciseCols.forEach(c => {
      const entries = parseExerciseCell(row[c.col - 1]);
      if (entries.length > 0) exercises.push({ name: c.name, entries: entries });
    });
    const bodyweight = parseBodyweight(row[weightCol - 1]);
    const syncedAt = syncedAtCol ? row[syncedAtCol - 1] : '';
    const healthIds = healthIdsCol ? parseHealthIds_(row[healthIdsCol - 1]) : [];
    rows.push({
      rowNum: rowNum,
      date: date,
      exercises: exercises,
      bodyweight: bodyweight,
      syncedAt: syncedAt ? String(syncedAt).trim() : '',
      healthIds: healthIds
    });
  });
  return { rows: rows, syncedAtCol: syncedAtCol, healthIdsCol: healthIdsCol };
}

function parseHealthIds_(raw) {
  if (!raw) return [];
  const text = String(raw).trim();
  if (!text) return [];
  try {
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed.filter(s => typeof s === 'string') : [];
  } catch (err) {
    console.warn('parseHealthIds_: could not parse "' + text + '": ' + err);
    return [];
  }
}

function writeHealthIds(rowNum, healthIdsCol, names) {
  const sheet = getSheet_();
  sheet.getRange(rowNum, healthIdsCol).setValue(JSON.stringify(names));
}

function toDate_(value) {
  if (value instanceof Date) return value;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function ymd(date) {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

function ymdCompact(date) {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(date, tz, 'yyyyMMdd');
}

function markRowSynced(rowNum, syncedAtCol, isoTimestamp) {
  const sheet = getSheet_();
  sheet.getRange(rowNum, syncedAtCol).setValue(isoTimestamp);
}

function clearRowSynced(rowNum, syncedAtCol) {
  const sheet = getSheet_();
  sheet.getRange(rowNum, syncedAtCol).setValue('');
}

function onEditMarkDirty(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getSheetId() !== getSheet_().getSheetId()) return;

  const firstRow = e.range.getRow();
  if (firstRow < 2) return;
  const lastRow = e.range.getLastRow();
  const firstCol = e.range.getColumn();
  const lastCol = e.range.getLastColumn();

  const { map } = getHeaderMap_(sheet);
  const syncedAtCol = map[SYNCED_AT_COLUMN_HEADER];
  if (!syncedAtCol) return;

  const managedCols = MANAGED_COLUMN_HEADERS.map(h => map[h]).filter(c => c);
  const editedCols = [];
  for (let c = firstCol; c <= lastCol; c++) editedCols.push(c);
  const onlyManaged = editedCols.every(c => managedCols.indexOf(c) !== -1);
  if (onlyManaged) return;

  const numRows = lastRow - firstRow + 1;
  const blanks = new Array(numRows).fill(['']);
  sheet.getRange(firstRow, syncedAtCol, numRows, 1).setValues(blanks);
}

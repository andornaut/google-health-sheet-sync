// Module-level caches reset on every Apps Script execution (each invocation
// gets a fresh V8 context), so they amortize repeated lookups within one
// sync pass without leaking state across passes.
let cachedSheet_ = null;
let cachedTz_ = null;

function getSheet_() {
  if (cachedSheet_) return cachedSheet_;
  cachedSheet_ = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  return cachedSheet_;
}

function getTz_() {
  if (cachedTz_) return cachedTz_;
  cachedTz_ = Session.getScriptTimeZone();
  return cachedTz_;
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim()] = i + 1; });
  return { map: map, headers: headers };
}

// Format an edited range using column header names: `Header[row]` for one
// cell, `H1,H2[row]` for multi-column, `Header[r1-r2]` for multi-row. Falls
// back to the range's bounding box when no cells had content (all-empty
// edits like clearing already-blank cells).
function describeEditRange_(headers, touched, firstRow, lastRow, firstCol, lastCol) {
  let cells = touched;
  if (cells.length === 0) {
    cells = [];
    for (let r = firstRow; r <= lastRow; r++) {
      for (let c = firstCol; c <= lastCol; c++) cells.push({ col: c, row: r });
    }
  }
  const seen = {};
  const headerList = [];
  let minRow = Infinity, maxRow = -Infinity;
  for (let i = 0; i < cells.length; i++) {
    const { col, row } = cells[i];
    const name = String(headers[col - 1] || '').trim() || 'col' + col;
    if (!seen[name]) { seen[name] = true; headerList.push(name); }
    if (row < minRow) minRow = row;
    if (row > maxRow) maxRow = row;
  }
  const rowDesc = minRow === maxRow ? String(minRow) : minRow + '-' + maxRow;
  return headerList.join(',') + '[' + rowDesc + ']';
}

function ensureManagedColumns() {
  const sheet = getSheet_();
  const { map } = getHeaderMap_(sheet);
  MANAGED_COLUMN_HEADERS.forEach(header => {
    let col = map[header];
    if (!col) {
      col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header);
      map[header] = col;
    }
    sheet.hideColumns(col);
  });
}

function readRows() {
  const sheet = getSheet_();
  const { map, headers } = getHeaderMap_(sheet);
  if (!map[DATE_COLUMN_HEADER]) throw new Error('Missing column: ' + DATE_COLUMN_HEADER);
  if (!map[WEIGHT_COLUMN_HEADER]) throw new Error('Missing column: ' + WEIGHT_COLUMN_HEADER);

  const exerciseSyncedAtCol = map[EXERCISE_SYNCED_AT_COLUMN_HEADER] || null;
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER] || null;
  const healthIdsCol = map[HEALTH_IDS_COLUMN_HEADER] || null;
  const exerciseFirstEditedAtCol = map[EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER] || null;
  const exercisesLastEditedAtCol = map[EXERCISES_LAST_EDITED_AT_COLUMN_HEADER] || null;
  const weightEditedAtCol = map[WEIGHT_EDITED_AT_COLUMN_HEADER] || null;
  const matchedHealthSessionCol = map[MATCHED_HEALTH_SESSION_COLUMN_HEADER] || null;
  const dateCol = map[DATE_COLUMN_HEADER];
  const weightCol = map[WEIGHT_COLUMN_HEADER];

  const exerciseCols = [];
  const seenExerciseNames = {};
  const duplicateExerciseNames = {};
  headers.forEach((h, i) => {
    const name = String(h).trim();
    if (!name) return;
    if (name === DATE_COLUMN_HEADER || name === WEIGHT_COLUMN_HEADER) return;
    if (MANAGED_COLUMN_HEADERS.indexOf(name) !== -1) return;
    if (seenExerciseNames[name]) {
      duplicateExerciseNames[name] = true;
      return;
    }
    seenExerciseNames[name] = true;
    exerciseCols.push({ name: name, col: i + 1 });
  });
  const duplicates = Object.keys(duplicateExerciseNames);
  if (duplicates.length > 0) {
    throw new Error('Duplicate exercise column header(s): ' + duplicates.sort().join(', ')
      + '. Each exercise column header must be unique.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      rows: [],
      exerciseSyncedAtCol: exerciseSyncedAtCol,
      weightSyncedAtCol: weightSyncedAtCol,
      weightCol: weightCol,
      healthIdsCol: healthIdsCol,
      exerciseFirstEditedAtCol: exerciseFirstEditedAtCol,
      exercisesLastEditedAtCol: exercisesLastEditedAtCol,
      weightEditedAtCol: weightEditedAtCol,
      matchedHealthSessionCol: matchedHealthSessionCol
    };
  }

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
    const exerciseSyncedAt = exerciseSyncedAtCol ? row[exerciseSyncedAtCol - 1] : '';
    const weightSyncedAt = weightSyncedAtCol ? row[weightSyncedAtCol - 1] : '';
    const healthIds = healthIdsCol ? parseHealthIds_(row[healthIdsCol - 1]) : [];
    const exerciseFirstEditedAt = exerciseFirstEditedAtCol ? toDate_(row[exerciseFirstEditedAtCol - 1]) : null;
    const exercisesLastEditedAt = exercisesLastEditedAtCol ? toDate_(row[exercisesLastEditedAtCol - 1]) : null;
    const weightEditedAt = weightEditedAtCol ? toDate_(row[weightEditedAtCol - 1]) : null;
    const matchedHealthSession = matchedHealthSessionCol
      ? String(row[matchedHealthSessionCol - 1] || '').trim()
      : '';
    rows.push({
      rowNum: rowNum,
      date: date,
      exercises: exercises,
      bodyweight: bodyweight,
      exerciseSyncedAt: exerciseSyncedAt ? String(exerciseSyncedAt).trim() : '',
      weightSyncedAt: weightSyncedAt ? String(weightSyncedAt).trim() : '',
      healthIds: healthIds,
      exerciseFirstEditedAt: exerciseFirstEditedAt,
      exercisesLastEditedAt: exercisesLastEditedAt,
      weightEditedAt: weightEditedAt,
      matchedHealthSession: matchedHealthSession
    });
  });
  return {
    rows: rows,
    exerciseSyncedAtCol: exerciseSyncedAtCol,
    weightSyncedAtCol: weightSyncedAtCol,
    weightCol: weightCol,
    healthIdsCol: healthIdsCol,
    exerciseFirstEditedAtCol: exerciseFirstEditedAtCol,
    exercisesLastEditedAtCol: exercisesLastEditedAtCol,
    weightEditedAtCol: weightEditedAtCol,
    matchedHealthSessionCol: matchedHealthSessionCol
  };
}

function writeMatchedHealthSession(rowNum, matchedHealthSessionCol, name) {
  if (!matchedHealthSessionCol) return;
  const sheet = getSheet_();
  sheet.getRange(rowNum, matchedHealthSessionCol).setValue(name || '');
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
  return Utilities.formatDate(date, getTz_(), 'yyyy-MM-dd');
}

function markRowExerciseSynced(rowNum, exerciseSyncedAtCol, isoTimestamp) {
  if (!exerciseSyncedAtCol) return;
  const sheet = getSheet_();
  sheet.getRange(rowNum, exerciseSyncedAtCol).setValue(isoTimestamp);
}

function clearRowExerciseSynced(rowNum, exerciseSyncedAtCol) {
  if (!exerciseSyncedAtCol) return;
  const sheet = getSheet_();
  sheet.getRange(rowNum, exerciseSyncedAtCol).setValue('');
}

function markRowWeightSynced(rowNum, weightSyncedAtCol, isoTimestamp) {
  if (!weightSyncedAtCol) return;
  const sheet = getSheet_();
  sheet.getRange(rowNum, weightSyncedAtCol).setValue(isoTimestamp);
}

function clearRowWeightSynced(rowNum, weightSyncedAtCol) {
  if (!weightSyncedAtCol) return;
  const sheet = getSheet_();
  sheet.getRange(rowNum, weightSyncedAtCol).setValue('');
}

// Classify the edited range and apply phase-isolated dirty marking (clear the
// relevant Synced At stamp(s), advance edit timestamps, bump PENDING_DIRTY_KEY).
// Returns true when the row was marked dirty (so syncOnEdit knows to run an
// immediate sync), false on every no-op/early-return path.
function onEditMarkDirty(e) {
  if (!e || !e.range) return false;
  const sheet = e.range.getSheet();
  if (sheet.getSheetId() !== getSheet_().getSheetId()) return false;

  const firstRow = e.range.getRow();
  if (firstRow < 2) return false;
  const lastRow = e.range.getLastRow();
  const firstCol = e.range.getColumn();
  const lastCol = e.range.getLastColumn();

  const { map, headers } = getHeaderMap_(sheet);
  const exerciseSyncedAtCol = map[EXERCISE_SYNCED_AT_COLUMN_HEADER];
  if (!exerciseSyncedAtCol) return false;
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER] || null;
  const exerciseFirstEditedAtCol = map[EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER] || null;
  const exercisesLastEditedAtCol = map[EXERCISES_LAST_EDITED_AT_COLUMN_HEADER] || null;
  const weightEditedAtCol = map[WEIGHT_EDITED_AT_COLUMN_HEADER] || null;
  const dateCol = map[DATE_COLUMN_HEADER] || null;
  const weightCol = map[WEIGHT_COLUMN_HEADER] || null;

  // Classify per phase by walking the cell values (not just column indices).
  // A range that spans empty cells — e.g. pasting only Date + Weight on a
  // new row leaves the exercise columns in between empty — must not be
  // treated as an exercise edit. Only cells with content count. This also
  // implicitly handles the "every cell empty" case (clearing already-blank
  // cells, pasting empty data): no flag gets set and we return early.
  // Deletions of real content are also skipped — use Force Resync to
  // remove a row's datapoint after clearing it.
  //
  // Weight-column edits affect only the weight datapoint and must NOT
  // advance the row's exercise timestamps. Exercise columns (any non-
  // managed column that isn't Date or Weight) advance Exercises Last Edited At.
  // The Date column is metadata that doesn't itself trigger a sync;
  // typing a Date alone on a new row is a no-op until exercise or weight
  // content lands.
  const managedCols = MANAGED_COLUMN_HEADERS.map(h => map[h]).filter(c => c);
  const newValues = e.range.getValues();
  let exerciseRelevant = false;
  let weightRelevant = false;
  const touched = [];
  for (let i = 0; i < newValues.length; i++) {
    for (let j = 0; j < newValues[i].length; j++) {
      const v = newValues[i][j];
      if (v === '' || v === null || v === undefined) continue;
      const c = firstCol + j;
      touched.push({ col: c, row: firstRow + i });
      if (managedCols.indexOf(c) !== -1) continue;
      if (c === dateCol) continue;
      if (c === weightCol) weightRelevant = true;
      else exerciseRelevant = true;
    }
  }
  const desc = describeEditRange_(headers, touched, firstRow, lastRow, firstCol, lastCol);
  if (!exerciseRelevant && !weightRelevant) {
    console.info('syncOnEdit: ' + desc + ' no-op (date-only/empty)');
    return false;
  }

  const phases = [];
  if (exerciseRelevant) phases.push('exercise');
  if (weightRelevant) phases.push('weight');
  console.info('syncOnEdit: ' + desc + ' dirty=[' + phases.join(',') + ']');

  markPendingDirty_();
  // No lock: these are single-cell writes that race safely with an in-flight
  // sync. syncOneRow_'s per-phase concurrent-edit guards re-check at stamp
  // time and defer if our update landed during processing.
  //
  // exerciseFirstEditedAt is sticky-written on exercise-relevant edits only
  // (sets if blank, otherwise leaves it alone). A weight or Date edit must
  // not seed it, otherwise the exercise interval's startTime would be
  // anchored before any exercise content was typed.
  // exercisesLastEditedAt is overwritten only on exercise-relevant edits —
  // it drives the exercise interval's endTime.
  // weightEditedAt is overwritten on every weight-relevant edit — it
  // drives the weight sample time and the weight phase's concurrent-edit
  // guard, so it should reflect the latest weight cell change.
  writeEditMarkers_(sheet, firstRow, lastRow,
    exerciseRelevant ? exerciseSyncedAtCol : null,
    weightRelevant ? weightSyncedAtCol : null,
    exerciseRelevant ? exerciseFirstEditedAtCol : null,
    exerciseRelevant ? exercisesLastEditedAtCol : null,
    weightRelevant ? weightEditedAtCol : null);
  return true;
}

function clearStampColumn_(sheet, firstRow, numRows, col) {
  const range = sheet.getRange(firstRow, col, numRows, 1);
  const values = range.getValues();
  let needsClear = false;
  const blanks = [];
  for (let i = 0; i < numRows; i++) {
    if (values[i][0] !== '') needsClear = true;
    blanks.push(['']);
  }
  if (needsClear) range.setValues(blanks);
}

function writeEditMarkers_(sheet, firstRow, lastRow, exerciseSyncedAtCol, weightSyncedAtCol, exerciseFirstEditedAtCol, exercisesLastEditedAtCol, weightEditedAtCol) {
  const numRows = lastRow - firstRow + 1;
  const nowIso = new Date().toISOString();

  if (exerciseSyncedAtCol) clearStampColumn_(sheet, firstRow, numRows, exerciseSyncedAtCol);
  if (weightSyncedAtCol) clearStampColumn_(sheet, firstRow, numRows, weightSyncedAtCol);

  if (exerciseFirstEditedAtCol) {
    const firstRange = sheet.getRange(firstRow, exerciseFirstEditedAtCol, numRows, 1);
    const firstValues = firstRange.getValues();
    let needsWrite = false;
    const writes = [];
    for (let i = 0; i < numRows; i++) {
      if (firstValues[i][0] === '' || firstValues[i][0] === null) {
        writes.push([nowIso]);
        needsWrite = true;
      } else {
        writes.push([firstValues[i][0]]);
      }
    }
    if (needsWrite) firstRange.setValues(writes);
  }

  if (exercisesLastEditedAtCol) {
    const lastRange = sheet.getRange(firstRow, exercisesLastEditedAtCol, numRows, 1);
    const stamps = [];
    for (let i = 0; i < numRows; i++) stamps.push([nowIso]);
    lastRange.setValues(stamps);
  }

  if (weightEditedAtCol) {
    const weightRange = sheet.getRange(firstRow, weightEditedAtCol, numRows, 1);
    const stamps = [];
    for (let i = 0; i < numRows; i++) stamps.push([nowIso]);
    weightRange.setValues(stamps);
  }
}

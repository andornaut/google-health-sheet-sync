// Module-level caches reset on every Apps Script execution (each invocation
// gets a fresh V8 context), so they amortize repeated lookups within one
// sync pass without leaking state across passes.
let cachedSheet_ = null;
let cachedTz_ = null;

function getSheet_() {
  if (cachedSheet_) {
    return cachedSheet_;
  }
  cachedSheet_ = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  return cachedSheet_;
}

function getTz_() {
  if (cachedTz_) {
    return cachedTz_;
  }
  cachedTz_ = Session.getScriptTimeZone();
  return cachedTz_;
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim()] = i + 1;
  });
  return { headers, map };
}

// Format an edited range using column header names: `Header[row]` for one
// cell, `H1,H2[row]` for multi-column, `Header[r1-r2]` for multi-row. Falls
// back to the range's bounding box when no cells had content (all-empty
// edits like clearing already-blank cells).
function describeEditRange_(
  headers,
  touched,
  firstRow,
  lastRow,
  firstCol,
  lastCol,
) {
  let cells = touched;
  if (cells.length === 0) {
    cells = [];
    for (let r = firstRow; r <= lastRow; r++) {
      for (let c = firstCol; c <= lastCol; c++) {
        cells.push({ col: c, row: r });
      }
    }
  }
  const seen = {};
  const headerList = [];
  let minRow = Infinity,
    maxRow = -Infinity;
  for (let i = 0; i < cells.length; i++) {
    const { col, row } = cells[i];
    const name = String(headers[col - 1] || "").trim() || `col${col}`;
    if (!seen[name]) {
      seen[name] = true;
      headerList.push(name);
    }
    if (row < minRow) {
      minRow = row;
    }
    if (row > maxRow) {
      maxRow = row;
    }
  }
  const rowDesc = minRow === maxRow ? String(minRow) : `${minRow}-${maxRow}`;
  return `${headerList.join(",")}[${rowDesc}]`;
}

function ensureManagedColumns() {
  const sheet = getSheet_();
  const { map } = getHeaderMap_(sheet);
  MANAGED_COLUMN_HEADERS.forEach((header) => {
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
  const { headers, map } = getHeaderMap_(sheet);
  if (!map[DATE_COLUMN_HEADER]) {
    throw new Error(`Missing column: ${DATE_COLUMN_HEADER}`);
  }
  if (!map[WEIGHT_COLUMN_HEADER]) {
    throw new Error(`Missing column: ${WEIGHT_COLUMN_HEADER}`);
  }

  const exerciseSyncedAtCol = map[EXERCISE_SYNCED_AT_COLUMN_HEADER] || null;
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER] || null;
  const healthIdsCol = map[HEALTH_IDS_COLUMN_HEADER] || null;
  const exerciseFirstEditedAtCol =
    map[EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER] || null;
  const exercisesLastEditedAtCol =
    map[EXERCISES_LAST_EDITED_AT_COLUMN_HEADER] || null;
  const weightEditedAtCol = map[WEIGHT_EDITED_AT_COLUMN_HEADER] || null;
  const matchedHealthSessionCol =
    map[MATCHED_HEALTH_SESSION_COLUMN_HEADER] || null;
  const dateCol = map[DATE_COLUMN_HEADER];
  const weightCol = map[WEIGHT_COLUMN_HEADER];

  const exerciseCols = [];
  const seenExerciseNames = {};
  const duplicateExerciseNames = {};
  headers.forEach((h, i) => {
    const name = String(h).trim();
    if (!name) {
      return;
    }
    if (name === DATE_COLUMN_HEADER || name === WEIGHT_COLUMN_HEADER) {
      return;
    }
    if (MANAGED_COLUMN_HEADERS.indexOf(name) !== -1) {
      return;
    }
    if (seenExerciseNames[name]) {
      duplicateExerciseNames[name] = true;
      return;
    }
    seenExerciseNames[name] = true;
    exerciseCols.push({ col: i + 1, name });
  });
  const duplicates = Object.keys(duplicateExerciseNames);
  if (duplicates.length > 0) {
    throw new Error(
      `Duplicate exercise column header(s): ${duplicates
        .sort()
        .join(", ")}. Each exercise column header must be unique.`,
    );
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      allHealthIds: [],
      allMatchedSessions: [],
      exerciseFirstEditedAtCol,
      exerciseSyncedAtCol,
      exercisesLastEditedAtCol,
      healthIdsCol,
      matchedHealthSessionCol,
      rows: [],
      weightCol,
      weightEditedAtCol,
      weightSyncedAtCol,
    };
  }

  const width = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();

  // Every column that is not Date, not Weight, and not managed, whether or not
  // it has a header. Deliberately broader than exerciseCols (which requires a
  // non-blank header) because this feeds hasExerciseText, whose only job is to
  // answer "is there any text here at all". Blanking or deleting an exercise
  // column's HEADER would otherwise make every historical row look emptied and
  // hand the backstop's reconciliation a whole sheet of datapoints to delete.
  // The cost is that a scratch column with notes in it suppresses backstop
  // reconciliation for that row; single-cell clears still reconcile via onEdit,
  // and "Resync selected rows" still works.
  const managedColNums = MANAGED_COLUMN_HEADERS.map((h) => map[h]).filter(
    (c) => c,
  );
  const textCols = [];
  for (let c = 1; c <= width; c++) {
    if (c === dateCol || c === weightCol) {
      continue;
    }
    if (managedColNums.indexOf(c) !== -1) {
      continue;
    }
    textCols.push(c);
  }
  const hasText = (v) =>
    v !== null && v !== undefined && String(v).trim() !== "";

  const rows = [];
  // Every data row's Created Health IDs, including the rows dropped below for a
  // blank/unparseable Date. A dated row losing its Date must not make its
  // datapoints look untracked: orphan reconciliation would delete them and
  // foreign matching would offer them to another row as a "foreign" session.
  // Callers that need ownership (not row content) use this, not `rows`.
  const allHealthIds = [];
  // Companion to allHealthIds, same rationale: `{ name, rowNum }` for every data
  // row carrying a Matched Health Session, including dropped ones. A dropped row
  // can never be ready to sync, so the foreign session it borrowed must stay
  // excluded, otherwise another row claims the same session's interval.
  const allMatchedSessions = [];
  values.forEach((row, idx) => {
    const rowNum = idx + 2;
    const healthIds = healthIdsCol
      ? parseHealthIds_(row[healthIdsCol - 1])
      : [];
    healthIds.forEach((n) => allHealthIds.push(n));
    if (matchedHealthSessionCol) {
      const matched = String(row[matchedHealthSessionCol - 1] || "").trim();
      if (matched) {
        allMatchedSessions.push({ name: matched, rowNum });
      }
    }
    const dateVal = row[dateCol - 1];
    if (!dateVal) {
      return;
    }
    const date = toDate_(dateVal);
    if (!date) {
      return;
    }
    const exercises = [];
    exerciseCols.forEach((c) => {
      const entries = parseExerciseCell(row[c.col - 1]);
      if (entries.length > 0) {
        exercises.push({ entries, name: c.name });
      }
    });
    const bodyweight = parseBodyweight(row[weightCol - 1]);
    const exerciseSyncedAt = exerciseSyncedAtCol
      ? row[exerciseSyncedAtCol - 1]
      : "";
    const weightSyncedAt = weightSyncedAtCol ? row[weightSyncedAtCol - 1] : "";
    const exerciseFirstEditedAt = exerciseFirstEditedAtCol
      ? toDate_(row[exerciseFirstEditedAtCol - 1])
      : null;
    const exercisesLastEditedAt = exercisesLastEditedAtCol
      ? toDate_(row[exercisesLastEditedAtCol - 1])
      : null;
    const weightEditedAt = weightEditedAtCol
      ? toDate_(row[weightEditedAtCol - 1])
      : null;
    const matchedHealthSession = matchedHealthSessionCol
      ? String(row[matchedHealthSessionCol - 1] || "").trim()
      : "";
    rows.push({
      bodyweight,
      date,
      exerciseFirstEditedAt,
      exerciseSyncedAt: exerciseSyncedAt ? String(exerciseSyncedAt).trim() : "",
      exercises,
      exercisesLastEditedAt,
      // Raw-text presence, NOT parse results. "The parser produced nothing" and
      // "the cell is empty" are different claims, and only the second one means
      // the user cleared something. selectStaleDataPointRows_ deletes on that
      // second claim, so it must not be inferred from the first: an unparseable
      // cell (a reformat to "185 lb", a bodyweight outside the plausible
      // bounds, a typo) parses to nothing while plainly still holding data.
      hasExerciseText: textCols.some((c) => hasText(row[c - 1])),

      hasWeightText: hasText(row[weightCol - 1]),

      healthIds,

      matchedHealthSession,

      rowNum,

      weightEditedAt,
      weightSyncedAt: weightSyncedAt ? String(weightSyncedAt).trim() : "",
    });
  });
  return {
    allHealthIds,
    allMatchedSessions,
    exerciseFirstEditedAtCol,
    exerciseSyncedAtCol,
    exercisesLastEditedAtCol,
    healthIdsCol,
    matchedHealthSessionCol,
    rows,
    weightCol,
    weightEditedAtCol,
    weightSyncedAtCol,
  };
}

function writeMatchedHealthSession(rowNum, matchedHealthSessionCol, name) {
  if (!matchedHealthSessionCol) {
    return;
  }
  const sheet = getSheet_();
  sheet.getRange(rowNum, matchedHealthSessionCol).setValue(name || "");
}

function parseHealthIds_(raw) {
  const text = String(raw || "").trim();
  if (!text) {
    return [];
  }
  try {
    const parsed = JSON.parse(text);
    return Array.isArray(parsed)
      ? parsed.filter((s) => typeof s === "string")
      : [];
  } catch (err) {
    console.warn(`parseHealthIds_: could not parse "${text}": ${err}`);
    return [];
  }
}

function writeHealthIds(rowNum, healthIdsCol, names) {
  const sheet = getSheet_();
  sheet.getRange(rowNum, healthIdsCol).setValue(JSON.stringify(names));
}

function toDate_(value) {
  if (value instanceof Date) {
    return value;
  }
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function ymd(date) {
  return Utilities.formatDate(date, getTz_(), "yyyy-MM-dd");
}

function markRowExerciseSynced(rowNum, exerciseSyncedAtCol, isoTimestamp) {
  if (!exerciseSyncedAtCol) {
    return;
  }
  const sheet = getSheet_();
  sheet.getRange(rowNum, exerciseSyncedAtCol).setValue(isoTimestamp);
}

function clearRowExerciseSynced(rowNum, exerciseSyncedAtCol) {
  if (!exerciseSyncedAtCol) {
    return;
  }
  const sheet = getSheet_();
  sheet.getRange(rowNum, exerciseSyncedAtCol).setValue("");
}

function markRowWeightSynced(rowNum, weightSyncedAtCol, isoTimestamp) {
  if (!weightSyncedAtCol) {
    return;
  }
  const sheet = getSheet_();
  sheet.getRange(rowNum, weightSyncedAtCol).setValue(isoTimestamp);
}

function clearRowWeightSynced(rowNum, weightSyncedAtCol) {
  if (!weightSyncedAtCol) {
    return;
  }
  const sheet = getSheet_();
  sheet.getRange(rowNum, weightSyncedAtCol).setValue("");
}

// Classify the edited range and apply phase-isolated dirty marking (clear the
// relevant Synced At stamp(s), advance edit timestamps, bump PENDING_DIRTY_KEY).
// Returns true when the row was marked dirty (so syncOnEdit knows to run an
// immediate sync), false on every no-op/early-return path.
function onEditMarkDirty(e) {
  if (!e || !e.range) {
    return false;
  }
  const sheet = e.range.getSheet();
  if (sheet.getSheetId() !== getSheet_().getSheetId()) {
    return false;
  }

  const firstRow = e.range.getRow();
  if (firstRow < 2) {
    return false;
  }
  const rangeLastRow = e.range.getLastRow();
  const firstCol = e.range.getColumn();
  const lastCol = e.range.getLastColumn();
  // Bound the read to the data range. A range's LAST row can sit far past the
  // data while its content sits in an earlier row (Ctrl+Shift+Down from row 2),
  // and reading tens of thousands of cells inside an onEdit trigger is the cost
  // this avoids. It is NOT what keeps blank rows from being stamped: marking is
  // per row and only rows that actually held content are marked, and a row past
  // the data holds nothing by definition. A true whole-column selection starts
  // at row 1 and has already exited at the firstRow < 2 guard above.
  const lastRow = Math.min(rangeLastRow, sheet.getLastRow());
  if (lastRow < firstRow) {
    return false;
  }

  const { headers, map } = getHeaderMap_(sheet);
  const exerciseSyncedAtCol = map[EXERCISE_SYNCED_AT_COLUMN_HEADER];
  if (!exerciseSyncedAtCol) {
    return false;
  }
  const weightSyncedAtCol = map[WEIGHT_SYNCED_AT_COLUMN_HEADER] || null;
  const exerciseFirstEditedAtCol =
    map[EXERCISE_FIRST_EDITED_AT_COLUMN_HEADER] || null;
  const exercisesLastEditedAtCol =
    map[EXERCISES_LAST_EDITED_AT_COLUMN_HEADER] || null;
  const weightEditedAtCol = map[WEIGHT_EDITED_AT_COLUMN_HEADER] || null;
  const dateCol = map[DATE_COLUMN_HEADER] || null;
  const weightCol = map[WEIGHT_COLUMN_HEADER] || null;

  // Classify per phase by walking the cell values (not just column indices).
  // A range that spans empty cells (e.g. pasting only Date + Weight on a
  // new row leaves the exercise columns in between empty) must not be
  // treated as an exercise edit. Only cells with content count.
  //
  // The one exception is a SINGLE-CELL clear of real content, detected via
  // e.oldValue (Apps Script supplies it for single-cell edits only). That is a
  // latency optimization, not the correctness path: it reaches the delete paths
  // in syncOneRow_ within seconds. Correctness for every other shape of clear
  // (multi-cell, paste, a mixed range that blanks one cell while writing
  // another) belongs to selectStaleDataPointRows_ in the backstop, which
  // compares recorded datapoints against current content instead of trying to
  // infer intent from an event that does not carry the old values.
  //
  // Weight-column edits affect only the weight datapoint and must NOT
  // advance the row's exercise timestamps. Exercise columns (any non-
  // managed column that isn't Date or Weight) advance Exercises Last Edited At.
  // The Date column is metadata that doesn't itself trigger a sync;
  // typing a Date alone on a new row is a no-op until exercise or weight
  // content lands.
  const managedCols = MANAGED_COLUMN_HEADERS.map((h) => map[h]).filter(
    (c) => c,
  );
  const numRows = lastRow - firstRow + 1;
  // Read the CLAMPED range, not e.range: a Ctrl+Shift+Down selection from row 2
  // reports every row of the sheet, and pulling all of them into memory inside
  // an onEdit trigger costs more than the one read this replaces.
  const newValues = sheet
    .getRange(firstRow, firstCol, numRows, lastCol - firstCol + 1)
    .getValues();
  const isEmptyValue = (v) => v === "" || v === null || v === undefined;
  const singleCell = firstRow === rangeLastRow && firstCol === lastCol;
  const clearedContent =
    singleCell &&
    isEmptyValue(newValues[0] && newValues[0][0]) &&
    !isEmptyValue(e.oldValue);
  // Relevance is tracked PER ROW, not per range. A multi-row edit is not
  // uniform: pasting a block whose content sits in its first row leaves the
  // rest of the range empty, and stamping those rows would clear their Synced
  // At and overwrite their Exercises Last Edited At with now, inflating the
  // recorded endTime of a row the user never touched and forcing a needless
  // re-sync. The per-cell walk below already knows which rows had content.
  const exerciseRows = new Set();
  const weightRows = new Set();
  const touched = [];
  for (let i = 0; i < numRows; i++) {
    for (let j = 0; j < newValues[i].length; j++) {
      const v = newValues[i][j];
      if (isEmptyValue(v) && !clearedContent) {
        continue;
      }
      const c = firstCol + j;
      const rowNum = firstRow + i;
      touched.push({ col: c, row: rowNum });
      if (managedCols.indexOf(c) !== -1) {
        continue;
      }
      if (c === dateCol) {
        continue;
      }
      if (c === weightCol) {
        weightRows.add(rowNum);
        continue;
      }
      // Exercise-relevant only if this is a real exercise column. readRows
      // skips blank-header columns when building exerciseCols, so a scratch
      // column parked to the right of Weight contributes no content, so marking
      // the row dirty for it would advance Exercises Last Edited At, stretch
      // the 'edit' interval's endTime, and churn the datapoint for an edit
      // that changes nothing the sync reads. Blankness is decided exactly as
      // readRows decides it (`String(h).trim()`), so the two can't disagree on
      // a header cell holding a falsy-but-real value like 0; a column past the
      // header row's width has no header at all and counts as blank.
      const headerName =
        c - 1 < headers.length ? String(headers[c - 1]).trim() : "";
      if (headerName) {
        exerciseRows.add(rowNum);
      }
    }
  }
  const desc = describeEditRange_(
    headers,
    touched,
    firstRow,
    lastRow,
    firstCol,
    lastCol,
  );
  if (exerciseRows.size === 0 && weightRows.size === 0) {
    console.info(`syncOnEdit: ${desc} no-op (date-only/empty)`);
    return false;
  }

  const phases = [];
  if (exerciseRows.size > 0) {
    phases.push("exercise");
  }
  if (weightRows.size > 0) {
    phases.push("weight");
  }
  console.info(
    `syncOnEdit: ${desc}${
      clearedContent ? " cleared" : ""
    } dirty=[${phases.join(",")}]`,
  );

  markPendingDirty_();
  // No lock: these are single-cell writes that race safely with an in-flight
  // sync. syncOneRow_'s per-phase concurrent-edit guards re-check at stamp
  // time and defer if our update landed during processing.
  //
  // exerciseFirstEditedAt is sticky-written on exercise-relevant edits only
  // (sets if blank, otherwise leaves it alone). A weight or Date edit must
  // not seed it, otherwise the exercise interval's startTime would be
  // anchored before any exercise content was typed.
  // exercisesLastEditedAt is overwritten only on exercise-relevant edits: it
  // drives the exercise interval's endTime.
  // weightEditedAt is overwritten on every weight-relevant edit: it
  // drives the weight sample time and the weight phase's concurrent-edit
  // guard, so it should reflect the latest weight cell change.
  writeEditMarkers_(sheet, {
    exerciseFirstEditedAtCol,
    exerciseRows,
    exerciseSyncedAtCol,
    exercisesLastEditedAtCol,
    weightEditedAtCol,
    weightRows,
    weightSyncedAtCol,
  });
  return true;
}

// The block of `col` spanned by `rows`, as [range, values, offset]. Reading the
// spanned block rather than each row keeps this at one read per column for the
// contiguous ranges edits almost always are, while never touching a row outside
// `rows` (values for those are read and written back unchanged).
function rowBlock_(sheet, col, rows) {
  let first = Infinity;
  let last = -Infinity;
  rows.forEach((r) => {
    if (r < first) {
      first = r;
    }
    if (r > last) {
      last = r;
    }
  });
  const range = sheet.getRange(first, col, last - first + 1, 1);
  return { first, range, values: range.getValues() };
}

// Overwrite `col` with `value` for the rows in `rows`, leaving every other row
// in the spanned block exactly as it was. Skips the write entirely when nothing
// would change, so clearing an already-clear stamp costs no write.
function stampRows_(sheet, col, rows, value) {
  const block = rowBlock_(sheet, col, rows);
  let changed = false;
  rows.forEach((r) => {
    const i = r - block.first;
    if (block.values[i][0] === value) {
      return;
    }
    block.values[i][0] = value;
    changed = true;
  });
  if (changed) {
    block.range.setValues(block.values);
  }
}

// Sticky variant: write `value` into `col` only where the cell is currently
// blank, for the rows in `rows`. Used for Exercise First Edited At, which must
// keep the FIRST edit's timestamp so the exercise interval's startTime doesn't
// drift forward as more sets are typed in.
function seedRows_(sheet, col, rows, value) {
  const block = rowBlock_(sheet, col, rows);
  let changed = false;
  rows.forEach((r) => {
    const i = r - block.first;
    const current = block.values[i][0];
    if (current !== "" && current !== null && current !== undefined) {
      return;
    }
    block.values[i][0] = value;
    changed = true;
  });
  if (changed) {
    block.range.setValues(block.values);
  }
}

// Apply the dirty markers for one edit. `marks.exerciseRows` / `marks.weightRows`
// are the row numbers that actually had content in the relevant columns, so each
// phase touches only its own rows and only its own columns:
//   - a weight edit must not advance the exercise timestamps, or a bodyweight
//     change would drag the exercise interval's endTime forward;
//   - a row inside the edited range that had no content in a phase's columns
//     must not be marked for that phase at all.
function writeEditMarkers_(sheet, marks) {
  const nowIso = new Date().toISOString();

  if (marks.exerciseRows.size > 0) {
    if (marks.exerciseSyncedAtCol) {
      stampRows_(sheet, marks.exerciseSyncedAtCol, marks.exerciseRows, "");
    }
    // Sticky: the first exercise edit anchors the interval's startTime.
    if (marks.exerciseFirstEditedAtCol) {
      seedRows_(
        sheet,
        marks.exerciseFirstEditedAtCol,
        marks.exerciseRows,
        nowIso,
      );
    }
    // Advances every time: drives the interval's endTime and the exercise
    // phase's concurrent-edit guard.
    if (marks.exercisesLastEditedAtCol) {
      stampRows_(
        sheet,
        marks.exercisesLastEditedAtCol,
        marks.exerciseRows,
        nowIso,
      );
    }
  }

  if (marks.weightRows.size > 0) {
    if (marks.weightSyncedAtCol) {
      stampRows_(sheet, marks.weightSyncedAtCol, marks.weightRows, "");
    }
    // Advances every time: drives the weight sample time and the weight phase's
    // concurrent-edit guard, so it must reflect the latest weight cell change.
    if (marks.weightEditedAtCol) {
      stampRows_(sheet, marks.weightEditedAtCol, marks.weightRows, nowIso);
    }
  }
}

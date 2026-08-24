// Orchestration tests for the stateful glue that the pure-helper tests in
// Parser.test.gs can't reach: onEditMarkDirty's column-aware dirty marking,
// syncDirtyRows' dirty-flag lifecycle (drain / error-retained / concurrent-edit),
// syncOneRow_'s phase dispatch + idempotency skip, and weight orphan
// reconciliation. Runs against the in-memory Apps Script fakes in Harness.gs,
// which withSyncTestHarness_ swaps in for the real services and hands over as
// H; the Health API functions are stubbed per-test via globalThis, the same
// pattern resolveForeignMatches_'s tests use for listStrengthOnDate.
function runSyncTests() {
  reportTestResults_(
    "Orchestration tests",
    withSyncTestHarness_(runSyncTestsBody_),
  );
}

// Returns the PASS/FAIL list, and takes the harness handle rather than reaching
// for a global. Reporting happens outside withSyncTestHarness_ because the fake
// SpreadsheetApp's getUi throws by design.
function runSyncTestsBody_(H) {
  const results = [];
  const t = (name, fn) => {
    try {
      fn();
      results.push(`PASS ${name}`);
    } catch (err) {
      results.push(`FAIL ${name}: ${(err && err.stack) || err}`);
    }
  };
  // JSON.stringify serializes NaN and Infinity as null, so a plain stringify
  // comparison cannot tell an expected null from a leaked NaN. The replacer
  // tags non-finite numbers so an assertion on a timestamp or a datapoint
  // payload fails on a NaN instead of matching a null the fixture meant.
  const show = (v) =>
    JSON.stringify(v, (_k, val) =>
      typeof val === "number" && !Number.isFinite(val) ? `#${val}` : val,
    );
  const eq = (a, b, msg) => {
    const sa = show(a),
      sb = show(b);
    if (sa !== sb) {
      throw new Error(`${msg || "mismatch"} expected ${sb} got ${sa}`);
    }
  };
  const ok = (cond, msg) => {
    if (!cond) {
      throw new Error(msg || "expected truthy");
    }
  };

  const SHEET = H.sheet;
  const PROPS = H.scriptProps;
  const LOCK = H.lockState;
  const ACTIVE = H.activeSheetRef;

  // Shared column layout for the fake sheet. 1-based column numbers in COL so
  // they can be passed straight to getRange.
  const HEADERS = [
    "Date",
    "Weight",
    "Bench",
    "Exercise Synced At",
    "Weight Synced At",
    "Created Health IDs",
    "Exercise First Edited At",
    "Exercises Last Edited At",
    "Weight Edited At",
    "Matched Health Session",
  ];
  const COL = {};
  HEADERS.forEach((h, i) => {
    COL[h] = i + 1;
  });

  const TOASTS = H.toasts;

  // A test needing a non-standard header row calls resetGrid directly; every
  // other one goes through reset. Both must clear the same harness state, so
  // there is one body and reset only supplies the grid.
  const resetGrid = (grid) => {
    SHEET._setGrid(grid);
    SHEET._setSelection(null);
    ACTIVE.sheet = SHEET;
    PROPS._clear();
    TOASTS.length = 0;
    LOCK.held = false;
  };
  const reset = (dataRows) =>
    resetGrid([HEADERS.slice()].concat(dataRows || []));
  const cell = (header) => SHEET.getRange(2, COL[header]).getValue();

  // Swap a set of globals (Health API stubs), run fn, restore. Returns fn's value.
  const withStubs = (stubs, fn) => {
    const saved = {};
    Object.keys(stubs).forEach((k) => {
      saved[k] = globalThis[k];
      globalThis[k] = stubs[k];
    });
    try {
      return fn();
    } finally {
      Object.keys(saved).forEach((k) => {
        globalThis[k] = saved[k];
      });
    }
  };
  const NO_FOREIGN = { listStrengthOnDate: () => [] };

  // ---- onEditMarkDirty: column-aware dirty marking -----------------------

  t(
    "onEditMarkDirty exercise edit clears exercise stamp, seeds edit timestamps, leaves weight alone",
    () => {
      reset([
        ["2026-01-15", "", "135x5", "PREV-EX", "PREV-WT", "", "", "", "", ""],
      ]);
      const marked = onEditMarkDirty({ range: SHEET.getRange(2, COL.Bench) });
      ok(marked === true, "returns true");
      eq(cell("Exercise Synced At"), "", "exercise synced cleared");
      eq(cell("Weight Synced At"), "PREV-WT", "weight synced untouched");
      ok(cell("Exercise First Edited At") !== "", "first edited seeded");
      ok(cell("Exercises Last Edited At") !== "", "last edited seeded");
      eq(cell("Weight Edited At"), "", "weight edited untouched");
      ok(PROPS.getProperty("pendingDirty") !== null, "pending flag set");
    },
  );

  t(
    "onEditMarkDirty weight edit clears weight stamp, sets weight edited, leaves exercise timestamps alone",
    () => {
      reset([
        ["2026-01-15", "185", "", "PREV-EX", "PREV-WT", "", "", "", "", ""],
      ]);
      const marked = onEditMarkDirty({ range: SHEET.getRange(2, COL.Weight) });
      ok(marked === true, "returns true");
      eq(cell("Weight Synced At"), "", "weight synced cleared");
      eq(cell("Exercise Synced At"), "PREV-EX", "exercise synced untouched");
      ok(cell("Weight Edited At") !== "", "weight edited set");
      eq(
        cell("Exercise First Edited At"),
        "",
        "exercise first edited NOT seeded by weight edit",
      );
      eq(
        cell("Exercises Last Edited At"),
        "",
        "exercise last edited NOT advanced by weight edit",
      );
    },
  );

  // Clearing real content is an edit: it must reach the delete paths in
  // syncOneRow_ (bodyweight cleared -> DELETE, exercises cleared ->
  // delete-only). oldValue is supplied for single-cell edits only, which is
  // what lets a real clear be told apart from clearing an already-blank cell.
  t(
    "onEditMarkDirty single-cell clear of an exercise value marks the row dirty",
    () => {
      reset([["2026-01-15", "", "", "PREV-EX", "PREV-WT", "", "", "", "", ""]]);
      const marked = onEditMarkDirty({
        oldValue: "135x5x3",
        range: SHEET.getRange(2, COL.Bench),
      });
      ok(marked === true, "returns true");
      eq(cell("Exercise Synced At"), "", "exercise synced cleared");
      eq(cell("Weight Synced At"), "PREV-WT", "weight synced untouched");
      ok(cell("Exercises Last Edited At") !== "", "last edited advanced");
    },
  );

  t(
    "onEditMarkDirty single-cell clear of the bodyweight marks the row dirty",
    () => {
      reset([["2026-01-15", "", "", "PREV-EX", "PREV-WT", "", "", "", "", ""]]);
      const marked = onEditMarkDirty({
        oldValue: "185",
        range: SHEET.getRange(2, COL.Weight),
      });
      ok(marked === true, "returns true");
      eq(cell("Weight Synced At"), "", "weight synced cleared");
      eq(cell("Exercise Synced At"), "PREV-EX", "exercise synced untouched");
      ok(cell("Weight Edited At") !== "", "weight edited set");
    },
  );

  t("onEditMarkDirty clearing an already-blank cell stays a no-op", () => {
    reset([["2026-01-15", "", "", "PREV-EX", "PREV-WT", "", "", "", "", ""]]);
    ok(
      onEditMarkDirty({ oldValue: "", range: SHEET.getRange(2, COL.Bench) }) ===
        false,
      "no oldValue content",
    );
    ok(
      onEditMarkDirty({ range: SHEET.getRange(2, COL.Bench) }) === false,
      "no oldValue at all",
    );
    eq(cell("Exercise Synced At"), "PREV-EX", "stamp untouched");
  });

  // A multi-cell edit carries no oldValue, so onEditMarkDirty cannot tell a
  // clear from a range that was already blank and deliberately does not guess.
  // Reconciling a cleared row is selectStaleDataPointRows_'s job in the
  // backstop, which reads state instead of inferring intent; see its tests.
  t(
    "onEditMarkDirty multi-cell clear stays a no-op (the backstop reconciles it)",
    () => {
      reset([["2026-01-15", "", "", "PREV-EX", "PREV-WT", "", "", "", "", ""]]);
      const range = SHEET.getRange(2, COL.Weight, 1, 2); // Weight + Bench
      ok(onEditMarkDirty({ range }) === false, "returns false");
      eq(cell("Exercise Synced At"), "PREV-EX", "stamps untouched");
      eq(cell("Weight Synced At"), "PREV-WT", "stamps untouched");
      ok(PROPS.getProperty("pendingDirty") === null, "no pending flag");
    },
  );

  // Clearing a whole row leaves it genuinely empty. Writing markers back into
  // it would leave a permanently non-blank row still counting toward
  // getLastRow(), the phantom row the clamp exists to prevent.
  t(
    "onEditMarkDirty writes nothing back into a row cleared of everything",
    () => {
      reset([
        [
          "2026-01-15",
          "185",
          "135x5",
          "PREV-EX",
          "PREV-WT",
          '["ex/E1"]',
          "",
          "",
          "",
          "",
        ],
        [
          "2026-01-16",
          "186",
          "225x5",
          "PREV-EX",
          "PREV-WT",
          "[]",
          "",
          "",
          "",
          "",
        ],
      ]);
      for (let c = 1; c <= HEADERS.length; c++) {
        SHEET.getRange(2, c).setValue("");
      }
      ok(
        onEditMarkDirty({ range: SHEET.getRange(2, 1, 1, HEADERS.length) }) ===
          false,
        "no-op",
      );
      eq(
        SHEET.getRange(2, 1, 1, HEADERS.length).getValues()[0].join("|"),
        new Array(HEADERS.length).fill("").join("|"),
        "the cleared row stayed empty",
      );
      eq(
        SHEET.getLastRow(),
        3,
        "still one trailing data row, no phantom row 2",
      );
    },
  );

  // The rule that makes the clear handling safe: a range carrying ANY content
  // is classified by the cells that have content, so pasting Date + Weight
  // across the exercise columns between them is a weight edit only.
  t(
    "onEditMarkDirty a paste spanning blank exercise columns is not an exercise edit",
    () => {
      reset([["", "", "", "PREV-EX", "PREV-WT", "", "", "", "", ""]]);
      SHEET.getRange(2, COL.Date).setValue("2026-01-15");
      SHEET.getRange(2, COL.Weight).setValue("185");
      ok(
        onEditMarkDirty({ range: SHEET.getRange(2, COL.Date, 1, 3) }) === true,
        "Date..Bench pasted",
      );
      eq(cell("Weight Synced At"), "", "weight synced cleared");
      eq(cell("Exercise Synced At"), "PREV-EX", "exercise synced untouched");
      eq(
        cell("Exercises Last Edited At"),
        "",
        "blank exercise cell did not advance the edit timestamp",
      );
    },
  );

  // A selection can reach far past the data (Ctrl+Shift+Down from the first
  // data row). Only rows that held content may be marked, and the data range
  // must not grow. Asserted as an end property: per-row marking is what
  // guarantees it, and onEditMarkDirty's read-size clamp is unobservable here.
  t(
    "onEditMarkDirty clamps a range reaching past the data and skips one entirely below it",
    () => {
      reset([
        ["2026-01-15", "", "135x5", "PREV-EX", "PREV-WT", "", "", "", "", ""],
      ]);
      ok(
        onEditMarkDirty({ range: SHEET.getRange(2, COL.Bench, 999, 1) }) ===
          true,
        "data row still marked",
      );
      eq(cell("Exercise Synced At"), "", "the one data row was cleared");
      eq(
        SHEET.getRange(3, COL["Exercises Last Edited At"]).getValue(),
        "",
        "no marker written into the blank row below the data",
      );
      ok(
        onEditMarkDirty({ range: SHEET.getRange(50, COL.Bench, 10, 1) }) ===
          false,
        "a multi-row range entirely below the data is a no-op",
      );
      eq(SHEET.getLastRow(), 2, "the data range did not grow");
    },
  );

  // readRows ignores blank-header columns, so a scratch column parked to the
  // right of Weight has no exercise content to sync. Marking the row dirty for
  // it would stretch the 'edit' interval and recreate the datapoint for nothing.
  t("onEditMarkDirty ignores a blank-header scratch column", () => {
    const scratchCol = HEADERS.length + 1;
    resetGrid([
      HEADERS.concat([""]),
      [
        "2026-01-15",
        "",
        "135x5",
        "PREV-EX",
        "PREV-WT",
        "",
        "",
        "",
        "",
        "",
        "note to self",
      ],
    ]);
    ok(
      onEditMarkDirty({ range: SHEET.getRange(2, scratchCol) }) === false,
      "typing there is a no-op",
    );
    ok(
      onEditMarkDirty({
        oldValue: "note to self",
        range: SHEET.getRange(2, scratchCol),
      }) === false,
      "clearing it is a no-op too",
    );
    eq(cell("Exercise Synced At"), "PREV-EX", "exercise stamp untouched");
    eq(cell("Exercises Last Edited At"), "", "edit timestamp not advanced");
  });

  // Blankness must be decided the same way readRows decides it, or a header
  // cell holding a falsy-but-real value becomes a column whose content is read
  // by readRows while its edits never mark the row dirty.
  t(
    "onEditMarkDirty treats a numeric-zero header as a real exercise column",
    () => {
      const headers = HEADERS.slice();
      headers[2] = 0; // the exercise column's header is the number 0
      resetGrid([
        headers,
        ["2026-01-15", "", "135x5", "PREV-EX", "PREV-WT", "", "", "", "", ""],
      ]);
      ok(
        onEditMarkDirty({ range: SHEET.getRange(2, 3) }) === true,
        "marks the row dirty",
      );
      eq(
        SHEET.getRange(2, COL["Exercise Synced At"]).getValue(),
        "",
        "exercise stamp cleared",
      );
    },
  );

  t("onEditMarkDirty date-only edit is a no-op", () => {
    reset([["2026-01-15", "", "", "", "", "", "", "", "", ""]]);
    const marked = onEditMarkDirty({ range: SHEET.getRange(2, COL.Date) });
    ok(marked === false, "returns false");
    ok(PROPS.getProperty("pendingDirty") === null, "no pending flag");
  });

  t("onEditMarkDirty ignores edits on a different sheet", () => {
    reset([["2026-01-15", "185", "", "", "", "", "", "", "", ""]]);
    const otherSheetRange = { getSheet: () => ({ getSheetId: () => 99999 }) };
    ok(
      onEditMarkDirty({ range: otherSheetRange }) === false,
      "returns false for foreign sheet",
    );
  });

  // A multi-row edit is not uniform. Pasting a block whose content sits in its
  // first row leaves the rest of the range empty, and those rows must be left
  // alone: clearing their Synced At forces a needless re-sync, and overwriting
  // their Exercises Last Edited At inflates the recorded endTime of a row the
  // user never touched (up to the 2h cap, for a row still on its own date).
  t(
    "onEditMarkDirty marks only the rows that had content, not the whole range",
    () => {
      reset([
        [
          "2026-01-15",
          "",
          "",
          "PREV-EX",
          "PREV-WT",
          "",
          "FIRST-2",
          "LAST-2",
          "",
          "",
        ],
        [
          "2026-01-16",
          "",
          "",
          "PREV-EX",
          "PREV-WT",
          "",
          "FIRST-3",
          "LAST-3",
          "",
          "",
        ],
        [
          "2026-01-17",
          "",
          "",
          "PREV-EX",
          "PREV-WT",
          "",
          "FIRST-4",
          "LAST-4",
          "",
          "",
        ],
      ]);
      SHEET.getRange(2, COL.Bench).setValue("135x5"); // only row 2 gets content
      ok(
        onEditMarkDirty({ range: SHEET.getRange(2, COL.Bench, 3, 1) }) === true,
        "marked",
      );
      eq(
        SHEET.getRange(2, COL["Exercise Synced At"]).getValue(),
        "",
        "row 2 dirtied",
      );
      ok(
        SHEET.getRange(2, COL["Exercises Last Edited At"]).getValue() !==
          "LAST-2",
        "row 2 advanced",
      );
      for (const r of [3, 4]) {
        eq(
          SHEET.getRange(r, COL["Exercise Synced At"]).getValue(),
          "PREV-EX",
          `row ${r} stayed synced`,
        );
        eq(
          SHEET.getRange(r, COL["Exercises Last Edited At"]).getValue(),
          `LAST-${r}`,
          `row ${r} kept its real last-edit time`,
        );
        eq(
          SHEET.getRange(r, COL["Exercise First Edited At"]).getValue(),
          `FIRST-${r}`,
          `row ${r} kept its interval anchor`,
        );
      }
    },
  );

  // The marked rows need not be contiguous: pasting a block with a blank middle
  // row leaves a gap. The marker write reads the spanned block in one call, so
  // the rows inside the span that are NOT marked have to be written back
  // unchanged rather than swept up with their neighbours.
  t(
    "onEditMarkDirty leaves an unmarked row sitting between two marked ones alone",
    () => {
      reset([
        [
          "2026-01-15",
          "",
          "",
          "PREV-EX",
          "PREV-WT",
          "",
          "FIRST-2",
          "LAST-2",
          "",
          "",
        ],
        [
          "2026-01-16",
          "",
          "",
          "PREV-EX",
          "PREV-WT",
          "",
          "FIRST-3",
          "LAST-3",
          "",
          "",
        ],
        [
          "2026-01-17",
          "",
          "",
          "PREV-EX",
          "PREV-WT",
          "",
          "FIRST-4",
          "LAST-4",
          "",
          "",
        ],
      ]);
      SHEET.getRange(2, COL.Bench).setValue("135x5"); // row 3 deliberately left blank
      SHEET.getRange(4, COL.Bench).setValue("225x5");
      ok(
        onEditMarkDirty({ range: SHEET.getRange(2, COL.Bench, 3, 1) }) === true,
        "marked",
      );
      eq(
        SHEET.getRange(2, COL["Exercise Synced At"]).getValue(),
        "",
        "row 2 dirtied",
      );
      eq(
        SHEET.getRange(4, COL["Exercise Synced At"]).getValue(),
        "",
        "row 4 dirtied",
      );
      eq(
        SHEET.getRange(3, COL["Exercise Synced At"]).getValue(),
        "PREV-EX",
        "the gap row stayed synced",
      );
      eq(
        SHEET.getRange(3, COL["Exercises Last Edited At"]).getValue(),
        "LAST-3",
        "gap row kept its last-edit time",
      );
      eq(
        SHEET.getRange(3, COL["Exercise First Edited At"]).getValue(),
        "FIRST-3",
        "gap row kept its anchor",
      );
    },
  );

  t("onEditMarkDirty marks only the rows whose bodyweight changed", () => {
    reset([
      ["2026-01-15", "", "", "PREV-EX", "PREV-WT", "", "", "", "WEDIT-2", ""],
      ["2026-01-16", "", "", "PREV-EX", "PREV-WT", "", "", "", "WEDIT-3", ""],
    ]);
    SHEET.getRange(3, COL.Weight).setValue("186"); // only row 3 gets content
    ok(
      onEditMarkDirty({ range: SHEET.getRange(2, COL.Weight, 2, 1) }) === true,
      "marked",
    );
    eq(
      SHEET.getRange(3, COL["Weight Synced At"]).getValue(),
      "",
      "row 3 dirtied",
    );
    ok(
      SHEET.getRange(3, COL["Weight Edited At"]).getValue() !== "WEDIT-3",
      "row 3 advanced",
    );
    eq(
      SHEET.getRange(2, COL["Weight Synced At"]).getValue(),
      "PREV-WT",
      "row 2 stayed synced",
    );
    eq(
      SHEET.getRange(2, COL["Weight Edited At"]).getValue(),
      "WEDIT-2",
      "row 2 kept its edit time",
    );
  });

  // The two phases are marked independently within one range: a row that only
  // had a bodyweight change must not get exercise markers, and vice versa.
  t(
    "onEditMarkDirty keeps the two phases on their own rows within one range",
    () => {
      reset([
        [
          "2026-01-15",
          "",
          "",
          "PREV-EX",
          "PREV-WT",
          "",
          "",
          "LAST-2",
          "WEDIT-2",
          "",
        ],
        [
          "2026-01-16",
          "",
          "",
          "PREV-EX",
          "PREV-WT",
          "",
          "",
          "LAST-3",
          "WEDIT-3",
          "",
        ],
      ]);
      SHEET.getRange(2, COL.Weight).setValue("186"); // row 2: weight only
      SHEET.getRange(3, COL.Bench).setValue("135x5"); // row 3: exercise only
      ok(
        onEditMarkDirty({ range: SHEET.getRange(2, COL.Weight, 2, 2) }) ===
          true,
        "marked",
      );
      eq(
        SHEET.getRange(2, COL["Weight Synced At"]).getValue(),
        "",
        "row 2 weight dirtied",
      );
      eq(
        SHEET.getRange(2, COL["Exercise Synced At"]).getValue(),
        "PREV-EX",
        "row 2 exercise untouched",
      );
      eq(
        SHEET.getRange(2, COL["Exercises Last Edited At"]).getValue(),
        "LAST-2",
        "row 2 exercise time kept",
      );
      eq(
        SHEET.getRange(3, COL["Exercise Synced At"]).getValue(),
        "",
        "row 3 exercise dirtied",
      );
      eq(
        SHEET.getRange(3, COL["Weight Synced At"]).getValue(),
        "PREV-WT",
        "row 3 weight untouched",
      );
      eq(
        SHEET.getRange(3, COL["Weight Edited At"]).getValue(),
        "WEDIT-3",
        "row 3 weight time kept",
      );
    },
  );

  t(
    "onEditMarkDirty keeps Exercise First Edited At sticky but advances Exercises Last Edited At",
    () => {
      reset([
        [
          "2026-01-15",
          "",
          "135x5",
          "",
          "",
          "",
          "STICKY-FIRST",
          "OLD-LAST",
          "",
          "",
        ],
      ]);
      onEditMarkDirty({ range: SHEET.getRange(2, COL.Bench) });
      eq(
        cell("Exercise First Edited At"),
        "STICKY-FIRST",
        "first edited stays sticky",
      );
      ok(
        cell("Exercises Last Edited At") !== "OLD-LAST",
        "last edited advanced",
      );
    },
  );

  // ---- readRows: the full-sheet id/session lists -------------------------

  // The contract every ownership decision depends on. A row with no parseable
  // Date is not syncable and is absent from `rows`, but it still OWNS its
  // datapoints and still holds its foreign session, so both must appear in the
  // full-sheet lists. Deriving either from `rows` instead makes orphan
  // reconciliation delete live datapoints and lets foreign matching hand the
  // same session to another row.
  t(
    "readRows reports ids and matched sessions of rows it drops for a blank Date",
    () => {
      const undatedWeight = "users/me/dataTypes/weight/dataPoints/W-undated";
      const datedExercise = "users/me/dataTypes/exercise/dataPoints/E-dated";
      reset([
        [
          "",
          "185",
          "",
          "SYNC",
          "SYNC",
          JSON.stringify([undatedWeight]),
          "",
          "",
          "",
          "foreign/undated",
        ],
        [
          "2026-01-16",
          "",
          "135x5",
          "SYNC",
          "SYNC",
          JSON.stringify([datedExercise]),
          "",
          "",
          "",
          "foreign/dated",
        ],
      ]);
      const r = readRows();
      eq(
        r.rows.map((row) => row.rowNum),
        [3],
        "only the dated row is syncable",
      );
      eq(
        r.allHealthIds,
        [undatedWeight, datedExercise],
        "ids from BOTH rows, dropped one included",
      );
      eq(
        r.allMatchedSessions,
        [
          { name: "foreign/undated", rowNum: 2 },
          { name: "foreign/dated", rowNum: 3 },
        ],
        "matched sessions from BOTH rows",
      );
    },
  );

  // Two columns sharing a header would silently merge two exercises into one,
  // so readRows refuses. It is unrecoverable: re-thrown for the owner email,
  // with the dirty flag left set so the backlog syncs once the sheet is fixed.
  t("readRows refuses duplicate exercise column headers", () => {
    const headers = HEADERS.slice();
    headers[3] = "Bench"; // was Exercise Synced At; now a duplicate of col 3
    headers[4] = "Exercise Synced At";
    headers[5] = "Weight Synced At";
    reset([["2026-01-15", "185", "135x5", "", "SYNC", "", "", "", "", ""]]);
    SHEET.getRange(1, 1, 1, headers.length).setValues([headers]);
    PROPS.setProperty("pendingDirty", "GEN1");
    let thrown = null;
    try {
      withStubs(NO_FOREIGN, () => syncDirtyRows(0));
    } catch (err) {
      thrown = err;
    }
    ok(thrown !== null, "throws");
    ok(
      String(thrown).indexOf("Bench") !== -1,
      `names the duplicated header: ${thrown}`,
    );
    ok(
      PROPS.getProperty("pendingDirty") !== null,
      "dirty flag kept for after the fix",
    );
  });

  // ---- syncDirtyRows: lifecycle ------------------------------------------

  t("syncDirtyRows returns null when the lock is held", () => {
    reset([["2026-01-15", "185", "", "SYNC", "", "", "", "", "", ""]]);
    LOCK.held = true;
    const r = withStubs(NO_FOREIGN, () => syncDirtyRows(0));
    eq(r, null, "skips when lock held");
  });

  t("syncDirtyRows with no dirty rows returns zero counts", () => {
    reset([["2026-01-15", "185", "", "SYNC", "SYNC", "[]", "", "", "", ""]]);
    const r = withStubs(NO_FOREIGN, () => syncDirtyRows(0));
    eq(r, { errors: 0, ok: 0 }, "nothing to do");
  });

  t(
    "syncDirtyRows weight first-sync POSTs, stamps, persists ID, drains the flag",
    () => {
      reset([["2026-01-15", "185", "", "SYNC", "", "", "", "", "", ""]]);
      PROPS.setProperty("pendingDirty", "GEN1");
      const calls = [];
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: (utc, off, lbs) => {
            calls.push(["createWeightAt", lbs]);
            return "users/me/dataTypes/weight/dataPoints/W1";
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "one row synced");
      eq(calls, [["createWeightAt", 185]], "POSTed bodyweight 185");
      eq(
        cell("Created Health IDs"),
        JSON.stringify(["users/me/dataTypes/weight/dataPoints/W1"]),
        "ID persisted",
      );
      ok(cell("Weight Synced At") !== "", "weight synced stamped");
      ok(
        PROPS.getProperty("pendingDirty") === null,
        "flag cleared after a clean drain",
      );
    },
  );

  t("syncDirtyRows keeps the dirty flag set when a row errors", () => {
    reset([["2026-01-15", "185", "", "SYNC", "", "", "", "", "", ""]]);
    const r = withStubs(
      Object.assign({}, NO_FOREIGN, {
        createWeightAt: () => {
          throw new Error("boom");
        },
      }),
      () => syncDirtyRows(0),
    );
    eq(r.ok, 0, "no rows synced");
    eq(r.errors, 1, "one error");
    ok(
      PROPS.getProperty("pendingDirty") !== null,
      "flag retained so the next poll retries",
    );
    eq(cell("Weight Synced At"), "", "weight stamp not written on failure");
  });

  // ---- Failure flags gate the Synced At stamp ------------------------------
  // A failed write must leave the phase unstamped and the row dirty. Stamping
  // it anyway is silent data loss: the row stops being dirty, nothing retries,
  // and the sheet claims a datapoint Health does not have. Only the weight POST
  // path was covered; these are the other four ways a phase can fail.

  t(
    "syncDirtyRows leaves the weight phase unstamped when the PATCH fails",
    () => {
      const wName = "users/me/dataTypes/weight/dataPoints/W1";
      reset([
        [
          "2026-01-15",
          "186",
          "",
          "SYNC",
          "",
          JSON.stringify([wName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          getDataPoint: () => ({
            weight: {
              sampleTime: {
                physicalTime: "2026-01-15T17:00:00Z",
                utcOffset: "-18000s",
              },
            },
          }),
          patchWeight: () => {
            throw new Error("simulated PATCH failure");
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 0, "the row did not sync");
      eq(r.errors, 1, "the failure is counted");
      eq(cell("Weight Synced At"), "", "no stamp on a failed PATCH");
      eq(
        cell("Created Health IDs"),
        JSON.stringify([wName]),
        "the prior id is kept so the next pass retries the same datapoint",
      );
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "flag retained so the next poll retries",
      );
    },
  );

  t(
    "syncDirtyRows leaves the weight phase unstamped when a prior delete fails",
    () => {
      const wName = "users/me/dataTypes/weight/dataPoints/W1";
      // Bodyweight cleared, so the phase takes the delete branch.
      reset([
        [
          "2026-01-15",
          "",
          "",
          "SYNC",
          "",
          JSON.stringify([wName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          deleteDataPointsByName: () => {
            throw new Error("simulated delete failure");
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 0, "the row did not sync");
      eq(r.errors, 1, "the failure is counted");
      eq(cell("Weight Synced At"), "", "no stamp on a failed delete");
      eq(
        cell("Created Health IDs"),
        JSON.stringify([wName]),
        "the undeleted id stays tracked so the next pass retries it",
      );
    },
  );

  t(
    "syncDirtyRows leaves the exercise phase unstamped when a prior delete fails",
    () => {
      const eName = "users/me/dataTypes/exercise/dataPoints/E1";
      reset([
        [
          "2026-01-15",
          "",
          "135x5x3",
          "",
          "SYNC",
          JSON.stringify([eName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      let created = 0;
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => {
            created++;
            return "users/me/dataTypes/exercise/dataPoints/E2";
          },
          deleteDataPointsByName: () => {
            throw new Error("simulated delete failure");
          },
          getDataPoint: () => null,
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 0, "the row did not sync");
      eq(r.errors, 1, "the failure is counted");
      eq(created, 0, "a failed delete blocks the recreate");
      eq(cell("Exercise Synced At"), "", "no stamp on a failed delete");
      eq(
        cell("Created Health IDs"),
        JSON.stringify([eName]),
        "the undeleted id stays tracked so the next pass retries it",
      );
    },
  );

  t(
    "syncDirtyRows leaves the exercise phase unstamped when the create fails",
    () => {
      reset([["2026-01-15", "", "135x5x3", "", "SYNC", "", "", "", "", ""]]);
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => {
            throw new Error("simulated create failure");
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 0, "the row did not sync");
      eq(r.errors, 1, "the failure is counted");
      eq(cell("Exercise Synced At"), "", "no stamp on a failed create");
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "flag retained so the next poll retries",
      );
    },
  );

  t(
    "syncDirtyRows preserves a concurrent-edit generation rather than clearing it",
    () => {
      reset([["2026-01-15", "185", "", "SYNC", "", "", "", "", "", ""]]);
      PROPS.setProperty("pendingDirty", "GEN1");
      // A stub that advances the generation mid-pass simulates an edit landing
      // after readRows snapshotted. End-of-pass must NOT clear the flag.
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: () => {
            PROPS.setProperty("pendingDirty", "GEN2");
            return "users/me/dataTypes/weight/dataPoints/W1";
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "row synced");
      eq(
        PROPS.getProperty("pendingDirty"),
        "GEN2",
        "concurrent-edit generation kept",
      );
    },
  );

  // An unexpected throw out of syncOneRow_ (sheet I/O, not a Health API call)
  // must not abort the pass: rows are processed newest-first, so aborting would
  // strand every older row behind the failure on every subsequent pass. The row
  // is isolated, the rest of the pass completes, and a summary throw still
  // routes the failure through the unrecoverable path (owner email, flag kept).
  t(
    "syncDirtyRows isolates an unexpected per-row failure and still reports it",
    () => {
      const older = new Date(Date.now() - 24 * 60 * 60 * 1000);
      reset([
        [older, "185", "", "SYNC", "", "", "", "", "", ""], // row 2, older
        [new Date(), "186", "", "SYNC", "", "", "", "", "", ""], // row 3, newest -> processed first
      ]);
      PROPS.setProperty("pendingDirty", "GEN1");
      const synced = [];
      let thrown = null;
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: (utc, off, lbs) =>
            `users/me/dataTypes/weight/dataPoints/W-${lbs}`,
          writeHealthIds: (rowNum, _col, _names) => {
            if (rowNum === 3) {
              throw new Error("simulated Spreadsheets service failure");
            }
            if (synced.indexOf(rowNum) === -1) {
              synced.push(rowNum);
            } // called twice per row
          },
        }),
        () => {
          try {
            syncDirtyRows(0);
          } catch (err) {
            thrown = err;
          }
        },
      );
      ok(thrown !== null, "throws so Apps Script emails the owner");
      ok(
        String(thrown).indexOf("row 3") !== -1,
        `message names the failing row: ${thrown}`,
      );
      eq(synced, [2], "the older row still synced instead of being stranded");
      ok(
        SHEET.getRange(2, COL["Weight Synced At"]).getValue() !== "",
        "older row stamped",
      );
      eq(
        SHEET.getRange(3, COL["Weight Synced At"]).getValue(),
        "",
        "failing row left dirty",
      );
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "dirty flag kept so the next poll retries",
      );
    },
  );

  // The unrecoverable path skips end-of-pass flag resolution, so it only
  // preserves a flag that is already there. runSyncNow is the one entry point
  // that can start a pass with no flag set; without an explicit
  // markPendingDirty_() before the summary throw, the failed row is left dirty
  // with nothing scheduled to retry it (flushPending short-circuits, and the
  // backstop only re-dirties exercise rows with sendable content).
  t(
    "syncDirtyRows SETS the dirty flag when an unexpected failure hits a pass that had none",
    () => {
      reset([["2026-01-15", "185", "", "SYNC", "", "", "", "", "", ""]]);
      ok(
        PROPS.getProperty("pendingDirty") === null,
        "precondition: no flag at pass start",
      );
      let thrown = null;
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: () => "users/me/dataTypes/weight/dataPoints/W1",
          writeHealthIds: () => {
            throw new Error("simulated Spreadsheets service failure");
          },
        }),
        () => {
          try {
            syncDirtyRows(0);
          } catch (err) {
            thrown = err;
          }
        },
      );
      ok(thrown !== null, "failure reported");
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "flag created, so a later poll retries the row",
      );
    },
  );

  // No failure count stops the pass. A row that throws deterministically would
  // do so on every pass, so any early stop strands the rows behind it. Rows
  // are processed newest-first, and the same rows lead every time. Orphan
  // reconciliation, not a stopping rule, is what handles the datapoints created
  // by rows whose id write failed.
  const outageGrid = (count) => {
    const day = 24 * 60 * 60 * 1000;
    const now = Date.now();
    const grid = [];
    for (let i = count; i >= 1; i--) {
      // ascending dates down the sheet
      grid.push([
        new Date(now - i * day),
        `1${80 + i}`,
        "",
        "SYNC",
        "",
        "",
        "",
        "",
        "",
        "",
      ]);
    }
    return grid;
  };

  // Every ready row is attempted even when all of them fail, and the toast
  // truncates, so a broad failure must lead with the counts rather than a wall
  // of row errors that pushes the summary off the end.
  t(
    "syncDirtyRows attempts every ready row and leads the summary with the counts",
    () => {
      reset(outageGrid(8));
      let created = 0;
      let thrown = null;
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: () => {
            created++;
            return `users/me/dataTypes/weight/dataPoints/W${created}`;
          },
          writeHealthIds: () => {
            throw new Error("simulated Spreadsheets outage");
          },
        }),
        () => {
          try {
            syncDirtyRows(0);
          } catch (err) {
            thrown = err;
          }
        },
      );
      eq(created, 8, "no row was skipped because earlier ones failed");
      const msg = String(thrown.message);
      ok(
        msg.indexOf("0 synced, 8 error(s)") === 0,
        `counts come first: ${msg}`,
      );
      ok(
        msg.indexOf("+3 more (see Executions)") !== -1,
        `row list trimmed with the remainder named: ${msg}`,
      );
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "flag kept so the next pass retries",
      );
    },
  );

  // The summary throw replaces the normal return, so anything not interpolated
  // into its message is lost. deferredCount is the easiest to forget: it only
  // appears when the row cap and an unexpected failure coincide.
  t(
    "syncDirtyRows summary reports the deferred backlog alongside the failure",
    () => {
      const day = 24 * 60 * 60 * 1000;
      const now = Date.now();
      const grid = [];
      for (let i = MAX_ROWS_PER_SYNC + 1; i >= 1; i--) {
        grid.push([
          new Date(now - i * day),
          "185",
          "",
          "SYNC",
          "",
          "",
          "",
          "",
          "",
          "",
        ]);
      }
      reset(grid);
      const newestRow = grid.length + 1; // ascending dates, so the last row is newest
      let thrown = null;
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: () => "users/me/dataTypes/weight/dataPoints/W",
          writeHealthIds: (rowNum) => {
            if (rowNum === newestRow) {
              throw new Error("simulated Spreadsheets service failure");
            }
          },
        }),
        () => {
          try {
            syncDirtyRows(0);
          } catch (err) {
            thrown = err;
          }
        },
      );
      ok(
        String(thrown).indexOf("1 deferred by the row cap") !== -1,
        `deferred backlog surfaced, not swallowed by the throw: ${thrown}`,
      );
    },
  );

  // The cap is what keeps a pass inside Apps Script's 6-minute execution limit.
  // A reported deferred count doesn't prove it was applied (deferredCount is
  // computed before the truncation), so assert the work actually stopped.
  t(
    "syncDirtyRows processes at most MAX_ROWS_PER_SYNC rows and defers the rest",
    () => {
      const extra = 2;
      const day = 24 * 60 * 60 * 1000;
      const now = Date.now();
      const grid = [];
      for (let i = MAX_ROWS_PER_SYNC + extra; i >= 1; i--) {
        grid.push([
          new Date(now - i * day),
          "185",
          "",
          "SYNC",
          "",
          "",
          "",
          "",
          "",
          "",
        ]);
      }
      reset(grid);
      let created = 0;
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: () => {
            created++;
            return `users/me/dataTypes/weight/dataPoints/W${created}`;
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(created, MAX_ROWS_PER_SYNC, "the cap bounded the pass");
      eq(r.ok, MAX_ROWS_PER_SYNC, "every processed row synced");
      eq(r.deferred, extra, "the rest deferred to the next pass");
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "flag kept for the deferred backlog",
      );
    },
  );

  // ---- syncOneRow_ via syncDirtyRows: exercise idempotency ---------------

  t(
    "syncDirtyRows skips the delete+recreate when the prior exercise datapoint is unchanged",
    () => {
      const startMs = Date.UTC(2026, 0, 15, 17, 0, 0);
      const endMs = Date.UTC(2026, 0, 15, 17, 30, 0);
      const exercises = [
        {
          entries: [{ assisted: false, reps: 5, sets: 3, weight: 135 }],
          name: "Bench",
        },
      ];
      const priorNotes = buildNotes(endMs - startMs, exercises);
      const priorName = "users/me/dataTypes/exercise/dataPoints/E1";
      reset([
        [
          "2026-01-15",
          "",
          "135x5x3",
          "",
          "SYNC",
          JSON.stringify([priorName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const calls = [];
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => {
            calls.push(["create"]);
            return "users/me/dataTypes/exercise/dataPoints/E2";
          },
          deleteDataPointsByName: (names) => {
            calls.push(["delete", names]);
          },
          getDataPoint: () => ({
            exercise: {
              interval: {
                endTime: new Date(endMs).toISOString(),
                startTime: new Date(startMs).toISOString(),
              },
              notes: priorNotes,
            },
          }),
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "row counted ok");
      eq(calls, [], "no delete or create issued for an unchanged exercise");
      eq(
        cell("Created Health IDs"),
        JSON.stringify([priorName]),
        "resource name preserved",
      );
      ok(cell("Exercise Synced At") !== "", "exercise synced stamped");
    },
  );

  t("syncDirtyRows recreates the exercise datapoint when notes changed", () => {
    const startMs = Date.UTC(2026, 0, 15, 17, 0, 0);
    const endMs = Date.UTC(2026, 0, 15, 17, 30, 0);
    const priorName = "users/me/dataTypes/exercise/dataPoints/E1";
    // Prior notes describe a different set/rep scheme than the current cell.
    reset([
      [
        "2026-01-15",
        "",
        "135x5x3",
        "",
        "SYNC",
        JSON.stringify([priorName]),
        "",
        "",
        "",
        "",
      ],
    ]);
    const calls = [];
    const r = withStubs(
      Object.assign({}, NO_FOREIGN, {
        createExerciseAt: () => {
          calls.push(["create"]);
          return "users/me/dataTypes/exercise/dataPoints/E2";
        },
        deleteDataPointsByName: (names) => {
          calls.push(["delete", names.slice()]);
        },
        getDataPoint: () => ({
          exercise: {
            interval: {
              endTime: new Date(endMs).toISOString(),
              startTime: new Date(startMs).toISOString(),
            },
            notes: "Bench, 999 lbs, 1 set of 1",
          },
        }),
      }),
      () => syncDirtyRows(0),
    );
    eq(r.ok, 1, "row counted ok");
    eq(
      calls,
      [["delete", [priorName]], ["create"]],
      "deletes old then creates new",
    );
    eq(
      cell("Created Health IDs"),
      JSON.stringify(["users/me/dataTypes/exercise/dataPoints/E2"]),
      "new resource name recorded",
    );
  });

  t(
    "syncDirtyRows recreates rather than skipping when a row has two prior exercise ids",
    () => {
      // The idempotency skip is deliberately limited to a single prior id:
      // multiple priors are consolidated by recreating, otherwise the extras stay
      // live in Health with nothing left to reconcile them against.
      const startMs = Date.UTC(2026, 0, 15, 17, 0, 0);
      const endMs = Date.UTC(2026, 0, 15, 17, 30, 0);
      const priorNotes = buildNotes(endMs - startMs, [
        {
          entries: [{ assisted: false, reps: 5, sets: 3, weight: 135 }],
          name: "Bench",
        },
      ]);
      const first = "users/me/dataTypes/exercise/dataPoints/E1";
      const second = "users/me/dataTypes/exercise/dataPoints/E2";
      reset([
        [
          "2026-01-15",
          "",
          "135x5x3",
          "",
          "SYNC",
          JSON.stringify([first, second]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const calls = [];
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => {
            calls.push(["create"]);
            return "users/me/dataTypes/exercise/dataPoints/E3";
          },
          deleteDataPointsByName: (names) => {
            calls.push(["delete", names.slice()]);
          },
          getDataPoint: () => ({
            exercise: {
              interval: {
                endTime: new Date(endMs).toISOString(),
                startTime: new Date(startMs).toISOString(),
              },
              notes: priorNotes,
            },
          }),
        }),
        () => syncDirtyRows(0),
      );
      eq(
        calls,
        [["delete", [first]], ["delete", [second]], ["create"]],
        "both priors deleted despite matching content, one datapoint left",
      );
      eq(
        cell("Created Health IDs"),
        JSON.stringify(["users/me/dataTypes/exercise/dataPoints/E3"]),
      );
    },
  );

  // Clearing every exercise cell must remove the datapoint. This is the path
  // the single-cell-clear rule in onEditMarkDirty exists to reach: the backstop
  // won't re-dirty a row with no sendable content, and orphan reconciliation
  // won't touch a datapoint the sheet still tracks, so nothing else collects it.
  t(
    "syncDirtyRows deletes without recreating when the exercise content is gone",
    () => {
      const exName = "users/me/dataTypes/exercise/dataPoints/E1";
      reset([
        [
          "2026-01-15",
          "",
          "",
          "",
          "SYNC",
          JSON.stringify([exName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const calls = [];
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => {
            calls.push(["create"]);
            return "users/me/dataTypes/exercise/dataPoints/E2";
          },
          deleteDataPointsByName: (names) => {
            calls.push(["delete", names.slice()]);
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "row counted ok");
      eq(calls, [["delete", [exName]]], "deleted, nothing recreated");
      eq(cell("Created Health IDs"), JSON.stringify([]), "id dropped");
      ok(cell("Exercise Synced At") !== "", "exercise stamped");
    },
  );

  // ---- weight phase: the PATCH branch ------------------------------------

  // (prior id, bodyweight set) is the branch that keeps sampleTime, createTime
  // and the resource name stable across re-syncs. The PATCH body is rejected
  // without sampleTime, so it is read from a GET and echoed back verbatim.
  t(
    "syncDirtyRows PATCHes a re-synced bodyweight in place and keeps the resource name",
    () => {
      const wName = "users/me/dataTypes/weight/dataPoints/W1";
      const sampleTime = {
        physicalTime: "2026-01-15T17:00:00Z",
        utcOffset: "-18000s",
      };
      reset([
        [
          "2026-01-15",
          "186",
          "",
          "SYNC",
          "",
          JSON.stringify([wName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const calls = [];
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: () => {
            calls.push(["create"]);
            return "users/me/dataTypes/weight/dataPoints/W2";
          },
          deleteDataPointsByName: (names) => {
            calls.push(["delete", names.slice()]);
          },
          getDataPoint: (name) => {
            calls.push(["get", name]);
            return { weight: { sampleTime } };
          },
          patchWeight: (name, st, lbs) => {
            calls.push(["patch", name, st, lbs]);
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "row synced");
      eq(
        calls,
        [
          ["get", wName],
          ["patch", wName, sampleTime, 186],
        ],
        "PATCH only: no delete, no create",
      );
      eq(
        cell("Created Health IDs"),
        JSON.stringify([wName]),
        "resource name unchanged",
      );
      ok(cell("Weight Synced At") !== "", "weight stamped");
    },
  );

  // Without sampleTime the PATCH body 500s, so a failed GET must leave the row
  // dirty rather than issuing a request that cannot succeed.
  t(
    "syncDirtyRows skips the PATCH and stays dirty when the prior-weight GET fails",
    () => {
      const wName = "users/me/dataTypes/weight/dataPoints/W1";
      reset([
        [
          "2026-01-15",
          "186",
          "",
          "SYNC",
          "",
          JSON.stringify([wName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const calls = [];
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          getDataPoint: () => {
            throw new Error("simulated 500");
          },
          patchWeight: () => {
            calls.push("patch");
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.errors, 1, "row counted as failed");
      eq(calls, [], "no PATCH attempted");
      eq(cell("Weight Synced At"), "", "row left dirty for the next pass");
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "flag kept so the next poll retries",
      );
    },
  );

  // A 404 means the datapoint was removed in the Health app. PATCHing it would
  // fail on every future pass, so the stale id is dropped and the row POSTs a
  // fresh datapoint instead of wedging.
  t(
    "syncDirtyRows drops a 404 prior weight id and POSTs a fresh datapoint",
    () => {
      const gone = "users/me/dataTypes/weight/dataPoints/W-gone";
      reset([
        [
          "2026-01-15",
          "186",
          "",
          "SYNC",
          "",
          JSON.stringify([gone]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: () => "users/me/dataTypes/weight/dataPoints/W-new",
          getDataPoint: () => {
            const err = new Error("Health API GET ... -> 404: not found");
            err.statusCode = 404;
            throw err;
          },
          patchWeight: () => {
            throw new Error("must not PATCH a datapoint that is gone");
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "row synced");
      eq(
        cell("Created Health IDs"),
        JSON.stringify(["users/me/dataTypes/weight/dataPoints/W-new"]),
        "stale id replaced",
      );
      ok(cell("Weight Synced At") !== "", "weight stamped");
    },
  );

  // ---- foreign-match alignment, end to end --------------------------------

  // What the pure resolveForeignMatches_ / resolveRowTiming_ tests can't show:
  // the plan reaching createExerciseAt as the borrowed interval, and the
  // aligned session's name being recorded for the next pass's exclusion
  // bookkeeping. Date and timestamp cells hold Date objects so toDate_ reads
  // them verbatim. A 'yyyy-MM-dd' string is UTC midnight, i.e. the PREVIOUS
  // civil day in the harness time zone, which would fail the on-row-date test.
  const JAN15_NOON_EST = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
  const JAN15_5PM_EST = new Date(Date.UTC(2026, 0, 15, 22, 0, 0));
  const JAN15_530PM_EST = new Date(Date.UTC(2026, 0, 15, 22, 30, 0));

  t(
    "syncDirtyRows borrows a matched foreign interval and records the session",
    () => {
      reset([
        [
          JAN15_NOON_EST,
          "",
          "135x5x3",
          "",
          "SYNC",
          "",
          JAN15_5PM_EST,
          JAN15_530PM_EST,
          "",
          "",
        ],
      ]);
      const cand = {
        endUtcMs: Date.UTC(2026, 0, 15, 23, 20, 0),
        endUtcOffsetSeconds: -5 * 3600,
        name: "foreign/A",
        startUtcMs: Date.UTC(2026, 0, 15, 22, 5, 0),
        startUtcOffsetSeconds: -5 * 3600,
      };
      let created = null;
      const r = withStubs(
        {
          createExerciseAt: (startUtcMs, startOff, endUtcMs, endOff, notes) => {
            created = { endOff, endUtcMs, notes, startOff, startUtcMs };
            return "users/me/dataTypes/exercise/dataPoints/E1";
          },
          listStrengthOnDate: () => [cand],
        },
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "row synced");
      eq(
        created.startUtcMs,
        cand.startUtcMs,
        "foreign start borrowed verbatim, not the edit time",
      );
      eq(created.endUtcMs, cand.endUtcMs, "foreign end borrowed verbatim");
      eq(created.startOff, cand.startUtcOffsetSeconds);
      eq(created.endOff, cand.endUtcOffsetSeconds);
      eq(
        created.notes,
        buildNotes(cand.endUtcMs - cand.startUtcMs, [
          {
            entries: [{ assisted: false, reps: 5, sets: 3, weight: 135 }],
            name: "Bench",
          },
        ]),
        "the session length in the notes follows the borrowed interval",
      );
      eq(
        cell("Matched Health Session"),
        "foreign/A",
        "match recorded for the next pass",
      );
    },
  );

  // The clear is what stops a stale name from shielding a session forever: the
  // aligned-elsewhere exclusion list is built from this column.
  t(
    "syncDirtyRows clears the recorded foreign session when a row no longer matches",
    () => {
      reset([
        [
          JAN15_NOON_EST,
          "",
          "135x5x3",
          "",
          "SYNC",
          "",
          JAN15_5PM_EST,
          JAN15_530PM_EST,
          "",
          "foreign/OLD",
        ],
      ]);
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => "users/me/dataTypes/exercise/dataPoints/E1",
        }),
        () => syncDirtyRows(0),
      );
      eq(cell("Matched Health Session"), "", "stale match cleared");
    },
  );

  // ---- syncOneRow_: per-phase concurrent-edit guards ----------------------

  // An edit landing while the row is being pushed must not be stamped synced:
  // the datapoint was built from a stale snapshot. The id is still persisted
  // (the datapoint exists and must stay tracked), but the stamp is deferred and
  // the dirty flag survives so a later pass re-pushes the current content.
  t(
    "syncOneRow_ defers the Exercise Synced At stamp when an exercise edit lands mid-pass",
    () => {
      const editedAt = new Date(Date.UTC(2026, 0, 15, 22, 0, 0));
      reset([
        ["2026-01-15", "", "135x5x3", "", "SYNC", "", "", editedAt, "", ""],
      ]);
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => {
            SHEET.getRange(2, COL["Exercises Last Edited At"]).setValue(
              new Date(editedAt.getTime() + 60 * 1000),
            );
            return "users/me/dataTypes/exercise/dataPoints/E1";
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "the push itself succeeded");
      eq(
        cell("Exercise Synced At"),
        "",
        "stamp deferred, so the row stays dirty",
      );
      eq(
        cell("Created Health IDs"),
        JSON.stringify(["users/me/dataTypes/exercise/dataPoints/E1"]),
        "id still persisted",
      );
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "flag kept so the next poll re-pushes",
      );
    },
  );

  // The guards are phase-isolated and both read a timestamp column rather than
  // the content cell, which is what catches an "edit the value, edit it back"
  // change that a value comparison would miss.
  t(
    "syncOneRow_ defers the Weight Synced At stamp when a weight edit lands mid-pass",
    () => {
      const editedAt = new Date(Date.UTC(2026, 0, 15, 22, 0, 0));
      reset([["2026-01-15", "185", "", "SYNC", "", "", "", "", editedAt, ""]]);
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: () => {
            SHEET.getRange(2, COL["Weight Edited At"]).setValue(
              new Date(editedAt.getTime() + 60 * 1000),
            );
            return "users/me/dataTypes/weight/dataPoints/W1";
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "the push itself succeeded");
      eq(
        cell("Weight Synced At"),
        "",
        "stamp deferred, so the row stays dirty",
      );
      eq(
        cell("Created Health IDs"),
        JSON.stringify(["users/me/dataTypes/weight/dataPoints/W1"]),
        "id still persisted",
      );
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "flag kept so the next poll re-pushes",
      );
    },
  );

  // Clearing a bodyweight whose datapoint was already deleted in the Health app
  // must not wedge the row: retrying a delete that 404s forever would keep the
  // row dirty and the pending flag set, re-issuing the same call every poll.
  t(
    "syncDirtyRows treats a 404 on the weight delete as already deleted",
    () => {
      const priorName = "users/me/dataTypes/weight/dataPoints/W-gone";
      reset([
        [
          "2026-01-15",
          "",
          "",
          "SYNC",
          "",
          JSON.stringify([priorName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          deleteDataPointsByName: () => {
            const err = new Error("Health API POST ... -> 404: not found");
            err.statusCode = 404;
            throw err;
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "row counted ok");
      eq(
        cell("Created Health IDs"),
        JSON.stringify([]),
        "stale weight id dropped",
      );
      ok(
        cell("Weight Synced At") !== "",
        "weight stamped instead of retrying forever",
      );
      ok(PROPS.getProperty("pendingDirty") === null, "queue drained");
    },
  );

  // Deletes go one name per call so a 404 is attributed to the name that is
  // actually gone. A batched :batchDelete fails as a unit, so treating that
  // failure as "all deleted" would drop still-live siblings from the sheet.
  t(
    "syncDirtyRows keeps a live sibling when only one prior weight id is gone",
    () => {
      const gone = "users/me/dataTypes/weight/dataPoints/W-gone";
      const live = "users/me/dataTypes/weight/dataPoints/W-live";
      reset([
        [
          "2026-01-15",
          "",
          "",
          "SYNC",
          "",
          JSON.stringify([gone, live]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const deleted = [];
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          deleteDataPointsByName: (names) => {
            if (names[0] === gone) {
              const err = new Error("Health API POST ... -> 404: not found");
              err.statusCode = 404;
              throw err;
            }
            deleted.push(names[0]);
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "row counted ok");
      eq(
        deleted,
        [live],
        "the live sibling was actually deleted, not assumed gone",
      );
      eq(cell("Created Health IDs"), JSON.stringify([]), "both ids resolved");
    },
  );

  // The accepted tradeoff of running every pass to completion. A row whose id
  // write fails has already created its datapoint, so the id is lost and the
  // NEXT pass creates another: a leak no failure threshold can bound (stopping
  // early only trades it for stranding the rows behind it). Orphan
  // reconciliation closes the loop, but only within
  // ORPHAN_RECONCILE_LOOKBACK_DAYS, which is what this covers (the rows are
  // dated today/yesterday). Candidates are listed by the datapoint's own civil
  // date, i.e. the row's Date, so a leak on an older row is never reclaimed;
  // that residual is documented, not tested, since no code path closes it.
  t(
    "an untracked create is re-made next pass and later reclaimed by reconciliation",
    () => {
      const tracked = "users/me/dataTypes/weight/dataPoints/W-tracked";
      reset([
        [
          new Date(Date.now() - 24 * 60 * 60 * 1000),
          "185",
          "",
          "SYNC",
          "",
          "",
          "",
          "",
          "",
          "",
        ],
        [
          new Date(),
          "",
          "",
          "SYNC",
          "SYNC",
          JSON.stringify([tracked]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const created = [];
      const brokenWrite = Object.assign({}, NO_FOREIGN, {
        createWeightAt: () => {
          const name = `users/me/dataTypes/weight/dataPoints/W-leaked${created.length + 1}`;
          created.push(name);
          return name;
        },
        writeHealthIds: () => {
          throw new Error("simulated Spreadsheets service failure");
        },
      });
      withStubs(brokenWrite, () => {
        try {
          syncDirtyRows(0);
        } catch {
          /* reported */
        }
      });
      withStubs(brokenWrite, () => {
        try {
          syncDirtyRows(0);
        } catch {
          /* reported */
        }
      });
      eq(
        created.length,
        2,
        "each pass created a datapoint it could not record",
      );
      eq(cell("Created Health IDs"), "", "neither id reached the sheet");

      // The backstop reclaims both: they carry our client id and no row
      // references them. The tracked datapoint on the other row establishes
      // ownership and must survive.
      const deleted = [];
      let listed = false;
      withStubs(
        {
          deleteDataPointsByName: (names) => {
            deleted.push(...names);
          },
          listStrengthOnDate: () => [],
          listWeightOnDate: () => {
            if (listed) {
              return [];
            } // one civil day's worth, not once per lookback day
            listed = true;
            return [tracked, created[0], created[1]].map((name) => ({
              googleWebClientId: "ours",
              name,
            }));
          },
        },
        () => backstop(),
      );
      eq(
        deleted,
        created,
        "both leaked datapoints reclaimed, tracked one spared",
      );
    },
  );

  // ---- weight orphan reconciliation --------------------------------------

  t(
    "reconcileWeightOrphans_ deletes our untracked weight datapoint, spares tracked/foreign",
    () => {
      const tracked = "users/me/dataTypes/weight/dataPoints/W-tracked";
      const orphan = "users/me/dataTypes/weight/dataPoints/W-orphan";
      const device = "users/me/dataTypes/weight/dataPoints/W-device";
      const deleted = [];
      withStubs(
        {
          deleteDataPointsByName: (names) => {
            deleted.push(...names);
          },
          listWeightOnDate: () => [
            { googleWebClientId: "ours", name: tracked },
            { googleWebClientId: "ours", name: orphan },
            { googleWebClientId: null, name: device },
          ],
        },
        () => {
          reconcileWeightOrphans_(
            [tracked],
            Date.UTC(2026, 0, 15, 12, 0, 0),
            1,
          );
        },
      );
      eq(deleted, [orphan], "only our untracked weight datapoint is deleted");
    },
  );

  // A row whose Date cell is blank is dropped by readRows, but its datapoints
  // are still live and still referenced by the sheet. Reconciling against the
  // full-sheet id list (readRows' allHealthIds) rather than the surviving rows
  // is what keeps the backstop from deleting them.
  t("backstop spares datapoints of a row whose Date is blank", () => {
    const liveEx = "users/me/dataTypes/exercise/dataPoints/E-undated";
    const liveWt = "users/me/dataTypes/weight/dataPoints/W-undated";
    const trackedEx = "users/me/dataTypes/exercise/dataPoints/E-tracked";
    const trackedWt = "users/me/dataTypes/weight/dataPoints/W-tracked";
    reset([
      [
        "",
        "185",
        "135x5x3",
        "SYNC",
        "SYNC",
        JSON.stringify([liveEx, liveWt]),
        "",
        "",
        "",
        "",
      ],
      [
        new Date(),
        "",
        "225x5x3",
        "SYNC",
        "SYNC",
        JSON.stringify([trackedEx, trackedWt]),
        "",
        "",
        "",
        "",
      ],
    ]);
    const deleted = [];
    withStubs(
      {
        deleteDataPointsByName: (names) => {
          deleted.push(...names);
        },
        listStrengthOnDate: () => [
          { googleWebClientId: "ours", name: liveEx },
          { googleWebClientId: "ours", name: trackedEx },
        ],
        listWeightOnDate: () => [
          { googleWebClientId: "ours", name: liveWt },
          { googleWebClientId: "ours", name: trackedWt },
        ],
      },
      () => backstop(),
    );
    eq(
      deleted,
      [],
      "nothing deleted: the undated row still owns both datapoints",
    );
  });

  // ---- backstop: re-dirties BOTH matched and unmatched recent rows --------

  // backstop() uses the real Date.now() (not injectable) for its lookback
  // window, so rows must be dated "now". A Date object in the Date cell is read
  // back verbatim by toDate_, sidestepping the UTC-midnight civil-date boundary
  // that a 'yyyy-MM-dd' string would hit in a non-UTC script time zone.
  t(
    "backstop re-dirties BOTH matched and unmatched recent rows and sets the pending flag",
    () => {
      // Distinct ascending dates (yesterday, today) so the date validation pass
      // doesn't reject the sheet; both fall inside BACKSTOP_LOOKBACK_DAYS = 2.
      const today = new Date();
      const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);
      reset([
        [yesterday, "", "135x5x3", "SYNC", "", "", "", "", "", ""], // row 2: unmatched
        [today, "", "225x5x3", "SYNC", "", "", "", "", "", "foreign/F1"], // row 3: matched
      ]);
      withStubs(
        { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
        () => backstop(),
      );
      eq(
        SHEET.getRange(2, COL["Exercise Synced At"]).getValue(),
        "",
        "unmatched row re-dirtied",
      );
      eq(
        SHEET.getRange(3, COL["Exercise Synced At"]).getValue(),
        "",
        "matched row re-dirtied",
      );
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "pending flag set for the next poll",
      );
    },
  );

  // ---- backstop: reconciling cleared content ------------------------------

  // The end-to-end claim: a multi-cell clear that onEditMarkDirty deliberately
  // ignores is still reconciled, because the backstop compares tracked ids
  // against content. Two passes, exactly as it runs in production: the backstop
  // re-dirties, the following poll deletes.
  t(
    "a multi-cell clear that onEdit ignores is reconciled by the backstop",
    () => {
      const exName = "users/me/dataTypes/exercise/dataPoints/E1";
      const wtName = "users/me/dataTypes/weight/dataPoints/W1";
      reset([
        [
          new Date(),
          "185",
          "135x5x3",
          "SYNC",
          "SYNC",
          JSON.stringify([exName, wtName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      // The user selects Weight..Bench and hits Delete. onEdit does not mark it.
      SHEET.getRange(2, COL.Weight).setValue("");
      SHEET.getRange(2, COL.Bench).setValue("");
      ok(
        onEditMarkDirty({ range: SHEET.getRange(2, COL.Weight, 1, 2) }) ===
          false,
        "onEdit ignores the multi-cell clear",
      );
      eq(cell("Exercise Synced At"), "SYNC", "still stamped after the edit");

      withStubs(
        { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
        () => backstop(),
      );
      eq(
        cell("Exercise Synced At"),
        "",
        "backstop re-dirtied the exercise phase",
      );
      eq(cell("Weight Synced At"), "", "backstop re-dirtied the weight phase");
      ok(
        PROPS.getProperty("pendingDirty") !== null,
        "pending flag set for the poll",
      );

      const deleted = [];
      const r = withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => {
            throw new Error("must not recreate a cleared row");
          },
          createWeightAt: () => {
            throw new Error("must not recreate a cleared bodyweight");
          },
          deleteDataPointsByName: (names) => {
            deleted.push(...names);
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(r.ok, 1, "row synced");
      eq(
        deleted.sort(),
        [exName, wtName].sort(),
        "both stale datapoints deleted",
      );
      eq(cell("Created Health IDs"), JSON.stringify([]), "ids dropped");
    },
  );

  // The exact gap reported against the earlier onEdit-only design: a paste that
  // blanks the bodyweight while WRITING another cell. onEdit classifies by the
  // cells that have content, so the weight phase is never marked and its stamp
  // survives. The backstop is what closes it, because the row's state (weight
  // id tracked, Weight cell empty) says the datapoint should not exist.
  t(
    "a paste that blanks the bodyweight while writing another cell is reconciled",
    () => {
      const wtName = "users/me/dataTypes/weight/dataPoints/W1";
      reset([
        [
          new Date(),
          "185",
          "",
          "SYNC",
          "SYNC",
          JSON.stringify([wtName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      SHEET.getRange(2, COL.Weight).setValue(""); // blanked by the paste
      SHEET.getRange(2, COL.Bench).setValue("135x5x3"); // written by the paste
      onEditMarkDirty({ range: SHEET.getRange(2, COL.Date, 1, 3) });
      eq(
        cell("Weight Synced At"),
        "SYNC",
        "onEdit leaves the weight phase stamped",
      );

      withStubs(
        { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
        () => backstop(),
      );
      eq(
        cell("Weight Synced At"),
        "",
        "backstop caught the blanked bodyweight",
      );
      const deleted = [];
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => "users/me/dataTypes/exercise/dataPoints/E1",
          createWeightAt: () => {
            throw new Error("must not recreate a blanked bodyweight");
          },
          deleteDataPointsByName: (names) => {
            deleted.push(...names);
          },
        }),
        () => syncDirtyRows(0),
      );
      eq(deleted, [wtName], "the weight datapoint was deleted");
    },
  );

  // Accepted boundary: clearing a row INCLUDING its Date parks it (readRows
  // drops the row), so its content is not authoritative and it is never
  // reconciled as cleared. Recovery is orphan reconciliation, within its
  // lookback.
  // The managed columns keep the stamps and the tracked id, which is what makes
  // the assertion falsifiable: were the row still read, it would look emptied
  // (id tracked, no exercise text) and be re-dirtied for deletion.
  t("a row cleared including its Date is never reconciled as cleared", () => {
    const exName = "users/me/dataTypes/exercise/dataPoints/E1";
    reset([
      [
        new Date(),
        "185",
        "135x5x3",
        "SYNC",
        "SYNC",
        JSON.stringify([exName]),
        "",
        "",
        "",
        "",
      ],
    ]);
    for (const header of ["Date", "Weight", "Bench"]) {
      SHEET.getRange(2, COL[header]).setValue("");
    }
    withStubs(
      { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
      () => backstop(),
    );
    eq(cell("Exercise Synced At"), "SYNC", "parked row not re-dirtied");
    eq(cell("Weight Synced At"), "SYNC", "neither phase re-dirtied");
    eq(
      cell("Created Health IDs"),
      JSON.stringify([exName]),
      "its datapoint is still tracked",
    );
    ok(PROPS.getProperty("pendingDirty") === null, "nothing queued to delete");
  });

  // The scenario that makes the raw-text guard load-bearing: blanking ONE
  // exercise column header stops readRows building `exercises` for EVERY
  // historical row at once. onEditMarkDirty ignores row-1 edits, so nothing
  // warns. Without the guard the next backstop re-dirties the whole sheet and
  // the following poll deletes every exercise datapoint in it, unrecoverably
  // (the ids are dropped and only rows inside the lookback are re-pushed).
  t(
    "blanking an exercise column header does not make every row look cleared",
    () => {
      const grid = [];
      for (let i = 5; i >= 1; i--) {
        grid.push([
          new Date(Date.now() - i * 24 * 60 * 60 * 1000),
          "185",
          "135x5x3",
          "SYNC",
          "SYNC",
          JSON.stringify([`users/me/dataTypes/exercise/dataPoints/E${i}`]),
          "",
          "",
          "",
          "",
        ]);
      }
      reset(grid);
      SHEET.getRange(1, COL.Bench).setValue(""); // header blanked
      withStubs(
        { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
        () => backstop(),
      );
      for (let r = 2; r <= 6; r++) {
        eq(
          SHEET.getRange(r, COL["Exercise Synced At"]).getValue(),
          "SYNC",
          `row ${r} still stamped; its cells still hold text`,
        );
      }
      ok(
        PROPS.getProperty("pendingDirty") === null,
        "nothing queued for deletion",
      );
    },
  );

  // Same shape via the weight parser: a reformat it rejects returns null, which
  // is not the same claim as an empty cell.
  t(
    "a bodyweight the parser rejects does not look like a cleared bodyweight",
    () => {
      reset([
        [
          new Date(),
          "185 lb",
          "",
          "SYNC",
          "SYNC",
          JSON.stringify(["users/me/dataTypes/weight/dataPoints/W1"]),
          "",
          "",
          "",
          "",
        ],
      ]);
      eq(
        parseBodyweight("185 lb"),
        null,
        "precondition: the parser rejects it",
      );
      withStubs(
        { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
        () => backstop(),
      );
      eq(
        cell("Weight Synced At"),
        "SYNC",
        "datapoint left alone; the cell still holds text",
      );
    },
  );

  // The bound on the destructive branch: past it, reconcile NOTHING rather than
  // part of it, so the sheet is left exactly as found.
  t(
    "backstop reconciles nothing when more rows look cleared than a person would clear",
    () => {
      const grid = [];
      for (let i = STALE_RECONCILE_MAX_ROWS + 1; i >= 1; i--) {
        grid.push([
          new Date(Date.now() - i * 24 * 60 * 60 * 1000),
          "",
          "",
          "SYNC",
          "SYNC",
          JSON.stringify([`users/me/dataTypes/exercise/dataPoints/E${i}`]),
          "",
          "",
          "",
          "",
        ]);
      }
      reset(grid);
      withStubs(
        { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
        () => backstop(),
      );
      for (let r = 2; r <= grid.length + 1; r++) {
        eq(
          SHEET.getRange(r, COL["Exercise Synced At"]).getValue(),
          "SYNC",
          `row ${r} untouched`,
        );
      }
      ok(PROPS.getProperty("pendingDirty") === null, "nothing queued");
    },
  );

  t("backstop still reconciles a stale set at the limit", () => {
    const grid = [];
    for (let i = STALE_RECONCILE_MAX_ROWS; i >= 1; i--) {
      grid.push([
        new Date(Date.now() - i * 24 * 60 * 60 * 1000),
        "",
        "",
        "SYNC",
        "SYNC",
        JSON.stringify([`users/me/dataTypes/exercise/dataPoints/E${i}`]),
        "",
        "",
        "",
        "",
      ]);
    }
    reset(grid);
    withStubs(
      { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
      () => backstop(),
    );
    eq(
      SHEET.getRange(2, COL["Exercise Synced At"]).getValue(),
      "",
      "reconciled at the limit",
    );
  });

  // Deliberately unbounded by BACKSTOP_LOOKBACK_DAYS: a clear on an old row is
  // exactly the case nothing else recovers, and the scan makes no API calls.
  t(
    "backstop reconciles a cleared row older than the foreign-match lookback",
    () => {
      const exName = "users/me/dataTypes/exercise/dataPoints/E-old";
      const old = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
      reset([
        [old, "", "", "SYNC", "SYNC", JSON.stringify([exName]), "", "", "", ""],
      ]);
      withStubs(
        { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
        () => backstop(),
      );
      eq(
        cell("Exercise Synced At"),
        "",
        "stale row re-dirtied regardless of age",
      );
    },
  );

  // Self-limiting: once the sync drops the id the row stops matching, so a
  // steady-state sheet must not be re-dirtied on every backstop run.
  t("backstop leaves a consistent row alone", () => {
    const exName = "users/me/dataTypes/exercise/dataPoints/E1";
    const old = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
    reset([
      [
        old,
        "185",
        "135x5x3",
        "SYNC",
        "SYNC",
        JSON.stringify([exName]),
        "",
        "",
        "",
        "",
      ],
    ]);
    withStubs(
      { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
      () => backstop(),
    );
    eq(cell("Exercise Synced At"), "SYNC", "untouched");
    ok(PROPS.getProperty("pendingDirty") === null, "no pending flag");
  });

  // ---- resyncSelectedRows: selection guards --------------------------------

  // The selection is spreadsheet-global, so row numbers from another tab would
  // re-dirty unrelated rows on the synced tab.
  t("resyncSelectedRows refuses a selection on another tab", () => {
    reset([["2026-01-15", "", "135x5", "SYNC", "SYNC", "", "", "", "", ""]]);
    SHEET._setSelection([[2, 1]]);
    ACTIVE.sheet = H.otherSheet;
    try {
      withStubs(NO_FOREIGN, () => resyncSelectedRows());
    } finally {
      ACTIVE.sheet = SHEET;
    }
    eq(cell("Exercise Synced At"), "SYNC", "stamps not cleared");
    ok(PROPS.getProperty("pendingDirty") === null, "no pending flag");
    ok(
      TOASTS.length === 1 && TOASTS[0].indexOf("tab first") !== -1,
      `told the user why nothing happened: ${JSON.stringify(TOASTS)}`,
    );
  });

  // Clicking a column header selects the full sheet height. Without clamping,
  // every empty row below the data would be re-dirtied and stamped.
  // Asserts on which rows were targeted, not on the resulting cell values:
  // clearing a stamp writes '', which is indistinguishable from an untouched
  // empty cell below the data, so a value-based assertion passes either way.
  t(
    "resyncSelectedRows clamps a whole-column selection to the data range",
    () => {
      reset([
        ["2026-01-15", "", "135x5", "SYNC", "SYNC", "[]", "", "", "", ""],
        ["2026-01-16", "", "225x5", "SYNC", "SYNC", "[]", "", "", "", ""],
      ]);
      SHEET._setSelection([[1, 1000]]); // whole column, header row included
      const cleared = [];
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          clearRowExerciseSynced: (rowNum) => {
            cleared.push(rowNum);
          },
          clearRowWeightSynced: () => {},
          createExerciseAt: () => "users/me/dataTypes/exercise/dataPoints/E",
        }),
        () => resyncSelectedRows(),
      );
      eq(
        cleared,
        [2, 3],
        "only the two data rows were re-dirtied, not the full sheet height",
      );
    },
  );

  t(
    "resyncSelectedRows toasts instead of throwing when nothing is selected",
    () => {
      reset([["2026-01-15", "", "135x5", "SYNC", "SYNC", "", "", "", "", ""]]);
      SHEET._setSelection(null);
      withStubs(NO_FOREIGN, () => resyncSelectedRows());
      eq(cell("Exercise Synced At"), "SYNC", "stamps not cleared");
      ok(
        TOASTS.length === 1 &&
          TOASTS[0].indexOf("No data rows selected") !== -1,
        `told the user why nothing happened: ${JSON.stringify(TOASTS)}`,
      );
    },
  );

  // The re-dirty is followed by an immediate sync, so the observable effect is
  // the recreated datapoint rather than a cleared stamp.
  t(
    "resyncSelectedRows re-dirties and re-pushes the selected rows on the synced tab",
    () => {
      reset([
        ["2026-01-15", "", "135x5", "SYNC", "SYNC", "[]", "", "", "", ""],
      ]);
      SHEET._setSelection([[2, 1]]);
      const newName = "users/me/dataTypes/exercise/dataPoints/E9";
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          createExerciseAt: () => newName,
        }),
        () => resyncSelectedRows(),
      );
      eq(
        cell("Created Health IDs"),
        JSON.stringify([newName]),
        "exercise datapoint recreated",
      );
      ok(cell("Exercise Synced At") !== "SYNC", "exercise stamp refreshed");
    },
  );

  // resyncAllRows' happy path: clearing BOTH stamp columns across EVERY data
  // row, then marking the sheet dirty. Only its date-validation abort was
  // covered, so it could have cleared nothing, or only the first row, and the
  // suite would not have noticed. It is the documented manual recovery path.
  t("resyncAllRows clears both stamps on every data row and re-syncs", () => {
    reset([
      ["2026-01-15", "", "135x5", "SYNC-E1", "SYNC-W1", "[]", "", "", "", ""],
      ["2026-01-16", "", "145x5", "SYNC-E2", "SYNC-W2", "[]", "", "", "", ""],
      ["2026-01-17", "", "155x5", "SYNC-E3", "SYNC-W3", "[]", "", "", "", ""],
    ]);
    let created = 0;
    withStubs(
      Object.assign({}, NO_FOREIGN, {
        createExerciseAt: () => {
          created++;
          return `users/me/dataTypes/exercise/dataPoints/E${created}`;
        },
      }),
      () => resyncAllRows(),
    );
    eq(created, 3, "every data row was re-pushed, not just the first");
    for (let r = 2; r <= 4; r++) {
      ok(
        String(SHEET.getRange(r, COL["Exercise Synced At"]).getValue()).indexOf(
          "SYNC-E",
        ) === -1,
        `row ${r} exercise stamp was cleared and rewritten`,
      );
      ok(
        String(SHEET.getRange(r, COL["Weight Synced At"]).getValue()).indexOf(
          "SYNC-W",
        ) === -1,
        `row ${r} weight stamp was cleared and rewritten`,
      );
    }
  });

  // hasExerciseText is an OR across every exercise column: one column still
  // holding text means the row was not emptied. Every other fixture has exactly
  // one exercise column, where "some" and "every" agree, so this is the only
  // test that can tell them apart. It only bites where the parse yields nothing
  // sendable but a cell still holds text, which is the case the raw-text check
  // exists for: unparseable content in one column and a blank in another must
  // not read as a cleared row, or the backstop deletes the row's datapoint.
  t(
    "unparseable text in one of two exercise columns is not a cleared row",
    () => {
      const eName = "users/me/dataTypes/exercise/dataPoints/E1";
      resetGrid([
        [
          "Date",
          "Weight",
          "Bench",
          "Squat",
          "Exercise Synced At",
          "Weight Synced At",
          "Created Health IDs",
          "Exercise First Edited At",
          "Exercises Last Edited At",
          "Weight Edited At",
          "Matched Health Session",
        ],
        [
          // Outside BACKSTOP_LOOKBACK_DAYS, so the foreign re-review leaves the
          // row alone and only the cleared-content reconciliation can act on it.
          new Date(Date.now() - 10 * 24 * 60 * 60 * 1000),
          "",
          "bench press day", // holds text, parses to nothing sendable
          "", // Squat never filled in
          "SYNC",
          "SYNC",
          JSON.stringify([eName]),
          "",
          "",
          "",
          "",
        ],
      ]);
      const deleted = [];
      withStubs(
        {
          deleteDataPointsByName: (names) => deleted.push(names.slice()),
          listStrengthOnDate: () => [],
          listWeightOnDate: () => [],
        },
        () => backstop(),
      );
      eq(
        deleted,
        [],
        "the backstop only re-dirties; it never deletes directly",
      );
      eq(
        SHEET.getRange(2, 5).getValue(),
        "SYNC",
        "not re-dirtied: the Bench cell still holds text, so nothing was cleared",
      );
    },
  );

  // A multi-cell range must never take the single-cell clear path. Apps Script
  // only supplies e.oldValue for single-cell edits, so the oldValue test alone
  // happens to be enough today; the singleCell conjunction is what keeps that
  // true if the platform ever starts reporting an oldValue for a range. The
  // consequence of losing it is a paste being read as a clear, which deletes
  // datapoints.
  t(
    "a multi-cell range carrying an oldValue is not read as a cleared cell",
    () => {
      // Weight and Bench both blank, as they would be right after a two-cell
      // clear, with an oldValue the platform does not actually supply for a range.
      reset([["2026-01-15", "", "", "SYNC", "SYNC", "[]", "", "", "", ""]]);
      const marked = onEditMarkDirty({
        oldValue: "135x5",
        range: SHEET.getRange(2, COL.Weight, 1, 2),
      });
      ok(marked === false, "a blank multi-cell range marks nothing");
      eq(cell("Exercise Synced At"), "SYNC", "exercise stamp untouched");
      eq(cell("Weight Synced At"), "SYNC", "weight stamp untouched");
    },
  );

  // ---- whitespace tolerance -----------------------------------------------
  // Both trims below are what make a plausible typo survivable, and neither was
  // covered: every fixture uses exact header names and exact cell values.

  // A header typed with a trailing space is the common version of this. Without
  // the trim in getHeaderMap_ the lookup misses, and the sync aborts with
  // "Weight Synced At column missing. Run setup." on a sheet that looks correct.
  t("headers with surrounding whitespace are still recognized", () => {
    reset([["2026-01-15", "185", "135x5", "", "", "", "", "", "", ""]]);
    HEADERS.forEach((h, i) => {
      SHEET.getRange(1, i + 1).setValue(` ${h} `);
    });
    const wName = "users/me/dataTypes/weight/dataPoints/W1";
    const eName = "users/me/dataTypes/exercise/dataPoints/E1";
    const r = withStubs(
      Object.assign({}, NO_FOREIGN, {
        createExerciseAt: () => eName,
        createWeightAt: () => wName,
      }),
      () => syncDirtyRows(0),
    );
    eq(r.errors, 0, "no structural error from the padded headers");
    eq(r.ok, 1, "the row synced");
    eq(cell("Created Health IDs"), JSON.stringify([wName, eName]));
  });

  // Clearing a cell by typing a space is the other one. hasText decides whether
  // the backstop treats the row as emptied, so without the trim the row keeps a
  // datapoint for a workout the cell no longer describes.
  t("a cell holding only whitespace counts as cleared", () => {
    const eName = "users/me/dataTypes/exercise/dataPoints/E1";
    reset([
      [
        new Date(Date.now() - 10 * 24 * 60 * 60 * 1000),
        "",
        " ", // "cleared" by typing a space
        "SYNC",
        "SYNC",
        JSON.stringify([eName]),
        "",
        "",
        "",
        "",
      ],
    ]);
    withStubs(
      { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
      () => backstop(),
    );
    eq(
      cell("Exercise Synced At"),
      "",
      "re-dirtied, so the next sync deletes the datapoint",
    );
  });

  // ---- trigger-entry date validation --------------------------------------

  t(
    "syncOnEdit throws on a date validation violation (uncaught -> owner email)",
    () => {
      reset([
        ["2026-01-15", "", "135x5", "", "", "", "", "", "", ""],
        ["2026-01-15", "", "225x5", "", "", "", "", "", "", ""], // duplicate date
      ]);
      let thrown = null;
      try {
        syncOnEdit({ range: SHEET.getRange(2, COL.Bench) });
      } catch (err) {
        thrown = err;
      }
      ok(thrown !== null, "throws out of syncOnEdit");
      ok(
        String(thrown).indexOf("date validation failed") !== -1,
        `message names the validation: ${thrown}`,
      );
      // Thrown before dirty marking, so the edit is not recorded.
      eq(cell("Exercises Last Edited At"), "", "edit not dirty-marked");
      ok(PROPS.getProperty("pendingDirty") === null, "no pending flag");
    },
  );

  t(
    "flushPending logs and skips (no throw) on a date validation violation",
    () => {
      reset([
        ["2026-01-16", "185", "", "SYNC", "", "", "", "", "", ""],
        ["2026-01-15", "186", "", "SYNC", "", "", "", "", "", ""], // out of order
      ]);
      PROPS.setProperty("pendingDirty", "GEN1");
      const calls = [];
      withStubs(
        Object.assign({}, NO_FOREIGN, {
          createWeightAt: () => {
            calls.push("create");
            return "users/me/dataTypes/weight/dataPoints/W1";
          },
        }),
        () => flushPending(),
      );
      eq(calls, [], "no sync work attempted");
      eq(
        PROPS.getProperty("pendingDirty"),
        "GEN1",
        "dirty flag left so the backlog syncs after the fix",
      );
    },
  );

  t("backstop skips on a date validation violation", () => {
    reset([["2024-06-01", "", "135x5x3", "SYNC", "", "", "", "", "", ""]]); // year below MIN
    withStubs(
      { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
      () => backstop(),
    );
    eq(cell("Exercise Synced At"), "SYNC", "row not re-dirtied");
    ok(PROPS.getProperty("pendingDirty") === null, "no pending flag");
  });

  t(
    "manual resyncAllRows toasts and aborts on a date validation violation",
    () => {
      reset([
        ["2026-01-15", "", "135x5", "SYNC", "SYNC", "", "", "", "", ""],
        ["2026-01-15", "", "225x5", "SYNC", "SYNC", "", "", "", "", ""], // duplicate date
      ]);
      withStubs(NO_FOREIGN, () => resyncAllRows());
      eq(cell("Exercise Synced At"), "SYNC", "stamps not cleared");
      eq(cell("Weight Synced At"), "SYNC", "weight stamps not cleared");
      ok(PROPS.getProperty("pendingDirty") === null, "no pending flag");
      ok(
        TOASTS.length === 1 &&
          TOASTS[0].indexOf("Date validation failed") !== -1,
        `surfaced the violation to the user: ${JSON.stringify(TOASTS)}`,
      );
    },
  );

  t("backstop skips entirely when the lock is held", () => {
    reset([[new Date(), "", "135x5x3", "SYNC", "", "", "", "", "", ""]]);
    LOCK.held = true;
    withStubs(
      { listStrengthOnDate: () => [], listWeightOnDate: () => [] },
      () => backstop(),
    );
    eq(
      SHEET.getRange(2, COL["Exercise Synced At"]).getValue(),
      "SYNC",
      "row untouched when lock held",
    );
    ok(
      PROPS.getProperty("pendingDirty") === null,
      "no pending flag when skipped",
    );
  });

  return results;
}

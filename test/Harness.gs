// In-memory fakes for the Apps Script services the orchestration code touches
// (SpreadsheetApp, PropertiesService, LockService), and the wrapper that swaps
// them in for the duration of a suite.
//
// Pushed to Apps Script, unlike run.js, so `Sync ▸ Run tests` runs the
// orchestration suite there as well as locally. run.js loads this same file
// into its VM sandbox, so the fakes have one definition rather than two that
// drift apart.
//
// ES2019: Apps Script's V8 runtime is old enough that optional chaining and
// nullish coalescing do not parse here. Object keys are alphabetical because
// the .gs lint config requires it, not because the order means anything.

function makeFakeSheet_(sheetId) {
  let grid = []; // grid[r0][c0], 0-indexed; auto-grows on write
  const isEmpty = (v) => v === "" || v === null || v === undefined;
  const ensure = (r0, c0) => {
    while (grid.length <= r0) {
      grid.push([]);
    }
    for (const row of grid) {
      while (row.length <= c0) {
        row.push("");
      }
    }
  };
  const getCell = (r0, c0) =>
    grid[r0] && grid[r0][c0] !== undefined ? grid[r0][c0] : "";
  const setCell = (r0, c0, v) => {
    ensure(r0, c0);
    grid[r0][c0] = v;
  };

  const sheet = {
    _selection: null,
    _setGrid(rows) {
      grid = rows.map((r) => r.slice());
    },
    // Selection fake for resyncSelectedRows: _setSelection([[firstRow, numRows], ...]),
    // or null to simulate "nothing selected".
    _setSelection(specs) {
      sheet._selection =
        specs === null
          ? null
          : {
              getRanges: () =>
                specs.map(([row, numRows]) =>
                  sheet.getRange(row, 1, numRows, 1),
                ),
            };
    },
    getActiveRangeList: () => sheet._selection,
    getLastColumn() {
      let last = 0;
      for (let r0 = 0; r0 < grid.length; r0++) {
        for (let c0 = (grid[r0] || []).length - 1; c0 >= 0; c0--) {
          if (!isEmpty(grid[r0][c0])) {
            if (c0 + 1 > last) {
              last = c0 + 1;
            }
            break;
          }
        }
      }
      return last;
    },
    getLastRow() {
      let last = 0;
      for (let r0 = 0; r0 < grid.length; r0++) {
        if (grid[r0].some((v) => !isEmpty(v))) {
          last = r0 + 1;
        }
      }
      return last;
    },
    getName: () => `Sheet${sheetId}`,
    getRange(row, col, numRows, numCols) {
      const rows = numRows || 1;
      const cols = numCols || 1;
      return {
        getColumn: () => col,
        getLastColumn: () => col + cols - 1,
        getLastRow: () => row + rows - 1,
        getNumRows: () => rows,
        getRow: () => row,
        getSheet: () => sheet,
        getValue: () => getCell(row - 1, col - 1),
        getValues: () => {
          const out = [];
          for (let i = 0; i < rows; i++) {
            const r = [];
            for (let j = 0; j < cols; j++) {
              r.push(getCell(row - 1 + i, col - 1 + j));
            }
            out.push(r);
          }
          return out;
        },
        setValue: (v) => {
          setCell(row - 1, col - 1, v);
        },
        setValues: (vals) => {
          // The real service rejects a mismatch; without this check an
          // off-by-one in a bulk write passes locally and throws in Apps Script.
          if (vals.length !== rows || vals.some((r) => r.length !== cols)) {
            throw new Error(
              `setValues: data is ${vals.length}x${(vals[0] || []).length}, range is ${rows}x${cols}`,
            );
          }
          for (let i = 0; i < rows; i++) {
            for (let j = 0; j < cols; j++) {
              setCell(row - 1 + i, col - 1 + j, vals[i][j]);
            }
          }
        },
      };
    },
    getSheetId: () => sheetId,
    hideColumns: () => {},
  };
  return sheet;
}

function makeFakeStore_() {
  let m = {};
  return {
    _clear: () => {
      m = {};
    },
    deleteProperty: (k) => {
      delete m[k];
    },
    getProperty: (k) => (k in m ? m[k] : null),
    setProperty: (k, v) => {
      m[k] = String(v);
    },
  };
}

// A fresh set of fakes plus the service objects that front them. Every call is
// independent, so one run never inherits grid or property state from an earlier
// one. `_syncTestFake` is the marker withSyncTestHarness_ verifies.
function buildSyncTestHarness_() {
  const sheet = makeFakeSheet_(1);
  // A second tab the sync never manages, so tests can put the selection on the
  // "wrong" sheet the way a user with multiple tabs would.
  const otherSheet = makeFakeSheet_(2);
  const activeSheetRef = { sheet };
  const toasts = [];
  const scriptProps = makeFakeStore_();
  const lockState = { held: false };

  const spreadsheet = {
    getActiveSheet: () => activeSheetRef.sheet,
    getSheets: () => [sheet, otherSheet],
    getUi: () => {
      throw new Error("no UI");
    },
    toast: (msg) => {
      toasts.push(String(msg));
    },
  };
  const makeLock = () => {
    let owned = false;
    return {
      releaseLock: () => {
        if (owned) {
          lockState.held = false;
          owned = false;
        }
      },
      tryLock: () => {
        if (lockState.held && !owned) {
          return false;
        }
        lockState.held = true;
        owned = true;
        return true;
      },
    };
  };

  return {
    handle: {
      activeSheetRef,
      lockState,
      otherSheet,
      scriptProps,
      sheet,
      toasts,
    },
    services: {
      LockService: {
        _syncTestFake: true,
        getScriptLock: makeLock,
        getUserLock: makeLock,
      },
      PropertiesService: {
        _syncTestFake: true,
        getScriptProperties: () => scriptProps,
        getUserProperties: makeFakeStore_,
      },
      // Not used by the orchestration suite either: no test runs setup() or
      // installTriggers(). Faked for the same reason as UrlFetchApp below: in
      // Apps Script the real service is in scope, and a future test reaching
      // it would delete and recreate the project's real triggers mid-run.
      ScriptApp: {
        _syncTestFake: true,
        getProjectTriggers: () => {
          throw new Error(
            "ScriptApp: the orchestration suite must not touch real triggers.",
          );
        },
        newTrigger: () => {
          throw new Error(
            "ScriptApp: the orchestration suite must not touch real triggers.",
          );
        },
      },
      SpreadsheetApp: {
        _syncTestFake: true,
        flush: () => {},
        getActiveSpreadsheet: () => spreadsheet,
        getUi: () => {
          throw new Error("no UI");
        },
      },
      // Not used by the orchestration suite: every test stubs the Health API
      // functions above httpJson_. It is faked so that a test which forgets one
      // fails here instead of reaching the real Health account. Locally the
      // sandbox simply has no UrlFetchApp, so the same slip is a ReferenceError;
      // in Apps Script the real service is in scope and would send the request.
      UrlFetchApp: {
        _syncTestFake: true,
        fetch: () => {
          throw new Error(
            "UrlFetchApp: the orchestration suite must stub the Health API; " +
              "an unstubbed call would reach the real account.",
          );
        },
      },
    },
  };
}

// Swaps the fakes in over the real services, calls fn(handle), and restores in
// a finally. Returns fn's value.
//
// The swap is verified before fn runs, and a failure throws rather than
// degrading: in Apps Script the real SpreadsheetApp and PropertiesService reach
// the live spreadsheet and the real script properties, and this suite writes
// rows, clears stamps and deletes the pending-dirty flag. An assignment that
// silently failed to shadow would run all of that against real data. Refusing
// to start is the only safe failure, so do not soften this into a warning.
function withSyncTestHarness_(fn) {
  const built = buildSyncTestHarness_();
  const names = Object.keys(built.services);
  const saved = {};
  names.forEach((n) => {
    saved[n] = globalThis[n];
    globalThis[n] = built.services[n];
  });
  try {
    // Read through the bare identifiers the source files use, not through
    // globalThis: those are the bindings that have to resolve to the fakes.
    if (
      SpreadsheetApp._syncTestFake !== true ||
      PropertiesService._syncTestFake !== true ||
      LockService._syncTestFake !== true ||
      UrlFetchApp._syncTestFake !== true ||
      ScriptApp._syncTestFake !== true
    ) {
      throw new Error(
        "withSyncTestHarness_: the Apps Script services did not shadow; " +
          "refusing to run the orchestration suite against live data.",
      );
    }
    return fn(built.handle);
  } finally {
    names.forEach((n) => {
      globalThis[n] = saved[n];
    });
  }
}

// One suite's PASS/FAIL list, counted. Shared so the suites and runAllTests
// word their summaries the same way.
function summarizeTestResults_(results) {
  const failures = results.filter((r) => r.startsWith("FAIL "));
  const passed = results.length - failures.length;
  return {
    failures,
    summary: `${results.length} tests: ${passed} passed, ${failures.length} failed`,
  };
}

// alert() blocks until the dialog is dismissed, and getUi() succeeds whenever
// the bound spreadsheet is open, including for a run started from the editor's
// Run button. So an execution shows at most one: two suites that each alerted
// left the second one waiting behind a modal in a tab nobody was looking at,
// which reads as a hung execution. Keep it to one per entry point.
function showTestAlert_(text) {
  try {
    SpreadsheetApp.getUi().alert(text);
  } catch {
    /* no UI: the container is closed, or this is not a bound execution */
  }
}

// Full dump plus the single alert, for a suite run on its own.
function reportTestResults_(title, results) {
  const msg = `${results.join("\n")}\n\n${summarizeTestResults_(results).summary}`;
  console.log(msg);
  showTestAlert_(`${title}\n\n${msg}`);
}

// Menu entry point for `Sync ▸ Run tests`, and the only place that knows about
// both suites. No arguments, because the Apps Script editor's Run button
// supplies none.
//
// Reports compactly rather than dumping both suites: Apps Script truncates an
// over-long execution log, and one suite's full PASS list already approaches
// the limit, so dumping both drops the second summary, which is the line worth
// reading. Run a suite on its own for its full list.
function runAllTests() {
  const suites = [
    { results: runParserTestsBody_(), title: "Parser / pure-helper tests" },
    {
      results: withSyncTestHarness_(runSyncTestsBody_),
      title: "Orchestration tests",
    },
  ];
  const lines = [];
  suites.forEach((suite) => {
    const counted = summarizeTestResults_(suite.results);
    lines.push(`${suite.title}: ${counted.summary}`);
    counted.failures.forEach((f) => lines.push(`  ${f}`));
  });
  const msg = lines.join("\n");
  console.log(msg);
  showTestAlert_(msg);
}

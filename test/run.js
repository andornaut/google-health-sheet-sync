const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const srcFiles = ['Config.gs', 'Parser.gs', 'Format.gs', 'Sheet.gs', 'HealthApi.gs', 'Main.gs'];
const testFiles = ['Parser.test.gs', 'Sync.test.gs'];

// Silence the code's diagnostic chatter (warn/info) so the suite output is just
// the PASS/FAIL lines, which are emitted via console.log and captured below.
const quietConsole = Object.assign({}, console, { warn: () => {}, info: () => {} });

// Minimal Utilities.formatDate stub backed by Node's Intl. Covers the format
// strings used in src/*.gs; throws on anything else so a new usage shows up
// loudly instead of silently returning wrong data.
function formatDateStub(date, tz, format) {
  const zone = (tz === 'GMT' || tz === 'UTC') ? 'UTC' : tz;
  if (format === 'Z') {
    if (zone === 'UTC') return '+0000';
    const parts = new Intl.DateTimeFormat('en-US', { timeZone: zone, timeZoneName: 'longOffset', year: 'numeric' }).formatToParts(date);
    const tzn = parts.find(p => p.type === 'timeZoneName').value;
    if (tzn === 'GMT' || tzn === 'UTC') return '+0000';
    const m = tzn.match(/GMT([+-])(\d{2}):(\d{2})/);
    if (!m) throw new Error('formatDateStub: unparseable offset "' + tzn + '"');
    return m[1] + m[2] + m[3];
  }
  if (format === 'yyyy MM dd HH mm ss') {
    const parts = new Intl.DateTimeFormat('en-US', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).formatToParts(date);
    const get = t => parts.find(p => p.type === t).value;
    let h = get('hour');
    if (h === '24') h = '00';
    return get('year') + ' ' + get('month') + ' ' + get('day') + ' ' + h + ' ' + get('minute') + ' ' + get('second');
  }
  if (format === "yyyy-MM-dd'T'HH:mm:ss'Z'") {
    if (zone !== 'UTC') throw new Error('formatDateStub: ISO Z format only stubbed for GMT/UTC');
    return date.toISOString().slice(0, 19) + 'Z';
  }
  if (format === 'yyyy-MM-dd') {
    const parts = new Intl.DateTimeFormat('en-CA', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(date);
    const get = t => parts.find(p => p.type === t).value;
    return get('year') + '-' + get('month') + '-' + get('day');
  }
  throw new Error('formatDateStub: unmocked format "' + format + '"');
}

let scriptTimeZone = 'America/Toronto';
const Utilities = {
  formatDate: formatDateStub,
  sleep: () => {}
};
const Session = {
  getScriptTimeZone: () => scriptTimeZone
};

// ---------------------------------------------------------------------------
// Minimal in-memory fakes for the Apps Script services the orchestration code
// (syncDirtyRows, syncOneRow_, onEditMarkDirty) touches. These let Sync.test.gs
// exercise the stateful glue — dirty-flag lifecycle, phase dispatch, idempotency
// — that the pure-helper tests can't reach. The Health API functions themselves
// are stubbed per-test via globalThis (same pattern as listStrengthOnDate), so
// no UrlFetchApp fake is needed. Exposed to the sandbox as SYNC_TEST_HARNESS_.
// ---------------------------------------------------------------------------
function makeFakeSheet(sheetId) {
  let grid = []; // grid[r0][c0], 0-indexed; auto-grows on write
  const isEmpty = v => v === '' || v === null || v === undefined;
  const ensure = (r0, c0) => {
    while (grid.length <= r0) grid.push([]);
    for (const row of grid) while (row.length <= c0) row.push('');
  };
  const getCell = (r0, c0) => (grid[r0] && grid[r0][c0] !== undefined ? grid[r0][c0] : '');
  const setCell = (r0, c0, v) => { ensure(r0, c0); grid[r0][c0] = v; };

  const sheet = {
    _setGrid(rows) { grid = rows.map(r => r.slice()); },
    // Selection fake for resyncSelectedRows: _setSelection([[firstRow, numRows], ...]),
    // or null to simulate "nothing selected".
    _setSelection(specs) {
      sheet._selection = specs === null ? null : {
        getRanges: () => specs.map(([row, numRows]) => sheet.getRange(row, 1, numRows, 1))
      };
    },
    _selection: null,
    getActiveRangeList: () => sheet._selection,
    getSheetId: () => sheetId,
    getName: () => 'Sheet' + sheetId,
    hideColumns: () => {},
    getLastRow() {
      let last = 0;
      for (let r0 = 0; r0 < grid.length; r0++) {
        if (grid[r0].some(v => !isEmpty(v))) last = r0 + 1;
      }
      return last;
    },
    getLastColumn() {
      let last = 0;
      for (let r0 = 0; r0 < grid.length; r0++) {
        for (let c0 = (grid[r0] || []).length - 1; c0 >= 0; c0--) {
          if (!isEmpty(grid[r0][c0])) { if (c0 + 1 > last) last = c0 + 1; break; }
        }
      }
      return last;
    },
    getRange(row, col, numRows, numCols) {
      numRows = numRows || 1;
      numCols = numCols || 1;
      return {
        getRow: () => row,
        getColumn: () => col,
        getNumRows: () => numRows,
        getLastRow: () => row + numRows - 1,
        getLastColumn: () => col + numCols - 1,
        getSheet: () => sheet,
        getValue: () => getCell(row - 1, col - 1),
        setValue: v => { setCell(row - 1, col - 1, v); },
        getValues: () => {
          const out = [];
          for (let i = 0; i < numRows; i++) {
            const r = [];
            for (let j = 0; j < numCols; j++) r.push(getCell(row - 1 + i, col - 1 + j));
            out.push(r);
          }
          return out;
        },
        setValues: vals => {
          for (let i = 0; i < numRows; i++) {
            for (let j = 0; j < numCols; j++) setCell(row - 1 + i, col - 1 + j, vals[i][j]);
          }
        }
      };
    }
  };
  return sheet;
}

function makeFakeStore() {
  let m = {};
  return {
    getProperty: k => (k in m ? m[k] : null),
    setProperty: (k, v) => { m[k] = String(v); },
    deleteProperty: k => { delete m[k]; },
    _clear: () => { m = {}; }
  };
}

const fakeSheet = makeFakeSheet(1);
// A second tab the sync never manages, so tests can put the selection on the
// "wrong" sheet the way a user with multiple tabs would.
const otherFakeSheet = makeFakeSheet(2);
const activeSheetRef = { sheet: fakeSheet };
const toasts = [];
const fakeSpreadsheet = {
  getSheets: () => [fakeSheet, otherFakeSheet],
  getActiveSheet: () => activeSheetRef.sheet,
  toast: msg => { toasts.push(String(msg)); },
  getUi: () => { throw new Error('no UI'); }
};
const scriptProps = makeFakeStore();
const lockState = { held: false };
const makeLock = () => {
  let owned = false;
  return {
    tryLock: () => {
      if (lockState.held && !owned) return false;
      lockState.held = true;
      owned = true;
      return true;
    },
    releaseLock: () => { if (owned) { lockState.held = false; owned = false; } }
  };
};

const SpreadsheetApp = {
  getActiveSpreadsheet: () => fakeSpreadsheet,
  getUi: () => { throw new Error('no UI'); },
  flush: () => {}
};
const PropertiesService = {
  getScriptProperties: () => scriptProps,
  getUserProperties: () => makeFakeStore()
};
const LockService = {
  getScriptLock: makeLock,
  getUserLock: makeLock
};

const sandbox = {
  console: quietConsole,
  SpreadsheetApp: SpreadsheetApp,
  PropertiesService: PropertiesService,
  LockService: LockService,
  Utilities: Utilities,
  Session: Session,
  setTestTimeZone: tz => { scriptTimeZone = tz; },
  SYNC_TEST_HARNESS_: {
    sheet: fakeSheet,
    otherSheet: otherFakeSheet,
    activeSheetRef: activeSheetRef,
    toasts: toasts,
    scriptProps: scriptProps,
    lockState: lockState
  }
};
vm.createContext(sandbox);

for (const f of srcFiles) {
  const code = fs.readFileSync(path.join(root, 'src', f), 'utf8');
  vm.runInContext(code, sandbox, { filename: `src/${f}` });
}
for (const f of testFiles) {
  const code = fs.readFileSync(path.join(root, 'test', f), 'utf8');
  vm.runInContext(code, sandbox, { filename: `test/${f}` });
}

const logs = [];
const origLog = quietConsole.log;
quietConsole.log = (...args) => { logs.push(args.join(' ')); origLog(...args); };
try {
  vm.runInContext('runParserTests();', sandbox);
  vm.runInContext('runSyncTests();', sandbox);
} finally {
  quietConsole.log = origLog;
}

const output = logs.join('\n');
const passed = (output.match(/^PASS /gm) || []).length;
const failed = (output.match(/^FAIL /gm) || []).length;
console.log(`\n${passed + failed} tests: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  process.exit(1);
}

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const srcFiles = ['Config.gs', 'Parser.gs', 'Format.gs', 'Sheet.gs', 'HealthApi.gs', 'Main.gs'];
const testFiles = ['Parser.test.gs'];

const quietConsole = Object.assign({}, console, { warn: () => {} });

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

const sandbox = {
  console: quietConsole,
  SpreadsheetApp: { getUi: () => { throw new Error('no UI'); } },
  Utilities: Utilities,
  Session: Session,
  setTestTimeZone: tz => { scriptTimeZone = tz; }
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
const origLog = console.log;
console.log = (...args) => { logs.push(args.join(' ')); origLog(...args); };
try {
  vm.runInContext('runParserTests();', sandbox);
} finally {
  console.log = origLog;
}

const output = logs.join('\n');
if (/^FAIL /m.test(output)) {
  process.exit(1);
}

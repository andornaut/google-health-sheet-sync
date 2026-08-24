const fs = require("fs");
const path = require("path");
const vm = require("vm");

const root = path.resolve(__dirname, "..");
const srcFiles = [
  "Config.gs",
  "Parser.gs",
  "Format.gs",
  "Sheet.gs",
  "HealthApi.gs",
  "Main.gs",
];
// Harness.gs first: it declares the fakes the orchestration suite runs against.
// Unlike this file it is pushed to Apps Script, so the fakes have one
// definition and `Sync ▸ Run tests` exercises the same ones.
const testFiles = ["Harness.gs", "Parser.test.gs", "Sync.test.gs"];

// Silence the code's diagnostic chatter (warn/info) so the suite output is just
// the PASS/FAIL lines, which are emitted via console.log and captured below.
const quietConsole = Object.assign({}, console, {
  warn: () => {},
  info: () => {},
});

// Minimal Utilities.formatDate stub backed by Node's Intl. Covers the format
// strings used in src/*.gs; throws on anything else so a new usage shows up
// loudly instead of silently returning wrong data.
function formatDateStub(date, tz, format) {
  const zone = tz === "GMT" || tz === "UTC" ? "UTC" : tz;
  if (format === "Z") {
    if (zone === "UTC") {
      return "+0000";
    }
    const parts = new Intl.DateTimeFormat("en-US", {
      timeZone: zone,
      timeZoneName: "longOffset",
      year: "numeric",
    }).formatToParts(date);
    const tzn = parts.find((p) => p.type === "timeZoneName").value;
    if (tzn === "GMT" || tzn === "UTC") {
      return "+0000";
    }
    const m = tzn.match(/GMT([+-])(\d{2}):(\d{2})/);
    if (!m) {
      throw new Error(`formatDateStub: unparseable offset "${tzn}"`);
    }
    return m[1] + m[2] + m[3];
  }
  if (format === "yyyy MM dd HH mm ss" || format === "yyyy-MM-dd HH:mm:ss") {
    const parts = new Intl.DateTimeFormat("en-US", {
      timeZone: zone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: false,
    }).formatToParts(date);
    const get = (t) => parts.find((p) => p.type === t).value;
    let h = get("hour");
    if (h === "24") {
      h = "00";
    }
    const ymdParts = [get("year"), get("month"), get("day")];
    const hmsParts = [h, get("minute"), get("second")];
    return format === "yyyy MM dd HH mm ss"
      ? ymdParts.concat(hmsParts).join(" ")
      : `${ymdParts.join("-")} ${hmsParts.join(":")}`;
  }
  if (format === "yyyy-MM-dd'T'HH:mm:ss'Z'") {
    if (zone !== "UTC") {
      throw new Error("formatDateStub: ISO Z format only stubbed for GMT/UTC");
    }
    return `${date.toISOString().slice(0, 19)}Z`;
  }
  if (format === "yyyy-MM-dd") {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: zone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).formatToParts(date);
    const get = (t) => parts.find((p) => p.type === t).value;
    return `${get("year")}-${get("month")}-${get("day")}`;
  }
  throw new Error(`formatDateStub: unmocked format "${format}"`);
}

// Matches appsscript.json, so a Date cell's civil date resolves the same way
// locally and in Apps Script. The suites assert on that resolution, so the two
// have to agree.
const scriptTimeZone = "America/Toronto";
const Utilities = {
  formatDate: formatDateStub,
  sleep: () => {},
};
const Session = {
  getScriptTimeZone: () => scriptTimeZone,
};

// Apps Script services the suites do not fake for themselves. SpreadsheetApp,
// PropertiesService and LockService are not here: Harness.gs declares those
// fakes and withSyncTestHarness_ installs them for the duration of the
// orchestration suite, so this sandbox starts without them and a source file
// reaching for one outside that window fails loudly.
const sandbox = {
  console: quietConsole,
  Utilities,
  Session,
};
vm.createContext(sandbox);

for (const f of srcFiles) {
  const code = fs.readFileSync(path.join(root, "src", f), "utf8");
  vm.runInContext(code, sandbox, { filename: `src/${f}` });
}
for (const f of testFiles) {
  const code = fs.readFileSync(path.join(root, "test", f), "utf8");
  vm.runInContext(code, sandbox, { filename: `test/${f}` });
}

const logs = [];
const origLog = quietConsole.log;
quietConsole.log = (...args) => {
  logs.push(args.join(" "));
  // .apply rather than a spread call: console.log's declared signature is not
  // a bare rest parameter, so spreading an unknown-length array into it is a
  // type error (TS2556) even though the call is valid at runtime.
  origLog.apply(quietConsole, args);
};
try {
  vm.runInContext("runParserTests();", sandbox);
  vm.runInContext("runSyncTests();", sandbox);
} finally {
  quietConsole.log = origLog;
}

const output = logs.join("\n");
const passed = (output.match(/^PASS /gm) || []).length;
const failed = (output.match(/^FAIL /gm) || []).length;
console.log(`\n${passed + failed} tests: ${passed} passed, ${failed} failed`);
if (failed > 0) {
  process.exit(1);
}

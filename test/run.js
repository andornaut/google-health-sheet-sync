const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const srcFiles = ['Config.gs', 'Parser.gs', 'Format.gs', 'Sheet.gs', 'HealthApi.gs', 'Main.gs', 'Debug.gs'];
const testFiles = ['Parser.test.gs'];

const quietConsole = Object.assign({}, console, { warn: () => {} });
const sandbox = {
  console: quietConsole,
  SpreadsheetApp: { getUi: () => { throw new Error('no UI'); } }
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

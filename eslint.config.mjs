import js from "@eslint/js";
import globals from "globals";

export default [
  {
    ignores: ["node_modules/**"],
  },
  js.configs.recommended,
  // The Apps Script sources and the suites that run against them. `.gs` is
  // JavaScript under another extension, so it has to be named: eslint reads
  // `.js` and nothing else unless told.
  //
  // Apps Script loads every file into one shared script scope, which eslint has
  // no model for: it sees each file alone, so a function defined in Sheet.gs and
  // called from Main.gs reads as undefined here, and one called only from
  // another file reads as unused. Both checks are therefore scoped to what a
  // single file can settle: `no-undef` off, and `no-unused-vars` on locals and
  // parameters rather than on the 130 top-level declarations. Undefined symbols
  // across the project are jsconfig.json's to catch, where `checkJs` and the
  // google-apps-script types see every file at once.
  //
  // A caught error goes unreported because swallowing it is the point: the UI
  // calls fail outside a spreadsheet and the suites carry on. A stub argument
  // prefixed with `_` is declared to match the arity of the function it stands
  // in for, which is what makes it a stub.
  {
    files: ["src/**/*.gs", "test/**/*.gs"],
    languageOptions: {
      ecmaVersion: 2019,
      globals: {
        ...globals.browser,
      },
      sourceType: "script",
    },
    rules: {
      "no-undef": "off",
      "no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", caughtErrors: "none", vars: "local" },
      ],
    },
  },
  // The local test runner, which Apps Script never loads: `.claspignore` keeps
  // both out of `clasp push` because they use CommonJS `require`.
  {
    files: ["test/run.js", "test/mutate.js"],
    languageOptions: {
      ecmaVersion: "latest",
      globals: {
        ...globals.node,
      },
      sourceType: "commonjs",
    },
    rules: {
      "no-unused-vars": ["error", { caughtErrors: "none" }],
    },
  },
];

import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

import js from "@eslint/js";
import globals from "globals";

import { plugins, sourceRules, toolingRules } from "./eslint.config.base.mjs";

// Apps Script loads every .gs file into one shared script scope, which eslint
// models per file: a function defined in Sheet.gs and called from Main.gs reads
// as undefined, and one called only from another file reads as unused. Nothing
// else covers this - TypeScript does not recognise the .gs extension, so
// jsconfig.json's checkJs never sees these files at all.
//
// So reconstruct the scope here rather than turning no-undef off. Every
// top-level declaration becomes a global, read from the sources at lint time so
// there is no list to keep up to date, and a name that is not declared anywhere
// is still reported. A `let` or `var` is writable because the module-level
// caches are assigned on first use; a function or const is not.
const sharedScope = () => {
  const names = {};
  for (const dir of ["src", "test"]) {
    for (const file of readdirSync(dir).filter((f) => f.endsWith(".gs"))) {
      const source = readFileSync(join(dir, file), "utf8");
      for (const [, keyword, name] of source.matchAll(
        /^(function|const|let|var)\s+([A-Za-z_$][\w$]*)/gm,
      )) {
        names[name] =
          keyword === "let" || keyword === "var" ? "writable" : "readonly";
      }
    }
  }
  return names;
};

// The services Apps Script puts in scope. Listed rather than taken wholesale
// from a plugin, so reaching for a service this project has never used says so.
// The three the test harness swaps out are still real globals here: it assigns
// them through globalThis, which is a property write, not a redeclaration.
const appsScriptGlobals = {
  CacheService: "readonly",
  globalThis: "readonly",
  HtmlService: "readonly",
  LockService: "readonly",
  OAuth2: "readonly",
  PropertiesService: "readonly",
  ScriptApp: "readonly",
  Session: "readonly",
  SpreadsheetApp: "readonly",
  UrlFetchApp: "readonly",
  Utilities: "readonly",
};

export default [
  {
    ignores: ["node_modules/**"],
  },
  js.configs.recommended,
  // The Apps Script sources and the suites that run against them. `.gs` is
  // JavaScript under another extension, so it has to be named: eslint reads
  // `.js` and nothing else unless told.
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
        ...appsScriptGlobals,
        ...sharedScope(),
      },
      sourceType: "script",
    },
    plugins,
    rules: {
      ...sourceRules,
      // The shared scope is declared above, so every file also declares what it
      // contributes to it. That is the arrangement, not a redeclaration.
      "no-redeclare": ["error", { builtinGlobals: false }],
      // Reported per file, so a top-level declaration used only from another
      // one reads as unused. Locals and parameters are what a single file can
      // settle, and they are where an actual mistake shows up.
      "no-unused-vars": ["error", { argsIgnorePattern: "^_", vars: "local" }],
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
    plugins,
    rules: {
      ...toolingRules,
    },
  },
  // test/mutate.js only. Its mutation specs hold excerpts matched against the
  // sources verbatim, so wrapping one stops it matching and the mutation it
  // describes silently stops being tested. Scoped to this file rather than the
  // repository, and to strings rather than to max-len, so a long line of code
  // here is still reported.
  //
  // no-template-curly-in-string is off for the same reason: the sources use
  // template literals, so an excerpt quoting one holds a literal `${...}` that
  // must stay unevaluated to match.
  {
    files: ["test/mutate.js"],
    rules: {
      "max-len": ["error", { code: 120, ignoreStrings: true }],
      "no-template-curly-in-string": "off",
    },
  },
];

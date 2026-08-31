function runParserTests() {
  reportTestResults_("Parser / pure-helper tests", runParserTestsBody_());
}

// Returns the PASS/FAIL list rather than reporting it, so runAllTests can
// summarize both suites without two dumps and two blocking alerts.
function runParserTestsBody_() {
  const results = [];
  const t = (name, fn) => {
    try {
      fn();
      results.push(`PASS ${name}`);
    } catch (err) {
      results.push(`FAIL ${name}: ${err}`);
    }
  };
  // JSON.stringify serializes NaN and Infinity as null, so a plain stringify
  // comparison cannot tell a rejected-input null from a leaked NaN. The
  // replacer tags non-finite numbers so "-> null" assertions on the numeric
  // parsers actually pin the guard that produces the null.
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
  const throws = (fn, re, msg) => {
    let err = null;
    try {
      fn();
    } catch (e) {
      err = e;
    }
    if (!err) {
      throw new Error(`${msg || "expected a throw"}, but none was thrown`);
    }
    if (re && !re.test(String(err))) {
      throw new Error(`${msg || "wrong error"}: ${err}`);
    }
  };
  // Swap globals (the Health API entry points) for the duration of fn, then
  // restore; returns fn's value. Apps Script declares top-level functions on the
  // global scope, but the VM sandbox these tests run in treats the function name
  // like a const binding at the outer scope, so direct reassignment throws.
  // Stash and restore via globalThis instead.
  const withGlobals = (stubs, fn) => {
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

  t("empty cell -> []", () => eq(parseExerciseCell(""), []));
  t("null cell -> []", () => eq(parseExerciseCell(null), []));
  t('single weight "135" (reps/sets unknown)', () =>
    eq(parseExerciseCell("135"), [
      { assisted: false, reps: null, sets: null, weight: 135 },
    ]),
  );
  t('weight x reps "135x5" (sets unknown)', () =>
    eq(parseExerciseCell("135x5"), [
      { assisted: false, reps: 5, sets: null, weight: 135 },
    ]),
  );
  t('weight x reps x sets "135x5x3"', () =>
    eq(parseExerciseCell("135x5x3"), [
      { assisted: false, reps: 5, sets: 3, weight: 135 },
    ]),
  );
  t('assisted prefix "*135x5x3"', () =>
    eq(parseExerciseCell("*135x5x3"), [
      { assisted: true, reps: 5, sets: 3, weight: 135 },
    ]),
  );
  t("mixed comma + newline", () =>
    eq(parseExerciseCell("135x5x3, 145x3x2\n*155x1"), [
      { assisted: false, reps: 5, sets: 3, weight: 135 },
      { assisted: false, reps: 3, sets: 2, weight: 145 },
      { assisted: true, reps: 1, sets: null, weight: 155 },
    ]),
  );
  t('whitespace tolerance "  135 x 5 x 3  "', () =>
    eq(parseExerciseCell("  135 x 5 x 3  "), [
      { assisted: false, reps: 5, sets: 3, weight: 135 },
    ]),
  );
  t('uppercase X "135X5X3"', () =>
    eq(parseExerciseCell("135X5X3"), [
      { assisted: false, reps: 5, sets: 3, weight: 135 },
    ]),
  );
  t("junk line skipped", () =>
    eq(parseExerciseCell("garbage\n135x5"), [
      { assisted: false, reps: 5, sets: null, weight: 135 },
    ]),
  );
  // Zero reps means not performed and is dropped; zero sets is a start-only
  // marker and is retained (the notes layer suppresses it).
  t("zero-reps dropped, zero-sets retained, valid kept", () =>
    eq(parseExerciseCell("135x5x3, 90x6x0, 95x0"), [
      { assisted: false, reps: 5, sets: 3, weight: 135 },
      { assisted: false, reps: 6, sets: 0, weight: 90 },
    ]),
  );
  // A multi-segment entry is what separates "any value is invalid" from "every
  // value is invalid": with a single number the two agree, so these are the only
  // cases that pin the rejection to ANY negative rather than all of them.
  t('negative weight rejected in a multi-segment entry "-135x5"', () =>
    eq(parseExerciseCell("-135x5"), []),
  );
  t('negative reps rejected "135x-2"', () =>
    eq(parseExerciseCell("135x-2"), []),
  );
  t('negative sets rejected "135x5x-3"', () =>
    eq(parseExerciseCell("135x5x-3"), []),
  );
  t('decimal weight allowed "22.5" (reps/sets unknown)', () =>
    eq(parseExerciseCell("22.5"), [
      { assisted: false, reps: null, sets: null, weight: 22.5 },
    ]),
  );
  t('decimal reps rejected "135x5.5"', () =>
    eq(parseExerciseCell("135x5.5"), []),
  );
  t('decimal sets rejected "135x5x2.5"', () =>
    eq(parseExerciseCell("135x5x2.5"), []),
  );
  t('too many x segments rejected "135x5x3x2"', () =>
    eq(parseExerciseCell("135x5x3x2"), []),
  );
  // The empty part a trailing separator leaves behind must be skipped: parsed,
  // it is Number("") = 0, a zero-weight entry that reaches the notes.
  t('trailing separator "135x5x3," yields no zero-weight entry', () =>
    eq(parseExerciseCell("135x5x3,"), [
      { assisted: false, reps: 5, sets: 3, weight: 135 },
    ]),
  );
  t('semicolon-separated cell "95x5x2;85x5x5"', () =>
    eq(parseExerciseCell("95x5x2;85x5x5"), [
      { assisted: false, reps: 5, sets: 2, weight: 95 },
      { assisted: false, reps: 5, sets: 5, weight: 85 },
    ]),
  );

  t('bodyweight parse "185"', () => eq(parseBodyweight("185"), 185));
  t('bodyweight parse "185.4"', () => eq(parseBodyweight("185.4"), 185.4));
  t("bodyweight empty -> null", () => eq(parseBodyweight(""), null));
  t("bodyweight junk -> null", () => eq(parseBodyweight("heavy"), null));
  t("bodyweight zero -> null", () => eq(parseBodyweight("0"), null));
  t('bodyweight at cap "499" -> 499', () => eq(parseBodyweight("499"), 499));
  t('bodyweight above cap "500" -> null', () =>
    eq(parseBodyweight("500"), null),
  );
  t('bodyweight at floor "50" -> 50', () => eq(parseBodyweight("50"), 50));
  t('bodyweight below floor "49" -> null', () =>
    eq(parseBodyweight("49"), null),
  );

  t("formatEntryNote_ single set", () =>
    eq(
      formatEntryNote_("Bench press", {
        assisted: false,
        reps: 5,
        sets: 1,
        weight: 135,
      }),
      "Bench press, 135 lbs, 1 set of 5",
    ),
  );
  t("formatEntryNote_ assisted suffix", () =>
    eq(
      formatEntryNote_("Pull up", {
        assisted: true,
        reps: 5,
        sets: 3,
        weight: 25,
      }),
      "Pull up, 25 lbs, 3 sets of 5 (assisted)",
    ),
  );
  t("formatEntryNote_ decimal weight", () =>
    eq(
      formatEntryNote_("Lateral raise", {
        assisted: false,
        reps: 10,
        sets: 3,
        weight: 22.5,
      }),
      "Lateral raise, 22.5 lbs, 3 sets of 10",
    ),
  );
  t("formatEntryNote_ weight + reps (sets unknown)", () =>
    eq(
      formatEntryNote_("Bench press", {
        assisted: false,
        reps: 5,
        sets: null,
        weight: 135,
      }),
      "Bench press, 135 lbs, 5 reps",
    ),
  );
  t("formatEntryNote_ weight + 1 rep (sets unknown)", () =>
    eq(
      formatEntryNote_("Bench press", {
        assisted: false,
        reps: 1,
        sets: null,
        weight: 225,
      }),
      "Bench press, 225 lbs, 1 rep",
    ),
  );
  t("formatEntryNote_ weight only + assisted", () =>
    eq(
      formatEntryNote_("Pull up", {
        assisted: true,
        reps: null,
        sets: null,
        weight: 25,
      }),
      "Pull up, 25 lbs (assisted)",
    ),
  );

  // Commas are the field delimiter WITHIN an entry, so the period that ends
  // each line is what keeps entry boundaries readable, and the duration has to
  // stand on its own line: glued on as ", 45 minute session" it reads as one
  // more attribute of the last exercise.
  t(
    "buildNotes one period-terminated line per entry, duration on its own line",
    () => {
      const notes = buildNotes(45 * 60 * 1000, [
        {
          entries: [
            { assisted: false, reps: 5, sets: 3, weight: 135 },
            { assisted: false, reps: 3, sets: 2, weight: 145 },
          ],
          name: "Bench press",
        },
        {
          entries: [{ assisted: false, reps: 5, sets: 3, weight: 225 }],
          name: "Squat",
        },
      ]);
      eq(
        notes,
        "Bench press, 135 lbs, 3 sets of 5.\n" +
          "Bench press, 145 lbs, 2 sets of 3.\n" +
          "Squat, 225 lbs, 3 sets of 5.\n" +
          "45 minute session.",
      );
    },
  );

  t("buildNotes suppresses zero-set entries (start-only markers)", () => {
    const notes = buildNotes(10 * 60 * 1000, [
      {
        entries: [
          { assisted: false, reps: 5, sets: 0, weight: 200 },
          { assisted: false, reps: 5, sets: 2, weight: 200 },
        ],
        name: "Bench press",
      },
    ]);
    eq(notes, "Bench press, 200 lbs, 2 sets of 5.\n10 minute session.");
  });

  t(
    "buildNotes with only zero-set entries -> just the session sentence",
    () => {
      const notes = buildNotes(10 * 60 * 1000, [
        {
          entries: [{ assisted: false, reps: 5, sets: 0, weight: 200 }],
          name: "Bench press",
        },
      ]);
      eq(notes, "10 minute session.");
    },
  );

  const oneSet = [
    {
      entries: [{ assisted: false, reps: 5, sets: 5, weight: 190 }],
      name: "Bench press",
    },
  ];
  // 90s is 1.5 minutes: rounded it reads "2 minute session", truncated "1".
  t("buildNotes rounds the session duration to the nearest minute", () =>
    eq(
      buildNotes(90 * 1000, oneSet),
      "Bench press, 190 lbs, 5 sets of 5.\n2 minute session.",
    ),
  );
  t("buildNotes keeps the session line at one minute", () =>
    eq(buildNotes(60000, []), "1 minute session."),
  );
  t("buildNotes omits the session line when the duration rounds to zero", () =>
    eq(buildNotes(29 * 1000, oneSet), "Bench press, 190 lbs, 5 sets of 5."),
  );

  t("parseHealthIds_ empty/null", () => {
    eq(parseHealthIds_(""), []);
    eq(parseHealthIds_(null), []);
    eq(parseHealthIds_(undefined), []);
    eq(parseHealthIds_("   "), []);
  });
  t("parseHealthIds_ valid JSON array of strings", () =>
    eq(
      parseHealthIds_(
        '["users/me/dataTypes/exercise/dataPoints/abc","users/me/dataTypes/weight/dataPoints/def"]',
      ),
      [
        "users/me/dataTypes/exercise/dataPoints/abc",
        "users/me/dataTypes/weight/dataPoints/def",
      ],
    ),
  );
  t("parseHealthIds_ filters non-string elements", () =>
    eq(parseHealthIds_('["a", 1, null, "b", {}]'), ["a", "b"]),
  );
  t("parseHealthIds_ non-array JSON -> []", () =>
    eq(parseHealthIds_('{"x":1}'), []),
  );
  t("parseHealthIds_ malformed JSON -> []", () =>
    eq(parseHealthIds_("not json"), []),
  );

  t("parseExerciseEditTimes_ empty/blank -> {}", () => {
    eq(parseExerciseEditTimes_(""), {});
    eq(parseExerciseEditTimes_(null), {});
    eq(parseExerciseEditTimes_(undefined), {});
    eq(parseExerciseEditTimes_("   "), {});
  });
  t("parseExerciseEditTimes_ valid map round-trips", () =>
    eq(
      parseExerciseEditTimes_(
        '{"Bench press":{"first":"2026-08-30T17:09:18.000Z","last":"2026-08-30T17:20:00.000Z"}}',
      ),
      {
        "Bench press": {
          first: "2026-08-30T17:09:18.000Z",
          last: "2026-08-30T17:20:00.000Z",
        },
      },
    ),
  );
  t("parseExerciseEditTimes_ drops entries with no usable timestamp", () =>
    eq(
      parseExerciseEditTimes_(
        '{"A":{"first":"x"},"B":{},"C":null,"D":5,"E":{"first":1,"last":2}}',
      ),
      { A: { first: "x" } },
    ),
  );
  // A cell holding an array or a scalar is not the shape the merge writes, and
  // treating it as one would let mergeExerciseEditTimes_ carry junk keys
  // forward into every later edit of the row.
  t("parseExerciseEditTimes_ non-object JSON -> {}", () => {
    eq(parseExerciseEditTimes_("[1,2]"), {});
    eq(parseExerciseEditTimes_('"str"'), {});
    eq(parseExerciseEditTimes_("7"), {});
    eq(parseExerciseEditTimes_('[{"first":"T1"}]'), {});
  });
  t("parseExerciseEditTimes_ malformed JSON -> {}", () =>
    eq(parseExerciseEditTimes_("not json"), {}),
  );

  t("mergeExerciseEditTimes_ seeds first and last on a new exercise", () =>
    eq(mergeExerciseEditTimes_({}, ["Bench press"], "T1"), {
      "Bench press": { first: "T1", last: "T1" },
    }),
  );
  // first is sticky: this is what keeps an exercise attributed to the workout
  // it was originally logged in when it is corrected during a later one.
  t("mergeExerciseEditTimes_ keeps first, advances last", () =>
    eq(
      mergeExerciseEditTimes_(
        { "Bench press": { first: "T1", last: "T1" } },
        ["Bench press"],
        "T2",
      ),
      { "Bench press": { first: "T1", last: "T2" } },
    ),
  );
  t("mergeExerciseEditTimes_ leaves untouched exercises alone", () =>
    eq(
      mergeExerciseEditTimes_(
        {
          "Bench press": { first: "T1", last: "T1" },
          Squat: { first: "T1", last: "T1" },
        },
        ["Squat"],
        "T2",
      ),
      {
        "Bench press": { first: "T1", last: "T1" },
        Squat: { first: "T1", last: "T2" },
      },
    ),
  );
  t("mergeExerciseEditTimes_ does not mutate its input", () => {
    const prior = { "Bench press": { first: "T1", last: "T1" } };
    mergeExerciseEditTimes_(prior, ["Bench press", "Squat"], "T2");
    eq(prior, { "Bench press": { first: "T1", last: "T1" } });
  });
  t("mergeExerciseEditTimes_ tolerates a null prior", () =>
    eq(mergeExerciseEditTimes_(null, ["Squat"], "T1"), {
      Squat: { first: "T1", last: "T1" },
    }),
  );

  t("extractDataPointName_ null/undefined -> null", () => {
    eq(extractDataPointName_(null), null);
    eq(extractDataPointName_(undefined), null);
    eq(extractDataPointName_({}), null);
  });
  t("extractDataPointName_ LRO response shape", () =>
    eq(
      extractDataPointName_({
        done: true,
        response: { name: "users/me/dataTypes/exercise/dataPoints/xyz" },
      }),
      "users/me/dataTypes/exercise/dataPoints/xyz",
    ),
  );
  t("extractDataPointName_ direct name with dataPoints segment", () =>
    eq(
      extractDataPointName_({
        name: "users/me/dataTypes/weight/dataPoints/abc",
      }),
      "users/me/dataTypes/weight/dataPoints/abc",
    ),
  );
  t("extractDataPointName_ ignores non-datapoint name", () =>
    eq(extractDataPointName_({ name: "operations/123" }), null),
  );

  t("parseOffsetSeconds_ empty -> 0", () => {
    eq(parseOffsetSeconds_(""), 0);
    eq(parseOffsetSeconds_(null), 0);
    eq(parseOffsetSeconds_(undefined), 0);
  });
  t('parseOffsetSeconds_ "3600s" -> 3600', () =>
    eq(parseOffsetSeconds_("3600s"), 3600),
  );
  t('parseOffsetSeconds_ "-18000s" -> -18000', () =>
    eq(parseOffsetSeconds_("-18000s"), -18000),
  );
  t("parseOffsetSeconds_ bare number string", () =>
    eq(parseOffsetSeconds_("7200"), 7200),
  );
  t("parseOffsetSeconds_ garbage -> 0", () =>
    eq(parseOffsetSeconds_("abc"), 0),
  );

  t("formatSyncResult_ null -> lock-held message", () =>
    eq(
      formatSyncResult_(null, "Synced"),
      "Sync skipped (another run holds the lock). Try again shortly.",
    ),
  );
  t("formatSyncResult_ ok + deferred", () =>
    eq(
      formatSyncResult_({ deferred: 25, errors: 0, ok: 75 }, "Synced"),
      "Synced 75 row(s), 25 deferred.",
    ),
  );
  t("formatSyncResult_ ok + errors + deferred", () =>
    eq(
      formatSyncResult_({ deferred: 25, errors: 5, ok: 70 }, "Synced"),
      "Synced 70 row(s), 5 error(s), 25 deferred.\n\nSee Executions for details.",
    ),
  );

  t("toDate_ passes Date through", () => {
    const d = new Date("2026-01-15T12:00:00Z");
    if (toDate_(d) !== d) {
      throw new Error("expected same Date instance");
    }
  });
  // Strict identity, not eq: JSON.stringify serializes an Invalid Date as
  // null, so eq(toDate_("x"), null) holds whether or not the guard exists.
  t("toDate_ invalid -> null", () => {
    if (toDate_("not a date") !== null || toDate_("") !== null) {
      throw new Error("expected null for unparseable input");
    }
  });

  t("splitHealthIdsByType_ empty/null", () => {
    eq(splitHealthIdsByType_(null), { exercise: [], other: [], weight: [] });
    eq(splitHealthIdsByType_([]), { exercise: [], other: [], weight: [] });
  });
  t("splitHealthIdsByType_ buckets weight vs exercise", () =>
    eq(
      splitHealthIdsByType_([
        "users/me/dataTypes/weight/dataPoints/w1",
        "users/me/dataTypes/exercise/dataPoints/e1",
        "users/me/dataTypes/weight/dataPoints/w2",
      ]),
      {
        exercise: ["users/me/dataTypes/exercise/dataPoints/e1"],
        other: [],
        weight: [
          "users/me/dataTypes/weight/dataPoints/w1",
          "users/me/dataTypes/weight/dataPoints/w2",
        ],
      },
    ),
  );
  t("splitHealthIdsByType_ unknown type -> other", () =>
    eq(splitHealthIdsByType_(["users/me/dataTypes/sleep/dataPoints/s1"]), {
      exercise: [],
      other: ["users/me/dataTypes/sleep/dataPoints/s1"],
      weight: [],
    }),
  );
  t("splitHealthIdsByType_ malformed name -> other", () =>
    eq(
      splitHealthIdsByType_([
        "not-a-resource-name",
        "users/me/dataTypes/weight/dataPoints/ok",
      ]),
      {
        exercise: [],
        other: ["not-a-resource-name"],
        weight: ["users/me/dataTypes/weight/dataPoints/ok"],
      },
    ),
  );

  t("toMeName_ leaves an already-me name unchanged", () =>
    eq(
      toMeName_("users/me/dataTypes/exercise/dataPoints/xyz"),
      "users/me/dataTypes/exercise/dataPoints/xyz",
    ),
  );

  t("isNotFoundError_ true for 404", () =>
    eq(isNotFoundError_({ statusCode: 404 }), true),
  );
  t("isNotFoundError_ false for 500", () =>
    eq(isNotFoundError_({ statusCode: 500 }), false),
  );
  t("isNotFoundError_ false for null", () => eq(isNotFoundError_(null), false));

  // The cap itself is pinned through resolveForeignMatches_ ("window is
  // clamped"); nothing else observes that a sub-max span passes through.
  t("capExerciseDurationToMax_ leaves sub-max durations untouched", () =>
    eq(capExerciseDurationToMax_(30 * 60 * 1000), 30 * 60 * 1000),
  );

  t("humanizeMs_ seconds rounds", () => eq(humanizeMs_(1500), "2s"));
  t("humanizeMs_ exact minute", () => eq(humanizeMs_(60 * 1000), "1m"));
  t("humanizeMs_ minutes + seconds", () =>
    eq(humanizeMs_(90 * 1000), "1m 30s"),
  );
  t("humanizeMs_ negative clamps to 0ms", () => eq(humanizeMs_(-5), "0ms"));

  t("buildSampleTimeFromUtc_ formats physical + civil time", () =>
    eq(buildSampleTimeFromUtc_(Date.UTC(2026, 0, 15, 17, 0, 0), -18000), {
      civilTime: {
        date: { day: 15, month: 1, year: 2026 },
        time: { hours: 12, minutes: 0, seconds: 0 },
      },
      physicalTime: "2026-01-15T17:00:00Z",
      utcOffset: "-18000s",
    }),
  );

  // DST behavior for America/Toronto in 2026:
  //   Spring forward: Mar 8, 02:00 EST -> 03:00 EDT  (clocks jump forward)
  //   Fall back:      Nov 1, 02:00 EDT -> 01:00 EST  (clocks jump back)
  const TORONTO = "America/Toronto";
  const EST = -5 * 3600;
  const EDT = -4 * 3600;

  t("localCivilToUtcMs_ winter EST (Jan)", () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 1, 15, 12, 0);
    eq(r.offsetSeconds, EST);
    eq(r.utcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
  });
  t("localCivilToUtcMs_ summer EDT (Jul)", () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 7, 15, 12, 0);
    eq(r.offsetSeconds, EDT);
    eq(r.utcMs, Date.UTC(2026, 6, 15, 16, 0, 0));
  });
  // Only a civil time whose first-pass offset differs from the offset at the
  // resolved instant takes the correction path, so the two post-cutover cases
  // are the DST coverage; a pre-cutover or next-day time resolves exactly like
  // any other winter or summer time.
  t(
    "localCivilToUtcMs_ spring-forward day post-cutover (Mar 8 03:00 = EDT)",
    () => {
      const r = localCivilToUtcMs_(TORONTO, 2026, 3, 8, 3, 0);
      eq(r.offsetSeconds, EDT);
      eq(r.utcMs, Date.UTC(2026, 2, 8, 7, 0, 0));
    },
  );
  t("localCivilToUtcMs_ fall-back day post-cutover (Nov 1 03:00 = EST)", () => {
    const r = localCivilToUtcMs_(TORONTO, 2026, 11, 1, 3, 0);
    eq(r.offsetSeconds, EST);
    eq(r.utcMs, Date.UTC(2026, 10, 1, 8, 0, 0));
  });

  t("getTzOffsetSeconds_ winter EST", () =>
    eq(
      getTzOffsetSeconds_(TORONTO, new Date(Date.UTC(2026, 0, 15, 17, 0, 0))),
      EST,
    ),
  );
  t("getTzOffsetSeconds_ summer EDT", () =>
    eq(
      getTzOffsetSeconds_(TORONTO, new Date(Date.UTC(2026, 6, 15, 16, 0, 0))),
      EDT,
    ),
  );
  // A half-hour zone is the only thing that exercises the minutes term; EST,
  // EDT and GMT are all whole hours, so the mins arithmetic could be dropped
  // entirely and every other offset test would still pass.
  t("getTzOffsetSeconds_ half-hour zone keeps the minutes", () =>
    eq(
      getTzOffsetSeconds_("Asia/Kolkata", new Date(Date.UTC(2026, 0, 15))),
      5 * 3600 + 30 * 60,
    ),
  );
  t("getTzOffsetSeconds_ negative half-hour zone", () =>
    eq(
      getTzOffsetSeconds_("America/St_Johns", new Date(Date.UTC(2026, 0, 15))),
      -(3 * 3600 + 30 * 60),
    ),
  );

  t("buildIntervalFromUtc_ formats interval", () =>
    eq(
      buildIntervalFromUtc_(
        Date.UTC(2026, 0, 15, 17, 0, 0),
        EST,
        Date.UTC(2026, 0, 15, 18, 0, 0),
        EST,
      ),
      {
        endTime: "2026-01-15T18:00:00Z",
        endUtcOffset: "-18000s",
        startTime: "2026-01-15T17:00:00Z",
        startUtcOffset: "-18000s",
      },
    ),
  );

  // syntheticExerciseInterval_ and resolveRowTiming_ use getTz_() which is
  // cached on first call and resolves to the test runner's default
  // (America/Toronto). Test dates use noon UTC so civilDateParts_ returns
  // the intended calendar day in EST (UTC-5).
  const JAN_15_NOON_UTC = new Date(Date.UTC(2026, 0, 15, 12, 0, 0));

  t("syntheticExerciseInterval_ ordinal 0 -> noon-1pm EST", () => {
    const r = syntheticExerciseInterval_(JAN_15_NOON_UTC, 0);
    eq(r.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
    eq(r.endUtcMs, Date.UTC(2026, 0, 15, 18, 0, 0));
    eq(r.startOffsetSeconds, EST);
    eq(r.endOffsetSeconds, EST);
  });
  t("syntheticExerciseInterval_ ordinal 1 -> 1pm-2pm EST", () => {
    const r = syntheticExerciseInterval_(JAN_15_NOON_UTC, 1);
    eq(r.startUtcMs, Date.UTC(2026, 0, 15, 18, 0, 0));
    eq(r.endUtcMs, Date.UTC(2026, 0, 15, 19, 0, 0));
  });
  t(
    "syntheticExerciseInterval_ clamps to final slot when end would spill past midnight",
    () => {
      // ordinal 12 would yield startHour=24/endHour=25; instead of throwing, it
      // clamps into the last 1h slot of the day (23:00-24:00 local = 4-5am UTC
      // next day in EST).
      const r = syntheticExerciseInterval_(JAN_15_NOON_UTC, 12);
      eq(r.startUtcMs, Date.UTC(2026, 0, 16, 4, 0, 0));
      eq(r.endUtcMs, Date.UTC(2026, 0, 16, 5, 0, 0));
      eq(r.startOffsetSeconds, EST);
      eq(r.endOffsetSeconds, EST);
    },
  );

  // Fixture times are 9-10am EST, deliberately NOT noon: synthetic timing for
  // ordinal 0 is noon-1pm, so a fixture there would assert values the synthetic
  // fallback also produces and the interval checks would hold either way.
  t("resolveRowTiming_ edit source preserves interval within bounds", () => {
    const first = new Date(Date.UTC(2026, 0, 15, 14, 0, 0));
    const last = new Date(Date.UTC(2026, 0, 15, 15, 0, 0));
    const r = resolveRowTiming_(
      {
        date: JAN_15_NOON_UTC,
        exerciseFirstEditedAt: first,
        exercisesLastEditedAt: last,
        weightEditedAt: first,
      },
      0,
      null,
    );
    eq(r.exerciseSource, "edit");
    eq(r.weightSource, "edit");
    eq(r.exercise.startUtcMs, first.getTime());
    eq(r.exercise.endUtcMs, last.getTime());
    eq(r.exercise.startOffsetSeconds, EST);
    eq(r.exercise.endOffsetSeconds, EST);
    eq(r.weight, { offsetSeconds: EST, utcMs: first.getTime() });
  });
  t(
    "resolveRowTiming_ edit source clamps too-short duration to MIN (10 min)",
    () => {
      const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
      const last = new Date(first.getTime() + 60 * 1000);
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: last,
        },
        0,
        null,
      );
      eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 10 * 60 * 1000);
    },
  );
  // A prior datapoint is required to reach the span test at all: with no prior,
  // exerciseEditIsUsable_ short-circuits on `|| !priorExercise` and 'edit' wins
  // whatever the span, so the MAX boundary would go unexercised.
  t(
    "resolveRowTiming_ edit source accepts a span right at MAX (120 min)",
    () => {
      const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
      const last = new Date(first.getTime() + MAX_EXERCISE_DURATION_MS);
      const prior = {
        exercise: {
          interval: {
            endTime: "2026-01-15T17:30:00Z",
            endUtcOffset: `${EST}s`,
            startTime: "2026-01-15T17:00:00Z",
            startUtcOffset: `${EST}s`,
          },
        },
      };
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: last,
        },
        0,
        prior,
      );
      eq(
        r.exerciseSource,
        "edit",
        "a span exactly at MAX still spans the workout",
      );
      eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 120 * 60 * 1000);
    },
  );
  // Past the cap the last edit is a later correction, not the end of the
  // workout. Using it would rebuild a short recorded session as a fabricated
  // 2 h one, so the row falls through to 'prior' (or 'synthetic' with no prior).
  t(
    "resolveRowTiming_ span past MAX falls through to prior, keeping the recorded interval",
    () => {
      const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
      const last = new Date(first.getTime() + 5 * 60 * 60 * 1000); // corrected hours later
      const prior = {
        exercise: {
          interval: {
            endTime: "2026-01-15T17:30:00Z",
            endUtcOffset: `${EST}s`,
            startTime: "2026-01-15T17:00:00Z",
            startUtcOffset: `${EST}s`,
          },
        },
      };
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: last,
        },
        0,
        prior,
      );
      eq(r.exerciseSource, "prior");
      eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
      eq(
        r.exercise.endUtcMs - r.exercise.startUtcMs,
        30 * 60 * 1000,
        "recorded 30 min kept, not stretched",
      );
    },
  );
  // The span test guards an interval we already recorded, so with no prior it
  // does not apply: the observed on-date start beats synthetic noon, and the
  // MAX clamp bounds the duration. This is the one path where that clamp works.
  // 9am EST, so "kept the observed start" is distinguishable from synthetic noon.
  t(
    "resolveRowTiming_ span past MAX with no prior still uses the observed start",
    () => {
      const first = new Date(Date.UTC(2026, 0, 15, 14, 0, 0));
      const last = new Date(first.getTime() + 5 * 60 * 60 * 1000);
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: last,
        },
        0,
        null,
      );
      eq(r.exerciseSource, "edit");
      eq(
        r.exercise.startUtcMs,
        first.getTime(),
        "keeps the 9am start rather than synthetic noon",
      );
      eq(
        r.exercise.endUtcMs - r.exercise.startUtcMs,
        MAX_EXERCISE_DURATION_MS,
        "clamped to the cap",
      );
    },
  );
  // Known consequence of protecting the recorded interval: timestamps cannot
  // tell "still logging this workout" from "correcting it later", so a sparsely
  // logged session stays at whatever the first sync recorded.
  t(
    "resolveRowTiming_ sparse logging past MAX keeps the start-only default",
    () => {
      const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
      const last = new Date(first.getTime() + 2.5 * 60 * 60 * 1000);
      const startOnly = {
        exercise: {
          interval: {
            endTime: "2026-01-15T17:10:00Z",
            endUtcOffset: `${EST}s`,
            startTime: "2026-01-15T17:00:00Z",
            startUtcOffset: `${EST}s`,
          },
        },
      };
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: last,
        },
        0,
        startOnly,
      );
      eq(r.exerciseSource, "prior");
      eq(r.exercise.endUtcMs - r.exercise.startUtcMs, MIN_EXERCISE_DURATION_MS);
    },
  );
  // A midnight-crossing workout keeps the 'edit' path: the span stays small
  // even though the last edit lands on the following civil date.
  t(
    "resolveRowTiming_ edit source survives a midnight-crossing workout",
    () => {
      const first = new Date(Date.UTC(2026, 0, 16, 4, 45, 0)); // 11:45pm EST Jan 15
      const last = new Date(Date.UTC(2026, 0, 16, 5, 15, 0)); // 12:15am EST Jan 16
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: last,
        },
        0,
        null,
      );
      eq(r.exerciseSource, "edit");
      eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 30 * 60 * 1000);
    },
  );
  t("resolveRowTiming_ synthetic source when no edit timestamps", () => {
    const r = resolveRowTiming_(
      {
        date: JAN_15_NOON_UTC,
        exerciseFirstEditedAt: null,
        exercisesLastEditedAt: null,
      },
      0,
      null,
    );
    eq(r.exerciseSource, "synthetic");
    eq(r.weightSource, "synthetic");
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
    eq(r.exercise.endUtcMs, Date.UTC(2026, 0, 15, 18, 0, 0));
  });
  t(
    "resolveRowTiming_ weight falls back to synthetic when weightEditedAt is missing",
    () => {
      // Exercise-only row (e.g. set bodyweight separately or not at all).
      // exerciseFirstEditedAt no longer feeds weight, so without weightEditedAt
      // the weight phase falls through to synthetic noon on row.date.
      const first = new Date(Date.UTC(2026, 0, 15, 20, 0, 0));
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: null,
          weightEditedAt: null,
        },
        0,
        null,
      );
      eq(r.weightSource, "synthetic");
      eq(r.weight.utcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
    },
  );
  t(
    "resolveRowTiming_ weight uses weightEditedAt on weight-only row with no exerciseFirstEditedAt",
    () => {
      const wEdit = new Date(Date.UTC(2026, 0, 15, 22, 0, 0));
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: null,
          exercisesLastEditedAt: null,
          weightEditedAt: wEdit,
        },
        0,
        null,
      );
      eq(r.exerciseSource, "synthetic");
      eq(r.weightSource, "edit");
      eq(r.weight, { offsetSeconds: EST, utcMs: wEdit.getTime() });
      eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0)); // exercise synthetic
    },
  );

  // Off-date edits: row.date is JAN-15 but the edit timestamps are on
  // JAN-20 (5 days later). The trust rule kicks in: edit-derived timing
  // is rejected because exerciseFirstEditedAt's civil date != row.date.
  const JAN_20_3PM_EST = new Date(Date.UTC(2026, 0, 20, 20, 0, 0));
  const JAN_20_4PM_EST = new Date(Date.UTC(2026, 0, 20, 21, 0, 0));

  t("resolveRowTiming_ off-date edit with no prior -> synthetic", () => {
    const r = resolveRowTiming_(
      {
        date: JAN_15_NOON_UTC,
        exerciseFirstEditedAt: JAN_20_3PM_EST,
        exercisesLastEditedAt: JAN_20_4PM_EST,
        weightEditedAt: JAN_20_3PM_EST,
      },
      0,
      null,
    );
    eq(r.exerciseSource, "synthetic");
    eq(r.weightSource, "synthetic");
    eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0)); // noon EST on row.date
    eq(r.weight.utcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
  });

  t(
    "resolveRowTiming_ off-date edit with prior exercise -> reuses prior interval",
    () => {
      const priorStart = Date.UTC(2026, 0, 15, 18, 30, 0);
      const priorEnd = Date.UTC(2026, 0, 15, 19, 15, 0);
      const priorExercise = {
        exercise: {
          interval: {
            endTime: "2026-01-15T19:15:00Z",
            endUtcOffset: "-18000s",
            startTime: "2026-01-15T18:30:00Z",
            startUtcOffset: "-18000s",
          },
        },
      };
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: JAN_20_3PM_EST,
          exercisesLastEditedAt: JAN_20_4PM_EST,
        },
        0,
        priorExercise,
      );
      eq(r.exerciseSource, "prior");
      eq(r.exercise.startUtcMs, priorStart);
      eq(r.exercise.endUtcMs, priorEnd);
      eq(r.exercise.startOffsetSeconds, EST);
      eq(r.exercise.endOffsetSeconds, EST);
    },
  );

  t(
    "resolveRowTiming_ same-date edit beats prior exercise (live-workout endTime advance)",
    () => {
      const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
      const last = new Date(Date.UTC(2026, 0, 15, 18, 0, 0));
      const priorExercise = {
        exercise: {
          interval: {
            endTime: "2026-01-15T16:30:00Z",
            endUtcOffset: "-18000s",
            startTime: "2026-01-15T16:00:00Z",
            startUtcOffset: "-18000s",
          },
        },
      };
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: last,
        },
        0,
        priorExercise,
      );
      eq(r.exerciseSource, "edit");
      eq(r.exercise.startUtcMs, first.getTime());
      eq(r.exercise.endUtcMs, last.getTime());
    },
  );

  // 'foreign' is the highest-priority exercise timing source: an overlapping
  // foreign session's manual start/stop is better evidence than our
  // edit-derived window or an interval we recorded earlier, so it must win over
  // BOTH 'edit' and 'prior', and its interval is borrowed verbatim.
  const FOREIGN_INTERVAL = {
    endUtcMs: Date.UTC(2026, 0, 15, 22, 45, 0),
    endUtcOffsetSeconds: EST,
    startUtcMs: Date.UTC(2026, 0, 15, 21, 30, 0),
    startUtcOffsetSeconds: EST,
  };
  t(
    "resolveRowTiming_ foreign match is borrowed verbatim and beats same-date edits",
    () => {
      const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
      const last = new Date(Date.UTC(2026, 0, 15, 18, 0, 0));
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: last,
        },
        0,
        null,
        FOREIGN_INTERVAL,
      );
      eq(r.exerciseSource, "foreign");
      eq(r.exercise.startUtcMs, FOREIGN_INTERVAL.startUtcMs);
      eq(r.exercise.endUtcMs, FOREIGN_INTERVAL.endUtcMs);
      eq(r.exercise.startOffsetSeconds, EST);
      eq(r.exercise.endOffsetSeconds, EST);
    },
  );
  t("resolveRowTiming_ foreign match beats a prior datapoint", () => {
    const prior = {
      exercise: {
        interval: {
          endTime: "2026-01-15T16:30:00Z",
          endUtcOffset: `${EST}s`,
          startTime: "2026-01-15T16:00:00Z",
          startUtcOffset: `${EST}s`,
        },
      },
    };
    const r = resolveRowTiming_(
      {
        date: JAN_15_NOON_UTC,
        exerciseFirstEditedAt: null,
        exercisesLastEditedAt: null,
      },
      0,
      prior,
      FOREIGN_INTERVAL,
    );
    eq(r.exerciseSource, "foreign");
    eq(r.exercise.startUtcMs, FOREIGN_INTERVAL.startUtcMs);
    eq(r.exercise.endUtcMs, FOREIGN_INTERVAL.endUtcMs);
  });
  // The two phases are independent: borrowing an exercise interval must not
  // drag the bodyweight sample time with it.
  t(
    "resolveRowTiming_ foreign match leaves the weight sample time alone",
    () => {
      const wEdit = new Date(Date.UTC(2026, 0, 15, 22, 0, 0)); // 5pm EST
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: null,
          exercisesLastEditedAt: null,
          weightEditedAt: wEdit,
        },
        0,
        null,
        FOREIGN_INTERVAL,
      );
      eq(r.weightSource, "edit");
      eq(r.weight, { offsetSeconds: EST, utcMs: wEdit.getTime() });
    },
  );

  t(
    "resolveRowTiming_ malformed prior exercise falls through to synthetic",
    () => {
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: JAN_20_3PM_EST,
          exercisesLastEditedAt: JAN_20_4PM_EST,
        },
        0,
        { exercise: {} },
      );
      eq(r.exerciseSource, "synthetic");
      eq(r.exercise.startUtcMs, Date.UTC(2026, 0, 15, 17, 0, 0));
    },
  );

  t(
    "resolveRowTiming_ single edit (start-only) -> 10 min default duration",
    () => {
      // Only one exercise edit: first == last, no observed end. The MIN floor
      // (10 min) doubles as the start-only default.
      const first = new Date(Date.UTC(2026, 0, 15, 17, 0, 0));
      const r = resolveRowTiming_(
        {
          date: JAN_15_NOON_UTC,
          exerciseFirstEditedAt: first,
          exercisesLastEditedAt: first,
        },
        0,
        null,
      );
      eq(r.exerciseSource, "edit");
      eq(r.exercise.startUtcMs, first.getTime());
      eq(r.exercise.endUtcMs - r.exercise.startUtcMs, 10 * 60 * 1000);
    },
  );

  // hasSendableExercises_: a zero-set-only row has nothing to send.
  t("hasSendableExercises_ false for empty", () =>
    eq(hasSendableExercises_([]), false),
  );
  t("hasSendableExercises_ false when all entries are zero-set", () =>
    eq(
      hasSendableExercises_([
        {
          entries: [{ assisted: false, reps: 5, sets: 0, weight: 200 }],
          name: "Bench",
        },
      ]),
      false,
    ),
  );
  t("hasSendableExercises_ true for a real set", () =>
    eq(
      hasSendableExercises_([
        {
          entries: [{ assisted: false, reps: 5, sets: 1, weight: 200 }],
          name: "Bench",
        },
      ]),
      true,
    ),
  );
  t("hasSendableExercises_ true for unknown-sets entry (sets null)", () =>
    eq(
      hasSendableExercises_([
        {
          entries: [{ assisted: false, reps: 5, sets: null, weight: 200 }],
          name: "Bench",
        },
      ]),
      true,
    ),
  );
  t("hasSendableExercises_ true when mixed zero-set and real", () =>
    eq(
      hasSendableExercises_([
        {
          entries: [
            { assisted: false, reps: 5, sets: 0, weight: 200 },
            { assisted: false, reps: 5, sets: 2, weight: 200 },
          ],
          name: "Bench",
        },
      ]),
      true,
    ),
  );

  // selectBackstopRows_: recent + sendable + not-yet-matched rows only.
  const bsNow = Date.UTC(2026, 0, 15, 17, 0, 0); // noon EST Jan 15
  const bsDate = (ymd) => new Date(Date.UTC(2026, 0, ymd, 17, 0, 0));
  const sendable = [
    {
      entries: [{ assisted: false, reps: 5, sets: 2, weight: 200 }],
      name: "Bench",
    },
  ];
  const zeroOnly = [
    {
      entries: [{ assisted: false, reps: 5, sets: 0, weight: 200 }],
      name: "Bench",
    },
  ];
  t(
    "selectBackstopRows_ picks recent unmatched sendable rows, drops matched/old/empty",
    () => {
      const rows = [
        {
          date: bsDate(15),
          exercises: sendable,
          matchedHealthSession: "",
          rowNum: 2,
        }, // today, unmatched -> pick
        // yesterday, unmatched -> pick
        {
          date: bsDate(14),
          exercises: sendable,
          matchedHealthSession: "",
          rowNum: 3,
        },
        {
          date: bsDate(15),
          exercises: sendable,
          matchedHealthSession: "foreign/x",
          rowNum: 4,
        }, // matched -> skip
        // 2 days back (outside lookback=2) -> skip
        {
          date: bsDate(13),
          exercises: sendable,
          matchedHealthSession: "",
          rowNum: 5,
        },
        // no sendable content -> skip
        {
          date: bsDate(15),
          exercises: zeroOnly,
          matchedHealthSession: "",
          rowNum: 6,
        },
      ];
      eq(
        selectBackstopRows_(rows, bsNow, 2).map((r) => r.rowNum),
        [2, 3],
      );
    },
  );

  t(
    "selectBackstopRows_ wantMatched=true picks recent matched sendable rows, drops unmatched/old/empty",
    () => {
      const rows = [
        // today, matched -> pick
        {
          date: bsDate(15),
          exercises: sendable,
          matchedHealthSession: "foreign/a",
          rowNum: 2,
        },
        // yesterday, matched -> pick
        {
          date: bsDate(14),
          exercises: sendable,
          matchedHealthSession: "foreign/b",
          rowNum: 3,
        },
        {
          date: bsDate(15),
          exercises: sendable,
          matchedHealthSession: "",
          rowNum: 4,
        }, // unmatched -> skip
        // outside lookback -> skip
        {
          date: bsDate(13),
          exercises: sendable,
          matchedHealthSession: "foreign/c",
          rowNum: 5,
        },
        // no sendable content -> skip
        {
          date: bsDate(15),
          exercises: zeroOnly,
          matchedHealthSession: "foreign/d",
          rowNum: 6,
        },
      ];
      eq(
        selectBackstopRows_(rows, bsNow, 2, true).map((r) => r.rowNum),
        [2, 3],
      );
    },
  );

  // selectStaleDataPointRows_: the state-based reconciliation path for cleared
  // content. A recorded datapoint that contradicts the row's cells is the one
  // unambiguous signal available, so this is what makes a clear recoverable no
  // matter how the user made it (onEditMarkDirty only sees single-cell clears).
  // Fixtures mirror what readRows produces, INCLUDING hasExerciseText /
  // hasWeightText. Those two are raw-cell facts, so a fixture that omits them
  // makes the guard vacuously true and the test pass for the wrong reason.
  const sRow = (o) =>
    Object.assign(
      {
        bodyweight: 185,
        exerciseSyncedAt: "SYNC",
        exercises: sendable,
        hasExerciseText: true,
        hasWeightText: true,
        healthIds: [],
        rowNum: 2,
        weightSyncedAt: "SYNC",
      },
      o,
    );
  // A row the user actually emptied: nothing parsed AND nothing in the cells.
  const cleared = (o) =>
    sRow(
      Object.assign(
        {
          bodyweight: null,
          exercises: [],
          hasExerciseText: false,
          hasWeightText: false,
        },
        o,
      ),
    );
  const EX_ID = "users/me/dataTypes/exercise/dataPoints/E1";
  const WT_ID = "users/me/dataTypes/weight/dataPoints/W1";

  t(
    "selectStaleDataPointRows_ nothing stale when content matches the tracked ids",
    () =>
      eq(selectStaleDataPointRows_([sRow({ healthIds: [EX_ID, WT_ID] })]), {
        exerciseRowNums: [],
        weightRowNums: [],
      }),
  );
  t(
    "selectStaleDataPointRows_ flags an exercise datapoint on an emptied row",
    () =>
      eq(selectStaleDataPointRows_([cleared({ healthIds: [EX_ID] })]), {
        exerciseRowNums: [2],
        weightRowNums: [],
      }),
  );
  t(
    "selectStaleDataPointRows_ flags a weight datapoint on an emptied row",
    () =>
      eq(selectStaleDataPointRows_([cleared({ healthIds: [WT_ID] })]), {
        exerciseRowNums: [],
        weightRowNums: [2],
      }),
  );
  t("selectStaleDataPointRows_ flags both phases independently", () =>
    eq(selectStaleDataPointRows_([cleared({ healthIds: [EX_ID, WT_ID] })]), {
      exerciseRowNums: [2],
      weightRowNums: [2],
    }),
  );
  // Nothing recorded means nothing to reconcile: a blank row is not stale.
  t("selectStaleDataPointRows_ ignores a row with no tracked ids", () =>
    eq(selectStaleDataPointRows_([cleared({ healthIds: [] })]), {
      exerciseRowNums: [],
      weightRowNums: [],
    }),
  );
  // An already-dirty row is picked up by the next pass regardless, so
  // re-dirtying it would only cost an extra write.
  t("selectStaleDataPointRows_ skips rows that are already dirty", () =>
    eq(
      selectStaleDataPointRows_([
        cleared({
          exerciseSyncedAt: "",
          healthIds: [EX_ID, WT_ID],
          weightSyncedAt: "",
        }),
      ]),
      { exerciseRowNums: [], weightRowNums: [] },
    ),
  );

  // The rule that keeps this from destroying history. "The parser produced
  // nothing" is not "the user cleared it": a cell holding text the parser
  // rejects still holds the user's data, so it must never be deleted. Blanking
  // an exercise column's HEADER makes readRows stop building `exercises` for
  // every historical row at once; a bulk reformat or a bounds change does the
  // same to `bodyweight`. Both leave the raw text in place, which is what these
  // two guard on.
  t(
    "selectStaleDataPointRows_ spares a row whose exercise cells still hold text",
    () =>
      eq(
        selectStaleDataPointRows_([
          cleared({ hasExerciseText: true, healthIds: [EX_ID] }),
        ]),
        { exerciseRowNums: [], weightRowNums: [] },
      ),
  );
  t(
    "selectStaleDataPointRows_ spares a row whose Weight cell still holds text",
    () =>
      eq(
        selectStaleDataPointRows_([
          cleared({ hasWeightText: true, healthIds: [WT_ID] }),
        ]),
        { exerciseRowNums: [], weightRowNums: [] },
      ),
  );
  // A zero-set-only row parses to no sendable content but its cell is not
  // empty, so the backstop leaves it alone; the single-cell onEdit clear path
  // is what reconciles that edit, and erring toward keeping a datapoint is the
  // safe direction.
  t("selectStaleDataPointRows_ spares a zero-set-only row (text present)", () =>
    eq(
      selectStaleDataPointRows_([
        sRow({ exercises: zeroOnly, healthIds: [EX_ID] }),
      ]),
      { exerciseRowNums: [], weightRowNums: [] },
    ),
  );

  // Disjoint from selectBackstopRows_ by construction: that requires sendable
  // content, this requires none, so the backstop can concat without dedup.
  t(
    "selectStaleDataPointRows_ and selectBackstopRows_ never select the same exercise row",
    () => {
      const rows = [
        Object.assign(sRow({ healthIds: [EX_ID], rowNum: 2 }), {
          date: bsDate(15),
          matchedHealthSession: "",
        }),
        Object.assign(cleared({ healthIds: [EX_ID], rowNum: 3 }), {
          date: bsDate(15),
          matchedHealthSession: "",
        }),
      ];
      eq(
        selectBackstopRows_(rows, bsNow, 2).map((r) => r.rowNum),
        [2],
        "foreign re-review takes the row with content",
      );
      eq(
        selectStaleDataPointRows_(rows).exerciseRowNums,
        [3],
        "stale takes the emptied row",
      );
    },
  );

  // exerciseUnchanged_: skip the recreate only when interval + notes all match.
  const priorEx = (startIso, endIso, notes) => ({
    exercise: { interval: { endTime: endIso, startTime: startIso }, notes },
  });
  const exStart = Date.UTC(2026, 0, 15, 17, 0, 0);
  const exEnd = Date.UTC(2026, 0, 15, 17, 30, 0);
  t("exerciseUnchanged_ true when interval and notes all match", () =>
    eq(
      exerciseUnchanged_(
        priorEx(
          "2026-01-15T17:00:00Z",
          "2026-01-15T17:30:00Z",
          "Bench: 200x5x2",
        ),
        exStart,
        exEnd,
        "Bench: 200x5x2",
      ),
      true,
    ),
  );
  t("exerciseUnchanged_ false when startTime differs", () =>
    eq(
      exerciseUnchanged_(
        priorEx(
          "2026-01-15T17:00:00Z",
          "2026-01-15T17:30:00Z",
          "Bench: 200x5x2",
        ),
        exStart - 60000,
        exEnd,
        "Bench: 200x5x2",
      ),
      false,
    ),
  );
  t("exerciseUnchanged_ false when endTime differs", () =>
    eq(
      exerciseUnchanged_(
        priorEx(
          "2026-01-15T17:00:00Z",
          "2026-01-15T17:30:00Z",
          "Bench: 200x5x2",
        ),
        exStart,
        exEnd + 60000,
        "Bench: 200x5x2",
      ),
      false,
    ),
  );
  t("exerciseUnchanged_ false when notes differ", () =>
    eq(
      exerciseUnchanged_(
        priorEx(
          "2026-01-15T17:00:00Z",
          "2026-01-15T17:30:00Z",
          "Bench: 200x5x2",
        ),
        exStart,
        exEnd,
        "Bench: 200x5x3",
      ),
      false,
    ),
  );
  t("exerciseUnchanged_ false when interval missing", () =>
    eq(
      exerciseUnchanged_({ exercise: { notes: "x" } }, exStart, exEnd, "x"),
      false,
    ),
  );

  // selectOrphanDataPointNames_: delete untracked datapoints from our own web
  // client, leave tracked / foreign / other-client / unattributable ones alone.
  const oCand = (name, clientId) => ({
    googleWebClientId: clientId || null,
    name,
  });
  t(
    "selectOrphanDataPointNames_ deletes our untracked datapoint (client derived from a tracked one)",
    () => {
      const candidates = [
        oCand("ex/tracked", "ours"), // tracked -> establishes "ours"
        oCand("ex/orphan", "ours"), // untracked, same client -> orphan
      ];
      eq(selectOrphanDataPointNames_(candidates, { "ex/tracked": true }), [
        "ex/orphan",
      ]);
    },
  );
  t("selectOrphanDataPointNames_ keeps tracked datapoints", () => {
    const candidates = [oCand("ex/tracked", "ours")];
    eq(selectOrphanDataPointNames_(candidates, { "ex/tracked": true }), []);
  });
  t(
    "selectOrphanDataPointNames_ keeps foreign datapoints (null client id)",
    () => {
      const candidates = [
        oCand("ex/tracked", "ours"),
        oCand("foreign/device", null), // device/first-party -> never an orphan
      ];
      eq(selectOrphanDataPointNames_(candidates, { "ex/tracked": true }), []);
    },
  );
  t(
    "selectOrphanDataPointNames_ keeps untracked datapoints from a different web client",
    () => {
      const candidates = [
        oCand("ex/tracked", "ours"),
        oCand("other/app", "theirs"), // another web app, not ours -> keep
      ];
      eq(selectOrphanDataPointNames_(candidates, { "ex/tracked": true }), []);
    },
  );
  t(
    "selectOrphanDataPointNames_ deletes nothing when ownership cannot be attributed",
    () => {
      const candidates = [oCand("ex/orphan", "ours")]; // no tracked candidate to derive "ours"
      eq(selectOrphanDataPointNames_(candidates, {}), []);
    },
  );

  // resolveForeignMatches_ tests. listStrengthOnDate is stubbed per-test so
  // we control the foreign candidate list without hitting the API.
  const withStubbedList = (stub, fn) =>
    withGlobals({ listStrengthOnDate: stub }, fn);

  // 2026-01-15 12:00 UTC = 2026-01-15 07:00 EST, civil date 2026-01-15.
  const FOREIGN_DATE = new Date(Date.UTC(2026, 0, 15, 12, 0, 0));
  const fRow_ = (overrides) =>
    Object.assign(
      {
        date: FOREIGN_DATE,
        exerciseFirstEditedAt: null,
        exercises: [
          {
            entries: [{ assisted: false, reps: 5, sets: 3, weight: 135 }],
            name: "Bench",
          },
        ],
        exercisesLastEditedAt: null,
        healthIds: [],
        matchedHealthSession: "",
        rowNum: 10,
      },
      overrides,
    );
  const fCand_ = (name, startUtcMs, endUtcMs) => ({
    endUtcMs,
    endUtcOffsetSeconds: EST,
    name,
    startUtcMs,
    startUtcOffsetSeconds: EST,
  });

  t(
    "resolveForeignMatches_ time-range matches on-date row to overlapping candidate",
    () => {
      const row = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)), // 5pm EST
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0)), // 6pm EST
      });
      const cand = fCand_(
        "foreign/A",
        Date.UTC(2026, 0, 15, 22, 0, 0),
        Date.UTC(2026, 0, 15, 23, 0, 0),
      );
      withStubbedList(
        () => [cand],
        () => {
          const plan = resolveForeignMatches_([], [], [row]);
          eq(plan[10] && plan[10].name, "foreign/A");
        },
      );
    },
  );

  t(
    "resolveForeignMatches_ matches across a civil-date boundary (midnight-crossing workout)",
    () => {
      // Edits 11:45pm EST Jan 15 -> 12:15am EST Jan 16 (exerciseFirstEditedAt is
      // still on row.date Jan 15). The window straddles midnight, so candidates
      // are probed for both Jan 15 and Jan 16. The foreign session was logged
      // just after midnight (12:00-12:30am EST Jan 16), a different civil date
      // than the row, and must still match on absolute-UTC overlap.
      const row = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 16, 4, 45, 0)),
        // 11:45pm EST Jan 15
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 16, 5, 15, 0)),
        rowNum: 10, // 12:15am EST Jan 16
      });
      const cand = fCand_(
        "foreign/after-midnight",
        Date.UTC(2026, 0, 16, 5, 0, 0),
        Date.UTC(2026, 0, 16, 5, 30, 0),
      ); // 12:00-12:30am EST Jan 16
      // Answer by civil date: a stub that returns the session for any date
      // would match even if only the window's start date were probed.
      withStubbedList(
        (date) => (ymd(date) === "2026-01-16" ? [cand] : []),
        () => {
          const plan = resolveForeignMatches_([], [], [row]);
          eq(plan[10] && plan[10].name, "foreign/after-midnight");
        },
      );
    },
  );

  t(
    "resolveForeignMatches_ off-date row gets no match (no ordinal fallback)",
    () => {
      // Row dated Jan 15 but edited Jan 20: off-date timestamps anchor no
      // trustworthy window. With the ordinal fallback removed the row gets no
      // alignment and falls through to its own synthetic/prior timing.
      const row = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 20, 22, 0, 0)),
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 20, 23, 0, 0)),
      });
      const cand = fCand_(
        "foreign/A",
        Date.UTC(2026, 0, 20, 22, 0, 0),
        Date.UTC(2026, 0, 20, 23, 0, 0),
      );
      withStubbedList(
        () => [cand],
        () => {
          const plan = resolveForeignMatches_([], [], [row]);
          eq(plan[10], undefined);
        },
      );
    },
  );

  t(
    "resolveForeignMatches_ no-timestamp row gets no match (no ordinal fallback)",
    () => {
      const row = fRow_({ rowNum: 10 });
      const cand = fCand_(
        "foreign/A",
        Date.UTC(2026, 0, 15, 22, 0, 0),
        Date.UTC(2026, 0, 15, 23, 0, 0),
      );
      withStubbedList(
        () => [cand],
        () => {
          const plan = resolveForeignMatches_([], [], [row]);
          eq(plan[10], undefined);
        },
      );
    },
  );

  t(
    "resolveForeignMatches_ excludes sync-created candidates (own datapoint not realigned)",
    () => {
      // On-date row whose window overlaps the candidate, but the candidate IS
      // the row's own prior datapoint, so it must be excluded rather than
      // aligned to itself on re-sync.
      const ownName = "users/me/dataTypes/exercise/dataPoints/123";
      const row = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0)),
        healthIds: [ownName],
        rowNum: 10,
      });
      const cand = fCand_(
        ownName,
        Date.UTC(2026, 0, 15, 22, 0, 0),
        Date.UTC(2026, 0, 15, 23, 0, 0),
      );
      withStubbedList(
        () => [cand],
        () => {
          const plan = resolveForeignMatches_([ownName], [], [row]);
          eq(plan[10], undefined);
        },
      );
    },
  );

  t(
    "resolveForeignMatches_ excludes candidates already aligned-elsewhere by a non-ready row",
    () => {
      // The ready row's window overlaps the candidate (so it would align absent
      // the exclusion), but row 5 already aligned to it. The exclusion is keyed
      // off the full-sheet allMatchedSessions list, so it holds whether row 5 is
      // merely not-ready this pass or was dropped by readRows for a blank Date.
      const readyRow = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0)),
        rowNum: 10,
      });
      const cand = fCand_(
        "foreign/A",
        Date.UTC(2026, 0, 15, 22, 0, 0),
        Date.UTC(2026, 0, 15, 23, 0, 0),
      );
      withStubbedList(
        () => [cand],
        () => {
          const plan = resolveForeignMatches_(
            [],
            [{ name: "foreign/A", rowNum: 5 }],
            [readyRow],
          );
          eq(plan[10], undefined);
        },
      );
    },
  );

  // The aligned-elsewhere exclusion covers only rows that are NOT ready. A
  // ready row re-syncing (the backstop's matched-row re-review) must be able to
  // borrow the session recorded against it again; excluding it would clear the
  // match and rebuild the datapoint on edit timing every cycle.
  t(
    "resolveForeignMatches_ lets a ready row re-borrow the session recorded against it",
    () => {
      const row = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0)),
        matchedHealthSession: "foreign/A",
        rowNum: 10,
      });
      const cand = fCand_(
        "foreign/A",
        Date.UTC(2026, 0, 15, 22, 0, 0),
        Date.UTC(2026, 0, 15, 23, 0, 0),
      );
      withStubbedList(
        () => [cand],
        () => {
          const plan = resolveForeignMatches_(
            [],
            [{ name: "foreign/A", rowNum: 10 }],
            [row],
          );
          eq(plan[10] && plan[10].name, "foreign/A");
        },
      );
    },
  );

  t(
    "resolveForeignMatches_ time-range window is clamped to MAX_EXERCISE_DURATION_MS",
    () => {
      // First edit 9am on row.date, last edit drifted 5 days forward. Unclamped
      // the window would span 5 days and catch the 5pm candidate; clamped it
      // ends at 9am + 2h + the 10min buffer, well before it.
      const row = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 14, 0, 0)),
        // 9am EST Jan 15
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 20, 14, 0, 0)),
        rowNum: 10, // 9am EST Jan 20
      });
      // Candidate at 5pm-6pm EST Jan 15, outside the clamped window but
      // inside the unclamped one.
      const cand = fCand_(
        "foreign/late",
        Date.UTC(2026, 0, 15, 22, 0, 0),
        Date.UTC(2026, 0, 15, 23, 0, 0),
      );
      withStubbedList(
        () => [cand],
        () => {
          const plan = resolveForeignMatches_([], [], [row]);
          eq(plan[10], undefined);
        },
      );
    },
  );

  t(
    "resolveForeignMatches_ time-range picks the best-overlap candidate when several exist",
    () => {
      // Row window 4:50pm-6:10pm EST (5pm-6pm edit + the 10min buffer each
      // side). candA: 7am-8am EST, no overlap. candB: 5pm-6pm EST, full overlap.
      const row = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0)),
      });
      const candA = fCand_(
        "foreign/early",
        Date.UTC(2026, 0, 15, 12, 0, 0),
        Date.UTC(2026, 0, 15, 13, 0, 0),
      );
      const candB = fCand_(
        "foreign/match",
        Date.UTC(2026, 0, 15, 22, 0, 0),
        Date.UTC(2026, 0, 15, 23, 0, 0),
      );
      withStubbedList(
        () => [candA, candB],
        () => {
          const plan = resolveForeignMatches_([], [], [row]);
          eq(plan[10] && plan[10].name, "foreign/match");
        },
      );
    },
  );

  // Two rows must never borrow the same foreign session: the claimed candidate
  // is spliced out of the pool, so the second row falls through to its own
  // timing rather than being aligned to a workout another row already owns.
  t(
    "resolveForeignMatches_ gives one candidate to a single row, not to both",
    () => {
      const rowA = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
        // 5:00pm EST
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 30, 0)),
        rowNum: 10,
      });
      const rowB = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 15, 0)),
        // 5:15pm EST
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 45, 0)),
        rowNum: 11,
      });
      const cand = fCand_(
        "foreign/only",
        Date.UTC(2026, 0, 15, 22, 0, 0),
        Date.UTC(2026, 0, 15, 23, 0, 0),
      );
      withStubbedList(
        () => [cand],
        () => {
          const plan = resolveForeignMatches_([], [], [rowA, rowB]);
          eq(
            plan[10] && plan[10].name,
            "foreign/only",
            "assignment runs in rowNum order",
          );
          eq(
            plan[11],
            undefined,
            "the second row gets nothing rather than the same session",
          );
        },
      );
    },
  );

  t(
    "resolveForeignMatches_ assigns each of two rows its own overlapping candidate",
    () => {
      const morningRow = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 14, 0, 0)),
        // 9:00am EST
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 14, 30, 0)),
        rowNum: 10,
      });
      const eveningRow = fRow_({
        exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
        // 5:00pm EST
        exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 30, 0)),
        rowNum: 11,
      });
      const morningCand = fCand_(
        "foreign/morning",
        Date.UTC(2026, 0, 15, 14, 0, 0),
        Date.UTC(2026, 0, 15, 15, 0, 0),
      );
      const eveningCand = fCand_(
        "foreign/evening",
        Date.UTC(2026, 0, 15, 22, 0, 0),
        Date.UTC(2026, 0, 15, 23, 0, 0),
      );
      withStubbedList(
        () => [morningCand, eveningCand],
        () => {
          const plan = resolveForeignMatches_([], [], [morningRow, eveningRow]);
          eq(plan[10] && plan[10].name, "foreign/morning");
          eq(plan[11] && plan[11].name, "foreign/evening");
        },
      );
    },
  );

  // A zero-set-only row records timing but produces no datapoint, so it must
  // not anchor a window and claim a candidate a real row could have used.
  t("resolveForeignMatches_ zero-set-only row anchors no window", () => {
    const row = fRow_({
      exerciseFirstEditedAt: new Date(Date.UTC(2026, 0, 15, 22, 0, 0)),
      exercises: [
        {
          entries: [{ assisted: false, reps: 5, sets: 0, weight: 200 }],
          name: "Bench",
        },
      ],
      exercisesLastEditedAt: new Date(Date.UTC(2026, 0, 15, 23, 0, 0)),
      rowNum: 10,
    });
    const cand = fCand_(
      "foreign/A",
      Date.UTC(2026, 0, 15, 22, 0, 0),
      Date.UTC(2026, 0, 15, 23, 0, 0),
    );
    withStubbedList(
      () => [cand],
      () => {
        eq(resolveForeignMatches_([], [], [row])[10], undefined);
      },
    );
  });

  // ---- Health API request/response shaping --------------------------------
  // httpJson_ is stubbed so these stay unit tests: no UrlFetchApp fake, just
  // the payload we build and the response we parse.

  t(
    "createExerciseAt sends a STRENGTH_TRAINING payload and returns the new name",
    () => {
      let sent = null;
      const created = withGlobals(
        {
          httpJson_: (method, url, payload) => {
            sent = { method, payload };
            return {
              done: true,
              response: { name: "users/1/dataTypes/exercise/dataPoints/E1" },
            };
          },
        },
        () =>
          createExerciseAt(
            Date.UTC(2026, 0, 15, 17, 0, 0),
            EST,
            Date.UTC(2026, 0, 15, 17, 30, 0),
            EST,
            "Bench press, 135 lbs, 3 sets of 5",
          ),
      );
      eq(created, "users/1/dataTypes/exercise/dataPoints/E1");
      eq(sent.method, "POST");
      eq(sent.payload.exercise.exerciseType, "STRENGTH_TRAINING");
      // No displayName: the server derives the card's title from exerciseType
      // for every type but OTHER, so sending one is at best redundant and at
      // worst our wording drifting from the server's.
      eq(
        Object.prototype.hasOwnProperty.call(
          sent.payload.exercise,
          "displayName",
        ),
        false,
        "displayName must not be sent",
      );
      eq(sent.payload.exercise.notes, "Bench press, 135 lbs, 3 sets of 5");
      // Neither activeDuration nor displayName is sent: the server computes
      // one from the interval and derives the other from exerciseType, so a
      // client copy of either can only drift from what Health actually stores.
      eq(
        Object.keys(sent.payload.exercise).sort(),
        ["exerciseType", "interval", "notes"],
        "only the fields the server does not write for itself",
      );
      eq(sent.payload.exercise.interval, {
        endTime: "2026-01-15T17:30:00Z",
        endUtcOffset: "-18000s",
        startTime: "2026-01-15T17:00:00Z",
        startUtcOffset: "-18000s",
      });
    },
  );

  // A create we can't track is treated as a FAILED create: throwing keeps the
  // row dirty so it retries, rather than stamping it synced and orphaning
  // whatever the server may have made.
  t("createExerciseAt throws when the create returns no resource name", () =>
    throws(
      () =>
        withGlobals({ httpJson_: () => ({ done: true }) }, () =>
          createExerciseAt(0, 0, 1000, 0, "notes"),
        ),
      /no datapoint name/,
      "a nameless create must throw",
    ),
  );

  t("createWeightAt converts pounds to grams and returns the new name", () => {
    let sent = null;
    const created = withGlobals(
      {
        httpJson_: (method, url, payload) => {
          sent = payload;
          return {
            done: true,
            response: { name: "users/1/dataTypes/weight/dataPoints/W1" },
          };
        },
      },
      () => createWeightAt(Date.UTC(2026, 0, 15, 17, 0, 0), EST, 185),
    );
    eq(created, "users/1/dataTypes/weight/dataPoints/W1");
    // Literal grams, not Math.round(185 * GRAMS_PER_LB): recomputing the
    // expected value with the constant under test asserts only that the
    // constant equals itself, so a wrong conversion factor would pass.
    eq(sent.weight.weightGrams, 83915);
    eq(sent.weight.sampleTime.physicalTime, "2026-01-15T17:00:00Z");
    eq(sent.weight.sampleTime.utcOffset, "-18000s");
  });

  t("createWeightAt throws when the create returns no resource name", () =>
    throws(
      () =>
        withGlobals({ httpJson_: () => ({ done: true }) }, () =>
          createWeightAt(0, 0, 185),
        ),
      /no datapoint name/,
      "a nameless create must throw",
    ),
  );

  // sampleTime is mandatory in the PATCH body (the server 500s without it) and
  // the resource is identified by the URL, which must use the literal `me`.
  t("patchWeight PATCHes the me-form URL and echoes sampleTime back", () => {
    const sampleTime = {
      physicalTime: "2026-01-15T17:00:00Z",
      utcOffset: "-18000s",
    };
    let sent = null;
    withGlobals(
      {
        httpJson_: (method, url, payload) => {
          sent = { method, payload, url };
          return {};
        },
      },
      () =>
        patchWeight(
          "users/123/dataTypes/weight/dataPoints/W1",
          sampleTime,
          // 187 lb converts to 84821.77 g, so rounding and truncation differ.
          // At 186 lb they agree, which left the conversion unpinned here.
          187,
        ),
    );
    eq(sent.method, "PATCH");
    // The whole URL, as a literal: a tail-anchored regex leaves HEALTH_API_BASE
    // unasserted, and interpolating the constant here would only assert that it
    // equals itself. This is the one place the base URL is pinned.
    eq(
      sent.url,
      "https://health.googleapis.com/v4/users/me/dataTypes/weight/dataPoints/W1",
      "the me-form resource URL under the configured API base",
    );
    eq(sent.payload, {
      weight: { sampleTime, weightGrams: 84822 },
    });
  });

  // The projection foreign matching and orphan attribution both consume: only
  // STRENGTH_TRAINING sessions with a usable interval, sorted by start, with a
  // null googleWebClientId for device / first-party sources.
  t(
    "listStrengthOnDate keeps usable STRENGTH_TRAINING only and maps the client id",
    () => {
      const point = (name, type, startTime, endTime, app) => ({
        dataSource: { application: app },
        exercise:
          startTime === null
            ? { exerciseType: type }
            : {
                exerciseType: type,
                interval: {
                  endTime,
                  endUtcOffset: "-18000s",
                  startTime,
                  startUtcOffset: "-18000s",
                },
              },
        name,
      });
      const out = withGlobals(
        {
          httpJson_: () => ({
            dataPoints: [
              point(
                "ex/late",
                "STRENGTH_TRAINING",
                "2026-01-15T22:00:00Z",
                "2026-01-15T23:00:00Z",
                { googleWebClientId: "ours" },
              ),
              point(
                "ex/early",
                "STRENGTH_TRAINING",
                "2026-01-15T14:00:00Z",
                "2026-01-15T15:00:00Z",
                null,
              ),
              point(
                "ex/running",
                "RUNNING",
                "2026-01-15T15:00:00Z",
                "2026-01-15T16:00:00Z",
                { googleWebClientId: "ours" },
              ),
              point("ex/no-interval", "STRENGTH_TRAINING", null, null, {
                googleWebClientId: "ours",
              }),
            ],
          }),
        },
        () => listStrengthOnDate(new Date(Date.UTC(2026, 0, 15, 17, 0, 0))),
      );
      eq(
        out.map((c) => c.name),
        ["ex/early", "ex/late"],
        "other exercise types and interval-less points dropped, sorted by start",
      );
      eq(
        out.map((c) => c.googleWebClientId),
        [null, "ours"],
        "a device session carries a null client id and is never treated as ours",
      );
      eq(out[1].startUtcMs, Date.UTC(2026, 0, 15, 22, 0, 0));
      eq(out[1].startUtcOffsetSeconds, EST);
    },
  );

  // findRowDateViolation_: trigger-entry date validation (increasing order,
  // no duplicate dates, year within [MIN_ROW_DATE_YEAR, MAX_ROW_DATE_YEAR]).
  // UTC noon keeps the civil date stable in the test time zone.
  const vDate = (y, m, d) => new Date(Date.UTC(y, m - 1, d, 12, 0, 0));
  const vRow = (rowNum, date) => ({ date, rowNum });

  // listWeightOnDate had no test at all, though listStrengthOnDate does. It
  // feeds weight orphan attribution, and selectOrphanDataPointNames_ only spares
  // a candidate whose googleWebClientId is null. Reporting our client id for a
  // foreign datapoint (a scale, the Health app, an assistant entry) makes the
  // backstop delete data we do not own.
  t("listWeightOnDate maps the client id and spares foreign datapoints", () => {
    let sentUrl = null;
    const out = withGlobals(
      {
        httpJson_: (method, url) => {
          sentUrl = url;
          return {
            dataPoints: [
              {
                dataSource: { application: { googleWebClientId: "ours" } },
                name: "users/1/dataTypes/weight/dataPoints/W1",
              },
              // Device / first-party: application is null.
              {
                dataSource: { application: null },
                name: "users/1/dataTypes/weight/dataPoints/W2",
              },
              { name: "users/1/dataTypes/weight/dataPoints/W3" }, // no dataSource at all
              { dataSource: { application: null } }, // no name: dropped
            ],
          };
        },
      },
      () => listWeightOnDate(new Date(Date.UTC(2026, 0, 15, 17, 0, 0))),
    );
    eq(out, [
      {
        googleWebClientId: "ours",
        name: "users/1/dataTypes/weight/dataPoints/W1",
      },
      {
        googleWebClientId: null,
        name: "users/1/dataTypes/weight/dataPoints/W2",
      },
      {
        googleWebClientId: null,
        name: "users/1/dataTypes/weight/dataPoints/W3",
      },
    ]);
    // The filter member is verified against the live API; a wrong one returns
    // 400 INVALID_DATA_POINT_FILTER rather than an empty list.
    eq(
      sentUrl.indexOf(encodeURIComponent("weight.sample_time.civil_time")) !==
        -1,
      true,
      sentUrl,
    );
  });

  // Paging: a day with more datapoints than one page must not silently truncate.
  // Under-listing makes orphan reconciliation miss datapoints and foreign
  // matching miss sessions, with no error either way.
  t("listWeightOnDate follows nextPageToken to the end", () => {
    const urls = [];
    const out = withGlobals(
      {
        httpJson_: (method, url) => {
          urls.push(url);
          return urls.length === 1
            ? {
                dataPoints: [
                  { name: "users/1/dataTypes/weight/dataPoints/W1" },
                ],
                nextPageToken: "PAGE2",
              }
            : {
                dataPoints: [
                  { name: "users/1/dataTypes/weight/dataPoints/W2" },
                ],
              };
        },
      },
      () => listWeightOnDate(new Date(Date.UTC(2026, 0, 15, 17, 0, 0))),
    );
    eq(urls.length, 2, "the second page was requested");
    eq(
      urls[1].indexOf("pageToken=PAGE2") !== -1,
      true,
      `the token is passed back: ${urls[1]}`,
    );
    eq(
      out.map((p) => p.name),
      [
        "users/1/dataTypes/weight/dataPoints/W1",
        "users/1/dataTypes/weight/dataPoints/W2",
      ],
      "both pages are returned",
    );
  });

  // Every delete path goes through this. Names are grouped by the data type
  // parsed out of the resource name, because batchDelete lives under a
  // per-type collection URL; a wrong type means the delete silently targets the
  // wrong collection.
  t("deleteDataPointsByName batches per data type and skips junk names", () => {
    const calls = [];
    withGlobals(
      {
        httpJson_: (method, url, payload) => {
          calls.push({ method, payload, url });
          return {};
        },
      },
      () =>
        deleteDataPointsByName([
          "users/1/dataTypes/exercise/dataPoints/E1",
          "users/1/dataTypes/weight/dataPoints/W1",
          "users/1/dataTypes/exercise/dataPoints/E2",
          "not-a-resource-name",
        ]),
    );
    eq(calls.length, 2, "one batch per data type, junk dropped");
    eq(
      calls[0].url,
      "https://health.googleapis.com/v4/users/me/dataTypes/exercise/dataPoints:batchDelete",
    );
    eq(calls[0].payload, {
      names: [
        "users/1/dataTypes/exercise/dataPoints/E1",
        "users/1/dataTypes/exercise/dataPoints/E2",
      ],
    });
    eq(
      calls[1].url,
      "https://health.googleapis.com/v4/users/me/dataTypes/weight/dataPoints:batchDelete",
    );
    eq(calls[1].payload, {
      names: ["users/1/dataTypes/weight/dataPoints/W1"],
    });
  });

  t("deleteDataPointsByName makes no request for an empty list", () => {
    let called = 0;
    withGlobals(
      {
        httpJson_: () => {
          called++;
          return {};
        },
      },
      () => {
        deleteDataPointsByName([]);
        deleteDataPointsByName(null);
      },
    );
    eq(called, 0, "nothing to delete means no API call");
  });

  t("findRowDateViolation_ empty rows -> null", () =>
    eq(findRowDateViolation_([]), null),
  );
  t("findRowDateViolation_ increasing in-range dates -> null", () =>
    eq(
      findRowDateViolation_([
        vRow(2, vDate(2025, 1, 1)),
        vRow(3, vDate(2026, 6, 1)),
        vRow(4, vDate(2049, 12, 31)),
      ]),
      null,
    ),
  );
  t("findRowDateViolation_ duplicate date flagged", () => {
    const v = findRowDateViolation_([
      vRow(2, vDate(2026, 1, 15)),
      vRow(3, vDate(2026, 1, 15)),
    ]);
    eq(/rows 2 and 3 share the date 2026-01-15/.test(v), true, v);
  });
  t(
    "findRowDateViolation_ same civil day, different times is a duplicate",
    () => {
      const v = findRowDateViolation_([
        vRow(2, new Date(Date.UTC(2026, 0, 15, 13, 0, 0))),
        vRow(3, new Date(Date.UTC(2026, 0, 15, 20, 0, 0))),
      ]);
      eq(/share the date 2026-01-15/.test(v), true, v);
    },
  );
  t("findRowDateViolation_ decreasing date flagged", () => {
    const v = findRowDateViolation_([
      vRow(2, vDate(2026, 1, 16)),
      vRow(3, vDate(2026, 1, 15)),
    ]);
    eq(
      /row 3 \(2026-01-15\) is dated before row 2 \(2026-01-16\)/.test(v),
      true,
      v,
    );
  });
  t("findRowDateViolation_ year below MIN flagged", () => {
    const v = findRowDateViolation_([vRow(2, vDate(2024, 12, 31))]);
    eq(
      /row 2: date 2024-12-31 is outside the allowed years 2025-2049/.test(v),
      true,
      v,
    );
  });
  t("findRowDateViolation_ year above MAX flagged", () => {
    const v = findRowDateViolation_([vRow(2, vDate(2050, 1, 1))]);
    eq(/outside the allowed years/.test(v), true, v);
  });

  // ---------------------------------------------------------------------------
  // httpJson_ transport: every other Health API test stubs httpJson_ itself, so
  // without these the retry policy, the 2xx/transient classification, and the
  // statusCode field are unexercised. That field is what isNotFoundError_ reads,
  // and the sync's 404-recovery paths (weight GET, weight DELETE, exercise
  // DELETE) all depend on it being set here.
  // ---------------------------------------------------------------------------
  const fakeFetch = (replies) => {
    const calls = [];
    const fetch = (url, options) => {
      calls.push({ options, url });
      const r = replies[Math.min(calls.length - 1, replies.length - 1)];
      if (r.throw) {
        throw new Error(r.throw);
      }
      return {
        getContentText: () => r.body,
        getResponseCode: () => r.code,
      };
    };
    return { calls, fetch };
  };
  // Runs fn with UrlFetchApp, the OAuth token and Utilities.sleep stubbed;
  // returns { result, err, calls, sleeps }. authHeaders_ caches its result in a
  // module-level binding no test can reach, so the first httpJson_ test in the
  // run is the only one that observes the token: a new test asserting a
  // different token has to be the first one, not merely added here.
  const withHttp = (replies, fn) => {
    const f = fakeFetch(replies);
    const sleeps = [];
    const out = withGlobals(
      {
        UrlFetchApp: { fetch: f.fetch },
        Utilities: Object.assign({}, Utilities, {
          sleep: (ms) => sleeps.push(ms),
        }),
        getHealthAccessToken_: () => "test-token",
      },
      () => {
        try {
          return { result: fn() };
        } catch (err) {
          return { err };
        }
      },
    );
    return Object.assign(out, { calls: f.calls, sleeps });
  };

  t("httpJson_ parses a 2xx body and sends the auth header", () => {
    const r = withHttp([{ body: '{"ok":1}', code: 200 }], () =>
      httpJson_("GET", "https://example/x"),
    );
    eq(r.result, { ok: 1 });
    eq(r.calls.length, 1, "no retry on success");
    eq(r.calls[0].options.method, "GET");
    eq(r.calls[0].options.headers.Authorization, "Bearer test-token");
    eq(r.calls[0].options.headers.Accept, "application/json");
    eq(r.calls[0].options.contentType, "application/json");
    eq(
      r.calls[0].options.muteHttpExceptions,
      true,
      "codes are read, not thrown",
    );
  });

  // authHeaders_ returns from its cache on every call after the first, which is
  // a separate return path from the one that builds the headers. Asserting the
  // header on a single call leaves that path free to come back empty, which in
  // production means every request but the first goes out unauthenticated.
  // Whether the token is re-fetched cannot be asserted here: the cache lives in
  // a module-level binding no test can reset, so it is already warm by now.
  t("httpJson_ authenticates the calls that come from the header cache", () => {
    const f = fakeFetch([{ body: "{}", code: 200 }]);
    withGlobals(
      {
        UrlFetchApp: { fetch: f.fetch },
        getHealthAccessToken_: () => "test-token",
      },
      () => {
        httpJson_("GET", "https://example/a");
        httpJson_("GET", "https://example/b");
      },
    );
    eq(
      f.calls.map((c) => c.options.headers.Authorization),
      ["Bearer test-token", "Bearer test-token"],
      "the second call is authenticated too",
    );
  });

  // An empty 2xx body is a success, not a parse error: batchDelete answers 200
  // with no content.
  t("httpJson_ maps an empty 2xx body to {}", () => {
    const r = withHttp([{ body: "", code: 200 }], () =>
      httpJson_("POST", "https://example/x:batchDelete", { names: ["a"] }),
    );
    eq(r.result, {});
    eq(r.calls[0].options.payload, '{"names":["a"]}');
  });

  t("httpJson_ omits the payload when none is given", () => {
    const r = withHttp([{ body: "{}", code: 200 }], () =>
      httpJson_("GET", "https://example/x"),
    );
    eq("payload" in r.calls[0].options, false, "no body on a GET");
  });

  // The 2xx boundary: 299 succeeds, 300 does not. The 300 body is valid JSON on
  // purpose: with unparseable content a misclassified 300 would throw in
  // JSON.parse and the assertion would pass for the wrong reason.
  t("httpJson_ treats 299 as success and 300 as an error", () => {
    eq(
      withHttp([{ body: "{}", code: 299 }], () => httpJson_("GET", "u")).result,
      {},
    );
    const r = withHttp([{ body: '{"ok":1}', code: 300 }], () =>
      httpJson_("GET", "u"),
    );
    eq(r.err.statusCode, 300, "300 must not be read as success");
  });

  // isNotFoundError_ reads statusCode off the thrown error. Producing it here is
  // what makes every 404-recovery path in the sync reachable.
  t("httpJson_ throws a 404 carrying statusCode, without retrying", () => {
    const r = withHttp([{ body: "gone", code: 404 }], () =>
      httpJson_("GET", "https://example/x"),
    );
    eq(r.err.statusCode, 404);
    eq(isNotFoundError_(r.err), true, "the 404-recovery paths key off this");
    eq(r.calls.length, 1, "a 404 is permanent, so it must not be retried");
    eq(r.sleeps, [], "no backoff on a permanent failure");
  });

  t("httpJson_ does not retry other 4xx errors", () => {
    const r = withHttp([{ body: "denied", code: 403 }], () =>
      httpJson_("GET", "https://example/x"),
    );
    eq(r.err.statusCode, 403);
    eq(r.calls.length, 1);
  });

  t("httpJson_ retries a 429 and returns the eventual success", () => {
    const r = withHttp(
      [
        { body: "slow down", code: 429 },
        { body: '{"ok":2}', code: 200 },
      ],
      () => httpJson_("GET", "https://example/x"),
    );
    eq(r.result, { ok: 2 });
    eq(r.calls.length, 2);
    eq(r.sleeps, [500], "one backoff before the retry");
  });

  t("httpJson_ retries a 500 and gives up after 4 attempts", () => {
    const r = withHttp([{ body: "boom", code: 500 }], () =>
      httpJson_("GET", "https://example/x"),
    );
    eq(r.calls.length, 4, "maxAttempts");
    eq(r.err.statusCode, 500);
    eq(r.sleeps, [500, 1000, 2000], "exponential backoff between attempts");
  });

  // The 5xx boundary: 599 is transient, 600 is not a server code.
  t("httpJson_ retries 599 but not 600", () => {
    eq(
      withHttp([{ body: "x", code: 599 }], () => httpJson_("GET", "u")).calls
        .length,
      4,
    );
    eq(
      withHttp([{ body: "x", code: 600 }], () => httpJson_("GET", "u")).calls
        .length,
      1,
    );
  });

  // A fetch that throws is a network fault: transient, so it retries.
  t("httpJson_ retries a thrown fetch and rethrows the last error", () => {
    const r = withHttp([{ throw: "socket closed" }], () =>
      httpJson_("GET", "https://example/x"),
    );
    eq(r.calls.length, 4);
    eq(/socket closed/.test(String(r.err)), true, String(r.err));
  });

  return results;
}

# Bug report: overlapping same-type sessions shadow each other instead of merging

## Summary

When two `STRENGTH_TRAINING` sessions from different sources overlap in time, the
Google Health app/AI displays only one of them and hides ("shadows") the other.
The shadowed session's data — including `exercise.notes` — never appears in the
card or to the in-app AI assistant, even though both datapoints exist and are
returned by the REST API.

## Use case

We sync strength-training reps/sets from a spreadsheet into Google Health via the
REST API (`users.dataTypes.dataPoints`, `STRENGTH_TRAINING`). Per-set data lives
in `exercise.notes` because the API has no structured field for it. On days where
a wearable also logged a Strength Training session for the same workout, the
device session carries HR/calories but no reps/sets, and our session carries
reps/sets but no HR. The user should see both.

## Observed behavior (probed 2026-06-03)

- Two same-type sessions overlapping by more than ~2–3 minutes → the app shows
  ONE card and hides the other. Overlap ≤ ~2 minutes (including adjacent/disjoint)
  → both render as separate cards.
- Which session survives is determined by SOURCE, not `recordingMethod`,
  `metricsSummary`, or start/creation order: a first-party/device session
  (`dataSource.application` = null) always shadows a third-party (OAuth client)
  session. Verified with our session placed both before and after the device
  session — the device survived in both cases.
- The shadowed datapoint still exists server-side: `GET` and `list` return it
  with its `notes` intact. Only the display/AI layer hides it.

## Impact

A third-party app cannot get its `STRENGTH_TRAINING` notes shown for a workout a
device also logged:

- We cannot write to the device's datapoint — exercise `PATCH` is a silent no-op,
  and cross-client writes return `403 DATA_POINT_NOT_OWNED_BY_CLIENT`.
- We cannot make our datapoint win the shared card (source priority).
- Avoiding the shadow requires placing our session at a non-overlapping time,
  which misrepresents when the workout occurred.

So user-entered data (reps/sets) is silently dropped from the UI whenever it
coincides with a device session.

## Proposed fix

When two same-type sessions overlap, MERGE their non-conflicting fields into the
single displayed card instead of shadowing one — e.g. show the device session's
HR/calories together with the overlapping session's `notes` (reps/sets). At
minimum, surface the shadowed session's `notes` on the surviving card so no
user-entered data is hidden.

## Environment

- Google Health REST API v4, `users.dataTypes.dataPoints`, `exercise` /
  `STRENGTH_TRAINING`.
- Reps/sets stored in `exercise.notes`.
- Related findings: exercise `PATCH` no-ops for all content fields;
  `metricsSummary` HR fields are not derived for app-created datapoints (only
  `caloriesKcal` is settable on create).

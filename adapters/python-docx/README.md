# python-docx adapter

Wraps [python-docx](https://python-docx.readthedocs.io/) behind adapter
protocol v1.

**What is library API vs glue:** `replaceFirstTextOccurrence` uses the
library's paragraph/run API for intra-run replacement; locating the first
matching paragraph is the only glue. A match spanning run boundaries is
declined (exit 2) instead of implementing a run-spanning rewrite in the
adapter — the suite measures the library, not the adapter author.

**Honest gaps:** `acceptAllTrackedChanges` / `rejectAllTrackedChanges`
report `unsupported` — python-docx has no revision API and does not expose
`w:ins`/`w:del` content in its object model. That asymmetry in the results
matrix is the point.

**Maintenance policy:** best-effort; the weekly CI run re-tests against the
latest PyPI release, so upstream breakage surfaces as `error` cells rather
than silent rot.

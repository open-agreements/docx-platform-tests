# Open XML SDK adapter

Wraps the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)
(`DocumentFormat.OpenXml`, MIT-licensed) behind adapter protocol v1.
This is Microsoft's own OOXML toolchain, so its column carries vendor-lineage
evidence: when a cell here disagrees with another implementation,
"Microsoft's library does X" is the strongest available signal short of
Word itself.

**What is library API vs glue:** the SDK is a strongly-typed DOM over the
package, not a document editor.

- `replaceFirstTextOccurrence`: opening the package, walking paragraphs in
  document order, and rewriting a single `w:t` are all library API; locating
  the first matching paragraph and the `w:t` containing the whole match is
  the only glue. A match spanning `w:t` boundaries is declined (exit 2)
  instead of implementing run-spanning matching in the adapter — the suite
  measures the library, not the adapter author.
- `acceptAllTrackedChanges` / `rejectAllTrackedChanges`: report
  `unsupported` — the SDK exposes `w:ins`/`w:del` as DOM nodes but provides
  no accept/reject operation; implementing the revision projections would be
  an adapter-side algorithm. A mostly-`unsupported` column is itself matrix
  signal: it documents what Microsoft's library does and does not give you.

**Maintenance policy:** best-effort. The package reference floats on the
latest `DocumentFormat.OpenXml` 3.x release and the weekly CI run rebuilds
from NuGet, so upstream breakage surfaces as `error` cells rather than
silent rot.

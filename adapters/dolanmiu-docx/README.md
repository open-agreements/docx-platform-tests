# dolanmiu/docx adapter

Evaluates [`dolanmiu/docx`](https://github.com/dolanmiu/docx) (`npm` package
`docx`) as an adapter-protocol v1 participant.

**What is library API vs glue:** this package is a document-generation library.
It can create DOCX content and includes `patchDocument` for replacing explicit
patch placeholders in an existing document, but it does not expose an editor API
for arbitrary paragraph-local literal search over an existing package.

- `replaceFirstTextOccurrence`: reports `unsupported`. The protocol requires
  locating the first literal occurrence in document order. Using `patchDocument`
  would require changing scenarios to contain placeholders, and locating
  arbitrary text across existing OOXML runs would be adapter-side search/mutation
  logic rather than glue.
- `acceptAllTrackedChanges` / `rejectAllTrackedChanges`: reports `unsupported`.
  `docx` can generate revision markup, but it has no public API that resolves
  existing revisions.

The column is intentionally mostly unsupported: it records that the dominant JS
generation library is not an existing-document editor under the suite's
glue-not-algorithms rule.

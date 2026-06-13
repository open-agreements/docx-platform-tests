# Scenario DSL, version 1.2

Changes in 1.2: composed scenarios can declare optional
`secondarySpecCitations` alongside the mandatory primary `specCitation`.

Changes in 1.1: outcome grading distinguishes serialization divergence from
conformance failure (the `pass-divergent` status — see "Outcome grading").
Manifests declaring `"dslVersion": "1.0"` or `"dslVersion": "1.1"` remain
valid.

A scenario is a directory under `scenarios/<group>/<scenarioId>/` containing:

| File | Role |
| --- | --- |
| `scenario.json` | The scenario manifest (this document specifies it). |
| `input/document.xml` | The source-of-truth input body. Reviewable in diffs. |
| `input.docx` | The interchange artifact adapters receive. Generated from `input/document.xml` by `npm run pack-fixtures` and committed; CI fails if it drifts from the fragment. |
| `expected/document.xml` | Present only when the scenario uses a `canonicalXmlEquals` assertion. |

Adapters consume full `.docx` packages, never bare XML: most implementations
(python-docx, LibreOffice, Word itself) can only open a complete package.
The committed fragment exists so reviewers diff ten lines of XML instead of
a zip file.

## scenario.json

All keys are deliberately self-disambiguating (3–4-word camelCase).

```json
{
  "dslVersion": "1.0",
  "scenarioId": "acceptInsertionsUnwrapsInsWrappers",
  "scenarioTitle": "Accepting insertions unwraps w:ins and keeps run content",
  "specCitation": {
    "standard": "ECMA-376",
    "edition": 5,
    "part": 1,
    "section": "17.13.5.18",
    "clauseTitle": "ins (Inserted Run Content)"
  },
  "secondarySpecCitations": [
    {
      "standard": "ECMA-376",
      "edition": 5,
      "part": 1,
      "section": "17.13.5.14",
      "clauseTitle": "del (Deleted Run Content)"
    }
  ],
  "wordBehaviorNote": null,
  "inputDocumentPath": "input.docx",
  "operationDescriptor": { "operationName": "acceptAllTrackedChanges" },
  "assertionList": [ ... ]
}
```

- `specCitation` is mandatory and is the primary clause the scenario is
  organized around. Every scenario must be defensible against the cited
  ECMA-376 clause, not against any implementation.
- `secondarySpecCitations` is optional. Use it when a composed scenario's
  assertions also derive from additional clauses (for example, nested
  `w:del` content inside a `w:ins` wrapper). Single-clause scenarios omit it,
  and the single-object `specCitation` form remains the DSL 1.x compatibility
  surface.
- `wordBehaviorNote` records cases where Microsoft Word's observed behavior
  deviates from a strict reading of the cited clause. When the two conflict,
  the suite treats documented Word behavior as canonical and the note says so
  explicitly. `null` when no deviation is known.

## Operations (v1.0 — closed enum)

| `operationName` | Parameters | Meaning |
| --- | --- | --- |
| `acceptAllTrackedChanges` | — | Accept every tracked revision in the document. |
| `rejectAllTrackedChanges` | — | Reject every tracked revision in the document. |
| `replaceFirstTextOccurrence` | `findText`, `replaceText` | Replace the first occurrence of `findText`, scanning paragraphs in document order. |

`replaceFirstTextOccurrence` match scope is **paragraph-local**: the first
paragraph (in document order) whose text contains `findText` is rewritten.
Matches that would span a paragraph boundary are out of scope for DSL 1.0 —
implementations whose primitives are paragraph-scoped would otherwise
legitimately diverge. The replacement is a plain edit, not a tracked change.

Adding an operation is a minor `dslVersion` bump; changing the semantics of
an existing one is a major bump.

## Assertions (v1.0)

Assertions are evaluated by the runner against the adapter's output
`word/document.xml`. Each assertion is reported individually, so one strict
assertion failing never masks a lenient one passing.

| `assertionKind` | Parameters | Evaluated against |
| --- | --- | --- |
| `xpathQueryCount` | `xpathExpression`, `expectedCount` | the raw output DOM |
| `xpathQueryExists` | `xpathExpression` | the raw output DOM (sugar for count ≥ 1) |
| `documentTextContainsAtOffset` | `expectedSubstring`, `expectedOffset` | the body text projection (below) |
| `canonicalXmlEquals` | `expectedDocumentPath` | the canonicalized output vs the canonicalized expected document |
| `schemaValidAgainstWml` | — | the output `word/document.xml` validated with `xmllint --schema` against `DPT_WML_SCHEMA_PATH` |

XPath expressions are evaluated with the prefix `w:` bound to
`http://schemas.openxmlformats.org/wordprocessingml/2006/main`.

`schemaValidAgainstWml` is an optional-tool assertion. The runner invokes
`xmllint --noout --schema "$DPT_WML_SCHEMA_PATH" <document.xml>` and expects
the schema path to be a WordprocessingML XSD entry point with imports
resolvable relative to that file. Set `DPT_XMLLINT_BIN` to override the
binary name/path. If a scenario declares this assertion without
`DPT_WML_SCHEMA_PATH`, the assertion fails with a setup diagnostic rather
than silently passing.

### Body text projection

`documentTextContainsAtOffset` offsets are character offsets into this exact
projection of the output document:

1. take every `w:p` under `w:body` in document order;
2. within each paragraph, concatenate the text content of every `w:t`
   element in document order — `w:delText` is **excluded**;
3. join paragraphs with a single `\n` (no trailing newline).

### Canonicalization (for `canonicalXmlEquals`)

Both documents are normalized before comparison:

- inter-element whitespace in element-only content is dropped;
- attributes are sorted by namespace URI + local name;
- namespace prefix spelling is ignored (elements compare by
  `{namespaceURI}localName`); redundant namespace declarations are dropped;
- the `rsid` attribute family (`w:rsidR`, `w:rsidRPr`, `w:rsidRDefault`,
  `w:rsidP`, `w:rsidDel`, …) and `w:id` on revision wrappers are stripped —
  these are implementation-chosen identifiers, not conformance signals;
- adjacent runs whose properties are identical are merged
  (`mergeAdjacentIdenticalRuns`): run splitting is legal in WordprocessingML
  and several implementations merge or split runs during edits.

Text content (including `xml:space="preserve"` whitespace) and element order
are **never** normalized.

### Assertion-strength rule

Prefer the weakest assertion that captures the conformance claim. As of
DSL 1.1, `canonicalXmlEquals` is a **serialization-audit assertion, not a
conformance assertion**: its failure alone grades a cell `pass-divergent`,
never `fail` (see "Outcome grading"). Every structural fact the cited
clause actually requires must therefore be expressed as a semantic
assertion (`xpathQueryCount` / `xpathQueryExists` /
`documentTextContainsAtOffset`) — a requirement carried only by
`canonicalXmlEquals` cannot fail an implementation. Use it to pin the
reference serialization for audit and regression visibility, always
alongside the semantic assertions that carry the claim.

### Outcome grading (v1.1)

A scenario cell is graded from the per-assertion results:

| Status | Condition |
| --- | --- |
| `pass` | every assertion passed |
| `pass-divergent` | every assertion **except** `canonicalXmlEquals` passed |
| `fail` | any non-`canonicalXmlEquals` assertion failed |

The conformance claim is carried by the semantic assertions (xpath and
text-projection); `canonicalXmlEquals` pins serialization granularity beyond
what the cited clause requires. WordprocessingML grants implementations real
serialization freedom — materializing formatting defaults (`pStyle`, `jc`,
empty `rPr`), regenerating section properties — that canonicalization
deliberately does not paper over. An implementation that satisfies the cited
clause but exercises that freedom on save is not failing the clause, and
labeling it `fail` would misreport the comparison; `pass-divergent` records
both facts. The per-assertion breakdown remains in `results/latest.json`, so
a `pass-divergent` cell is always auditable down to the differing canonical
forms.

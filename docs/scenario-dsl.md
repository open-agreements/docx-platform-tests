# Scenario DSL, version 1.3

Changes in 1.3: assertions may target a package part other than the main
document via an optional `assertedPart` selector (resolved by OPC
relationship, never by hardcoded part name); scenario XPath now binds the
`r:` prefix in addition to `w:`; a new `hyperlinkResolvesToExternalUrl`
assertion joins a hyperlink's display text to its relationship target. All
DSL <= 1.2 manifests remain valid — `assertedPart` defaults to the main
document part, reproducing prior behavior exactly.

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
| `input/{styles,numbering,comments,header-default,footer-default}.xml` | Optional sibling fragments packed as additional parts (below). Present only when a scenario needs a pre-existing styles/numbering/comments/header/footer part in its **input**. |
| `input.docx` | The interchange artifact adapters receive. Generated from the `input/` fragments by `npm run pack-fixtures` and committed; CI fails if it drifts from them. |
| `expected/document.xml` | Present only when the scenario uses a `canonicalXmlEquals` assertion. |

Adapters consume full `.docx` packages, never bare XML: most implementations
(python-docx, LibreOffice, Word itself) can only open a complete package.
The committed fragment exists so reviewers diff ten lines of XML instead of
a zip file.

### Multi-part input fixtures

Most scenarios pack a single-part document (`[Content_Types].xml`,
`_rels/.rels`, `word/document.xml`). A scenario that needs a pre-existing
non-main part in its input — a styles part defining `Heading1`, a comments
part its body anchors, a numbering part it references — adds the corresponding
sibling fragment under `input/`. The packer recognizes a **closed set**:

| Fragment | Packaged part | Relationship `Type` | Stable relationship `Id` |
| --- | --- | --- | --- |
| `input/styles.xml` | `word/styles.xml` | …/relationships/styles | `rId1` |
| `input/numbering.xml` | `word/numbering.xml` | …/relationships/numbering | `rId2` |
| `input/comments.xml` | `word/comments.xml` | …/relationships/comments | `rId3` |
| `input/header-default.xml` | `word/header-default.xml` | …/relationships/header | `rId4` |
| `input/footer-default.xml` | `word/footer-default.xml` | …/relationships/footer | `rId5` |

When any sibling is present the packer emits a `word/_rels/document.xml.rels`
(with one relationship per active fragment) and a `[Content_Types].xml`
Override per part. Each relationship id is **stable per part type** regardless
of which other fragments are present: a fixture author wiring a header writes
`<w:headerReference w:type="default" r:id="rId4"/>` into `input/document.xml`
and knows the packer emits the matching `rId4` relationship. Styles, numbering,
and comments are linked purely by relationship `Type` — their id is never
referenced from the body — so assertions resolve them with
`relationshipFromMainPart`, and headers/footers with
`headerReference`/`footerReference` (see "Asserted part" below).

The set is deliberately closed: only the `default` header/footer variant is
packed. The `first` and `even` variants (and additional sections) are
intentionally unsupported until a scenario needs them — adding one is a new
slot in `FIXTURE_PART_SLOTS`, not an open-ended convention.

With **no** siblings the package is byte-identical to the historical
three-entry form, so pre-existing fixtures never drift. `npm run
pack-fixtures --check` recomputes every package from its `input/` fragments and
compares the **decompressed part content** against the committed `input.docx`
(not raw zip bytes, which would couple the check to the build platform's
compression output), so drift in any part — the main document, any sibling,
`[Content_Types].xml`, `word/_rels/document.xml.rels`, or an added/removed
fragment — fails CI. Unrecognized `input/*.xml` files are rejected outright.

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
| `hyperlinkResolvesToExternalUrl` | `hyperlinkDisplayText`, `expectedTargetUrl` | a `w:hyperlink` whose display text equals `hyperlinkDisplayText` and whose `@r:id` resolves, in the main part's relationships, to an external target equal to `expectedTargetUrl` |

XPath expressions are evaluated with two bound prefixes: `w:` for
`http://schemas.openxmlformats.org/wordprocessingml/2006/main` and `r:` for
`http://schemas.openxmlformats.org/officeDocument/2006/relationships` (the
namespace of `@r:id` on hyperlinks and header/footer references). Relationship
ids, comment ids, and numbering ids are implementation-chosen, so assertions
never hardcode them; where a claim spans two parts (a reference resolving to a
target), a purpose-built assertion such as `hyperlinkResolvesToExternalUrl`
performs the id join inside the runner.

### Asserted part (`assertedPart`)

`xpathQueryCount`, `xpathQueryExists`, and `schemaValidAgainstWml` accept an
optional `assertedPart` selecting which package part to evaluate against.
Omitting it (or `{"partResolution":"mainDocumentPart"}`) targets
`word/document.xml`. `documentTextContainsAtOffset` and `canonicalXmlEquals`
are always evaluated against the main document part.

| `partResolution` | Extra fields | Resolves to |
| --- | --- | --- |
| `mainDocumentPart` | — | `word/document.xml` (the default) |
| `relationshipFromMainPart` | `relationshipTypeUri` | the single part reached from `word/_rels/document.xml.rels` by relationship `Type` (styles, numbering, comments) |
| `headerReference` | `headerReferenceType` | main-document `sectPr/headerReference[@w:type=…]` → its `@r:id` → relationship target |
| `footerReference` | `footerReferenceType` | main-document `sectPr/footerReference[@w:type=…]` → its `@r:id` → relationship target |

Part names (`styles.xml`, `header1.xml` vs `header3.xml`, …) are
implementation-chosen, so resolution is always by relationship, never by path.
Cardinality for `relationshipFromMainPart` is singleton-by-type: **zero**
matches fails the assertion with a "no part with relationship type …"
diagnostic (that failure is itself the conformance signal — e.g. "no comments
part was written"), and **more than one** match fails with a diagnostic rather
than guessing. `headerReference`/`footerReference` assume the document's single
section and resolve the first matching reference in document order; the
reference's `@r:id` must join to a relationship of the matching header/footer
Type (a reference wired to, say, a styles relationship is malformed and does
not resolve). A resolution failure is reported as an ordinary
(non-`canonicalXmlEquals`) assertion failure, so it participates in grading
like any semantic assertion.

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

1. take every `w:p` under `w:body` in document order — this includes
   paragraphs nested inside table cells (`w:tbl/w:tr/w:tc/w:p`), which appear
   at their document-order position; anything outside `w:body` is excluded;
2. within each paragraph, concatenate the text content of every `w:t`
   element in document order — `w:delText` is **excluded**;
3. join paragraphs with a single `\n` (no trailing newline).

Because empty and table-cell paragraphs contribute lines, prefer
`xpathQueryCount`/`xpathQueryExists` scoped to `w:tbl/w:tr/w:tc` over
`documentTextContainsAtOffset` when asserting table-cell content.

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

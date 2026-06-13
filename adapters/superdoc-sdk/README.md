# superdoc-sdk adapter

Adapter for `@superdoc-dev/sdk`, the headless SuperDoc Document API.

The SDK supplies the document semantics:

- `query.match` locates the first literal text match.
- `replace` mutates the matched text range.
- `trackChanges.list` enumerates existing revision marks.
- `trackChanges.decide` accepts or rejects each tracked change.
- `save` exports the mutated `.docx`.

Adapter glue is limited to protocol parsing, temporary SDK state isolation,
and DOCX package normalization. The suite's fixture packages are intentionally
minimal, while SuperDoc expects Word-style package support parts. Before
opening a document, the adapter copies the input to a temp file and adds only
missing `word/_rels/document.xml.rels`, `word/styles.xml`, and the styles
content type override needed for SuperDoc import/export.

`@superdoc-dev/sdk` is published as AGPL-3.0, so this optional adapter package
declares `AGPL-3.0` even though the top-level conformance suite is Apache-2.0.
It is installed only when its `install.sh` is run.

The suite protocol has bulk tracked-change operations. The SDK exposes
revision semantics as `trackChanges.list` plus one `trackChanges.decide` call
per SDK-reported tracked-change id, so the adapter bridges the protocol by
iterating those SDK ids. It does not inspect or rewrite revision XML to decide
what accepting or rejecting means.

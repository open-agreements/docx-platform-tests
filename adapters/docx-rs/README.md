# docx-rs adapter

Adapter for [`docx-rs`](https://github.com/bokuweb/docx-rs), pinned to
`docx-rs = 0.4.20`.

`docx-rs` is primarily a Rust/WebAssembly `.docx` writer. Its public API
includes builders for producing OOXML, revision-related data types, and read
helpers for package/XML data, but it does not expose a public operation that
accepts or rejects existing tracked changes. It also does not expose an
existing-document text replacement API that can satisfy this suite's arbitrary
literal replacement operation without adapter-side XML editing.

For protocol v1, the adapter therefore reports all current operations as
`unsupported`:

- `acceptAllTrackedChanges`
- `rejectAllTrackedChanges`
- `replaceFirstTextOccurrence`

The Cargo binary intentionally depends on `docx-rs` so CI verifies the pinned
crate still builds. `run-adapter.sh` executes that binary after `install.sh`
builds it.

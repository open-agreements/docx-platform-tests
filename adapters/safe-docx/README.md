# safe-docx adapter

Wraps the `safe-docx-conformance-adapter` bin shipped by
[`@usejunior/docx-core`](https://github.com/UseJunior/safe-docx)
(`packages/docx-core/src/cli/conformance-adapter.ts`).

**What is library API vs glue:** everything is library API — the bin itself
implements adapter protocol v1 natively and this directory only installs it
(`install.sh`). No adapter-side algorithms.

**Install source:** a pinned-SHA source build (`safe-docx.pin.json`) until
the bin lands in a published npm release; then this switches to the public
registry (issue #3).

**Maintenance policy:** maintained by the safe-docx project; safe-docx's own
CI runs this suite as a self-check, so breakage surfaces there first.

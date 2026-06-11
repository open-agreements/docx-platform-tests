# safe-docx adapter

Wraps the `safe-docx-conformance-adapter` bin shipped by
[`@usejunior/docx-core`](https://github.com/UseJunior/safe-docx)
(`packages/docx-core/src/cli/conformance-adapter.ts`).

**What is library API vs glue:** everything is library API — the bin itself
implements adapter protocol v1 natively and this directory only installs it
(`install.sh`). No adapter-side algorithms.

**Install source:** a source build from the tip of safe-docx `main`
(`safe-docx.pin.json`, `trackingBranchName`), so library changes are
re-tested by the suite without waiting for an npm release (supersedes the
npm switch planned in issue #3). `install.sh` resolves the branch to a
concrete commit and records it in `build-info.json`; the matrix reports the
adapter version as `<package version>+git.<commit>`. To reproduce a past
matrix, set `pinnedCommitSha` in `safe-docx.pin.json` — it takes precedence
over the tracking branch.

**Maintenance policy:** maintained by the safe-docx project; safe-docx's own
CI runs this suite as a self-check, so breakage surfaces there first.

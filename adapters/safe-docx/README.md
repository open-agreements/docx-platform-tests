# safe-docx adapter

Wraps the `safe-docx-conformance-adapter` bin shipped by
[`@usejunior/docx-core`](https://github.com/UseJunior/safe-docx)
(`packages/docx-core/src/cli/conformance-adapter.ts`).

**What is library API vs glue:** everything is library API — the bin itself
implements adapter protocol v1 natively and this directory only installs it
(`install.sh`). No adapter-side algorithms.

**Install source:** a source build from safe-docx (`safe-docx.pin.json`).
Scheduled and main-push suite runs resolve `trackingBranchName` (safe-docx
`main`) so library changes are re-tested without waiting for an npm release
(supersedes the npm switch planned in issue #3). Pull-request suite runs use
`pullRequestPinnedCommitSha` instead, keeping required checks hermetic when
safe-docx `main` moves. `install.sh` records the concrete commit in
`build-info.json`; the matrix reports the adapter version as
`<package version>+git.<commit>`. To reproduce a past matrix for every event,
set `pinnedCommitSha` in `safe-docx.pin.json` — it takes precedence over both
the tracking branch and the pull-request pin.

**Maintenance policy:** maintained by the safe-docx project; safe-docx's own
CI runs this suite as a self-check, so breakage surfaces there first.

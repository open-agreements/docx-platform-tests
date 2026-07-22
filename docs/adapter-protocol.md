# Adapter protocol, version 1

An adapter is a command-line program that applies one operation to one
document. The runner invokes it once per scenario:

```
<adapter-command> --protocol-version 1 \
  --operation <path/to/operation.json> \
  --input <path/to/input.docx> \
  --output <path/to/output.docx>
```

- `operation.json` is the scenario's `operationDescriptor` object, verbatim,
  written to a temporary file by the runner. The adapter never sees the
  scenario's assertions or expected output — it must not be able to teach to
  the test.
- `input.docx` is a complete WordprocessingML package.
- On success the adapter writes a complete package to the `--output` path.
  Assertions may inspect the main document part and related parts such as
  styles, numbering, comments, headers, footers, settings, content types, and
  relationships, so the
  output must be a loadable package with all required relationships intact.

## Exit codes

| Code | Meaning | Adapter obligations |
| --- | --- | --- |
| 0 | success | `--output` written |
| 1 | error | crash/exception; write diagnostics to stderr |
| 2 | **unsupported operation** | print a one-line reason to stdout; write no output. The wpt `NOTRUN` analog: declining honestly is always preferable to approximating. |
| 3 | protocol-version mismatch | the adapter does not speak the requested `--protocol-version` |

Any other exit code is recorded as `error`.

## Registration

Add the adapter under `adapters/<adapterName>/` with a README describing
what is library API and what is adapter glue, then register it in
`registry/adapters.json`:

```json
{
  "adapterName": "python-docx",
  "adapterCommand": ["python3", "adapters/python-docx/run_adapter.py"],
  "adapterVersionCommand": ["python3", "-c", "import docx; print(docx.__version__)"]
}
```

The declaration is documentation; at runtime the adapter's exit code is
authoritative (an adapter may declare an operation and still exit 2 for a
specific input it cannot handle).

## Published results contract

The runner writes `results/latest.json`, which is also published to GitHub
Pages. Consumers should treat the top-level `schemaVersion` as the results
format version and validate against `results/results.schema.json` (draft
2020-12), published alongside the data. CI checks that the committed schema is
current and validates every emitted `latest.json` before publication.
Each scenario result declares `oracleKind`. ECMA conformance uses `pass`,
`pass-divergent`, and `fail`; suite-declared metamorphic preservation evidence
uses `invariant-pass` and `invariant-fail` and is not an ECMA failure signal.

Schema-v2 snapshots did not carry `oracleKind`. The command
`npm run migrate-results -- --input <v2.json> --output <v3.json>` performs the
deterministic migration. It
accepts only the immutable snapshot vendored under `results/history/` and
authenticated by raw-fixture and canonical JSON SHA-256 digests, ordered
adapter/scenario identities, and cardinalities in
`registry/results-history.json`. After authentication it
infers `ecma-conformance` only when the current mapping remains pure normative
evidence. Missing, extra, duplicate, reordered, altered, or invariant-injected
rows are rejected rather than guessed. The history entry records the source
commit from which the fixture was vendored.

## Versioning

The protocol version is asserted by the adapter, not negotiated. Breaking
changes to flags, exit codes, or package expectations bump the version, and
the runner passes the version each scenario requires.

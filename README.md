# docx-platform-tests

A cross-implementation conformance test suite for WordprocessingML (`.docx`),
in the spirit of [web-platform-tests](https://github.com/web-platform-tests/wpt)
and [wpt.fyi](https://wpt.fyi/results/).

Every scenario in this suite asserts behavior that is derivable from
[ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
(Office Open XML), cited by edition, part, and section — never from any one
library's internal algorithm. Implementations participate through small
adapter binaries; the runner compares each adapter's output against the
scenario's assertions and publishes a results matrix.

## Scope

WordprocessingML conformance only, currently:

- tracked changes (`w:ins` / `w:del` accept and reject semantics),
- find-replace over run text,
- schema validity (assertion kind defined, runner support deferred — see
  `docs/scenario-dsl.md`).

SpreadsheetML and PresentationML are out of scope.

An honest `unsupported` outcome is part of the design: an implementation
with no tracked-changes API reports `unsupported` (exit code 2) for those
scenarios rather than failing, and the matrix shows the gap. That asymmetry
is information, not noise.

## Layout

```
docs/        scenario DSL and adapter protocol specifications (versioned)
scenarios/   one directory per scenario: scenario.json + input/ + expected/
adapters/    one directory per registered implementation adapter
runner/      the neutral runner (Node, @xmldom/xmldom; no implementation deps)
registry/    adapters.json — registered adapters and their invocation commands
results/     latest published results snapshot (also published to gh-pages)
```

## Running the suite

```bash
cd runner
npm ci
npm run check-fixtures   # verify input.docx packages match input/document.xml
npm run suite            # run all registered adapters, write ../results/latest.json
```

## Adding an implementation

Write an adapter that satisfies `docs/adapter-protocol.md` (a CLI taking
`--protocol-version 1 --operation operation.json --input input.docx
--output output.docx`), add it under `adapters/<name>/`, and register it in
`registry/adapters.json`. Adapters must decline operations they cannot
perform with exit code 2 — never approximate.

## Adding a scenario

A scenario must cite the ECMA-376 section its assertion derives from. Where
Microsoft Word's observed behavior deviates from a strict reading of the
spec, record it in the scenario's `wordBehaviorNote`. See
`docs/scenario-dsl.md` for the format and the assertion-strength rules.

## License

[Apache-2.0](LICENSE). ECMA-376 is cited by section number only; no spec
text is reproduced in this repository.

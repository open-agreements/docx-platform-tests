# Contributing

## Scenarios

A conformance scenario is accepted when its assertion is derivable from a
cited ECMA-376 clause — not from any particular implementation. A pure
metamorphic scenario may test a narrowly declared preservation property, but
it is reported and aggregated separately from ECMA conformance.
Requirements:

- `specCitation` names edition, part, and section; the PR description should
  quote nothing from the spec (cite by number only) but must explain the
  derivation.
- Do not mix conformance and metamorphic assertions in one scenario. Split
  them so invariant loss cannot grade an ECMA failure or contaminate another
  capability's evidence.
- Follow the assertion-strength rule in `docs/scenario-dsl.md`: weakest
  assertion that captures the claim; `canonicalXmlEquals` always paired with
  weaker assertions.
- Run `npm run pack-fixtures` in `runner/` after editing any
  `input/document.xml`; CI rejects drifted packages.
- Where Microsoft Word's observed behavior deviates from a strict reading of
  the clause, document it in `wordBehaviorNote` — the suite treats
  documented Word behavior as canonical in conflicts.

## Adapters

See `docs/adapter-protocol.md`. The bar for adapter code: glue, not
algorithms. If implementing an operation requires writing a non-trivial
algorithm the library does not provide, the honest answer is exit code 2
(`unsupported`) — the matrix exists to show that gap. Each adapter README
must state what is library API and what is glue, and name a maintenance
policy (best-effort is fine; the matrix shows staleness honestly).

Adapters install their own dependencies via an `install.sh` in their
directory (invoked by CI) and must not add dependencies to the runner.

## Runner

The runner stays neutral: it must not depend on any library that is itself
a registered implementation.

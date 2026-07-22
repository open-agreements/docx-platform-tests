# Capability registry

The capability registry gives neutral, stable names to the behavior exercised
by conformance scenarios. It is an index over evidence, not a declaration that
an implementation supports a feature forever.

## Files

- `registry/capabilities.json` defines capability identities, normative
  citations, dependencies, applicable axes, and package-part scope.
- `registry/scenario-capabilities.json` maps scenarios to capability axes and
  identifies the oracle classes used by each scenario.
- `registry/profiles.json` defines unweighted selections of capabilities and
  axes. Product-specific priority weights do not belong in the neutral suite.
- `registry/scenario-coverage.json` is generated coverage data.
- `results/capability-summary.json` aggregates the committed result snapshot
  while preserving raw scenario IDs, adapter versions, and denominators.

The three author-written files have JSON Schemas beside them. Generated files
must be updated with:

```bash
cd runner
npm run write-capability-index
```

CI runs `check-capability-index`, which validates schemas and rejects:

- duplicate or unknown capability IDs;
- missing, self-referential, or cyclic dependencies;
- unknown or inapplicable axes;
- unknown scenario mappings;
- scenarios without a mapping;
- scenario citations absent from their mapped capabilities;
- assertion/oracle mismatches, such as `canonicalXmlEquals` without a
  `serialization-specific` classification;
- unknown capability IDs in profiles; and
- stale generated coverage or result summaries.

## Axes

The registry keeps these dimensions independent:

| Axis | Question |
| --- | --- |
| `detect` | Can the implementation identify the construct? |
| `preserve` | Does unrelated work retain it? |
| `parse` | Can the implementation expose its semantics? |
| `validate` | Can malformed forms be rejected? |
| `generate` | Can the construct be created? |
| `edit` | Can it be intentionally modified? |
| `compare` | Can comparison represent the change? |
| `acceptReject` | Can revision resolution preserve its semantics? |
| `wordRoundtrip` | Has a Microsoft Word round-trip been exercised? |
| `crossPlatform` | Is behavior measured across registered adapters? |

An applicable axis with no mapped scenario remains visible as uncovered. This
prevents a broad feature label from implying that every operation is tested.
The scenario index does not infer `crossPlatform` evidence from eligibility to
run. The result aggregation derives that axis only for scenarios present in the
versioned result snapshot with outcomes from at least two registered adapters.

## Oracle classes

Scenario mappings distinguish:

- `normative-schema`;
- `normative-prose`;
- `normative-microsoft-extension`;
- `metamorphic-invariant`;
- `observed-word-behavior`;
- `cross-implementation-evidence`; and
- `serialization-specific`.

`normative-microsoft-extension` requires a structured `MS-DOCX` citation on
both the scenario and mapped capability. Observed Word behavior is not silently promoted to an ECMA requirement, and a
canonical serialization check does not turn valid XML freedom into a semantic
failure.

An ECMA-376 Part 3 citation can provide MCE context for a
`metamorphic-invariant` without making preservation of ignorable foreign
markup a normative ECMA oracle. In particular, an application that does not
understand ignorable markup may discard it. A preservation scenario therefore
asserts only its declared sentinels and classifies that evidence as
`metamorphic-invariant`, not `normative-prose`. Normative SDT structure and
opaque preservation belong in separate scenarios.

Mixed metamorphic/conformance scenarios are rejected. Result aggregation
groups by capability, axis, and `oracleKind`, so invariant failures cannot
enter conformance denominators or another capability's mapped evidence.

### Oracle-purity migration

An audit of every pre-existing mapping for the oracle-aware results contract
found two mixed scenarios. They were split without discarding their
preservation-oriented IDs:

| Historical scenario ID | Retained meaning | Normative companion |
| --- | --- | --- |
| `appendParagraphPreservesExistingContent` | prior paragraph preservation invariant | `appendParagraphAddsTrailingParagraph` |
| `replaceTextInsideTableCellPreservesStructure` | table and neighboring-cell preservation invariant | `replaceTextInsideTableCellUpdatesTargetText` |

The companions run the same operations against equivalent fixtures. This
keeps historical links stable while ensuring each result has one oracle kind.

## Repository boundary

This repository owns neutral capability definitions, neutral scenario mappings,
adapter protocols, and time-stamped results. Implementation repositories own
source paths, internal tests, formal proofs, conformance gaps, product priority
profiles, and release claims.

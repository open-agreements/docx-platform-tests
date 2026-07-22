import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { isDeepStrictEqual } from 'node:util';
import { Ajv2020 } from 'ajv/dist/2020.js';
import {
  buildCapabilitySummary,
  buildScenarioCoverage,
  loadAndValidateCapabilityRegistry,
  validateCapabilityRelationships,
  type CapabilityProfiles,
  type CapabilityRegistry,
  type ScenarioCapabilityRegistry,
} from './capability-registry.js';
import {
  loadAllScenarios,
  REPO_ROOT,
  validateScenarioDslCompatibility,
} from './scenarios.js';
import type { ResultsDocument } from './types.js';
import { gradeAssertionResults } from './oracle-grading.js';
import { RESULTS_SCHEMA } from './results-schema.js';
import { canonicalResultsDigest, normalizeResultsDocument } from './normalize-results.js';
import { renderMatrixHtml } from './render-matrix.js';

let failed = false;

function check(label: string, condition: boolean): void {
  console.log(`${condition ? 'PASS' : 'FAIL'} ${label}`);
  if (!condition) failed = true;
}

function expectThrows(label: string, expected: RegExp, action: () => void): void {
  try {
    action();
    check(label, false);
  } catch (error) {
    check(label, error instanceof Error && expected.test(error.message));
  }
}

const loaded = loadAndValidateCapabilityRegistry();
const scenarios = loadAllScenarios();
const registry = loaded.registry;
const mappings: ScenarioCapabilityRegistry = {
  schemaVersion: 1,
  registryVersion: registry.registryVersion,
  mappings: loaded.mappings,
};
const profiles: CapabilityProfiles = {
  schemaVersion: 1,
  registryVersion: registry.registryVersion,
  profiles: loaded.profiles,
};

check('committed registry validates', loaded.registry.capabilities.length === 23);
check('every committed scenario is mapped', new Set(loaded.mappings.map((mapping) => mapping.scenarioId)).size === scenarios.length);
const understatedDslScenario = structuredClone(
  scenarios.find(
    (scenario) => scenario.manifest.scenarioId === 'appendParagraphPreservesExistingContent'
  )!.manifest
);
understatedDslScenario.dslVersion = '1.7';
expectThrows('assertion kind newer than declared DSL is rejected', /requires DSL 1\.8/, () =>
  validateScenarioDslCompatibility(understatedDslScenario)
);
const auditedInvariantScenarioIds = [...new Set(
  loaded.mappings
    .filter((mapping) => mapping.oracleClasses.includes('metamorphic-invariant'))
    .map((mapping) => mapping.scenarioId)
)].sort();
check(
  'pre-existing mapping audit retains only declared pure invariants',
  isDeepStrictEqual(auditedInvariantScenarioIds, [
    'appendParagraphPreservesExistingContent',
    'replaceTextInsideTableCellPreservesStructure',
    'unrelatedTextEditPreservesOpaqueInlineContentControl',
  ])
);

const withoutScenario = structuredClone(mappings);
const removedScenarioId = withoutScenario.mappings[0].scenarioId;
withoutScenario.mappings = withoutScenario.mappings.filter((mapping) => mapping.scenarioId !== removedScenarioId);
expectThrows('unmapped scenario is rejected', /has no capability mapping/, () =>
  validateCapabilityRelationships(registry, withoutScenario, profiles, scenarios)
);

const unknownDependency = structuredClone(registry) as CapabilityRegistry;
unknownDependency.capabilities[0].dependencies.push('word.unknown.capability');
expectThrows('unknown dependency is rejected', /unknown dependency/, () =>
  validateCapabilityRelationships(unknownDependency, mappings, profiles, scenarios)
);

const cyclic = structuredClone(registry) as CapabilityRegistry;
cyclic.capabilities.find((capability) => capability.id === 'word.paragraphs.structure')!.dependencies = [
  'word.runs.formatting',
];
expectThrows('dependency cycle is rejected', /dependency cycle/, () =>
  validateCapabilityRelationships(cyclic, mappings, profiles, scenarios)
);

const wrongAxis = structuredClone(mappings);
wrongAxis.mappings[0].axis = 'wordRoundtrip';
expectThrows('inapplicable axis is rejected', /inapplicable axis/, () =>
  validateCapabilityRelationships(registry, wrongAxis, profiles, scenarios)
);

const wrongCitation = structuredClone(mappings);
const commentMapping = wrongCitation.mappings.find(
  (mapping) => mapping.scenarioId === 'commentPartRecordsAuthorInitialsAndText'
)!;
commentMapping.capabilityId = 'word.paragraphs.structure';
expectThrows('citation absent from mapped capability is rejected', /has no shared citation/, () =>
  validateCapabilityRelationships(registry, wrongCitation, profiles, scenarios)
);

const unrelatedExtraMapping = structuredClone(mappings);
unrelatedExtraMapping.mappings.push({
  scenarioId: 'commentPartRecordsAuthorInitialsAndText',
  capabilityId: 'word.paragraphs.structure',
  axis: 'edit',
  oracleClasses: ['normative-prose'],
});
expectThrows('each mapping must share a scenario citation', /has no shared citation/, () =>
  validateCapabilityRelationships(registry, unrelatedExtraMapping, profiles, scenarios)
);

const missingExtensionTraceability = structuredClone(scenarios);
const compatibilityScenario = missingExtensionTraceability.find(
  (scenario) => scenario.manifest.scenarioId === 'composeCompatibilityMode15WritesCompatSetting'
)!;
delete compatibilityScenario.manifest.microsoftExtensionCitations;
expectThrows(
  'Microsoft extension oracle without scenario citation is rejected',
  /no shared Microsoft extension citation/,
  () =>
    validateCapabilityRelationships(
      registry,
      mappings,
      profiles,
      missingExtensionTraceability
    )
);

const missingExtensionClassification = structuredClone(mappings);
const compatibilityMapping = missingExtensionClassification.mappings.find(
  (mapping) => mapping.scenarioId === 'composeCompatibilityMode15WritesCompatSetting'
)!;
compatibilityMapping.oracleClasses = compatibilityMapping.oracleClasses.filter(
  (oracleClass) => oracleClass !== 'normative-microsoft-extension'
);
expectThrows(
  'Microsoft extension citation without oracle classification is rejected',
  /without normative-microsoft-extension classification/,
  () =>
    validateCapabilityRelationships(
      registry,
      missingExtensionClassification,
      profiles,
      scenarios
    )
);

const missingSerializationOracle = structuredClone(mappings);
for (const mapping of missingSerializationOracle.mappings.filter(
  (candidate) => candidate.scenarioId === 'acceptInsertionsUnwrapsInsWrappers'
)) {
  mapping.oracleClasses = ['normative-prose'];
}
expectThrows('canonical XML without serialization oracle is rejected', /serialization-specific/, () =>
  validateCapabilityRelationships(registry, missingSerializationOracle, profiles, scenarios)
);

const mixedMetamorphicOracle = structuredClone(mappings);
const appendInvariant = mixedMetamorphicOracle.mappings.find(
  (mapping) => mapping.scenarioId === 'appendParagraphPreservesExistingContent'
)!;
appendInvariant.oracleClasses.push('normative-prose');
expectThrows('mixed metamorphic/conformance scenario is rejected', /mixes metamorphic-invariant/, () =>
  validateCapabilityRelationships(registry, mixedMetamorphicOracle, profiles, scenarios)
);

check(
  'metamorphic assertion failure grades invariant-fail, not ECMA fail',
  gradeAssertionResults(
    [{ assertionKind: 'xpathQueryCount', passed: false, detail: 'sentinel missing' }],
    'metamorphic-invariant'
  ) === 'invariant-fail'
);
check(
  'metamorphic assertion success grades invariant-pass',
  gradeAssertionResults(
    [{ assertionKind: 'xpathQueryCount', passed: true, detail: 'sentinel retained' }],
    'metamorphic-invariant'
  ) === 'invariant-pass'
);
check(
  'metamorphic canonical assertion failure grades invariant-fail',
  gradeAssertionResults(
    [{ assertionKind: 'canonicalXmlEquals', passed: false, detail: 'serialization changed' }],
    'metamorphic-invariant'
  ) === 'invariant-fail'
);

const unknownProfileCapability = structuredClone(profiles);
unknownProfileCapability.profiles[0].capabilityIds.push('word.unknown.capability');
expectThrows('unknown profile capability is rejected', /references unknown capability/, () =>
  validateCapabilityRelationships(registry, mappings, unknownProfileCapability, scenarios)
);

const coverage = buildScenarioCoverage(loaded) as {
  totals: { scenarios: number; mappings: number };
  unclassifiedScenarioIds: string[];
  capabilities: Array<{ capabilityId: string; axes: Array<{ axis: string; scenarioIds: string[] }> }>;
};
check('coverage retains raw scenario denominator', coverage.totals.scenarios === scenarios.length);
check('coverage retains raw mapping denominator', coverage.totals.mappings === loaded.mappings.length);
check('coverage has no unclassified scenarios', coverage.unclassifiedScenarioIds.length === 0);
check(
  'scenario registry does not claim cross-platform measurement',
  coverage.capabilities.every((capability) => {
    const axis = capability.axes.find((candidate) => candidate.axis === 'crossPlatform');
    return !axis || axis.scenarioIds.length === 0;
  })
);

const results = JSON.parse(
  readFileSync(join(REPO_ROOT, 'results', 'latest.json'), 'utf8')
) as ResultsDocument;
const resultsHistory = JSON.parse(
  readFileSync(join(REPO_ROOT, 'registry', 'results-history.json'), 'utf8')
);
const historicalV2 = JSON.parse(
  readFileSync(
    join(REPO_ROOT, resultsHistory.legacySchemas['2'].fixturePath),
    'utf8'
  )
);
check(
  'actual historical schema-v2 fixture matches its recorded digest',
  canonicalResultsDigest(historicalV2) === resultsHistory.legacySchemas['2'].canonicalSha256
);
check(
  'actual historical schema-v2 fixture normalizes deterministically',
  isDeepStrictEqual(normalizeResultsDocument(historicalV2, loaded), results)
);
const missingLegacyRow = structuredClone(historicalV2);
missingLegacyRow.results.pop();
expectThrows(
  'schema-v2 migration rejects a missing row',
  /scenario cardinality/,
  () => normalizeResultsDocument(missingLegacyRow, loaded)
);
const extraLegacyRow = structuredClone(historicalV2);
extraLegacyRow.results.push(structuredClone(extraLegacyRow.results[0]));
expectThrows(
  'schema-v2 migration rejects an extra row',
  /scenario cardinality/,
  () => normalizeResultsDocument(extraLegacyRow, loaded)
);
const duplicateLegacyRow = structuredClone(historicalV2);
duplicateLegacyRow.results[1] = structuredClone(duplicateLegacyRow.results[0]);
expectThrows(
  'schema-v2 migration rejects a duplicate row identity',
  /ordered scenario identities/,
  () => normalizeResultsDocument(duplicateLegacyRow, loaded)
);
const reorderedLegacyRows = structuredClone(historicalV2);
[reorderedLegacyRows.results[0], reorderedLegacyRows.results[1]] = [
  reorderedLegacyRows.results[1],
  reorderedLegacyRows.results[0],
];
expectThrows(
  'schema-v2 migration rejects reordered rows',
  /ordered scenario identities/,
  () => normalizeResultsDocument(reorderedLegacyRows, loaded)
);
const alteredLegacyRow = structuredClone(historicalV2);
alteredLegacyRow.results[0].scenarioTitle = 'Altered historical title';
expectThrows(
  'schema-v2 migration rejects altered row payload',
  /snapshot digest/,
  () => normalizeResultsDocument(alteredLegacyRow, loaded)
);
const injectedInvariantRow = structuredClone(historicalV2);
injectedInvariantRow.results[0].scenarioId = 'appendParagraphPreservesExistingContent';
expectThrows(
  'schema-v2 migration rejects injected invariant evidence',
  /ordered scenario identities/,
  () => normalizeResultsDocument(injectedInvariantRow, loaded)
);
const summary = buildCapabilitySummary(loaded, results) as {
  capabilities: Array<{
    axis: string;
    outcomes: Record<string, { denominator: number }>;
  }>;
};
const firstSummary = JSON.stringify(summary);
const secondSummary = JSON.stringify(buildCapabilitySummary(loaded, results));
check('capability aggregation is deterministic', firstSummary === secondSummary);
check('capability aggregation retains scenario IDs', firstSummary.includes('scenarioIds'));
check('capability aggregation retains denominators', firstSummary.includes('denominator'));
check(
  'cross-platform evidence is derived only from measured results',
  summary.capabilities.some((capability) => capability.axis === 'crossPlatform')
);
check(
  'result aggregation has no zero-denominator rows',
  summary.capabilities.every((capability) =>
    Object.values(capability.outcomes).every((outcome) => outcome.denominator > 0)
  )
);

const oracleResults = structuredClone(results);
oracleResults.implementations.push({
  adapterName: 'oracle-test-adapter',
  adapterVersion: 'synthetic',
});
oracleResults.results.push(
  {
    scenarioGroup: 'content-controls',
    scenarioId: 'unrelatedTextEditPreservesInlineContentControlStructure',
    scenarioTitle: 'normative structure',
    oracleKind: 'ecma-conformance',
    specCitation: scenarios.find(
      (scenario) =>
        scenario.manifest.scenarioId ===
        'unrelatedTextEditPreservesInlineContentControlStructure'
    )!.manifest.specCitation,
    outcomes: {
      'safe-docx': { status: 'pass' },
      'oracle-test-adapter': { status: 'pass' },
    },
  },
  {
    scenarioGroup: 'content-controls',
    scenarioId: 'unrelatedTextEditPreservesOpaqueInlineContentControl',
    scenarioTitle: 'opaque invariant',
    oracleKind: 'metamorphic-invariant',
    specCitation: scenarios.find(
      (scenario) =>
        scenario.manifest.scenarioId ===
        'unrelatedTextEditPreservesOpaqueInlineContentControl'
    )!.manifest.specCitation,
    outcomes: {
      'safe-docx': { status: 'invariant-pass' },
      'oracle-test-adapter': { status: 'invariant-fail' },
    },
  }
);
const oracleSummary = buildCapabilitySummary(loaded, oracleResults) as {
  capabilities: Array<{
    capabilityId: string;
    axis: string;
    oracleKind: string;
    scenarioIds: string[];
    outcomes: Record<string, { counts: Record<string, number> }>;
  }>;
};
const oracleMatrix = renderMatrixHtml(oracleResults);
check(
  'matrix renders separate conformance and invariant group totals',
  oracleMatrix.includes('<strong>ECMA conformance:</strong>') &&
    oracleMatrix.includes('<strong>Metamorphic invariants:</strong>') &&
    oracleMatrix.includes('oracle-test-adapter 1/1') &&
    oracleMatrix.includes('oracle-test-adapter 0/1') &&
    !oracleMatrix.includes('oracle-test-adapter 1/2')
);
const contentControlRows = oracleSummary.capabilities.filter(
  (row) => row.capabilityId === 'word.content-controls.inline' && row.axis === 'preserve'
);
check(
  'capability aggregation separates conformance from invariant outcomes',
  contentControlRows.some(
    (row) =>
      row.oracleKind === 'ecma-conformance' &&
      row.outcomes['oracle-test-adapter']?.counts.pass === 1
  ) &&
    contentControlRows.some(
      (row) =>
        row.oracleKind === 'metamorphic-invariant' &&
        row.outcomes['oracle-test-adapter']?.counts['invariant-fail'] === 1
    )
);
const textEditRows = oracleSummary.capabilities.filter(
  (row) => row.capabilityId === 'word.text.find-replace' && row.axis === 'edit'
);
check(
  'opaque invariant does not contaminate word.text.find-replace edit evidence',
  textEditRows.every(
    (row) => !row.scenarioIds.includes('unrelatedTextEditPreservesOpaqueInlineContentControl')
  )
);
const wrongInvariantStatus = structuredClone(oracleResults);
wrongInvariantStatus.results.find(
  (result) => result.scenarioId === 'unrelatedTextEditPreservesOpaqueInlineContentControl'
)!.outcomes['oracle-test-adapter']!.status = 'fail';
expectThrows('aggregation rejects ECMA fail status on invariant evidence', /invalid for metamorphic-invariant/, () =>
  buildCapabilitySummary(loaded, wrongInvariantStatus)
);
const validateResultsSchema = new Ajv2020({ strict: true }).compile(RESULTS_SCHEMA);
check(
  'results schema rejects ECMA fail status on invariant evidence',
  !validateResultsSchema(wrongInvariantStatus)
);
const wrongOracleKind = structuredClone(oracleResults);
wrongOracleKind.results.find(
  (result) => result.scenarioId === 'unrelatedTextEditPreservesOpaqueInlineContentControl'
)!.oracleKind = 'ecma-conformance';
expectThrows('aggregation rejects result oracle kind drift', /does not match registry/, () =>
  buildCapabilitySummary(loaded, wrongOracleKind)
);

const sparseResults = structuredClone(results);
const omittedAdapter = sparseResults.implementations.at(-1)!.adapterName;
for (const scenario of sparseResults.results) delete scenario.outcomes[omittedAdapter];
const sparseSummary = buildCapabilitySummary(loaded, sparseResults) as {
  capabilities: Array<{ outcomes: Record<string, { denominator: number }> }>;
};
check(
  'adapter without scenario outcomes is omitted from per-axis denominators',
  sparseSummary.capabilities.every((capability) => capability.outcomes[omittedAdapter] === undefined)
);

if (failed) process.exit(1);

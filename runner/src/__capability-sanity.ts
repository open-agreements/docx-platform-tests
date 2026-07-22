import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  buildCapabilitySummary,
  buildScenarioCoverage,
  loadAndValidateCapabilityRegistry,
  validateCapabilityRelationships,
  type CapabilityProfiles,
  type CapabilityRegistry,
  type ScenarioCapabilityRegistry,
} from './capability-registry.js';
import { loadAllScenarios, REPO_ROOT } from './scenarios.js';
import type { ResultsDocument } from './types.js';

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

check('committed registry validates', loaded.registry.capabilities.length === 21);
check('every committed scenario is mapped', new Set(loaded.mappings.map((mapping) => mapping.scenarioId)).size === scenarios.length);

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
expectThrows('citation absent from mapped capability is rejected', /absent from its mapped capabilities/, () =>
  validateCapabilityRelationships(registry, wrongCitation, profiles, scenarios)
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

const unknownProfileCapability = structuredClone(profiles);
unknownProfileCapability.profiles[0].capabilityIds.push('word.unknown.capability');
expectThrows('unknown profile capability is rejected', /references unknown capability/, () =>
  validateCapabilityRelationships(registry, mappings, unknownProfileCapability, scenarios)
);

const coverage = buildScenarioCoverage(loaded) as {
  totals: { scenarios: number; mappings: number; effectiveMappings: number };
  unclassifiedScenarioIds: string[];
  capabilities: Array<{ capabilityId: string; axes: Array<{ axis: string; scenarioIds: string[] }> }>;
};
check('coverage retains raw scenario denominator', coverage.totals.scenarios === scenarios.length);
check('coverage retains raw mapping denominator', coverage.totals.mappings === loaded.mappings.length);
check('coverage distinguishes derived cross-platform mappings', coverage.totals.effectiveMappings > coverage.totals.mappings);
check('coverage has no unclassified scenarios', coverage.unclassifiedScenarioIds.length === 0);
check(
  'cross-platform coverage is derived from neutral scenario execution',
  coverage.capabilities.every((capability) => {
    const axis = capability.axes.find((candidate) => candidate.axis === 'crossPlatform');
    return !axis || axis.scenarioIds.length > 0;
  })
);

const results = JSON.parse(
  readFileSync(join(REPO_ROOT, 'results', 'latest.json'), 'utf8')
) as ResultsDocument;
const firstSummary = JSON.stringify(buildCapabilitySummary(loaded, results));
const secondSummary = JSON.stringify(buildCapabilitySummary(loaded, results));
check('capability aggregation is deterministic', firstSummary === secondSummary);
check('capability aggregation retains scenario IDs', firstSummary.includes('scenarioIds'));
check('capability aggregation retains denominators', firstSummary.includes('denominator'));

if (failed) process.exit(1);

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { Ajv2020 } from 'ajv/dist/2020.js';
import { loadAllScenarios, REPO_ROOT } from './scenarios.js';
import type { LoadedScenario, OutcomeStatus, ResultsDocument, SpecCitation } from './types.js';

export const CAPABILITY_AXES = [
  'detect',
  'preserve',
  'parse',
  'validate',
  'generate',
  'edit',
  'compare',
  'acceptReject',
  'wordRoundtrip',
  'crossPlatform',
] as const;

export type CapabilityAxis = (typeof CAPABILITY_AXES)[number];

export type OracleClass =
  | 'normative-schema'
  | 'normative-prose'
  | 'metamorphic-invariant'
  | 'observed-word-behavior'
  | 'cross-implementation-evidence'
  | 'serialization-specific';

export interface Capability {
  id: string;
  title: string;
  family: string;
  description: string;
  standards: Array<Pick<SpecCitation, 'standard' | 'edition' | 'part' | 'section'>>;
  microsoftExtensions?: string[];
  dependencies: string[];
  applicableAxes: CapabilityAxis[];
  packageParts: string[];
}

export interface CapabilityRegistry {
  schemaVersion: 1;
  registryVersion: number;
  capabilities: Capability[];
}

export interface ScenarioCapabilityMapping {
  scenarioId: string;
  capabilityId: string;
  axis: CapabilityAxis;
  oracleClasses: OracleClass[];
}

export interface ScenarioCapabilityRegistry {
  schemaVersion: 1;
  registryVersion: number;
  mappings: ScenarioCapabilityMapping[];
}

export interface CapabilityProfile {
  id: string;
  title: string;
  description: string;
  capabilityIds: string[];
  axes: CapabilityAxis[];
}

export interface CapabilityProfiles {
  schemaVersion: 1;
  registryVersion: number;
  profiles: CapabilityProfile[];
}

export interface LoadedCapabilityRegistry {
  registry: CapabilityRegistry;
  mappings: ScenarioCapabilityMapping[];
  profiles: CapabilityProfile[];
}

const REGISTRY_ROOT = join(REPO_ROOT, 'registry');

function readJson(path: string): unknown {
  return JSON.parse(readFileSync(path, 'utf8'));
}

function validateJson(schemaName: string, dataName: string, data: unknown): void {
  const schema = readJson(join(REGISTRY_ROOT, schemaName));
  const ajv = new Ajv2020({ allErrors: true });
  const validate = ajv.compile(schema as Record<string, unknown>);
  if (validate(data)) return;
  const details = (validate.errors ?? [])
    .map((error) => `${error.instancePath || '/'} ${error.message}`)
    .join('\n');
  throw new Error(`${dataName} does not match ${schemaName}:\n${details}`);
}

function citationKey(citation: Pick<SpecCitation, 'standard' | 'edition' | 'part' | 'section'>): string {
  return `${citation.standard}|${citation.edition}|${citation.part}|${citation.section}`;
}

function assertUnique(values: string[], label: string): void {
  const duplicates = values.filter((value, index) => values.indexOf(value) !== index);
  if (duplicates.length > 0) {
    throw new Error(`${label} contains duplicate value(s): ${[...new Set(duplicates)].join(', ')}`);
  }
}

function assertAcyclic(capabilities: Capability[]): void {
  const dependencies = new Map(capabilities.map((capability) => [capability.id, capability.dependencies]));
  const visiting = new Set<string>();
  const visited = new Set<string>();
  const visit = (id: string, path: string[]): void => {
    if (visiting.has(id)) throw new Error(`capability dependency cycle: ${[...path, id].join(' -> ')}`);
    if (visited.has(id)) return;
    visiting.add(id);
    for (const dependency of dependencies.get(id) ?? []) visit(dependency, [...path, id]);
    visiting.delete(id);
    visited.add(id);
  };
  for (const capability of capabilities) visit(capability.id, []);
}

export function loadAndValidateCapabilityRegistry(): LoadedCapabilityRegistry {
  const registryData = readJson(join(REGISTRY_ROOT, 'capabilities.json'));
  const mappingData = readJson(join(REGISTRY_ROOT, 'scenario-capabilities.json'));
  const profileData = readJson(join(REGISTRY_ROOT, 'profiles.json'));
  validateJson('capabilities.schema.json', 'capabilities.json', registryData);
  validateJson('scenario-capabilities.schema.json', 'scenario-capabilities.json', mappingData);
  validateJson('profiles.schema.json', 'profiles.json', profileData);

  const registry = registryData as CapabilityRegistry;
  const mappingRegistry = mappingData as ScenarioCapabilityRegistry;
  const profileRegistry = profileData as CapabilityProfiles;
  return validateCapabilityRelationships(registry, mappingRegistry, profileRegistry);
}

export function validateCapabilityRelationships(
  registry: CapabilityRegistry,
  mappingRegistry: ScenarioCapabilityRegistry,
  profileRegistry: CapabilityProfiles,
  scenarios: LoadedScenario[] = loadAllScenarios()
): LoadedCapabilityRegistry {
  if (
    registry.registryVersion !== mappingRegistry.registryVersion ||
    registry.registryVersion !== profileRegistry.registryVersion
  ) {
    throw new Error('capabilities, mappings, and profiles must use the same registryVersion');
  }

  const capabilitiesById = new Map(registry.capabilities.map((capability) => [capability.id, capability]));
  assertUnique(registry.capabilities.map((capability) => capability.id), 'capabilities');
  for (const capability of registry.capabilities) {
    assertUnique(capability.dependencies, `${capability.id}.dependencies`);
    assertUnique(capability.applicableAxes, `${capability.id}.applicableAxes`);
    assertUnique(capability.packageParts, `${capability.id}.packageParts`);
    assertUnique(capability.standards.map(citationKey), `${capability.id}.standards`);
    for (const dependency of capability.dependencies) {
      if (!capabilitiesById.has(dependency)) {
        throw new Error(`${capability.id} has unknown dependency ${dependency}`);
      }
      if (dependency === capability.id) throw new Error(`${capability.id} cannot depend on itself`);
    }
  }
  assertAcyclic(registry.capabilities);

  const scenariosById = new Map(scenarios.map((scenario) => [scenario.manifest.scenarioId, scenario]));
  assertUnique(
    mappingRegistry.mappings.map((mapping) => `${mapping.scenarioId}|${mapping.capabilityId}|${mapping.axis}`),
    'scenario capability mappings'
  );
  for (const mapping of mappingRegistry.mappings) {
    const scenario = scenariosById.get(mapping.scenarioId);
    if (!scenario) throw new Error(`mapping references unknown scenario ${mapping.scenarioId}`);
    const capability = capabilitiesById.get(mapping.capabilityId);
    if (!capability) throw new Error(`mapping references unknown capability ${mapping.capabilityId}`);
    if (!capability.applicableAxes.includes(mapping.axis)) {
      throw new Error(`${mapping.scenarioId} maps ${mapping.capabilityId} to inapplicable axis ${mapping.axis}`);
    }
    const scenarioCitationKeys = new Set([
      citationKey(scenario.manifest.specCitation),
      ...(scenario.manifest.secondarySpecCitations ?? []).map(citationKey),
    ]);
    if (!capability.standards.some((citation) => scenarioCitationKeys.has(citationKey(citation)))) {
      throw new Error(`${mapping.scenarioId} mapping to ${mapping.capabilityId} has no shared citation`);
    }
  }

  const mappingsByScenario = new Map<string, ScenarioCapabilityMapping[]>();
  for (const mapping of mappingRegistry.mappings) {
    const current = mappingsByScenario.get(mapping.scenarioId) ?? [];
    current.push(mapping);
    mappingsByScenario.set(mapping.scenarioId, current);
  }
  for (const scenario of scenarios) {
    const scenarioId = scenario.manifest.scenarioId;
    const mappings = mappingsByScenario.get(scenarioId) ?? [];
    if (mappings.length === 0) throw new Error(`scenario ${scenarioId} has no capability mapping`);
    const mappedCitations = new Set(
      mappings.flatMap((mapping) => capabilitiesById.get(mapping.capabilityId)?.standards ?? []).map(citationKey)
    );
    const scenarioCitations = [
      scenario.manifest.specCitation,
      ...(scenario.manifest.secondarySpecCitations ?? []),
    ];
    for (const citation of scenarioCitations) {
      if (!mappedCitations.has(citationKey(citation))) {
        throw new Error(`${scenarioId} citation ${citationKey(citation)} is absent from its mapped capabilities`);
      }
    }
    const oracleClasses = new Set(mappings.flatMap((mapping) => mapping.oracleClasses));
    const assertionKinds = new Set(scenario.manifest.assertionList.map((assertion) => assertion.assertionKind));
    if (assertionKinds.has('canonicalXmlEquals') && !oracleClasses.has('serialization-specific')) {
      throw new Error(`${scenarioId} uses canonicalXmlEquals without serialization-specific classification`);
    }
    if (assertionKinds.has('schemaValidAgainstWml') && !oracleClasses.has('normative-schema')) {
      throw new Error(`${scenarioId} uses schema validation without normative-schema classification`);
    }
    if (scenario.manifest.wordBehaviorNote && !oracleClasses.has('observed-word-behavior')) {
      throw new Error(`${scenarioId} has a wordBehaviorNote without observed-word-behavior classification`);
    }
  }

  assertUnique(profileRegistry.profiles.map((profile) => profile.id), 'profiles');
  for (const profile of profileRegistry.profiles) {
    assertUnique(profile.capabilityIds, `${profile.id}.capabilityIds`);
    assertUnique(profile.axes, `${profile.id}.axes`);
    for (const capabilityId of profile.capabilityIds) {
      if (!capabilitiesById.has(capabilityId)) {
        throw new Error(`${profile.id} references unknown capability ${capabilityId}`);
      }
    }
  }

  return { registry, mappings: mappingRegistry.mappings, profiles: profileRegistry.profiles };
}

export function buildScenarioCoverage(loaded: LoadedCapabilityRegistry): object {
  const scenarios = loadAllScenarios();
  const mappedScenarioIds = new Set(loaded.mappings.map((mapping) => mapping.scenarioId));
  const capabilities = [...loaded.registry.capabilities]
    .sort((a, b) => a.id.localeCompare(b.id))
    .map((capability) => {
      const axes = capability.applicableAxes.map((axis) => ({
        axis,
        scenarioIds: loaded.mappings
          .filter((mapping) => mapping.capabilityId === capability.id && mapping.axis === axis)
          .map((mapping) => mapping.scenarioId)
          .sort(),
      }));
      return {
        capabilityId: capability.id,
        family: capability.family,
        axes,
        coveredAxisCount: axes.filter((axis) => axis.scenarioIds.length > 0).length,
        applicableAxisCount: axes.length,
      };
    });
  const coveredAxisCount = capabilities.reduce((sum, capability) => sum + capability.coveredAxisCount, 0);
  const applicableAxisCount = capabilities.reduce((sum, capability) => sum + capability.applicableAxisCount, 0);
  return {
    schemaVersion: 1,
    registryVersion: loaded.registry.registryVersion,
    totals: {
      capabilities: loaded.registry.capabilities.length,
      scenarios: scenarios.length,
      mappings: loaded.mappings.length,
      coveredAxes: coveredAxisCount,
      applicableAxes: applicableAxisCount,
      uncoveredAxes: applicableAxisCount - coveredAxisCount,
    },
    unclassifiedScenarioIds: scenarios
      .map((scenario) => scenario.manifest.scenarioId)
      .filter((scenarioId) => !mappedScenarioIds.has(scenarioId))
      .sort(),
    capabilities,
  };
}

const OUTCOME_STATUSES: OutcomeStatus[] = [
  'pass',
  'pass-divergent',
  'fail',
  'unsupported',
  'error',
  'protocol-mismatch',
];

export function buildCapabilitySummary(
  loaded: LoadedCapabilityRegistry,
  results: ResultsDocument
): object {
  const resultsByScenario = new Map(results.results.map((result) => [result.scenarioId, result]));
  const capabilitiesById = new Map(
    loaded.registry.capabilities.map((capability) => [capability.id, capability])
  );
  const measuredMappings = loaded.mappings.filter((mapping) => resultsByScenario.has(mapping.scenarioId));
  for (const mapping of [...measuredMappings]) {
    const capability = capabilitiesById.get(mapping.capabilityId);
    const scenarioResult = resultsByScenario.get(mapping.scenarioId);
    const measuredAdapterCount = results.implementations.filter(
      (implementation) => scenarioResult?.outcomes[implementation.adapterName] !== undefined
    ).length;
    if (
      mapping.axis !== 'crossPlatform' &&
      capability?.applicableAxes.includes('crossPlatform') &&
      measuredAdapterCount >= 2
    ) {
      measuredMappings.push({
        scenarioId: mapping.scenarioId,
        capabilityId: mapping.capabilityId,
        axis: 'crossPlatform',
        oracleClasses: ['cross-implementation-evidence'],
      });
    }
  }
  const grouped = new Map<string, ScenarioCapabilityMapping[]>();
  for (const mapping of measuredMappings) {
    const key = `${mapping.capabilityId}|${mapping.axis}`;
    grouped.set(key, [...(grouped.get(key) ?? []), mapping]);
  }
  const capabilities = [...grouped.entries()]
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([key, mappings]) => {
      const [capabilityId, axis] = key.split('|') as [string, CapabilityAxis];
      const scenarioIds = [...new Set(mappings.map((mapping) => mapping.scenarioId))].sort();
      const outcomes = Object.fromEntries(
        results.implementations.map((implementation) => {
          const counts = Object.fromEntries(OUTCOME_STATUSES.map((status) => [status, 0])) as Record<OutcomeStatus, number>;
          for (const scenarioId of scenarioIds) {
            const status = resultsByScenario.get(scenarioId)?.outcomes[implementation.adapterName]?.status;
            if (status) counts[status] += 1;
          }
          const nonzeroCounts = Object.fromEntries(
            OUTCOME_STATUSES.filter((status) => counts[status] > 0).map((status) => [status, counts[status]])
          );
          return [
            implementation.adapterName,
            {
              denominator: scenarioIds.length,
              passLike: counts.pass + counts['pass-divergent'],
              counts: nonzeroCounts,
            },
          ];
        })
      );
      return {
        capabilityId,
        axis,
        scenarioIds,
        outcomes,
      };
    });
  return {
    schemaVersion: 1,
    registryVersion: loaded.registry.registryVersion,
    sourceResults: {
      schemaVersion: results.schemaVersion,
      runTimestamp: results.runTimestamp,
      scenarioCount: results.results.length,
      implementations: results.implementations,
    },
    unmeasuredScenarioIds: [...new Set(
      loaded.mappings
        .map((mapping) => mapping.scenarioId)
        .filter((scenarioId) => !resultsByScenario.has(scenarioId))
    )].sort(),
    capabilities,
  };
}

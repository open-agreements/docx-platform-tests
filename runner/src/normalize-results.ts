import { isDeepStrictEqual } from 'node:util';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { createHash } from 'node:crypto';
import { Ajv2020 } from 'ajv/dist/2020.js';
import {
  loadAndValidateCapabilityRegistry,
  scenarioOracleKind,
  type LoadedCapabilityRegistry,
} from './capability-registry.js';
import { REPO_ROOT } from './scenarios.js';
import type { ResultsDocument } from './types.js';

interface JsonObject {
  [key: string]: unknown;
}

interface ResultsHistory {
  schemaVersion: 1;
  legacySchemas: {
    '2': {
      sourceCommit: string;
      fixturePath: string;
      fixtureSha256: string;
      canonicalSha256: string;
      expectedImplementationCount: number;
      expectedAdapterNames: string[];
      expectedScenarioCount: number;
      expectedScenarioIds: string[];
      oracleKindByScenario: Record<string, 'ecma-conformance'>;
    };
  };
}

function loadResultsHistory(): ResultsHistory {
  const schema = JSON.parse(
    readFileSync(join(REPO_ROOT, 'registry', 'results-history.schema.json'), 'utf8')
  );
  const history = JSON.parse(
    readFileSync(join(REPO_ROOT, 'registry', 'results-history.json'), 'utf8')
  );
  const validate = new Ajv2020({ allErrors: true }).compile(schema);
  if (!validate(history)) {
    throw new Error(`invalid results history: ${JSON.stringify(validate.errors)}`);
  }
  const typed = history as ResultsHistory;
  const evidence = typed.legacySchemas['2'];
  const fixtureBytes = readFileSync(join(REPO_ROOT, evidence.fixturePath));
  const fixtureDigest = createHash('sha256').update(fixtureBytes).digest('hex');
  if (fixtureDigest !== evidence.fixtureSha256) {
    throw new Error(
      `immutable schema-v2 fixture digest ${fixtureDigest} does not match ${evidence.fixtureSha256}`
    );
  }
  if (evidence.expectedAdapterNames.length !== evidence.expectedImplementationCount) {
    throw new Error('results history implementation cardinality does not match adapter identities');
  }
  if (evidence.expectedScenarioIds.length !== evidence.expectedScenarioCount) {
    throw new Error('results history scenario cardinality does not match scenario identities');
  }
  if (
    !isDeepStrictEqual(
      Object.keys(evidence.oracleKindByScenario).sort(),
      [...evidence.expectedScenarioIds].sort()
    )
  ) {
    throw new Error('results history oracle identities do not match expected scenario identities');
  }
  return typed;
}

function object(value: unknown, label: string): JsonObject {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`${label} must be an object`);
  }
  return value as JsonObject;
}

function canonicalJson(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(',')}]`;
  if (value && typeof value === 'object') {
    return `{${Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${canonicalJson((value as JsonObject)[key])}`)
      .join(',')}}`;
  }
  return JSON.stringify(value);
}

export function canonicalResultsDigest(value: unknown): string {
  return createHash('sha256').update(canonicalJson(value)).digest('hex');
}

function authenticateSchemaV2Snapshot(document: JsonObject, history: ResultsHistory): void {
  const evidence = history.legacySchemas['2'];
  if (!Array.isArray(document.implementations)) throw new Error('implementations must be an array');
  if (!Array.isArray(document.results)) throw new Error('results must be an array');

  if (document.implementations.length !== evidence.expectedImplementationCount) {
    throw new Error(
      `schema-v2 implementation cardinality ${document.implementations.length}; expected ${evidence.expectedImplementationCount}`
    );
  }
  const adapterNames = document.implementations.map((item, index) => {
    const implementation = object(item, `implementations[${index}]`);
    return implementation.adapterName;
  });
  if (!isDeepStrictEqual(adapterNames, evidence.expectedAdapterNames)) {
    throw new Error('schema-v2 ordered adapter identities do not match immutable history');
  }
  if (document.results.length !== evidence.expectedScenarioCount) {
    throw new Error(
      `schema-v2 scenario cardinality ${document.results.length}; expected ${evidence.expectedScenarioCount}`
    );
  }
  const scenarioIds = document.results.map((item, index) => {
    const result = object(item, `results[${index}]`);
    return result.scenarioId;
  });
  if (!isDeepStrictEqual(scenarioIds, evidence.expectedScenarioIds)) {
    throw new Error('schema-v2 ordered scenario identities do not match immutable history');
  }
  const digest = canonicalResultsDigest(document);
  if (digest !== evidence.canonicalSha256) {
    throw new Error(
      `schema-v2 snapshot digest ${digest} does not match immutable history ${evidence.canonicalSha256}`
    );
  }
}

export function normalizeResultsDocument(
  raw: unknown,
  loaded: LoadedCapabilityRegistry,
  history: ResultsHistory = loadResultsHistory()
): ResultsDocument {
  const document = object(raw, 'results document');
  if (document.schemaVersion === 3) return document as unknown as ResultsDocument;
  if (document.schemaVersion !== 2) {
    throw new Error(`unsupported results schemaVersion ${String(document.schemaVersion)}`);
  }
  authenticateSchemaV2Snapshot(document, history);
  const rawResults = document.results;
  if (!Array.isArray(rawResults)) throw new Error('results must be an array');

  const normalizedResults = rawResults.map((rawResult, index) => {
    const result = object(rawResult, `results[${index}]`);
    const scenarioId = result.scenarioId;
    if (typeof scenarioId !== 'string') throw new Error(`results[${index}].scenarioId must be a string`);
    if ('oracleKind' in result) {
      throw new Error(`${scenarioId} schema-v2 result unexpectedly declares oracleKind`);
    }
    const mappings = loaded.mappings.filter((mapping) => mapping.scenarioId === scenarioId);
    if (mappings.length === 0) {
      throw new Error(`${scenarioId} has no current registry history for schema-v2 inference`);
    }
    const oracleKind = scenarioOracleKind(loaded.mappings, scenarioId);
    if (oracleKind !== 'ecma-conformance') {
      throw new Error(
        `${scenarioId} is ambiguous legacy invariant evidence; schema-v2 cannot identify its oracle`
      );
    }
    if (history.legacySchemas['2'].oracleKindByScenario[scenarioId] !== 'ecma-conformance') {
      throw new Error(`${scenarioId} has no schema-v2 oracle history; refusing inference`);
    }
    return { ...result, oracleKind };
  });

  return {
    ...(document as unknown as Omit<ResultsDocument, 'schemaVersion' | 'results'>),
    schemaVersion: 3,
    results: normalizedResults as unknown as ResultsDocument['results'],
  };
}

function argValue(flag: string): string | undefined {
  const index = process.argv.indexOf(flag);
  return index === -1 ? undefined : process.argv[index + 1];
}

function main(): void {
  const latestPath = join(REPO_ROOT, 'results', 'latest.json');
  const loaded = loadAndValidateCapabilityRegistry();
  if (process.argv.includes('--check')) {
    const current = JSON.parse(readFileSync(latestPath, 'utf8')) as ResultsDocument;
    const evidence = loadResultsHistory().legacySchemas['2'];
    const historical = JSON.parse(
      readFileSync(join(REPO_ROOT, evidence.fixturePath), 'utf8')
    );
    const normalized = normalizeResultsDocument(historical, loaded);
    if (!isDeepStrictEqual(normalized, current)) {
      throw new Error('results/latest.json is not reproducible from the immutable schema-v2 fixture');
    }
    console.log('results/latest.json is reproducible from the immutable schema-v2 fixture');
    return;
  }

  const inputPath = argValue('--input');
  const outputPath = argValue('--output');
  if (!inputPath || !outputPath) {
    throw new Error('usage: normalize-results --input <schema-v2-or-v3.json> --output <schema-v3.json>');
  }
  const normalized = normalizeResultsDocument(
    JSON.parse(readFileSync(inputPath, 'utf8')),
    loaded
  );
  writeFileSync(outputPath, `${JSON.stringify(normalized, null, 2)}\n`);
  console.log(`wrote normalized schema-v3 results to ${outputPath}`);
}

if (process.argv[1]?.endsWith('normalize-results.ts')) main();

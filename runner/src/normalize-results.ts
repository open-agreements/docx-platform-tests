import { isDeepStrictEqual } from 'node:util';
import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
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
      sourceRunTimestamp: string;
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
  return history as ResultsHistory;
}

function object(value: unknown, label: string): JsonObject {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`${label} must be an object`);
  }
  return value as JsonObject;
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
  if (document.runTimestamp !== history.legacySchemas['2'].sourceRunTimestamp) {
    throw new Error(
      `schema-v2 run ${String(document.runTimestamp)} has no explicit oracle history`
    );
  }
  if (!Array.isArray(document.results)) throw new Error('results must be an array');

  const normalizedResults = document.results.map((rawResult, index) => {
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
    const legacyShape = {
      ...current,
      schemaVersion: 2,
      results: current.results.map(({ oracleKind: _oracleKind, ...result }) => result),
    };
    const normalized = normalizeResultsDocument(legacyShape, loaded);
    if (!isDeepStrictEqual(normalized, current)) {
      throw new Error('results/latest.json is not reproducible from its exact schema-v2 shape');
    }
    console.log('results/latest.json is reproducible from its exact schema-v2 shape');
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

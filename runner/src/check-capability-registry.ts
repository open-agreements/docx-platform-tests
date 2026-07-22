import { readFileSync, writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import {
  buildCapabilitySummary,
  buildScenarioCoverage,
  loadAndValidateCapabilityRegistry,
} from './capability-registry.js';
import { REPO_ROOT } from './scenarios.js';
import type { ResultsDocument } from './types.js';
import { normalizeResultsDocument } from './normalize-results.js';

const checkOnly = process.argv.includes('--check');
const loaded = loadAndValidateCapabilityRegistry();
const results = normalizeResultsDocument(
  JSON.parse(readFileSync(join(REPO_ROOT, 'results', 'latest.json'), 'utf8')),
  loaded
) as ResultsDocument;
const outputs = new Map<string, object>([
  [join(REPO_ROOT, 'registry', 'scenario-coverage.json'), buildScenarioCoverage(loaded)],
  [join(REPO_ROOT, 'results', 'capability-summary.json'), buildCapabilitySummary(loaded, results)],
]);

for (const [path, value] of outputs) {
  const serialized = JSON.stringify(value, null, 2) + '\n';
  const label = relative(REPO_ROOT, path);
  if (checkOnly) {
    if (readFileSync(path, 'utf8') !== serialized) {
      console.error(`${label} is stale; run \`npm run write-capability-index\` in runner/`);
      process.exitCode = 1;
    } else {
      console.log(`${label} is current`);
    }
  } else {
    writeFileSync(path, serialized);
    console.log(`wrote ${label}`);
  }
}

if (process.exitCode !== 1) {
  console.log(
    `capability registry valid: ${loaded.registry.capabilities.length} capabilities, ` +
      `${loaded.mappings.length} mappings, ${loaded.profiles.length} profile(s)`
  );
}

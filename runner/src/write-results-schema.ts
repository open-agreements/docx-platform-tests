import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { REPO_ROOT } from './scenarios.js';
import { RESULTS_SCHEMA } from './results-schema.js';

const schemaPath = join(REPO_ROOT, 'results', 'results.schema.json');
const schemaJson = JSON.stringify(RESULTS_SCHEMA, null, 2) + '\n';

if (process.argv.includes('--check')) {
  const existing = readFileSync(schemaPath, 'utf8');
  if (existing !== schemaJson) {
    console.error(
      'results/results.schema.json is stale; run `npm run write-results-schema` in runner/'
    );
    process.exit(1);
  }
  console.log('results/results.schema.json is current');
} else {
  mkdirSync(dirname(schemaPath), { recursive: true });
  writeFileSync(schemaPath, schemaJson);
  console.log('wrote results/results.schema.json');
}

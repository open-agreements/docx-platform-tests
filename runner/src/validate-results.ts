import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { Ajv2020 } from 'ajv/dist/2020.js';
import { REPO_ROOT } from './scenarios.js';

const schema = JSON.parse(
  readFileSync(join(REPO_ROOT, 'results', 'results.schema.json'), 'utf8')
);
const results = JSON.parse(readFileSync(join(REPO_ROOT, 'results', 'latest.json'), 'utf8'));

const ajv = new Ajv2020({ allErrors: true });
const validate = ajv.compile(schema);

if (!validate(results)) {
  for (const error of validate.errors ?? []) {
    console.error(`${error.instancePath || '/'} ${error.message}`);
  }
  process.exit(1);
}

console.log('results/latest.json validates against results/results.schema.json');

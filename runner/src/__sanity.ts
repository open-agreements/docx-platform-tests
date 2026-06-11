import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { evaluateAssertion, projectBodyText } from './assertions.js';
import type { ScenarioManifest } from './types.js';

// Simulated CORRECT post-accept output for acceptInsertionsUnwrapsInsWrappers,
// deliberately written with different whitespace/run-granularity than expected/.
const acceptedSplitRuns = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">Kept text </w:t></w:r><w:r><w:t xml:space="preserve">Inserted </w:t></w:r><w:r><w:t>text</w:t></w:r></w:p></w:body></w:document>`;

const dir = join('..', 'scenarios', 'tracked-changes', 'acceptInsertionsUnwrapsInsWrappers');
const manifest = JSON.parse(readFileSync(join(dir, 'scenario.json'), 'utf8')) as ScenarioManifest;
for (const a of manifest.assertionList) {
  const r = evaluateAssertion(a, acceptedSplitRuns, dir);
  console.log(`${r.passed ? 'PASS' : 'FAIL'} ${r.assertionKind}: ${r.detail.split('\n')[0]}`);
}

// A WRONG output (w:ins left in place) must fail.
const wrong = readFileSync(join(dir, 'input', 'document.xml'), 'utf8');
const wrongResults = manifest.assertionList.map(a => evaluateAssertion(a, wrong, dir));
console.log('wrong output fails at least one assertion:', wrongResults.some(r => !r.passed));

// Offset projection for the find-replace scenario.
const replaced = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Payment due in sixty days</w:t></w:r></w:p></w:body></w:document>`;
console.log('projection:', JSON.stringify(projectBodyText(replaced)), 'sixty at', projectBodyText(replaced).indexOf('sixty'));

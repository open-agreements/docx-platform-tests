import { spawnSync } from 'node:child_process';
import {
  mkdtempSync,
  readFileSync,
  writeFileSync,
  rmSync,
  existsSync,
  mkdirSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { extractDocumentXml } from './docx.js';
import { evaluateAssertion } from './assertions.js';
import { loadAllScenarios, REPO_ROOT } from './scenarios.js';
import type {
  AdapterRegistration,
  AdapterRegistry,
  LoadedScenario,
  ResultsDocument,
  ScenarioOutcome,
} from './types.js';

const PROTOCOL_VERSION = 1;

function adapterVersion(adapter: AdapterRegistration): string {
  if (!adapter.adapterVersionCommand?.length) return 'unknown';
  const [cmd, ...args] = adapter.adapterVersionCommand;
  const result = spawnSync(cmd, args, { cwd: REPO_ROOT, encoding: 'utf8' });
  return result.status === 0 ? result.stdout.trim() : 'unknown';
}

function runScenario(
  adapter: AdapterRegistration,
  scenario: LoadedScenario
): ScenarioOutcome {
  const workDir = mkdtempSync(join(tmpdir(), 'dpt-'));
  try {
    const operationPath = join(workDir, 'operation.json');
    const outputPath = join(workDir, 'output.docx');
    writeFileSync(
      operationPath,
      JSON.stringify(scenario.manifest.operationDescriptor, null, 2)
    );
    const [cmd, ...baseArgs] = adapter.adapterCommand;
    const result = spawnSync(
      cmd,
      [
        ...baseArgs,
        '--protocol-version',
        String(PROTOCOL_VERSION),
        '--operation',
        operationPath,
        '--input',
        join(scenario.scenarioDir, scenario.manifest.inputDocumentPath),
        '--output',
        outputPath,
      ],
      { cwd: REPO_ROOT, encoding: 'utf8', timeout: 120_000 }
    );

    if (result.error) {
      return { status: 'error', reason: String(result.error) };
    }
    if (result.status === 2) {
      return {
        status: 'unsupported',
        reason: result.stdout.trim() || 'no reason given',
      };
    }
    if (result.status === 3) {
      return {
        status: 'protocol-mismatch',
        reason: result.stdout.trim() || `adapter does not speak protocol v${PROTOCOL_VERSION}`,
      };
    }
    if (result.status !== 0) {
      return {
        status: 'error',
        reason: `exit ${result.status}: ${result.stderr.trim().slice(0, 2000)}`,
      };
    }
    if (!existsSync(outputPath)) {
      return { status: 'error', reason: 'adapter exited 0 but wrote no output package' };
    }

    const outputXml = extractDocumentXml(readFileSync(outputPath));
    const assertionResults = scenario.manifest.assertionList.map((assertion) =>
      evaluateAssertion(assertion, outputXml, scenario.scenarioDir)
    );
    return {
      status: assertionResults.every((r) => r.passed) ? 'pass' : 'fail',
      assertionResults,
    };
  } finally {
    rmSync(workDir, { recursive: true, force: true });
  }
}

const registry = JSON.parse(
  readFileSync(join(REPO_ROOT, 'registry', 'adapters.json'), 'utf8')
) as AdapterRegistry;
const scenarios = loadAllScenarios();

const results: ResultsDocument = {
  runTimestamp: new Date().toISOString(),
  dslVersion: '1.0',
  protocolVersion: PROTOCOL_VERSION,
  implementations: registry.adapters.map((adapter) => ({
    adapterName: adapter.adapterName,
    adapterVersion: adapterVersion(adapter),
  })),
  results: scenarios.map(({ manifest }) => ({
    scenarioId: manifest.scenarioId,
    scenarioTitle: manifest.scenarioTitle,
    specCitation: manifest.specCitation,
    outcomes: {},
  })),
};

for (const adapter of registry.adapters) {
  console.log(`\n=== adapter: ${adapter.adapterName}`);
  for (let i = 0; i < scenarios.length; i++) {
    const outcome = runScenario(adapter, scenarios[i]);
    results.results[i].outcomes[adapter.adapterName] = outcome;
    console.log(
      `  ${scenarios[i].manifest.scenarioId}: ${outcome.status}` +
        (outcome.reason ? ` (${outcome.reason})` : '')
    );
    for (const assertion of outcome.assertionResults ?? []) {
      if (!assertion.passed) {
        console.log(`    FAILED ${assertion.assertionKind}: ${assertion.detail}`);
      }
    }
  }
}

const resultsDir = join(REPO_ROOT, 'results');
mkdirSync(resultsDir, { recursive: true });
writeFileSync(
  join(resultsDir, 'latest.json'),
  JSON.stringify(results, null, 2) + '\n'
);
console.log(
  `\nwrote results/latest.json (${registry.adapters.length} adapter(s), ${scenarios.length} scenario(s))`
);

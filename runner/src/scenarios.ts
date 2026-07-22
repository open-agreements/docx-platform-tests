import { readdirSync, readFileSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { LoadedScenario, ScenarioAssertion, ScenarioManifest } from './types.js';

export const REPO_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
export const SCENARIOS_ROOT = join(REPO_ROOT, 'scenarios');

const ASSERTION_INTRODUCTION_VERSION: Record<
  ScenarioAssertion['assertionKind'],
  string
> = {
  xpathQueryCount: '1.0',
  xpathQueryExists: '1.0',
  documentTextContainsAtOffset: '1.0',
  canonicalXmlEquals: '1.0',
  schemaValidAgainstWml: '1.2',
  hyperlinkResolvesToExternalUrl: '1.3',
  commentExistsWithTextAndAnchor: '1.5',
  paragraphNumberingResolvesToFormat: '1.6',
  xpathElementTextEquals: '1.8',
  ignorableNamespaceDeclared: '1.8',
};

function versionTuple(version: string, label: string): [number, number] {
  const match = /^(\d+)\.(\d+)$/.exec(version);
  if (!match) throw new Error(`${label} has invalid DSL version '${version}'`);
  return [Number(match[1]), Number(match[2])];
}

function versionLessThan(left: string, right: string): boolean {
  const [leftMajor, leftMinor] = versionTuple(left, 'scenario');
  const [rightMajor, rightMinor] = versionTuple(right, 'assertion');
  return leftMajor < rightMajor || (leftMajor === rightMajor && leftMinor < rightMinor);
}

export function validateScenarioDslCompatibility(manifest: ScenarioManifest): void {
  versionTuple(manifest.dslVersion, manifest.scenarioId);
  for (const assertion of manifest.assertionList) {
    const minimum = ASSERTION_INTRODUCTION_VERSION[assertion.assertionKind];
    if (!minimum) {
      throw new Error(`${manifest.scenarioId} uses unknown assertion kind '${assertion.assertionKind}'`);
    }
    if (versionLessThan(manifest.dslVersion, minimum)) {
      throw new Error(
        `${manifest.scenarioId} uses ${assertion.assertionKind}, which requires DSL ${minimum}; declared ${manifest.dslVersion}`
      );
    }
  }
}

export function loadAllScenarios(): LoadedScenario[] {
  const scenarios: LoadedScenario[] = [];
  for (const group of readdirSync(SCENARIOS_ROOT, { withFileTypes: true })) {
    if (!group.isDirectory()) continue;
    const groupDir = join(SCENARIOS_ROOT, group.name);
    for (const entry of readdirSync(groupDir, { withFileTypes: true })) {
      if (!entry.isDirectory()) continue;
      const scenarioDir = join(groupDir, entry.name);
      const manifestPath = join(scenarioDir, 'scenario.json');
      if (!existsSync(manifestPath)) continue;
      const manifest = JSON.parse(
        readFileSync(manifestPath, 'utf8')
      ) as ScenarioManifest;
      validateScenarioDslCompatibility(manifest);
      if (manifest.scenarioId !== entry.name) {
        throw new Error(
          `scenario directory '${entry.name}' declares scenarioId '${manifest.scenarioId}'`
        );
      }
      scenarios.push({ manifest, scenarioDir });
    }
  }
  scenarios.sort((a, b) =>
    a.manifest.scenarioId.localeCompare(b.manifest.scenarioId)
  );
  return scenarios;
}

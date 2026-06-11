import { readdirSync, readFileSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { LoadedScenario, ScenarioManifest } from './types.js';

export const REPO_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
export const SCENARIOS_ROOT = join(REPO_ROOT, 'scenarios');

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

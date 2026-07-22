import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { REPO_ROOT } from './scenarios.js';
import type { ResultsDocument, ScenarioOracleKind } from './types.js';
import { loadAndValidateCapabilityRegistry } from './capability-registry.js';
import { normalizeResultsDocument } from './normalize-results.js';

const STATUS_LABEL: Record<string, string> = {
  pass: 'Pass',
  'pass-divergent': 'Pass (divergent serialization)',
  fail: 'Fail',
  'invariant-pass': 'Invariant pass',
  'invariant-fail': 'Invariant fail',
  unsupported: 'Unsupported',
  error: 'Error',
  'protocol-mismatch': 'Protocol mismatch',
};

function esc(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function groupTotals(
  results: ResultsDocument,
  group: string,
  oracleKind: ScenarioOracleKind,
  label: string
): string {
  const scenarios = results.results.filter(
    (scenario) => scenario.scenarioGroup === group && scenario.oracleKind === oracleKind
  );
  if (scenarios.length === 0) return '';
  const totals = results.implementations
    .map((implementation) => {
      const passLike = scenarios.filter((scenario) => {
        const status = scenario.outcomes[implementation.adapterName]?.status;
        return oracleKind === 'ecma-conformance'
          ? status === 'pass' || status === 'pass-divergent'
          : status === 'invariant-pass';
      }).length;
      return `${esc(implementation.adapterName)} ${passLike}/${scenarios.length}`;
    })
    .join(' · ');
  return `<br><small><strong>${label}:</strong> ${totals}</small>`;
}

export function renderMatrixHtml(results: ResultsDocument): string {
  const headerCells = results.implementations
    .map(
      (implementation) =>
        `<th scope="col">${esc(implementation.adapterName)}<br><small>${esc(implementation.adapterVersion)}</small></th>`
    )
    .join('');

  const rows = results.results
    .reduce((htmlRows, scenario, index) => {
      const previous = results.results[index - 1];
      if (!previous || previous.scenarioGroup !== scenario.scenarioGroup) {
        const conformanceTotals = groupTotals(
          results,
          scenario.scenarioGroup,
          'ecma-conformance',
          'ECMA conformance'
        );
        const invariantTotals = groupTotals(
          results,
          scenario.scenarioGroup,
          'metamorphic-invariant',
          'Metamorphic invariants'
        );
        htmlRows.push(`<tr class="group-row"><th scope="rowgroup" colspan="${
          results.implementations.length + 1
        }">${esc(scenario.scenarioGroup)}${conformanceTotals}${invariantTotals}</th></tr>`);
      }
      const citation = scenario.specCitation;
      const oracleLabel =
        scenario.oracleKind === 'metamorphic-invariant'
          ? 'Metamorphic invariant (not ECMA conformance)'
          : 'ECMA conformance';
      const cells = results.implementations
        .map((implementation) => {
          const outcome = scenario.outcomes[implementation.adapterName];
          const status = outcome?.status ?? 'error';
          const reason = outcome?.reason ? `<br><small>${esc(outcome.reason)}</small>` : '';
          return `<td class="cell-${esc(status)}">${esc(STATUS_LABEL[status] ?? status)}${reason}</td>`;
        })
        .join('');
      htmlRows.push(`<tr>
      <th scope="row"><code>${esc(scenario.scenarioId)}</code><br><small>${esc(scenario.scenarioTitle)}</small><br><small>${esc(oracleLabel)}</small><br><small class="citation">${esc(
        `${citation.standard} edition ${citation.edition}, Part ${citation.part} § ${citation.section} (${citation.clauseTitle})`
      )}</small></th>${cells}</tr>`);
      return htmlRows;
    }, [] as string[])
    .join('\n');

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>docx-platform-tests — cross-implementation results matrix</title>
<meta name="description" content="WordprocessingML conformance and explicitly labeled metamorphic-invariant results across implementations.">
<link rel="canonical" href="https://open-agreements.github.io/docx-platform-tests/results/">
<style>
  body { font: 16px/1.5 system-ui, sans-serif; margin: 2rem auto; max-width: 60rem; padding: 0 1rem; color: #1a1a1a; }
  table { border-collapse: collapse; width: 100%; margin: 1.5rem 0; }
  th, td { border: 1px solid #d9d9d9; padding: 0.5rem 0.7rem; text-align: left; vertical-align: top; }
  thead th { background: #f6f6f4; }
  .group-row th { background: #ece8df; font-size: 1.05rem; }
  .cell-pass { background: #e8f5e9; }
  .cell-pass-divergent { background: #f3f8e2; }
  .cell-fail, .cell-error, .cell-protocol-mismatch { background: #fdecea; }
  .cell-invariant-fail { background: #fff1d6; }
  .cell-invariant-pass { background: #e7f4ef; }
  .cell-unsupported { background: #f4f1ea; color: #6b6357; }
  .citation { color: #6b6357; }
  footer { margin-top: 2rem; font-size: 0.85rem; color: #6b6357; }
</style>
</head>
<body>
<h1>docx-platform-tests results</h1>
<p>Cross-implementation conformance and preservation-invariant results for WordprocessingML (<code>.docx</code>), in the
<a href="https://github.com/web-platform-tests/wpt">web-platform-tests</a> tradition. Each
conformance scenario asserts behavior derivable from its cited ECMA-376 clause. A row labeled
<strong>Metamorphic invariant</strong> instead reports a narrowly declared preservation property and
is not an ECMA conformance result. Every scenario runs unchanged against each registered adapter.
<strong>Unsupported</strong> means the adapter declined
the operation honestly (the NOTRUN analog): a gap in the matrix is information about the library,
not a failure of the suite. <strong>Pass (divergent serialization)</strong> means every assertion
derived from the cited clause passed, but the saved document did not match the reference
serialization's canonical XML — typically an implementation materializing formatting defaults on
save. That is legal WordprocessingML serialization freedom, not a conformance failure; the
per-assertion breakdown in <a href="./latest.json">latest.json</a> carries the detail.</p>
<table>
<thead><tr><th scope="col">Scenario</th>${headerCells}</tr></thead>
<tbody>
${rows}
</tbody>
</table>
<p>Data: <a href="./latest.json">results/latest.json</a> · Schema:
<a href="./results.schema.json">results/results.schema.json</a> ·
<a href="./capability-summary.json">Capability-axis aggregation</a> · Suite, scenario DSL, and adapter protocol:
<a href="https://github.com/open-agreements/docx-platform-tests">open-agreements/docx-platform-tests</a> (Apache-2.0)
· Canonical live matrix: <a href="https://open-agreements.github.io/docx-platform-tests/results/">open-agreements.github.io</a></p>
<footer>Run ${esc(results.runTimestamp)} · results schema v${results.schemaVersion} · DSL ${esc(
    results.dslVersion
  )} · adapter protocol v${results.protocolVersion}.
To add an implementation, see <a href="https://github.com/open-agreements/docx-platform-tests/blob/main/docs/adapter-protocol.md">docs/adapter-protocol.md</a>.</footer>
</body>
</html>
`;
}

function main(): void {
  const results = normalizeResultsDocument(
    JSON.parse(readFileSync(join(REPO_ROOT, 'results', 'latest.json'), 'utf8')),
    loadAndValidateCapabilityRegistry()
  );
  writeFileSync(join(REPO_ROOT, 'results', 'index.html'), renderMatrixHtml(results));
  console.log('wrote results/index.html');
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();

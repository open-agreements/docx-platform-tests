import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { REPO_ROOT } from './scenarios.js';
import type { ResultsDocument } from './types.js';

// Renders results/latest.json to a small static results/index.html — the
// suite-owned, implementation-neutral matrix view (the wpt.fyi analog).
// Published to gh-pages by CI alongside the JSON it is derived from.

const STATUS_LABEL: Record<string, string> = {
  pass: 'Pass',
  'pass-divergent': 'Pass (divergent serialization)',
  fail: 'Fail',
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

const results = JSON.parse(
  readFileSync(join(REPO_ROOT, 'results', 'latest.json'), 'utf8')
) as ResultsDocument;

const headerCells = results.implementations
  .map(
    (impl) =>
      `<th scope="col">${esc(impl.adapterName)}<br><small>${esc(impl.adapterVersion)}</small></th>`
  )
  .join('');

const rows = results.results
  .reduce(
    (htmlRows, scenario, index) => {
      const previous = results.results[index - 1];
      if (!previous || previous.scenarioGroup !== scenario.scenarioGroup) {
        const groupScenarios = results.results.filter(
          (candidate) => candidate.scenarioGroup === scenario.scenarioGroup
        );
        const totals = results.implementations
          .map((impl) => {
            const passLike = groupScenarios.filter((candidate) => {
              const status = candidate.outcomes[impl.adapterName]?.status;
              return status === 'pass' || status === 'pass-divergent';
            }).length;
            return `${esc(impl.adapterName)} ${passLike}/${groupScenarios.length}`;
          })
          .join(' · ');
        htmlRows.push(`<tr class="group-row"><th scope="rowgroup" colspan="${
          results.implementations.length + 1
        }">${esc(scenario.scenarioGroup)}<br><small>${totals}</small></th></tr>`);
      }
      const citation = scenario.specCitation;
      const cells = results.implementations
        .map((impl) => {
          const outcome = scenario.outcomes[impl.adapterName];
          const status = outcome?.status ?? 'error';
          const reason = outcome?.reason ? `<br><small>${esc(outcome.reason)}</small>` : '';
          return `<td class="cell-${esc(status)}">${esc(STATUS_LABEL[status] ?? status)}${reason}</td>`;
        })
        .join('');
      htmlRows.push(`<tr>
      <th scope="row"><code>${esc(scenario.scenarioId)}</code><br><small>${esc(scenario.scenarioTitle)}</small><br><small class="citation">${esc(
        `${citation.standard} edition ${citation.edition}, Part ${citation.part} § ${citation.section} (${citation.clauseTitle})`
      )}</small></th>${cells}</tr>`);
      return htmlRows;
    },
    [] as string[]
  )
  .join('\n');

const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>docx-platform-tests — cross-implementation results matrix</title>
<meta name="description" content="WordprocessingML conformance results across implementations: every scenario asserts behavior derivable from a cited ECMA-376 clause and runs unchanged against each registered adapter.">
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
  .cell-unsupported { background: #f4f1ea; color: #6b6357; }
  .citation { color: #6b6357; }
  footer { margin-top: 2rem; font-size: 0.85rem; color: #6b6357; }
</style>
</head>
<body>
<h1>docx-platform-tests results</h1>
<p>Cross-implementation conformance results for WordprocessingML (<code>.docx</code>), in the
<a href="https://github.com/web-platform-tests/wpt">web-platform-tests</a> tradition. Each scenario
asserts behavior derivable from the cited ECMA-376 clause — not from any one library — and runs
unchanged against every registered adapter. <strong>Unsupported</strong> means the adapter declined
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
<a href="./results.schema.json">results/results.schema.json</a> · Suite, scenario DSL, and adapter protocol:
<a href="https://github.com/open-agreements/docx-platform-tests">open-agreements/docx-platform-tests</a> (Apache-2.0)
· Canonical live matrix: <a href="https://open-agreements.github.io/docx-platform-tests/results/">open-agreements.github.io</a></p>
<footer>Run ${esc(results.runTimestamp)} · results schema v${results.schemaVersion} · DSL ${esc(
  results.dslVersion
)} · adapter protocol v${results.protocolVersion}.
To add an implementation, see <a href="https://github.com/open-agreements/docx-platform-tests/blob/main/docs/adapter-protocol.md">docs/adapter-protocol.md</a>.</footer>
</body>
</html>
`;

writeFileSync(join(REPO_ROOT, 'results', 'index.html'), html);
console.log('wrote results/index.html');

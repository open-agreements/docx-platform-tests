import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { REPO_ROOT } from './scenarios.js';
import type { ResultsDocument } from './types.js';

// Renders results/latest.json to a small static results/index.html — the
// suite-owned, implementation-neutral matrix view (the wpt.fyi analog).
// Published to gh-pages by CI alongside the JSON it is derived from.

const STATUS_LABEL: Record<string, string> = {
  pass: 'Pass',
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
  .map((scenario) => {
    const citation = scenario.specCitation;
    const cells = results.implementations
      .map((impl) => {
        const outcome = scenario.outcomes[impl.adapterName];
        const status = outcome?.status ?? 'error';
        const reason = outcome?.reason ? `<br><small>${esc(outcome.reason)}</small>` : '';
        return `<td class="cell-${esc(status)}">${esc(STATUS_LABEL[status] ?? status)}${reason}</td>`;
      })
      .join('');
    return `<tr>
      <th scope="row"><code>${esc(scenario.scenarioId)}</code><br><small>${esc(scenario.scenarioTitle)}</small><br><small class="citation">${esc(
        `${citation.standard} edition ${citation.edition}, Part ${citation.part} § ${citation.section} (${citation.clauseTitle})`
      )}</small></th>${cells}</tr>`;
  })
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
  .cell-pass { background: #e8f5e9; }
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
not a failure of the suite.</p>
<table>
<thead><tr><th scope="col">Scenario</th>${headerCells}</tr></thead>
<tbody>
${rows}
</tbody>
</table>
<p>Data: <a href="./latest.json">results/latest.json</a> · Suite, scenario DSL, and adapter protocol:
<a href="https://github.com/open-agreements/docx-platform-tests">open-agreements/docx-platform-tests</a> (BSD-3-Clause)
· Narrative comparisons: <a href="https://usejunior.com/engineering/safe-docx/cross-implementation">usejunior.com</a></p>
<footer>Run ${esc(results.runTimestamp)} · DSL ${esc(results.dslVersion)} · adapter protocol v${results.protocolVersion}.
To add an implementation, see <a href="https://github.com/open-agreements/docx-platform-tests/blob/main/docs/adapter-protocol.md">docs/adapter-protocol.md</a>.</footer>
</body>
</html>
`;

writeFileSync(join(REPO_ROOT, 'results', 'index.html'), html);
console.log('wrote results/index.html');

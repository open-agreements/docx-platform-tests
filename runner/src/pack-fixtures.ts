import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join, relative } from 'node:path';
import { packMinimalDocx, extractDocumentXml } from './docx.js';
import { loadAllScenarios, REPO_ROOT } from './scenarios.js';

const checkMode = process.argv.includes('--check');
let drifted = 0;

for (const { manifest, scenarioDir } of loadAllScenarios()) {
  const fragmentPath = join(scenarioDir, 'input', 'document.xml');
  const packagePath = join(scenarioDir, manifest.inputDocumentPath);
  const fragment = readFileSync(fragmentPath, 'utf8');
  if (checkMode) {
    if (!existsSync(packagePath)) {
      console.error(`MISSING ${relative(REPO_ROOT, packagePath)}`);
      drifted++;
      continue;
    }
    const packaged = extractDocumentXml(readFileSync(packagePath));
    if (packaged !== fragment) {
      console.error(
        `DRIFT ${relative(REPO_ROOT, packagePath)}: word/document.xml does not match input/document.xml`
      );
      drifted++;
    } else {
      console.log(`OK ${manifest.scenarioId}`);
    }
  } else {
    writeFileSync(packagePath, packMinimalDocx(fragment));
    console.log(`PACKED ${relative(REPO_ROOT, packagePath)}`);
  }
}

if (checkMode && drifted > 0) {
  console.error(
    `${drifted} fixture package(s) out of sync; run 'npm run pack-fixtures'`
  );
  process.exit(1);
}

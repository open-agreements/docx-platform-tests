import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strToU8, zipSync } from 'fflate';
import {
  evaluateAssertion,
  preprocessIgnorableMarkupForSchema,
  projectBodyText,
} from './assertions.js';
import { packageFromParts, packDocx, packMinimalDocx, loadPackage } from './docx.js';
import type { ScenarioAssertion, ScenarioManifest } from './types.js';

let failed = false;
function check(label: string, ok: boolean): void {
  console.log(`${ok ? 'PASS' : 'FAIL'} ${label}`);
  if (!ok) failed = true;
}

// Simulated CORRECT post-accept output for acceptInsertionsUnwrapsInsWrappers,
// deliberately written with different whitespace/run-granularity than expected/.
const acceptedSplitRuns = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t xml:space="preserve">Kept text </w:t></w:r><w:r><w:t xml:space="preserve">Inserted </w:t></w:r><w:r><w:t>text</w:t></w:r></w:p></w:body></w:document>`;

const dir = join('..', 'scenarios', 'tracked-changes', 'acceptInsertionsUnwrapsInsWrappers');
const manifest = JSON.parse(readFileSync(join(dir, 'scenario.json'), 'utf8')) as ScenarioManifest;
for (const a of manifest.assertionList) {
  const r = evaluateAssertion(a, packageFromParts({ 'word/document.xml': acceptedSplitRuns }), dir);
  check(`${r.assertionKind}: ${r.detail.split('\n')[0]}`, r.passed);
}

// A WRONG output (w:ins left in place) must fail.
const wrong = readFileSync(join(dir, 'input', 'document.xml'), 'utf8');
const wrongResults = manifest.assertionList.map((a) =>
  evaluateAssertion(a, packageFromParts({ 'word/document.xml': wrong }), dir)
);
check('wrong output fails at least one assertion', wrongResults.some((r) => !r.passed));

// Offset projection for the find-replace scenario.
const replaced = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Payment due in sixty days</w:t></w:r></w:p></w:body></w:document>`;
const projected = projectBodyText(replaced);
console.log('projection:', JSON.stringify(projected), 'sixty at', projected.indexOf('sixty'));
check('projection finds sixty', projected.indexOf('sixty') !== -1);

const contentControlRoot = join('..', 'scenarios', 'content-controls');
const normativeSdtDir = join(
  contentControlRoot,
  'unrelatedTextEditPreservesInlineContentControlStructure'
);
const invariantSdtDir = join(
  contentControlRoot,
  'unrelatedTextEditPreservesOpaqueInlineContentControl'
);
const normativeSdtManifest = JSON.parse(
  readFileSync(join(normativeSdtDir, 'scenario.json'), 'utf8')
) as ScenarioManifest;
const invariantSdtManifest = JSON.parse(
  readFileSync(join(invariantSdtDir, 'scenario.json'), 'utf8')
) as ScenarioManifest;
const inlineSdtOutput = readFileSync(
  join(normativeSdtDir, 'input', 'document.xml'),
  'utf8'
).replace('thirty', 'sixty');

function assertionPasses(
  manifest: ScenarioManifest,
  scenarioDir: string,
  xml: string,
  assertionIndex: number
): boolean {
  return evaluateAssertion(
    manifest.assertionList[assertionIndex],
    packageFromParts({ 'word/document.xml': xml }),
    scenarioDir
  ).passed;
}

check(
  'normative inline SDT reference output satisfies every assertion',
  normativeSdtManifest.assertionList.every((assertion) =>
    evaluateAssertion(
      assertion,
      packageFromParts({ 'word/document.xml': inlineSdtOutput }),
      normativeSdtDir
    ).passed
  )
);
check(
  'metamorphic inline SDT reference output satisfies every assertion',
  invariantSdtManifest.assertionList.every((assertion) =>
    evaluateAssertion(
      assertion,
      packageFromParts({ 'word/document.xml': inlineSdtOutput }),
      invariantSdtDir
    ).passed
  )
);

const splitRuns = inlineSdtOutput
  .replace(
    '<w:r><w:t xml:space="preserve">Edit target: sixty days. </w:t></w:r>',
    '<w:r><w:t>Edit target: six</w:t></w:r><w:r><w:t xml:space="preserve">ty days. </w:t></w:r>'
  )
  .replace(
    '<w:r><w:t>Controlled text sentinel</w:t></w:r>',
    '<w:r><w:t>Controlled </w:t></w:r><w:r><w:t>text sentinel</w:t></w:r>'
  );
check(
  'legal run splits satisfy paragraph and controlled-content text assertions',
  assertionPasses(normativeSdtManifest, normativeSdtDir, splitRuns, 0) &&
    assertionPasses(normativeSdtManifest, normativeSdtDir, splitRuns, 5) &&
    assertionPasses(invariantSdtManifest, invariantSdtDir, splitRuns, 0)
);

const targetSdt = inlineSdtOutput.match(/<w:sdt ext:opaqueAttribute[\s\S]*?<\/w:sdt>/)![0];
const duplicatedSdt = inlineSdtOutput.replace(targetSdt, `${targetSdt}${targetSdt}`);
check(
  'total SDT duplication is rejected by both oracle scenarios',
  !assertionPasses(normativeSdtManifest, normativeSdtDir, duplicatedSdt, 1) &&
    !assertionPasses(invariantSdtManifest, invariantSdtDir, duplicatedSdt, 1)
);

const opaqueChild = inlineSdtOutput.match(/\s*<ext:opaqueExtension[\s\S]*?<\/ext:opaqueExtension>/)![0];
const cleanedTargetSdt = targetSdt
  .replace(' ext:opaqueAttribute="opaque-attribute-sentinel"', '')
  .replace(opaqueChild.trim(), '');
const otherSdt = targetSdt
  .replace('w:val="54"', 'w:val="99"')
  .replace('dpt-inline-sdt', 'other-sdt')
  .replace('Controlled text sentinel', 'Other control');
const relocatedOpaque = inlineSdtOutput.replace(targetSdt, cleanedTargetSdt + otherSdt);
check(
  'opaque sentinels relocated to another SDT do not satisfy target-scoped assertions',
  !assertionPasses(invariantSdtManifest, invariantSdtDir, relocatedOpaque, 4) &&
    !assertionPasses(invariantSdtManifest, invariantSdtDir, relocatedOpaque, 5)
);

const reorderedKnownProperties = inlineSdtOutput.replace(
  /(<w:tag w:val="dpt-inline-sdt"\/>)([\s\S]*?<\/ext:opaqueExtension>)(\s*)(<w:id w:val="54"\/>)/,
  '$4$2$3$1'
);
check(
  'schema-invalid known sdtPr child order is rejected',
  !assertionPasses(normativeSdtManifest, normativeSdtDir, reorderedKnownProperties, 4)
);

const reorderedOpaqueChild = inlineSdtOutput.replace(
  /(<w:tag w:val="dpt-inline-sdt"\/>)([\s\S]*?<\/ext:opaqueExtension>)(\s*)(<w:id w:val="54"\/>)/,
  '$1$3$4$2'
);
check(
  'opaque child relative-position change is rejected as an invariant',
  !assertionPasses(invariantSdtManifest, invariantSdtDir, reorderedOpaqueChild, 5)
);
check(
  'opaque child removal is rejected as an invariant',
  !assertionPasses(
    invariantSdtManifest,
    invariantSdtDir,
    inlineSdtOutput.replace(opaqueChild, ''),
    5
  )
);

const aliasedPrefix = inlineSdtOutput
  .replace('xmlns:ext=', 'xmlns:opaque=')
  .replace('mc:Ignorable="ext"', 'mc:Ignorable="opaque"')
  .replaceAll('ext:', 'opaque:');
check(
  'foreign prefix alias rename preserves namespace-semantic assertions',
  invariantSdtManifest.assertionList.every((assertion) =>
    evaluateAssertion(
      assertion,
      packageFromParts({ 'word/document.xml': aliasedPrefix }),
      invariantSdtDir
    ).passed
  )
);
const extensionNamespace =
  'urn:open-agreements:docx-platform-tests:inline-content-control';
const locallyDeclaredIgnorable = inlineSdtOutput
  .replace(`\n  xmlns:ext="${extensionNamespace}"`, '')
  .replace('\n  mc:Ignorable="ext"', '')
  .replace(
    '<w:sdt ext:opaqueAttribute=',
    `<w:sdt xmlns:ext="${extensionNamespace}" mc:Ignorable="ext" ext:opaqueAttribute=`
  );
check(
  'target-local ignorable namespace declaration is equivalent',
  assertionPasses(invariantSdtManifest, invariantSdtDir, locallyDeclaredIgnorable, 3)
);
const siblingOnlyIgnorable = inlineSdtOutput
  .replace('\n  mc:Ignorable="ext"', '')
  .replace('<w:r><w:t xml:space="preserve">Edit target:', '<w:r mc:Ignorable="ext"><w:t xml:space="preserve">Edit target:');
check(
  'out-of-scope sibling mc:Ignorable declaration is rejected',
  !assertionPasses(invariantSdtManifest, invariantSdtDir, siblingOnlyIgnorable, 3)
);
const mismatchedIgnorable = inlineSdtOutput
  .replace('xmlns:ext=', 'xmlns:other="urn:wrong-extension" xmlns:ext=')
  .replace('mc:Ignorable="ext"', 'mc:Ignorable="other"');
check(
  'mismatched ignorable namespace binding is rejected',
  !assertionPasses(invariantSdtManifest, invariantSdtDir, mismatchedIgnorable, 3)
);

const sentinelMutations: Array<{
  label: string;
  manifest: ScenarioManifest;
  scenarioDir: string;
  assertionIndex: number;
  mutate: (xml: string) => string;
}> = [
  { label: 'edited paragraph text', manifest: normativeSdtManifest, scenarioDir: normativeSdtDir, assertionIndex: 0, mutate: (xml) => xml.replace('sixty', 'thirty') },
  { label: 'SDT alias', manifest: normativeSdtManifest, scenarioDir: normativeSdtDir, assertionIndex: 4, mutate: (xml) => xml.replace('Opaque inline control', 'Changed alias') },
  { label: 'SDT id', manifest: normativeSdtManifest, scenarioDir: normativeSdtDir, assertionIndex: 2, mutate: (xml) => xml.replace('w:val="54"', 'w:val="55"') },
  { label: 'SDT tag', manifest: normativeSdtManifest, scenarioDir: normativeSdtDir, assertionIndex: 2, mutate: (xml) => xml.replace('dpt-inline-sdt', 'other-sdt') },
  { label: 'controlled content', manifest: normativeSdtManifest, scenarioDir: normativeSdtDir, assertionIndex: 5, mutate: (xml) => xml.replace('Controlled text sentinel', 'Changed controlled text') },
  { label: 'mc:Ignorable namespace', manifest: invariantSdtManifest, scenarioDir: invariantSdtDir, assertionIndex: 3, mutate: (xml) => xml.replace('mc:Ignorable="ext"', '') },
  { label: 'opaque attribute', manifest: invariantSdtManifest, scenarioDir: invariantSdtDir, assertionIndex: 4, mutate: (xml) => xml.replace('opaque-attribute-sentinel', 'changed') },
  { label: 'opaque child', manifest: invariantSdtManifest, scenarioDir: invariantSdtDir, assertionIndex: 5, mutate: (xml) => xml.replace('extension-child-sentinel', 'changed') },
  { label: 'nested payload attribute', manifest: invariantSdtManifest, scenarioDir: invariantSdtDir, assertionIndex: 6, mutate: (xml) => xml.replace('nested-payload-sentinel', 'changed') },
  { label: 'nested payload text', manifest: invariantSdtManifest, scenarioDir: invariantSdtDir, assertionIndex: 6, mutate: (xml) => xml.replace('opaque payload sentinel', 'changed payload') },
  { label: 'first child', manifest: invariantSdtManifest, scenarioDir: invariantSdtDir, assertionIndex: 7, mutate: (xml) => xml.replace('first-child-sentinel', 'changed') },
  { label: 'last child', manifest: invariantSdtManifest, scenarioDir: invariantSdtDir, assertionIndex: 7, mutate: (xml) => xml.replace('last-child-sentinel', 'changed') },
];
for (const mutation of sentinelMutations) {
  const mutated = mutation.mutate(inlineSdtOutput);
  check(
    `inline SDT assertion rejects changed ${mutation.label} sentinel`,
    mutated !== inlineSdtOutput &&
      !assertionPasses(
        mutation.manifest,
        mutation.scenarioDir,
        mutated,
        mutation.assertionIndex
      )
  );
}

const mceProcessed = preprocessIgnorableMarkupForSchema(inlineSdtOutput);
const localMceProcessed = preprocessIgnorableMarkupForSchema(locallyDeclaredIgnorable);
check(
  'MCE preprocessing removes inherited and target-local ignorable extension markup',
  !mceProcessed.includes('opaqueAttribute') &&
    !mceProcessed.includes('opaqueExtension') &&
    !localMceProcessed.includes('opaqueAttribute') &&
    !localMceProcessed.includes('opaqueExtension') &&
    /<w:alias[^>]*\/>[\s\S]*<w:tag[^>]*\/>[\s\S]*<w:id[^>]*\/>/.test(mceProcessed) &&
    /<w:alias[^>]*\/>[\s\S]*<w:tag[^>]*\/>[\s\S]*<w:id[^>]*\/>/.test(localMceProcessed)
);

// --- DSL 1.3: multi-part assertion machinery ---

const STYLES_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const HYPERLINK_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const HEADER_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
const COMMENTS_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const NUMBERING_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const SETTINGS_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const SETTINGS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';

// A package whose main document references a styles part and an external
// hyperlink, plus a default-type header, all by relationship (arbitrary rIds).
const multiPart = packageFromParts({
  'word/document.xml': `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
  <w:p><w:hyperlink r:id="rIdLink"><w:r><w:t>Spec home</w:t></w:r></w:hyperlink></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Titled</w:t></w:r></w:p>
  <w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/></w:sectPr>
</w:body></w:document>`,
  'word/_rels/document.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="${STYLES_TYPE}" Target="styles.xml"/>
  <Relationship Id="rIdHdr" Type="${HEADER_TYPE}" Target="header3.xml"/>
  <Relationship Id="rIdLink" Type="${HYPERLINK_TYPE}" Target="https://ecma-international.org/" TargetMode="External"/>
</Relationships>`,
  'word/styles.xml': `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style></w:styles>`,
  'word/header3.xml': `<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:hdr>`,
});

function assertOn(a: ScenarioAssertion): boolean {
  return evaluateAssertion(a, multiPart, dir).passed;
}

// Styles part resolved by relationship type (not path); the style id lives there.
check(
  'relationshipFromMainPart resolves styles part',
  assertOn({
    assertionKind: 'xpathQueryExists',
    xpathExpression: "//w:style[@w:styleId='Heading1']",
    assertedPart: { partResolution: 'relationshipFromMainPart', relationshipTypeUri: STYLES_TYPE },
  })
);
// Same query against the main part must NOT find the style (proves part switch).
check(
  'main part does not contain the style definition',
  evaluateAssertion(
    { assertionKind: 'xpathQueryCount', xpathExpression: "//w:style[@w:styleId='Heading1']", expectedCount: 0 },
    multiPart,
    dir
  ).passed
);
// Header resolved two-hop through sectPr headerReference -> r:id -> rels (header3.xml, not header1.xml).
check(
  'headerReference resolves header part by r:id',
  assertOn({
    assertionKind: 'xpathQueryExists',
    xpathExpression: "//w:hdr/w:p/w:r/w:t[normalize-space(.)='Confidential']",
    assertedPart: { partResolution: 'headerReference', headerReferenceType: 'default' },
  })
);
// Missing relationship type -> assertion fails with a diagnostic (the conformance signal).
check(
  'missing numbering part fails, not throws',
  evaluateAssertion(
    {
      assertionKind: 'xpathQueryExists',
      xpathExpression: '//w:num',
      assertedPart: {
        partResolution: 'relationshipFromMainPart',
        relationshipTypeUri: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
      },
    },
    multiPart,
    dir
  ).passed === false
);
// Hyperlink resolves display text -> r:id -> external target.
check(
  'hyperlinkResolvesToExternalUrl matches',
  assertOn({
    assertionKind: 'hyperlinkResolvesToExternalUrl',
    hyperlinkDisplayText: 'Spec home',
    expectedTargetUrl: 'https://ecma-international.org/',
  })
);
// Wrong URL must fail.
check(
  'hyperlinkResolvesToExternalUrl rejects wrong url',
  assertOn({
    assertionKind: 'hyperlinkResolvesToExternalUrl',
    hyperlinkDisplayText: 'Spec home',
    expectedTargetUrl: 'https://example.com/',
  }) === false
);

// Package graph resolution: neither the main document nor settings part uses
// the conventional word/document.xml and word/settings.xml names.
function settingsGraphPackage(settingsContentType: string | null) {
  const settingsOverride = settingsContentType
    ? `<Override PartName="/config/preferences.xml" ContentType="${settingsContentType}"/>`
    : '';
  return packageFromParts({
    '[Content_Types].xml': `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/documents/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  ${settingsOverride}
</Types>`,
    '_rels/.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rootDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="documents/main.xml"/>
</Relationships>`,
    'documents/main.xml': `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Graph-aware body</w:t></w:r></w:p></w:body></w:document>`,
    'documents/_rels/main.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="settings-any-id" Type="${SETTINGS_TYPE}" Target="../config/preferences.xml"/>
</Relationships>`,
    'config/preferences.xml': `<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  });
}

const settingsAssertion: ScenarioAssertion = {
  assertionKind: 'xpathQueryExists',
  xpathExpression:
    "/w:settings/w:compat/w:compatSetting[@w:name='compatibilityMode' and @w:val='15']",
  assertedPart: {
    partResolution: 'relationshipFromMainPart',
    relationshipTypeUri: SETTINGS_TYPE,
    expectedContentType: SETTINGS_CONTENT_TYPE,
  },
};
check(
  'package graph resolves noncanonical main document and settings part',
  evaluateAssertion(settingsAssertion, settingsGraphPackage(SETTINGS_CONTENT_TYPE), dir).passed
);
check(
  'package graph main document drives body projection',
  evaluateAssertion(
    {
      assertionKind: 'documentTextContainsAtOffset',
      expectedSubstring: 'Graph-aware body',
      expectedOffset: 0,
    },
    settingsGraphPackage(SETTINGS_CONTENT_TYPE),
    dir
  ).passed
);
check(
  'wrong related-part content type fails',
  !evaluateAssertion(settingsAssertion, settingsGraphPackage('application/xml'), dir).passed
);
check(
  'missing related-part content type fails',
  !evaluateAssertion(settingsAssertion, settingsGraphPackage(null), dir).passed
);

const arbitraryExtensionZip = loadPackage(
  zipSync({
    '[Content_Types].xml': strToU8(`<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/payload/main.part" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/state/settings.part" ContentType="${SETTINGS_CONTENT_TYPE}"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="root-arbitrary" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="payload/main.part"/>
</Relationships>`),
    'payload/main.part': strToU8(`<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>ZIP graph body</w:t></w:r></w:p></w:body></w:document>`),
    'payload/_rels/main.part.rels': strToU8(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="settings-arbitrary" Type="${SETTINGS_TYPE}" Target="../state/settings.part"/>
</Relationships>`),
    'state/settings.part': strToU8(`<?xml version="1.0"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`),
    'media/binary.dat': new Uint8Array([0, 255, 1, 254, 2, 253]),
  })
);
check(
  'ZIP loader resolves non-XML-extension main and settings parts',
  evaluateAssertion(settingsAssertion, arbitraryExtensionZip, dir).passed
);
check(
  'ZIP loader projects body from non-XML-extension main part',
  evaluateAssertion(
    {
      assertionKind: 'documentTextContainsAtOffset',
      expectedSubstring: 'ZIP graph body',
      expectedOffset: 0,
    },
    arbitraryExtensionZip,
    dir
  ).passed
);
check(
  'ZIP loader does not decode unrelated binary media',
  arbitraryExtensionZip.rawParts.has('media/binary.dat') &&
    !arbitraryExtensionZip.parts.has('media/binary.dat')
);

let missingMainTargetRejected = false;
try {
  loadPackage(
    zipSync({
      '_rels/.rels': strToU8(`<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="missing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="absent/main.part"/>
</Relationships>`),
    })
  );
} catch (error) {
  missingMainTargetRejected =
    error instanceof Error && error.message.includes('resolves to missing part absent/main.part');
}
check('ZIP loader rejects a missing officeDocument target', missingMainTargetRejected);

// A headerReference whose r:id joins to a non-header relationship type is
// malformed and must NOT resolve (even though the id and part both exist).
const FOOTER_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
const wrongTypeHeader = packageFromParts({
  'word/document.xml': `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
  <w:sectPr>
    <w:headerReference w:type="default" r:id="rIdBad"/>
    <w:footerReference w:type="default" r:id="rIdFooter"/>
  </w:sectPr>
</w:body></w:document>`,
  'word/_rels/document.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdBad" Type="${STYLES_TYPE}" Target="styles.xml"/>
  <Relationship Id="rIdFooter" Type="${FOOTER_TYPE}" Target="footer1.xml"/>
</Relationships>`,
  'word/styles.xml': `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
  'word/footer1.xml': `<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Page footer</w:t></w:r></w:p></w:ftr>`,
});
check(
  'headerReference to wrong relationship type fails',
  evaluateAssertion(
    {
      assertionKind: 'xpathQueryExists',
      xpathExpression: '//w:hdr',
      assertedPart: { partResolution: 'headerReference', headerReferenceType: 'default' },
    },
    wrongTypeHeader,
    dir
  ).passed === false
);
check(
  'footerReference to correct relationship type resolves',
  evaluateAssertion(
    {
      assertionKind: 'xpathQueryExists',
      xpathExpression: "//w:ftr/w:p/w:r/w:t[normalize-space(.)='Page footer']",
      assertedPart: { partResolution: 'footerReference', footerReferenceType: 'default' },
    },
    wrongTypeHeader,
    dir
  ).passed
);

// --- Multi-part fixture packing (packDocx): pack -> load -> assert round-trip ---

// No-siblings packDocx must be byte-identical to the historical 3-entry package.
check(
  'packDocx with no siblings equals packMinimalDocx byte-for-byte',
  Buffer.from(packDocx(wrong)).equals(Buffer.from(packMinimalDocx(wrong)))
);

// A document that references a default header by the packer's stable rId (rId4),
// packed with styles + comments + header-default sibling fragments.
const multiPartDocumentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Titled</w:t></w:r></w:p>
  <w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Anchored</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
  <w:sectPr><w:headerReference w:type="default" r:id="rId4"/></w:sectPr>
</w:body></w:document>`;
const packedSiblings = new Map<string, string>([
  [
    'styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style></w:styles>`,
  ],
  [
    'comments.xml',
    `<?xml version="1.0"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer" w:initials="R"><w:p><w:r><w:t>Please clarify</w:t></w:r></w:p></w:comment></w:comments>`,
  ],
  [
    'header-default.xml',
    `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:hdr>`,
  ],
]);
const packed = loadPackage(packDocx(multiPartDocumentXml, packedSiblings));

// packDocx is deterministic: same inputs -> same bytes.
check(
  'packDocx is byte-deterministic across calls',
  Buffer.from(packDocx(multiPartDocumentXml, packedSiblings)).equals(
    Buffer.from(packDocx(multiPartDocumentXml, packedSiblings))
  )
);
// Styles part resolved by relationship Type (packed sibling, not hardcoded path).
check(
  'packed styles part resolves by relationship type',
  evaluateAssertion(
    {
      assertionKind: 'xpathQueryExists',
      xpathExpression: "//w:style[@w:styleId='Heading1']",
      assertedPart: { partResolution: 'relationshipFromMainPart', relationshipTypeUri: STYLES_TYPE },
    },
    packed,
    dir
  ).passed
);
// Comments part resolved by relationship Type.
check(
  'packed comments part resolves by relationship type',
  evaluateAssertion(
    {
      assertionKind: 'xpathQueryExists',
      xpathExpression: "//w:comment[@w:id='0']/w:p/w:r/w:t[normalize-space(.)='Please clarify']",
      assertedPart: {
        partResolution: 'relationshipFromMainPart',
        relationshipTypeUri: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
      },
    },
    packed,
    dir
  ).passed
);
// Header part resolved two-hop through the packer's stable rId4 wiring.
check(
  'packed header-default resolves two-hop via stable rId4',
  evaluateAssertion(
    {
      assertionKind: 'xpathQueryExists',
      xpathExpression: "//w:hdr/w:p/w:r/w:t[normalize-space(.)='Confidential']",
      assertedPart: { partResolution: 'headerReference', headerReferenceType: 'default' },
    },
    packed,
    dir
  ).passed
);

// --- DSL 1.5: comments assertion joins comments part IDs to main-document anchors ---
const commentPackage = packageFromParts({
  'word/document.xml': `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:commentRangeStart w:id="9"/><w:r><w:t>Review this phrase</w:t></w:r><w:commentRangeEnd w:id="9"/><w:r><w:commentReference w:id="9"/></w:r></w:p>
</w:body></w:document>`,
  'word/_rels/document.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rComments" Type="${COMMENTS_TYPE}" Target="comments.xml"/>
</Relationships>`,
  'word/comments.xml': `<?xml version="1.0"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="9" w:author="Reviewer" w:initials="RV"><w:p><w:r><w:t>Clarify this point</w:t></w:r></w:p></w:comment></w:comments>`,
});
check(
  'commentExistsWithTextAndAnchor matches text, author, and anchor',
  evaluateAssertion(
    {
      assertionKind: 'commentExistsWithTextAndAnchor',
      expectedCommentText: 'Clarify this point',
      expectedAuthorName: 'Reviewer',
      expectedAnchorText: 'Review this phrase',
    },
    commentPackage,
    dir
  ).passed
);
check(
  'commentExistsWithTextAndAnchor rejects wrong anchor',
  evaluateAssertion(
    {
      assertionKind: 'commentExistsWithTextAndAnchor',
      expectedCommentText: 'Clarify this point',
      expectedAnchorText: 'different phrase',
    },
    commentPackage,
    dir
  ).passed === false
);

// --- DSL 1.6: numbering assertion resolves direct numPr and style-carried numPr ---
const numberingXml = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>
  <w:num w:numId="10"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;
const stylesWithListNumber = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr></w:style>
</w:styles>`;
const numberingPackage = packageFromParts({
  'word/document.xml': `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr></w:pPr><w:r><w:t>Direct item</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr><w:r><w:t>Style item</w:t></w:r></w:p>
</w:body></w:document>`,
  'word/_rels/document.xml.rels': `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rStyles" Type="${STYLES_TYPE}" Target="styles.xml"/>
  <Relationship Id="rNumbering" Type="${NUMBERING_TYPE}" Target="numbering.xml"/>
</Relationships>`,
  'word/styles.xml': stylesWithListNumber,
  'word/numbering.xml': numberingXml,
});
check(
  'paragraphNumberingResolvesToFormat matches direct numPr',
  evaluateAssertion(
    {
      assertionKind: 'paragraphNumberingResolvesToFormat',
      anchorText: 'Direct item',
      expectedNumberFormat: 'decimal',
    },
    numberingPackage,
    dir
  ).passed
);
check(
  'paragraphNumberingResolvesToFormat matches style-carried numPr',
  evaluateAssertion(
    {
      assertionKind: 'paragraphNumberingResolvesToFormat',
      anchorText: 'Style item',
      expectedNumberFormat: 'decimal',
    },
    numberingPackage,
    dir
  ).passed
);
check(
  'paragraphNumberingResolvesToFormat rejects wrong format',
  evaluateAssertion(
    {
      assertionKind: 'paragraphNumberingResolvesToFormat',
      anchorText: 'Style item',
      expectedNumberFormat: 'bullet',
    },
    numberingPackage,
    dir
  ).passed === false
);

if (failed) {
  console.error('sanity checks FAILED');
  process.exit(1);
}
console.log('sanity checks passed');

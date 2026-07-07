import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { evaluateAssertion, projectBodyText } from './assertions.js';
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

// --- DSL 1.3: multi-part assertion machinery ---

const STYLES_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const HYPERLINK_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const HEADER_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';

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

if (failed) {
  console.error('sanity checks FAILED');
  process.exit(1);
}
console.log('sanity checks passed');

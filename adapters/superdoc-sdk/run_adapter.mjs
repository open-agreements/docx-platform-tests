#!/usr/bin/env node
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createSuperDocClient } from '@superdoc-dev/sdk';
import { DOMParser } from '@xmldom/xmldom';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

const PROTOCOL_VERSION = '1';

const CONTENT_TYPES = '[Content_Types].xml';
const DOCUMENT_RELS = 'word/_rels/document.xml.rels';
const STYLES = 'word/styles.xml';

const STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';
const STYLES_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const MINIMAL_DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDptSuperDocStyles" Type="${STYLES_REL_TYPE}" Target="styles.xml"/>
</Relationships>
`;

const MINIMAL_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr/></w:rPrDefault>
    <w:pPrDefault><w:pPr/></w:pPrDefault>
  </w:docDefaults>
</w:styles>
`;

function argValue(flag) {
  const index = process.argv.indexOf(flag);
  return index === -1 ? undefined : process.argv[index + 1];
}

function unsupported(reason) {
  console.log(reason);
  process.exit(2);
}

function insertBeforeClosingTag(xml, closingTag, insertion) {
  const index = xml.lastIndexOf(closingTag);
  if (index === -1) return xml;
  return `${xml.slice(0, index)}${insertion}${xml.slice(index)}`;
}

function addStylesContentType(xml) {
  if (xml.includes('PartName="/word/styles.xml"')) return xml;
  return insertBeforeClosingTag(
    xml,
    '</Types>',
    `  <Override PartName="/word/styles.xml" ContentType="${STYLES_CONTENT_TYPE}"/>\n`
  );
}

function addStylesRelationship(xml) {
  if (xml.includes(STYLES_REL_TYPE)) return xml;
  return insertBeforeClosingTag(
    xml,
    '</Relationships>',
    `  <Relationship Id="rIdDptSuperDocStyles" Type="${STYLES_REL_TYPE}" Target="styles.xml"/>\n`
  );
}

async function normalizeDocxForSuperDoc(inputPath, outputPath) {
  const entries = unzipSync(new Uint8Array(await readFile(inputPath)));
  if (!entries['word/document.xml']) {
    throw new Error('input package has no word/document.xml part');
  }
  if (!entries[CONTENT_TYPES]) {
    throw new Error('input package has no [Content_Types].xml part');
  }

  entries[CONTENT_TYPES] = strToU8(addStylesContentType(strFromU8(entries[CONTENT_TYPES])));
  if (entries[DOCUMENT_RELS]) {
    entries[DOCUMENT_RELS] = strToU8(addStylesRelationship(strFromU8(entries[DOCUMENT_RELS])));
  } else {
    entries[DOCUMENT_RELS] = strToU8(MINIMAL_DOCUMENT_RELS);
  }
  if (!entries[STYLES]) {
    entries[STYLES] = strToU8(MINIMAL_STYLES);
  }

  await writeFile(outputPath, zipSync(entries));
}

async function readDocumentXml(inputPath) {
  const entries = unzipSync(new Uint8Array(await readFile(inputPath)));
  const documentXml = entries['word/document.xml'];
  if (!documentXml) {
    throw new Error('input package has no word/document.xml part');
  }
  return strFromU8(documentXml);
}

function hasParagraphMarkRevision(documentXml) {
  const document = new DOMParser({
    onError: () => {},
  }).parseFromString(documentXml, 'application/xml');

  for (const pPr of document.getElementsByTagNameNS(W_NS, 'pPr')) {
    for (const rPr of elementChildren(pPr, W_NS, 'rPr')) {
      if (
        elementChildren(rPr, W_NS, 'del').length > 0 ||
        elementChildren(rPr, W_NS, 'ins').length > 0
      ) {
        return true;
      }
    }
  }
  return false;
}

function elementChildren(node, namespaceUri, localName) {
  return Array.from(node.childNodes).filter(
    (child) =>
      child.nodeType === 1 &&
      child.namespaceURI === namespaceUri &&
      child.localName === localName
  );
}

async function listAllTrackedChanges(doc) {
  const items = [];
  let offset = 0;
  const limit = 100;
  while (true) {
    const page = await doc.trackChanges.list({ in: 'all', limit, offset });
    items.push(...(page.items ?? []));
    const returned = page.page?.returned ?? page.items?.length ?? 0;
    if (returned < limit) break;
    offset += returned;
  }
  return items;
}

async function decideAllTrackedChanges(doc, decision) {
  for (let i = 0; i < 1000; i++) {
    const [item] = await listAllTrackedChanges(doc);
    if (!item) return;
    await doc.trackChanges.decide({ decision, target: { id: item.id } });
  }
  throw new Error('tracked change resolution did not converge after 1000 decisions');
}

async function replaceFirstTextOccurrence(doc, operation) {
  const match = await doc.query.match({
    select: {
      type: 'text',
      pattern: operation.findText,
      mode: 'contains',
      caseSensitive: true,
    },
    require: 'first',
  });
  const target = match.items?.[0]?.target;
  if (!target) {
    unsupported(`superdoc-sdk found no occurrence of '${operation.findText}'`);
  }
  await doc.replace({ target, text: operation.replaceText });
}

const protocolVersion = argValue('--protocol-version');
const operationPath = argValue('--operation');
const inputPath = argValue('--input');
const outputPath = argValue('--output');

if (protocolVersion !== PROTOCOL_VERSION) {
  console.log(`superdoc-sdk adapter speaks protocol v${PROTOCOL_VERSION}, got ${protocolVersion}`);
  process.exit(3);
}

if (!operationPath || !inputPath || !outputPath) {
  console.error('missing required adapter protocol arguments');
  process.exit(1);
}

const tempDir = await mkdtemp(join(tmpdir(), 'dpt-superdoc-'));
process.env.HOME = tempDir;

try {
  const operation = JSON.parse(await readFile(operationPath, 'utf8'));
  if (
    (operation.operationName === 'acceptAllTrackedChanges' ||
      operation.operationName === 'rejectAllTrackedChanges') &&
    hasParagraphMarkRevision(await readDocumentXml(inputPath))
  ) {
    unsupported(
      'SuperDoc SDK does not expose paragraph-mark tracked changes through trackChanges.list'
    );
  }

  const normalizedInputPath = join(tempDir, 'input.docx');
  await normalizeDocxForSuperDoc(inputPath, normalizedInputPath);

  const client = createSuperDocClient();
  await client.connect();
  try {
    const doc = await client.open({ doc: normalizedInputPath });
    try {
      if (operation.operationName === 'acceptAllTrackedChanges') {
        await decideAllTrackedChanges(doc, 'accept');
      } else if (operation.operationName === 'rejectAllTrackedChanges') {
        await decideAllTrackedChanges(doc, 'reject');
      } else if (operation.operationName === 'replaceFirstTextOccurrence') {
        await replaceFirstTextOccurrence(doc, operation);
      } else {
        unsupported(`superdoc-sdk adapter does not implement operation '${operation.operationName}'`);
      }
      await doc.save({ out: outputPath, force: true });
    } finally {
      await doc.close();
    }
  } finally {
    await client.dispose();
  }
} finally {
  await rm(tempDir, { recursive: true, force: true });
}

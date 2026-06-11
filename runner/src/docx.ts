import { strToU8, strFromU8, zipSync, unzipSync } from 'fflate';

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
`;

const PACKAGE_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
`;

// Fixed mtime keeps committed fixture packages stable across regenerations.
const FIXED_MTIME = new Date('2026-01-01T00:00:00Z');

export function packMinimalDocx(documentXml: string): Uint8Array {
  const entry = (text: string): [Uint8Array, { mtime: Date }] => [
    strToU8(text),
    { mtime: FIXED_MTIME },
  ];
  return zipSync({
    '[Content_Types].xml': entry(CONTENT_TYPES_XML),
    '_rels/.rels': entry(PACKAGE_RELS_XML),
    'word/document.xml': entry(documentXml),
  });
}

export function extractDocumentXml(docxBytes: Uint8Array): string {
  const entries = unzipSync(docxBytes);
  const part = entries['word/document.xml'];
  if (!part) {
    throw new Error('package has no word/document.xml part');
  }
  return strFromU8(part);
}

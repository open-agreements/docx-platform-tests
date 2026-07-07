import { strToU8, strFromU8, zipSync, unzipSync } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
import xpath from 'xpath';
import { WML_NS } from './canonicalize.js';

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HEADER_REL_TYPE = `${OFFICE_REL_NS}/header`;
const FOOTER_REL_TYPE = `${OFFICE_REL_NS}/footer`;

export const MAIN_DOCUMENT_PART = 'word/document.xml';

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
  const part = entries[MAIN_DOCUMENT_PART];
  if (!part) {
    throw new Error('package has no word/document.xml part');
  }
  return strFromU8(part);
}

/**
 * A loaded .docx package: every text part decoded to a string, keyed by part
 * name. Assertions resolve non-main parts (styles, numbering, comments,
 * headers, footers) through OPC relationships rather than by name.
 */
export interface LoadedPackage {
  parts: Map<string, string>;
  mainDocumentXml: string;
}

export function loadPackage(docxBytes: Uint8Array): LoadedPackage {
  const entries = unzipSync(docxBytes);
  const parts = new Map<string, string>();
  for (const [name, bytes] of Object.entries(entries)) {
    // Decode every text-shaped part; binary parts (media) are irrelevant to the
    // XML-only assertions and are simply skipped by not being queried.
    if (name.endsWith('.xml') || name.endsWith('.rels')) {
      parts.set(name, strFromU8(bytes));
    }
  }
  const mainDocumentXml = parts.get(MAIN_DOCUMENT_PART);
  if (mainDocumentXml === undefined) {
    throw new Error('package has no word/document.xml part');
  }
  return { parts, mainDocumentXml };
}

/** In-memory package for tests, no zip round-trip. `word/document.xml` required. */
export function packageFromParts(parts: Record<string, string>): LoadedPackage {
  const map = new Map(Object.entries(parts));
  const mainDocumentXml = map.get(MAIN_DOCUMENT_PART);
  if (mainDocumentXml === undefined) {
    throw new Error('packageFromParts requires a word/document.xml part');
  }
  return { parts: map, mainDocumentXml };
}

export type PartResolution =
  | { ok: true; partName: string; xml: string }
  | { ok: false; error: string };

interface RelationshipEntry {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

function parseMainPartRelationships(pkg: LoadedPackage): RelationshipEntry[] {
  const relsXml = pkg.parts.get('word/_rels/document.xml.rels');
  if (!relsXml) return [];
  const dom = new DOMParser().parseFromString(relsXml, 'text/xml');
  const nodes = dom.getElementsByTagNameNS(RELS_NS, 'Relationship');
  const out: RelationshipEntry[] = [];
  for (let i = 0; i < nodes.length; i++) {
    const el = nodes.item(i)!;
    out.push({
      id: el.getAttribute('Id') ?? '',
      type: el.getAttribute('Type') ?? '',
      target: el.getAttribute('Target') ?? '',
      external: (el.getAttribute('TargetMode') ?? 'Internal') === 'External',
    });
  }
  return out;
}

/**
 * Resolve an internal relationship target (relative to word/) to a package part
 * name, normalizing a leading `/` (package-absolute) and any `../` segments.
 */
export function resolveTargetToPartName(target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const segments = 'word'.split('/').concat(target.split('/'));
  const stack: string[] = [];
  for (const seg of segments) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') stack.pop();
    else stack.push(seg);
  }
  return stack.join('/');
}

/** Singleton lookup of a part by OPC relationship Type (styles, numbering, comments). */
export function resolvePartByRelationshipType(
  pkg: LoadedPackage,
  typeUri: string
): PartResolution {
  const matches = parseMainPartRelationships(pkg).filter(
    (r) => r.type === typeUri && !r.external
  );
  if (matches.length === 0) {
    return { ok: false, error: `no internal part with relationship type ${typeUri}` };
  }
  if (matches.length > 1) {
    return {
      ok: false,
      error: `expected a single part with relationship type ${typeUri}, found ${matches.length}`,
    };
  }
  const partName = resolveTargetToPartName(matches[0].target);
  const xml = pkg.parts.get(partName);
  if (xml === undefined) {
    return { ok: false, error: `relationship resolves to missing part ${partName}` };
  }
  return { ok: true, partName, xml };
}

const selectRefs = xpath.useNamespaces({ w: WML_NS, r: OFFICE_REL_NS });

/**
 * Two-hop resolution for headers/footers: main-document sectPr headerReference/
 * footerReference of the given type -> its r:id -> relationship target -> part.
 * The document's single section is assumed; the first matching reference in
 * document order is used.
 */
export function resolveHeaderFooterPart(
  pkg: LoadedPackage,
  referenceLocal: 'headerReference' | 'footerReference',
  referenceType: string
): PartResolution {
  const dom = new DOMParser().parseFromString(pkg.mainDocumentXml, 'text/xml');
  const refs = selectRefs(
    `//w:sectPr/w:${referenceLocal}[@w:type='${referenceType}']`,
    dom as never
  ) as Element[];
  if (refs.length === 0) {
    return {
      ok: false,
      error: `no ${referenceLocal}[type='${referenceType}'] in any sectPr`,
    };
  }
  const rid = refs[0].getAttributeNS(OFFICE_REL_NS, 'id');
  if (!rid) {
    return { ok: false, error: `${referenceLocal} has no r:id` };
  }
  const rel = parseMainPartRelationships(pkg).find((r) => r.id === rid && !r.external);
  if (!rel) {
    return { ok: false, error: `${referenceLocal} r:id ${rid} resolves to no internal relationship` };
  }
  // The r:id must join to a relationship of the matching header/footer type;
  // resolving a reference through, say, a styles relationship is malformed.
  const expectedType = referenceLocal === 'headerReference' ? HEADER_REL_TYPE : FOOTER_REL_TYPE;
  if (rel.type !== expectedType) {
    return {
      ok: false,
      error: `${referenceLocal} r:id ${rid} resolves to relationship type ${rel.type}, expected ${expectedType}`,
    };
  }
  const partName = resolveTargetToPartName(rel.target);
  const xml = pkg.parts.get(partName);
  if (xml === undefined) {
    return { ok: false, error: `relationship resolves to missing part ${partName}` };
  }
  return { ok: true, partName, xml };
}

/** External-hyperlink relationship target for a given r:id, or null. */
export function externalRelationshipTarget(
  pkg: LoadedPackage,
  rid: string
): string | null {
  const rel = parseMainPartRelationships(pkg).find((r) => r.id === rid);
  if (!rel || !rel.external) return null;
  return rel.target;
}

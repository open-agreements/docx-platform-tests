import { DOMParser } from '@xmldom/xmldom';

export const WML_NS =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

/**
 * Canonical tree: elements keyed by {namespaceURI}localName, attributes
 * sorted, implementation-chosen identifiers stripped, whitespace-only text
 * dropped in element-only content, adjacent identically-formatted runs
 * merged. See docs/scenario-dsl.md "Canonicalization".
 */
interface CanonElement {
  kind: 'element';
  name: string;
  attributes: Array<[string, string]>;
  children: CanonNode[];
}
interface CanonText {
  kind: 'text';
  text: string;
}
type CanonNode = CanonElement | CanonText;

const REVISION_WRAPPER_LOCALS = new Set([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'cellIns',
  'cellDel',
  'cellMerge',
]);

function isWml(node: { namespaceURI?: string | null }, local: string, name: string): boolean {
  return node.namespaceURI === WML_NS && local === name;
}

function canonAttributes(el: Element): Array<[string, string]> {
  const out: Array<[string, string]> = [];
  const isRevisionWrapper =
    el.namespaceURI === WML_NS && REVISION_WRAPPER_LOCALS.has(el.localName ?? '');
  for (let i = 0; i < el.attributes.length; i++) {
    const attr = el.attributes.item(i)!;
    if (attr.namespaceURI === XMLNS_NS || attr.name === 'xmlns') continue;
    const local = attr.localName ?? attr.name;
    if (attr.namespaceURI === WML_NS && local.startsWith('rsid')) continue;
    if (isRevisionWrapper && attr.namespaceURI === WML_NS && local === 'id') continue;
    // xml:space is consumed by text normalization, not compared.
    if (attr.namespaceURI === XML_NS && local === 'space') continue;
    out.push([`{${attr.namespaceURI ?? ''}}${local}`, attr.value]);
  }
  out.sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
  return out;
}

function preservesSpace(el: Element): boolean {
  return el.getAttributeNS(XML_NS, 'space') === 'preserve';
}

function canonElement(el: Element): CanonElement {
  const local = el.localName ?? el.nodeName;
  const children: CanonNode[] = [];
  let hasElementChild = false;
  for (let i = 0; i < el.childNodes.length; i++) {
    if (el.childNodes.item(i)!.nodeType === 1) hasElementChild = true;
  }
  for (let i = 0; i < el.childNodes.length; i++) {
    const child = el.childNodes.item(i)!;
    if (child.nodeType === 1) {
      children.push(canonElement(child as Element));
    } else if (child.nodeType === 3 || child.nodeType === 4) {
      const raw = child.nodeValue ?? '';
      if (hasElementChild && raw.trim() === '') continue; // element-only content
      children.push({ kind: 'text', text: raw });
    }
  }
  // Apply xml:space semantics to text-bearing WML leaves, then drop the attr.
  if (isWml(el, local, 't') || isWml(el, local, 'delText')) {
    const text = children
      .filter((c): c is CanonText => c.kind === 'text')
      .map((c) => c.text)
      .join('');
    const normalized = preservesSpace(el) ? text : text.trim();
    return {
      kind: 'element',
      name: `{${el.namespaceURI ?? ''}}${local}`,
      attributes: canonAttributes(el),
      children: [{ kind: 'text', text: normalized }],
    };
  }
  return {
    kind: 'element',
    name: `{${el.namespaceURI ?? ''}}${local}`,
    attributes: canonAttributes(el),
    children: mergeAdjacentIdenticalRuns(children),
  };
}

function runProperties(run: CanonElement): string {
  const rPr = run.children.find(
    (c): c is CanonElement => c.kind === 'element' && c.name === `{${WML_NS}}rPr`
  );
  return rPr ? JSON.stringify(rPr) : '';
}

function runContent(run: CanonElement): CanonNode[] {
  return run.children.filter(
    (c) => !(c.kind === 'element' && c.name === `{${WML_NS}}rPr`)
  );
}

/**
 * Run splitting is legal in WordprocessingML; implementations merge or split
 * runs freely during edits. Adjacent runs with identical properties compare
 * as one run, and their adjacent text leaves compare as one leaf.
 */
function mergeAdjacentIdenticalRuns(children: CanonNode[]): CanonNode[] {
  const out: CanonNode[] = [];
  for (const child of children) {
    const prev = out[out.length - 1];
    if (
      child.kind === 'element' &&
      child.name === `{${WML_NS}}r` &&
      prev?.kind === 'element' &&
      prev.name === `{${WML_NS}}r` &&
      runProperties(prev) === runProperties(child) &&
      prev.attributes.length === 0 &&
      child.attributes.length === 0
    ) {
      prev.children = collapseAdjacentTextLeaves([
        ...prev.children,
        ...runContent(child),
      ]);
      continue;
    }
    out.push(child);
  }
  return out;
}

function collapseAdjacentTextLeaves(children: CanonNode[]): CanonNode[] {
  const out: CanonNode[] = [];
  for (const child of children) {
    const prev = out[out.length - 1];
    if (
      child.kind === 'element' &&
      child.name === `{${WML_NS}}t` &&
      prev?.kind === 'element' &&
      prev.name === `{${WML_NS}}t`
    ) {
      const prevText = prev.children[0]?.kind === 'text' ? prev.children[0].text : '';
      const childText = child.children[0]?.kind === 'text' ? child.children[0].text : '';
      prev.children = [{ kind: 'text', text: prevText + childText }];
      continue;
    }
    out.push(child);
  }
  return out;
}

export function canonicalizeDocumentXml(documentXml: string): string {
  const dom = new DOMParser().parseFromString(documentXml, 'text/xml');
  const root = dom.documentElement;
  if (!root) throw new Error('document.xml has no root element');
  return JSON.stringify(canonElement(root as unknown as Element), null, 1);
}

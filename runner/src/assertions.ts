import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { DOMParser } from '@xmldom/xmldom';
import xpath from 'xpath';
import { canonicalizeDocumentXml, WML_NS } from './canonicalize.js';
import {
  externalRelationshipTarget,
  resolveHeaderFooterPart,
  resolvePartByRelationshipType,
  type LoadedPackage,
  type PartResolution,
} from './docx.js';
import type { AssertedPart, AssertionResult, ScenarioAssertion } from './types.js';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const OFFICE_REL_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const COMMENTS_REL_TYPE = `${OFFICE_REL_NS}/comments`;
const NUMBERING_REL_TYPE = `${OFFICE_REL_NS}/numbering`;
const STYLES_REL_TYPE = `${OFFICE_REL_NS}/styles`;
// Bound in every scenario xpath: w: for WordprocessingML, r: for relationship
// ids (@r:id on hyperlinks and header/footer references).
const selectRefs = xpath.useNamespaces({ w: WML_NS, r: OFFICE_REL_NS });
const WML_SCHEMA_PATH_ENV = 'DPT_WML_SCHEMA_PATH';
const XMLLINT_BIN_ENV = 'DPT_XMLLINT_BIN';

function parseDom(documentXml: string): Document {
  return new DOMParser().parseFromString(documentXml, 'text/xml') as unknown as Document;
}

/**
 * The body text projection defined in docs/scenario-dsl.md: per-paragraph
 * concatenation of w:t text (w:delText excluded, xml:space honored),
 * paragraphs joined with '\n'. Scoped to w:p under w:body in document order
 * (table-cell paragraphs are included; anything outside the body is not).
 */
export function projectBodyText(documentXml: string): string {
  const dom = parseDom(documentXml);
  const body = dom.getElementsByTagNameNS(WML_NS, 'body').item(0);
  const scope: Pick<Element, 'getElementsByTagNameNS'> = body ?? dom;
  const paragraphs = scope.getElementsByTagNameNS(WML_NS, 'p');
  const lines: string[] = [];
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs.item(i)!;
    const texts = p.getElementsByTagNameNS(WML_NS, 't');
    let line = '';
    for (let j = 0; j < texts.length; j++) {
      const t = texts.item(j)!;
      const raw = t.textContent ?? '';
      line += t.getAttributeNS(XML_NS, 'space') === 'preserve' ? raw : raw.trim();
    }
    lines.push(line);
  }
  return lines.join('\n');
}

/**
 * Resolve which part XML an xpath/schema assertion runs against. Absent selector
 * (or mainDocumentPart) is the main document. Relationship/header/footer
 * resolution failures surface as an assertion failure with a diagnostic — that
 * failure IS the conformance signal (e.g. "no comments part").
 */
function resolveAssertedPart(
  pkg: LoadedPackage,
  assertedPart: AssertedPart | undefined
): PartResolution {
  if (!assertedPart || assertedPart.partResolution === 'mainDocumentPart') {
    return { ok: true, partName: 'word/document.xml', xml: pkg.mainDocumentXml };
  }
  switch (assertedPart.partResolution) {
    case 'relationshipFromMainPart':
      return resolvePartByRelationshipType(pkg, assertedPart.relationshipTypeUri);
    case 'headerReference':
      return resolveHeaderFooterPart(pkg, 'headerReference', assertedPart.headerReferenceType);
    case 'footerReference':
      return resolveHeaderFooterPart(pkg, 'footerReference', assertedPart.footerReferenceType);
    default:
      return {
        ok: false,
        error: `unknown partResolution ${String((assertedPart as AssertedPart).partResolution)}`,
      };
  }
}

function validateAgainstWmlSchema(documentXml: string): AssertionResult {
  const schemaPath = process.env[WML_SCHEMA_PATH_ENV];
  if (!schemaPath) {
    return {
      assertionKind: 'schemaValidAgainstWml',
      passed: false,
      detail: `schema validation requires ${WML_SCHEMA_PATH_ENV} pointing to a WordprocessingML XSD entry point`,
    };
  }

  const xmllint = process.env[XMLLINT_BIN_ENV] ?? 'xmllint';
  const result = spawnSync(xmllint, ['--noout', '--schema', schemaPath, '-'], {
    input: documentXml,
    encoding: 'utf8',
    timeout: 120_000,
  });
  if (result.error) {
    return {
      assertionKind: 'schemaValidAgainstWml',
      passed: false,
      detail: `${xmllint} failed to run: ${String(result.error)}`,
    };
  }
  const output = `${result.stdout}${result.stderr}`.trim();
  return {
    assertionKind: 'schemaValidAgainstWml',
    passed: result.status === 0,
    detail:
      result.status === 0
        ? `valid against ${schemaPath}`
        : `xmllint exit ${result.status}: ${output.slice(0, 2000) || 'no diagnostics'}`,
  };
}

function hyperlinkText(hyperlink: Element): string {
  const texts = hyperlink.getElementsByTagNameNS(WML_NS, 't');
  let out = '';
  for (let i = 0; i < texts.length; i++) out += texts.item(i)!.textContent ?? '';
  return out;
}

function getWAttr(el: Element, localName: string): string | null {
  return el.getAttributeNS(WML_NS, localName) ?? el.getAttribute(`w:${localName}`);
}

function xpathLiteral(value: string): string {
  if (!value.includes("'")) return `'${value}'`;
  if (!value.includes('"')) return `"${value}"`;
  return `concat(${value
    .split("'")
    .map((part) => `'${part}'`)
    .join(`,"'",`)})`;
}

function elementText(el: Element): string {
  const texts = el.getElementsByTagNameNS(WML_NS, 't');
  let out = '';
  for (let i = 0; i < texts.length; i++) out += texts.item(i)!.textContent ?? '';
  return out;
}

function elementChildren(el: Element, localName: string): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (
      n.nodeType === 1 &&
      (n as Element).namespaceURI === WML_NS &&
      (n as Element).localName === localName
    ) {
      out.push(n as Element);
    }
  }
  return out;
}

function firstElementChild(el: Element, localName: string): Element | null {
  return elementChildren(el, localName)[0] ?? null;
}

function descendantElements(el: Element, localName: string): Element[] {
  const nodes = el.getElementsByTagNameNS(WML_NS, localName);
  const out: Element[] = [];
  for (let i = 0; i < nodes.length; i++) out.push(nodes.item(i)!);
  return out;
}

function evaluateHyperlinkResolvesToExternalUrl(
  assertion: ScenarioAssertion,
  pkg: LoadedPackage
): AssertionResult {
  const dom = parseDom(pkg.mainDocumentXml);
  const links = selectRefs('//w:hyperlink', dom as never) as Element[];
  const wanted = assertion.hyperlinkDisplayText ?? '';
  const wantedUrl = assertion.expectedTargetUrl ?? '';
  const seen: string[] = [];
  for (const link of links) {
    const text = hyperlinkText(link).trim();
    const rid = link.getAttributeNS(OFFICE_REL_NS, 'id');
    if (text !== wanted) continue;
    if (!rid) {
      seen.push(`'${text}' (no r:id)`);
      continue;
    }
    const target = externalRelationshipTarget(pkg, rid);
    seen.push(`'${text}' -> ${target ?? '(unresolved/internal)'}`);
    if (target === wantedUrl) {
      return {
        assertionKind: assertion.assertionKind,
        passed: true,
        detail: `hyperlink '${wanted}' resolves to external ${wantedUrl}`,
      };
    }
  }
  return {
    assertionKind: assertion.assertionKind,
    passed: false,
    detail: `no w:hyperlink with text '${wanted}' resolving to external ${wantedUrl}; saw [${seen.join(', ')}]`,
  };
}

function collectAnchorTextForCommentId(root: Element, commentId: string): {
  hasReference: boolean;
  anchorText: string;
} {
  let inside = false;
  let hasReference = false;
  let anchorText = '';

  function visit(node: Node): void {
    if (node.nodeType !== 1) return;
    const el = node as Element;
    if (el.namespaceURI === WML_NS) {
      if (el.localName === 'commentReference' && getWAttr(el, 'id') === commentId) {
        hasReference = true;
      }
      if (el.localName === 'commentRangeStart' && getWAttr(el, 'id') === commentId) {
        inside = true;
        return;
      }
      if (el.localName === 'commentRangeEnd' && getWAttr(el, 'id') === commentId) {
        inside = false;
        return;
      }
      if (inside && el.localName === 't') {
        anchorText += el.textContent ?? '';
      }
    }
    for (let child = el.firstChild; child; child = child.nextSibling) {
      visit(child);
    }
  }

  visit(root);
  return { hasReference, anchorText };
}

function evaluateCommentExistsWithTextAndAnchor(
  assertion: ScenarioAssertion,
  pkg: LoadedPackage
): AssertionResult {
  const commentsPart = resolvePartByRelationshipType(pkg, COMMENTS_REL_TYPE);
  if (!commentsPart.ok) {
    return { assertionKind: assertion.assertionKind, passed: false, detail: commentsPart.error };
  }

  const wantedText = assertion.expectedCommentText ?? '';
  const commentsDom = parseDom(commentsPart.xml);
  const comments = selectRefs('//w:comment', commentsDom as never) as Element[];
  const matchingComments = comments.filter((comment) => {
    if (!elementText(comment).includes(wantedText)) return false;
    if (assertion.expectedAuthorName === undefined) return true;
    return getWAttr(comment, 'author') === assertion.expectedAuthorName;
  });
  if (matchingComments.length === 0) {
    return {
      assertionKind: assertion.assertionKind,
      passed: false,
      detail:
        `no comment in ${commentsPart.partName} contains '${wantedText}'` +
        (assertion.expectedAuthorName ? ` by '${assertion.expectedAuthorName}'` : ''),
    };
  }

  const mainDom = parseDom(pkg.mainDocumentXml);
  const body = mainDom.getElementsByTagNameNS(WML_NS, 'body').item(0) ?? mainDom.documentElement;
  const seen: string[] = [];
  for (const comment of matchingComments) {
    const id = getWAttr(comment, 'id');
    if (!id) {
      seen.push('(matching comment with no w:id)');
      continue;
    }
    const anchor = collectAnchorTextForCommentId(body, id);
    seen.push(`id ${id}: reference=${anchor.hasReference}; anchor='${anchor.anchorText}'`);
    const anchorMatches =
      assertion.expectedAnchorText === undefined ||
      anchor.anchorText.includes(assertion.expectedAnchorText);
    if (anchor.hasReference && anchorMatches) {
      return {
        assertionKind: assertion.assertionKind,
        passed: true,
        detail:
          `comment '${wantedText}' is referenced by matching anchors` +
          (assertion.expectedAnchorText ? ` around '${assertion.expectedAnchorText}'` : ''),
      };
    }
  }
  return {
    assertionKind: assertion.assertionKind,
    passed: false,
    detail: `matching comment text exists, but no joined main-document anchor matched; saw [${seen.join(', ')}]`,
  };
}

interface EffectiveNumPr {
  numId: string;
  ilvl: string;
}

function paragraphEffectiveNumPr(paragraph: Element, stylesXml: string | null): EffectiveNumPr | null {
  const pPr = firstElementChild(paragraph, 'pPr');
  if (!pPr) return null;

  const directNumPr = firstElementChild(pPr, 'numPr');
  if (directNumPr) {
    const numId = firstElementChild(directNumPr, 'numId');
    const ilvl = firstElementChild(directNumPr, 'ilvl');
    const numIdVal = numId ? getWAttr(numId, 'val') : null;
    if (numIdVal) return { numId: numIdVal, ilvl: ilvl ? getWAttr(ilvl, 'val') ?? '0' : '0' };
  }

  const pStyle = firstElementChild(pPr, 'pStyle');
  const styleId = pStyle ? getWAttr(pStyle, 'val') : null;
  if (!styleId || !stylesXml) return null;

  const stylesDom = parseDom(stylesXml);
  const visited = new Set<string>();
  let current: string | null = styleId;
  while (current && !visited.has(current)) {
    visited.add(current);
    const styles = selectRefs(
      `//w:style[@w:type='paragraph' and @w:styleId=${xpathLiteral(current)}]`,
      stylesDom as never
    ) as Element[];
    const style = styles[0];
    if (!style) return null;
    const stylePPr = firstElementChild(style, 'pPr');
    const styleNumPr = stylePPr ? firstElementChild(stylePPr, 'numPr') : null;
    if (styleNumPr) {
      const numId = firstElementChild(styleNumPr, 'numId');
      const ilvl = firstElementChild(styleNumPr, 'ilvl');
      const numIdVal = numId ? getWAttr(numId, 'val') : null;
      if (numIdVal) return { numId: numIdVal, ilvl: ilvl ? getWAttr(ilvl, 'val') ?? '0' : '0' };
    }
    const basedOn = firstElementChild(style, 'basedOn');
    current = basedOn ? getWAttr(basedOn, 'val') : null;
  }
  return null;
}

function numberingFormat(numberingXml: string, numPr: EffectiveNumPr): string | null {
  const numberingDom = parseDom(numberingXml);
  const num = (selectRefs(
    `//w:num[@w:numId=${xpathLiteral(numPr.numId)}]`,
    numberingDom as never
  ) as Element[])[0];
  if (!num) return null;

  const override = elementChildren(num, 'lvlOverride').find((el) => getWAttr(el, 'ilvl') === numPr.ilvl);
  const overrideLvl = override ? firstElementChild(override, 'lvl') : null;
  const overrideNumFmt = overrideLvl ? firstElementChild(overrideLvl, 'numFmt') : null;
  if (overrideNumFmt) return getWAttr(overrideNumFmt, 'val');

  const abstractNumId = firstElementChild(num, 'abstractNumId');
  const abstractId = abstractNumId ? getWAttr(abstractNumId, 'val') : null;
  if (!abstractId) return null;
  const abstractNum = (selectRefs(
    `//w:abstractNum[@w:abstractNumId=${xpathLiteral(abstractId)}]`,
    numberingDom as never
  ) as Element[])[0];
  if (!abstractNum) return null;
  const lvl = elementChildren(abstractNum, 'lvl').find((el) => getWAttr(el, 'ilvl') === numPr.ilvl);
  if (!lvl) return null;
  const numFmt = firstElementChild(lvl, 'numFmt');
  return numFmt ? getWAttr(numFmt, 'val') : null;
}

function evaluateParagraphNumberingResolvesToFormat(
  assertion: ScenarioAssertion,
  pkg: LoadedPackage
): AssertionResult {
  const numberingPart = resolvePartByRelationshipType(pkg, NUMBERING_REL_TYPE);
  if (!numberingPart.ok) {
    return { assertionKind: assertion.assertionKind, passed: false, detail: numberingPart.error };
  }
  const stylesPart = resolvePartByRelationshipType(pkg, STYLES_REL_TYPE);
  const stylesXml = stylesPart.ok ? stylesPart.xml : null;
  const wantedAnchor = assertion.anchorText ?? '';
  const wantedFormat = assertion.expectedNumberFormat ?? '';
  const dom = parseDom(pkg.mainDocumentXml);
  const paragraphs = descendantElements(dom.documentElement, 'p').filter((p) =>
    elementText(p).includes(wantedAnchor)
  );
  const seen: string[] = [];
  for (const paragraph of paragraphs) {
    const numPr = paragraphEffectiveNumPr(paragraph, stylesXml);
    if (!numPr) {
      seen.push(`'${elementText(paragraph)}' has no effective numPr`);
      continue;
    }
    const fmt = numberingFormat(numberingPart.xml, numPr);
    seen.push(
      `'${elementText(paragraph)}' -> numId ${numPr.numId} ilvl ${numPr.ilvl} fmt ${fmt ?? '(unresolved)'}`
    );
    if (fmt === wantedFormat) {
      return {
        assertionKind: assertion.assertionKind,
        passed: true,
        detail: `paragraph '${wantedAnchor}' resolves to numbering format ${wantedFormat}`,
      };
    }
  }
  return {
    assertionKind: assertion.assertionKind,
    passed: false,
    detail: `no paragraph containing '${wantedAnchor}' resolved to numbering format ${wantedFormat}; saw [${seen.join(', ')}]`,
  };
}

export function evaluateAssertion(
  assertion: ScenarioAssertion,
  pkg: LoadedPackage,
  scenarioDir: string
): AssertionResult {
  switch (assertion.assertionKind) {
    case 'xpathQueryCount':
    case 'xpathQueryExists': {
      const part = resolveAssertedPart(pkg, assertion.assertedPart);
      if (!part.ok) {
        return { assertionKind: assertion.assertionKind, passed: false, detail: part.error };
      }
      const nodes = selectRefs(
        assertion.xpathExpression!,
        parseDom(part.xml) as never
      ) as unknown[];
      if (assertion.assertionKind === 'xpathQueryCount') {
        return {
          assertionKind: assertion.assertionKind,
          passed: nodes.length === assertion.expectedCount,
          detail: `${assertion.xpathExpression} matched ${nodes.length} node(s) in ${part.partName}; expected ${assertion.expectedCount}`,
        };
      }
      return {
        assertionKind: assertion.assertionKind,
        passed: nodes.length >= 1,
        detail: `${assertion.xpathExpression} matched ${nodes.length} node(s) in ${part.partName}; expected at least 1`,
      };
    }
    case 'documentTextContainsAtOffset': {
      const projection = projectBodyText(pkg.mainDocumentXml);
      const actualOffset = projection.indexOf(assertion.expectedSubstring!);
      const passed = actualOffset === assertion.expectedOffset;
      return {
        assertionKind: assertion.assertionKind,
        passed,
        detail: passed
          ? `'${assertion.expectedSubstring}' found at offset ${actualOffset}`
          : `'${assertion.expectedSubstring}' ${
              actualOffset === -1 ? 'not found' : `found at offset ${actualOffset}`
            }; expected offset ${assertion.expectedOffset}. Projection: ${JSON.stringify(projection)}`,
      };
    }
    case 'canonicalXmlEquals': {
      const expectedXml = readFileSync(
        join(scenarioDir, assertion.expectedDocumentPath!),
        'utf8'
      );
      const actualCanonical = canonicalizeDocumentXml(pkg.mainDocumentXml);
      const expectedCanonical = canonicalizeDocumentXml(expectedXml);
      const passed = actualCanonical === expectedCanonical;
      return {
        assertionKind: assertion.assertionKind,
        passed,
        detail: passed
          ? 'canonical forms are identical'
          : `canonical forms differ.\n--- expected\n${expectedCanonical}\n--- actual\n${actualCanonical}`,
      };
    }
    case 'schemaValidAgainstWml': {
      const part = resolveAssertedPart(pkg, assertion.assertedPart);
      if (!part.ok) {
        return { assertionKind: assertion.assertionKind, passed: false, detail: part.error };
      }
      return validateAgainstWmlSchema(part.xml);
    }
    case 'hyperlinkResolvesToExternalUrl': {
      return evaluateHyperlinkResolvesToExternalUrl(assertion, pkg);
    }
    case 'commentExistsWithTextAndAnchor': {
      return evaluateCommentExistsWithTextAndAnchor(assertion, pkg);
    }
    case 'paragraphNumberingResolvesToFormat': {
      return evaluateParagraphNumberingResolvesToFormat(assertion, pkg);
    }
    default:
      return {
        assertionKind: String((assertion as ScenarioAssertion).assertionKind),
        passed: false,
        detail: 'unknown assertion kind',
      };
  }
}

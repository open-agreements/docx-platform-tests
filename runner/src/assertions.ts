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
    default:
      return {
        assertionKind: String((assertion as ScenarioAssertion).assertionKind),
        passed: false,
        detail: 'unknown assertion kind',
      };
  }
}

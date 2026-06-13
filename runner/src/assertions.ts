import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { DOMParser } from '@xmldom/xmldom';
import xpath from 'xpath';
import { canonicalizeDocumentXml, WML_NS } from './canonicalize.js';
import type { AssertionResult, ScenarioAssertion } from './types.js';

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const selectWithW = xpath.useNamespaces({ w: WML_NS });
const WML_SCHEMA_PATH_ENV = 'DPT_WML_SCHEMA_PATH';
const XMLLINT_BIN_ENV = 'DPT_XMLLINT_BIN';

function parseDom(documentXml: string): Document {
  return new DOMParser().parseFromString(documentXml, 'text/xml') as unknown as Document;
}

/**
 * The body text projection defined in docs/scenario-dsl.md: per-paragraph
 * concatenation of w:t text (w:delText excluded, xml:space honored),
 * paragraphs joined with '\n'.
 */
export function projectBodyText(documentXml: string): string {
  const dom = parseDom(documentXml);
  const paragraphs = dom.getElementsByTagNameNS(WML_NS, 'p');
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

export function evaluateAssertion(
  assertion: ScenarioAssertion,
  outputDocumentXml: string,
  scenarioDir: string
): AssertionResult {
  switch (assertion.assertionKind) {
    case 'xpathQueryCount': {
      const nodes = selectWithW(
        assertion.xpathExpression!,
        parseDom(outputDocumentXml) as never
      ) as unknown[];
      const passed = nodes.length === assertion.expectedCount;
      return {
        assertionKind: assertion.assertionKind,
        passed,
        detail: `${assertion.xpathExpression} matched ${nodes.length} node(s); expected ${assertion.expectedCount}`,
      };
    }
    case 'xpathQueryExists': {
      const nodes = selectWithW(
        assertion.xpathExpression!,
        parseDom(outputDocumentXml) as never
      ) as unknown[];
      return {
        assertionKind: assertion.assertionKind,
        passed: nodes.length >= 1,
        detail: `${assertion.xpathExpression} matched ${nodes.length} node(s); expected at least 1`,
      };
    }
    case 'documentTextContainsAtOffset': {
      const projection = projectBodyText(outputDocumentXml);
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
      const actualCanonical = canonicalizeDocumentXml(outputDocumentXml);
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
      return validateAgainstWmlSchema(outputDocumentXml);
    }
    default:
      return {
        assertionKind: String((assertion as ScenarioAssertion).assertionKind),
        passed: false,
        detail: 'unknown assertion kind',
      };
  }
}

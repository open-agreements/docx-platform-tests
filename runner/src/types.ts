export interface SpecCitation {
  standard: string;
  edition: number;
  part: number;
  section: string;
  clauseTitle: string;
}

export interface MicrosoftExtensionCitation {
  standard: 'MS-DOCX';
  section: string;
  clauseTitle: string;
}

export interface OperationDescriptor {
  operationName: string;
  [parameter: string]: unknown;
}

/**
 * Which package part an assertion is evaluated against. Absent means the main
 * document part (word/document.xml) — every DSL <= 1.2 manifest omits it, so the
 * default preserves their behavior exactly. Parts other than the main document
 * are addressed by OPC relationship, never by hardcoded path, because part names
 * (styles.xml, header1.xml vs header3.xml, …) are implementation-chosen.
 */
export type AssertedPart =
  | { partResolution: 'mainDocumentPart'; expectedContentType?: string }
  | {
      partResolution: 'relationshipFromMainPart';
      relationshipTypeUri: string;
      expectedContentType?: string;
    }
  | { partResolution: 'headerReference'; headerReferenceType: string; expectedContentType?: string }
  | { partResolution: 'footerReference'; footerReferenceType: string; expectedContentType?: string };

export interface ScenarioAssertion {
  assertionKind:
    | 'xpathQueryCount'
    | 'xpathQueryExists'
    | 'xpathElementTextEquals'
    | 'ignorableNamespaceDeclared'
    | 'documentTextContainsAtOffset'
    | 'canonicalXmlEquals'
    | 'schemaValidAgainstWml'
    | 'hyperlinkResolvesToExternalUrl'
    | 'commentExistsWithTextAndAnchor'
    | 'paragraphNumberingResolvesToFormat';
  xpathExpression?: string;
  expectedCount?: number;
  expectedText?: string;
  expectedNamespaceUri?: string;
  expectedSubstring?: string;
  expectedOffset?: number;
  expectedDocumentPath?: string;
  // hyperlinkResolvesToExternalUrl
  hyperlinkDisplayText?: string;
  expectedTargetUrl?: string;
  // commentExistsWithTextAndAnchor
  expectedCommentText?: string;
  expectedAuthorName?: string;
  expectedAnchorText?: string;
  // paragraphNumberingResolvesToFormat
  anchorText?: string;
  expectedNumberFormat?: string;
  // Part selector for xpathQueryCount / xpathQueryExists / schemaValidAgainstWml.
  // documentTextContainsAtOffset and canonicalXmlEquals are always main-part.
  assertedPart?: AssertedPart;
}

export interface ScenarioManifest {
  dslVersion: string;
  scenarioId: string;
  scenarioTitle: string;
  specCitation: SpecCitation;
  secondarySpecCitations?: SpecCitation[];
  microsoftExtensionCitations?: MicrosoftExtensionCitation[];
  wordBehaviorNote: string | null;
  inputDocumentPath: string;
  operationDescriptor: OperationDescriptor;
  assertionList: ScenarioAssertion[];
}

export interface LoadedScenario {
  manifest: ScenarioManifest;
  scenarioDir: string;
}

export interface AdapterRegistration {
  adapterName: string;
  adapterCommand: string[];
  adapterVersionCommand?: string[];
}

export interface AdapterRegistry {
  protocolVersion: number;
  adapters: AdapterRegistration[];
}

export type OutcomeStatus =
  | 'pass'
  | 'pass-divergent'
  | 'fail'
  | 'invariant-pass'
  | 'invariant-fail'
  | 'unsupported'
  | 'error'
  | 'protocol-mismatch';

export interface AssertionResult {
  assertionKind: string;
  passed: boolean;
  detail: string;
}

export type ScenarioOracleKind = 'ecma-conformance' | 'metamorphic-invariant';

export interface ScenarioOutcome {
  status: OutcomeStatus;
  reason?: string;
  assertionResults?: AssertionResult[];
}

export interface ResultsDocument {
  schemaVersion: number;
  runTimestamp: string;
  dslVersion: string;
  protocolVersion: number;
  implementations: Array<{ adapterName: string; adapterVersion: string }>;
  results: Array<{
    scenarioGroup: string;
    scenarioId: string;
    scenarioTitle: string;
    oracleKind: ScenarioOracleKind;
    specCitation: SpecCitation;
    microsoftExtensionCitations?: MicrosoftExtensionCitation[];
    outcomes: Record<string, ScenarioOutcome>;
  }>;
}

export interface SpecCitation {
  standard: string;
  edition: number;
  part: number;
  section: string;
  clauseTitle: string;
}

export interface OperationDescriptor {
  operationName: string;
  [parameter: string]: unknown;
}

export interface ScenarioAssertion {
  assertionKind:
    | 'xpathQueryCount'
    | 'xpathQueryExists'
    | 'documentTextContainsAtOffset'
    | 'canonicalXmlEquals'
    | 'schemaValidAgainstWml';
  xpathExpression?: string;
  expectedCount?: number;
  expectedSubstring?: string;
  expectedOffset?: number;
  expectedDocumentPath?: string;
}

export interface ScenarioManifest {
  dslVersion: string;
  scenarioId: string;
  scenarioTitle: string;
  specCitation: SpecCitation;
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
  | 'unsupported'
  | 'error'
  | 'protocol-mismatch';

export interface AssertionResult {
  assertionKind: string;
  passed: boolean;
  detail: string;
}

export interface ScenarioOutcome {
  status: OutcomeStatus;
  reason?: string;
  assertionResults?: AssertionResult[];
}

export interface ResultsDocument {
  runTimestamp: string;
  dslVersion: string;
  protocolVersion: number;
  implementations: Array<{ adapterName: string; adapterVersion: string }>;
  results: Array<{
    scenarioId: string;
    scenarioTitle: string;
    specCitation: SpecCitation;
    outcomes: Record<string, ScenarioOutcome>;
  }>;
}

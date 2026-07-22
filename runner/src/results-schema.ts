export const RESULTS_SCHEMA_VERSION = 2;

export const RESULTS_SCHEMA = {
  $schema: 'https://json-schema.org/draft/2020-12/schema',
  $id: 'https://open-agreements.github.io/docx-platform-tests/results/results.schema.json',
  title: 'docx-platform-tests results document',
  type: 'object',
  additionalProperties: false,
  required: [
    'schemaVersion',
    'runTimestamp',
    'dslVersion',
    'protocolVersion',
    'implementations',
    'results',
  ],
  properties: {
    schemaVersion: {
      const: RESULTS_SCHEMA_VERSION,
      description: 'Version of the published results JSON contract.',
    },
    runTimestamp: {
      type: 'string',
      pattern: '^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}\\.\\d{3}Z$',
    },
    dslVersion: {
      type: 'string',
      pattern: '^\\d+\\.\\d+$',
    },
    protocolVersion: {
      type: 'integer',
      minimum: 1,
    },
    implementations: {
      type: 'array',
      items: { $ref: '#/$defs/implementation' },
    },
    results: {
      type: 'array',
      items: { $ref: '#/$defs/scenarioResult' },
    },
  },
  $defs: {
    implementation: {
      type: 'object',
      additionalProperties: false,
      required: ['adapterName', 'adapterVersion'],
      properties: {
        adapterName: { type: 'string', minLength: 1 },
        adapterVersion: { type: 'string' },
      },
    },
    specCitation: {
      type: 'object',
      additionalProperties: false,
      required: ['standard', 'edition', 'part', 'section', 'clauseTitle'],
      properties: {
        standard: { type: 'string', minLength: 1 },
        edition: { type: 'integer', minimum: 1 },
        part: { type: 'integer', minimum: 1 },
        section: { type: 'string', minLength: 1 },
        clauseTitle: { type: 'string', minLength: 1 },
      },
    },
    microsoftExtensionCitation: {
      type: 'object',
      additionalProperties: false,
      required: ['standard', 'section', 'clauseTitle'],
      properties: {
        standard: { const: 'MS-DOCX' },
        section: { type: 'string', pattern: '^\\d+(?:\\.\\d+)*$' },
        clauseTitle: { type: 'string', minLength: 1 },
      },
    },
    scenarioResult: {
      type: 'object',
      additionalProperties: false,
      required: ['scenarioGroup', 'scenarioId', 'scenarioTitle', 'specCitation', 'outcomes'],
      properties: {
        scenarioGroup: { type: 'string', minLength: 1 },
        scenarioId: { type: 'string', minLength: 1 },
        scenarioTitle: { type: 'string', minLength: 1 },
        specCitation: { $ref: '#/$defs/specCitation' },
        microsoftExtensionCitations: {
          type: 'array',
          minItems: 1,
          items: { $ref: '#/$defs/microsoftExtensionCitation' },
        },
        outcomes: {
          type: 'object',
          additionalProperties: { $ref: '#/$defs/scenarioOutcome' },
        },
      },
    },
    scenarioOutcome: {
      type: 'object',
      additionalProperties: false,
      required: ['status'],
      properties: {
        status: {
          enum: [
            'pass',
            'pass-divergent',
            'fail',
            'unsupported',
            'error',
            'protocol-mismatch',
          ],
        },
        reason: { type: 'string' },
        assertionResults: {
          type: 'array',
          items: { $ref: '#/$defs/assertionResult' },
        },
      },
    },
    assertionResult: {
      type: 'object',
      additionalProperties: false,
      required: ['assertionKind', 'passed', 'detail'],
      properties: {
        assertionKind: { type: 'string', minLength: 1 },
        passed: { type: 'boolean' },
        detail: { type: 'string' },
      },
    },
  },
} as const;

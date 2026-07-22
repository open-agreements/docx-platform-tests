import type { AssertionResult, OutcomeStatus, ScenarioOracleKind } from './types.js';

export function gradeAssertionResults(
  assertionResults: AssertionResult[],
  oracleKind: ScenarioOracleKind
): OutcomeStatus {
  if (oracleKind === 'metamorphic-invariant') {
    return assertionResults.some((result) => !result.passed)
      ? 'invariant-fail'
      : 'invariant-pass';
  }
  const semanticFailed = assertionResults.some(
    (result) => result.assertionKind !== 'canonicalXmlEquals' && !result.passed
  );
  const canonicalFailed = assertionResults.some(
    (result) => result.assertionKind === 'canonicalXmlEquals' && !result.passed
  );
  return semanticFailed ? 'fail' : canonicalFailed ? 'pass-divergent' : 'pass';
}

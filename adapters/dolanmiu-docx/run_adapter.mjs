#!/usr/bin/env node
import { readFile } from 'node:fs/promises';

const PROTOCOL_VERSION = '1';

function argValue(flag) {
  const index = process.argv.indexOf(flag);
  return index === -1 ? undefined : process.argv[index + 1];
}

function unsupported(reason) {
  console.log(reason);
  process.exit(2);
}

const protocolVersion = argValue('--protocol-version');
const operationPath = argValue('--operation');

if (protocolVersion !== PROTOCOL_VERSION) {
  console.log(`dolanmiu-docx adapter speaks protocol v${PROTOCOL_VERSION}, got ${protocolVersion}`);
  process.exit(3);
}

if (!operationPath || !argValue('--input') || !argValue('--output')) {
  console.error('missing required adapter protocol arguments');
  process.exit(1);
}

const operation = JSON.parse(await readFile(operationPath, 'utf8'));
const operationName = operation.operationName;

if (operationName === 'acceptAllTrackedChanges' || operationName === 'rejectAllTrackedChanges') {
  unsupported('dolanmiu/docx can generate revision markup but exposes no API to accept or reject existing tracked changes');
}

if (operationName === 'replaceFirstTextOccurrence') {
  unsupported(
    'dolanmiu/docx patchDocument targets explicit patch placeholders; protocol requires arbitrary paragraph-local literal search, which would be an adapter-side algorithm'
  );
}

unsupported(`dolanmiu-docx adapter does not implement operation '${operationName}'`);

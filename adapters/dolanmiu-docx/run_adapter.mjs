#!/usr/bin/env node
import { readFile, writeFile } from 'node:fs/promises';
import {
  AlignmentType,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from 'docx';

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
const outputPath = argValue('--output');

function textRun(text, formatting = undefined) {
  return new TextRun({
    text,
    bold: formatting?.bold,
    italics: formatting?.italic,
    size: formatting?.fontSizeHalfPoints,
  });
}

async function writeDocument(document) {
  await writeFile(outputPath, await Packer.toBuffer(document));
}

if (operationName === 'acceptAllTrackedChanges' || operationName === 'rejectAllTrackedChanges') {
  unsupported('dolanmiu/docx can generate revision markup but exposes no API to accept or reject existing tracked changes');
}

if (operationName === 'replaceFirstTextOccurrence') {
  unsupported(
    'dolanmiu/docx patchDocument targets explicit patch placeholders; protocol requires arbitrary paragraph-local literal search, which would be an adapter-side algorithm'
  );
}

if (operationName === 'composeDocumentWithParagraphs') {
  await writeDocument(
    new Document({
      sections: [
        {
          children: operation.paragraphDescriptorList.map(
            (descriptor) =>
              new Paragraph({
                children: [textRun(descriptor.paragraphText, descriptor.runFormatting)],
              })
          ),
        },
      ],
    })
  );
  process.exit(0);
}

if (operationName === 'composeDocumentWithTable') {
  await writeDocument(
    new Document({
      sections: [
        {
          children: [
            new Table({
              rows: operation.tableCellTextRows.map(
                (row) =>
                  new TableRow({
                    children: row.map(
                      (cellText) =>
                        new TableCell({
                          children: [new Paragraph(String(cellText))],
                        })
                    ),
                  })
              ),
            }),
          ],
        },
      ],
    })
  );
  process.exit(0);
}

if (operationName === 'composeDocumentWithHyperlink') {
  await writeDocument(
    new Document({
      sections: [
        {
          children: [
            new Paragraph({
              children: [
                new ExternalHyperlink({
                  children: [new TextRun({ text: operation.hyperlinkDisplayText, style: 'Hyperlink' })],
                  link: operation.hyperlinkTargetUrl,
                }),
              ],
            }),
          ],
        },
      ],
    })
  );
  process.exit(0);
}

if (operationName === 'composeDocumentWithHeaderText') {
  await writeDocument(
    new Document({
      sections: [
        {
          headers: { default: new Header({ children: [new Paragraph(operation.headerText)] }) },
          children: [new Paragraph(operation.bodyText)],
        },
      ],
    })
  );
  process.exit(0);
}

if (operationName === 'composeDocumentWithNumberedList') {
  if (operation.numberFormat !== 'decimal') {
    unsupported('dolanmiu-docx adapter only maps decimal numbered lists');
  }
  await writeDocument(
    new Document({
      numbering: {
        config: [
          {
            reference: 'dpt-numbering',
            levels: [
              {
                level: 0,
                format: LevelFormat.DECIMAL,
                text: '%1.',
                alignment: AlignmentType.START,
              },
            ],
          },
        ],
      },
      sections: [
        {
          children: operation.listItemTexts.map(
            (itemText, index) =>
              new Paragraph({
                text: itemText,
                numbering: {
                  reference: 'dpt-numbering',
                  level: operation.itemLevels?.[index] ?? 0,
                },
              })
          ),
        },
      ],
    })
  );
  process.exit(0);
}

unsupported(`dolanmiu-docx adapter does not implement operation '${operationName}'`);

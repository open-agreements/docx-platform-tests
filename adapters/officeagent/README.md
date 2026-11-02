# OfficeAgent.NET adapter

Wraps [OfficeAgent.NET](https://github.com/ilia-sokolov/OfficeAgent.NET)
(`OfficeAgent.Core` and `OfficeAgent.Word`, MIT-licensed) behind adapter
protocol v1. OfficeAgent is a .NET engine for agent-driven edits to existing
documents: callers find content-verified anchors, then commit a typed
`DocumentPlan` that applies completely or not at all.

**What is library API vs glue:** every edit is one `DocumentPlan` committed
through `OfficeAgentClient`. `FindAsync` resolves the text a descriptor names
(case-sensitive, first hit in document order), `InspectAsync` supplies
paragraph and comment ids, and `CreateBlank` starts the compose scenarios.
The adapter maps descriptor fields onto those calls and plan verbs:

| Operation | Library call |
| --- | --- |
| `replaceFirstTextOccurrence` | `changeText`, `Direct` mode |
| `formatFirstTextOccurrence` | `format` (bold, italic, underline, size) |
| `addCommentOnFirstTextOccurrence` | `comment` with `Add` |
| `removeAllComments` | `comment` with `Remove`, once per comment until inspection lists none |
| `insertParagraphAfterAnchorText`, `appendParagraphWithText` | `insert` after the anchor paragraph, or after the last body paragraph |
| `applyParagraphStyleToAnchor` | `format` with `styleId` on the whole paragraph |
| `setDefaultFooterText` | `headerFooter` with `scope: default` |
| `appendTableRow`, `deleteTableRowAtIndex` | `insertTableRows` at the end, `removeTableRows` |
| `acceptAllTrackedChanges`, `rejectAllTrackedChanges` | `revision` on path `all` |
| `composeDocumentWithParagraphs`, `composeDocumentWithNumberedList`, `composeDocumentWithHeaderText` | `CreateBlank`, then `changeText` / `insert` per paragraph, `format` for run formatting or `listStyle`, `headerFooter` |
| `composeDocumentWithTable` | `CreateBlank`, then `insertTable` (first row as the header row) |

Edits use `Direct` mode because the operations ask for plain edits;
OfficeAgent's default for Word is `Tracked`.

These exit 2, because the library has no verb for them and an adapter-side
implementation would be an algorithm rather than glue:

- `composeDocumentWithHyperlink`: no plan verb creates a hyperlink.
- `composeDocumentWithCompatibilityMode`: no plan verb writes `w:compat`.
- `mergeTableCellsInRow`: no plan verb merges cells.
- `applyNumberingToAnchorParagraph`: OfficeAgent numbers a paragraph by list
  style and writes the numbering definition itself; it cannot attach an
  existing `w:num` by id.
- `setTableCellText`: text edits target paragraphs, and inspection names the
  table a paragraph is in but not its cell.

**Known divergence:** OfficeAgent gives every paragraph a `w14:paraId` when it
saves, because its anchors are paragraph ids. Scenarios that compare
canonical XML therefore report `pass-divergent`: in each of them the output
differs from the expected document only by that attribute.

**Maintenance policy:** best-effort, maintained by the OfficeAgent.NET
project. The package references float on the latest 1.x release and the
weekly CI run rebuilds from NuGet, so upstream breakage surfaces as `error`
cells rather than silent rot.

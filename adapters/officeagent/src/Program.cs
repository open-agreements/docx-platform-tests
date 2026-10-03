// docx-platform-tests adapter protocol v1 entrypoint for OfficeAgent.NET.
//
// Every edit is one OfficeAgent DocumentPlan committed through OfficeAgentClient:
// FindAsync locates the text the descriptor names, InspectAsync supplies paragraph
// and comment ids, and the plan verbs (changeText, comment, insert, format,
// headerFooter, table rows, revision) do the editing. The adapter maps descriptor
// fields onto those calls and nothing more. Operations the library has no verb for
// exit 2 with the reason (see README.md).
using System.Reflection;
using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

const string ProtocolVersion = "1";

static string? Arg(string[] argv, string name)
{
    var index = Array.IndexOf(argv, name);
    return index >= 0 && index + 1 < argv.Length ? argv[index + 1] : null;
}

if (args.Contains("--print-library-version"))
{
    var informational = typeof(OfficeAgentClient).Assembly
        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
    Console.WriteLine($"OfficeAgent.NET {informational?.Split('+')[0] ?? "unknown"}");
    return 0;
}

var protocol = Arg(args, "--protocol-version");
if (protocol != ProtocolVersion)
{
    Console.WriteLine($"officeagent adapter speaks protocol v{ProtocolVersion}, got {protocol}");
    return 3;
}

var operationPath = Arg(args, "--operation");
var inputPath = Arg(args, "--input");
var outputPath = Arg(args, "--output");
if (operationPath is null || inputPath is null || outputPath is null)
{
    Console.Error.WriteLine("missing required --operation/--input/--output argument");
    return 1;
}

using var services = new ServiceCollection().AddWordFormat().AddOfficeAgent().BuildServiceProvider();
var client = services.GetRequiredService<OfficeAgentClient>();
using var descriptor = JsonDocument.Parse(File.ReadAllText(operationPath));
var operation = descriptor.RootElement;

try
{
    var output = await Run(client, operation, File.ReadAllBytes(inputPath));
    File.WriteAllBytes(outputPath, output);
    return 0;
}
catch (UnsupportedException unsupported)
{
    Console.WriteLine(unsupported.Message);
    return 2;
}
catch (Exception failure)
{
    Console.Error.WriteLine(failure.ToString());
    return 1;
}

static async Task<byte[]> Run(OfficeAgentClient client, JsonElement operation, byte[] input)
{
    string Text(string name) => operation.GetProperty(name).GetString()!;
    int Number(string name) => operation.GetProperty(name).GetInt32();

    switch (operation.GetProperty("operationName").GetString())
    {
        case "replaceFirstTextOccurrence":
            return await Commit(client, input, new ChangeTextOp
            {
                Target = await FirstOccurrence(client, input, Text("findText")),
                With = Text("replaceText"),
                Mode = ChangeMode.Direct,
            });

        case "formatFirstTextOccurrence":
            return await Commit(client, input,
                WithRunFormatting(await FirstOccurrence(client, input, Text("findText")), operation.GetProperty("runFormatting")));

        case "addCommentOnFirstTextOccurrence":
            return await Commit(client, input, new CommentOp
            {
                Target = await FirstOccurrence(client, input, Text("anchorText")),
                Action = CommentAction.Add,
                Text = Text("commentText"),
                Author = Text("commentAuthorName"),
                Initials = Text("commentAuthorInitials"),
            });

        case "removeAllComments":
            // Remove takes one comment at a time, and removing a thread's first comment
            // removes its replies, so re-inspect after each removal.
            var document = input;
            while ((await client.InspectAsync(Handle(document))).Nodes.FirstOrDefault(n => n.Kind == "comment") is { } comment)
                document = await Commit(client, document, new CommentOp
                {
                    Target = new NodeAnchor { Kind = "comment", Path = comment.Path },
                    Action = CommentAction.Remove,
                });
            return document;

        case "insertParagraphAfterAnchorText":
            return await Commit(client, input, new InsertOp
            {
                Target = await FirstOccurrence(client, input, Text("anchorText")),
                Position = InsertPosition.After,
                Text = Text("paragraphText"),
                Mode = ChangeMode.Direct,
            });

        case "appendParagraphWithText":
            var last = (await client.InspectAsync(Handle(input))).Paragraphs.LastOrDefault(p => p.In is null)
                ?? throw new InvalidOperationException("the document body has no paragraph to append after");
            return await Commit(client, input, new InsertOp
            {
                Target = new TextSpanAnchor { ParaId = last.ParaId, Expect = last.Text },
                Position = InsertPosition.After,
                Text = Text("paragraphText"),
                Mode = ChangeMode.Direct,
            });

        case "applyParagraphStyleToAnchor":
            return await Commit(client, input, new FormatOp
            {
                Target = await WholeParagraph(client, input, Text("anchorText")),
                StyleId = Text("paragraphStyleId"),
                Mode = ChangeMode.Direct,
            });

        case "setDefaultFooterText":
            return await Commit(client, input, new HeaderFooterOp { Footer = Text("footerText"), Scope = "default" });

        case "appendTableRow":
            return await Commit(client, input, new InsertTableRowsOp
            {
                Target = Table(Number("tableIndex")),
                Rows = new[] { Strings(operation.GetProperty("cellTexts")) },
                Position = TablePosition.End,
                Mode = ChangeMode.Direct,
            });

        case "deleteTableRowAtIndex":
            return await Commit(client, input, new RemoveTableRowsOp
            {
                Target = Table(Number("tableIndex")),
                RowIndices = new[] { Number("rowIndex") },
                Mode = ChangeMode.Direct,
            });

        case "setTableCellText":
            throw new UnsupportedException(
                "OfficeAgent edits text by paragraph anchor; inspection names the table a paragraph is in but not its cell, " +
                "and no plan verb takes cell coordinates for text");

        case "acceptAllTrackedChanges":
        case "rejectAllTrackedChanges":
            return await Commit(client, input, new RevisionOp
            {
                Target = new NodeAnchor { Kind = "revision", Path = "all" },
                Action = operation.GetProperty("operationName").GetString() == "acceptAllTrackedChanges"
                    ? RevisionAction.Accept
                    : RevisionAction.Reject,
            });

        case "composeDocumentWithParagraphs":
            return await Compose(client, operation.GetProperty("paragraphDescriptorList").EnumerateArray()
                .Select(p => (p.GetProperty("paragraphText").GetString()!,
                              p.TryGetProperty("runFormatting", out var formatting) ? formatting : (JsonElement?)null))
                .ToList(), listStyle: null);

        case "composeDocumentWithNumberedList":
            return await Compose(client, Strings(operation.GetProperty("listItemTexts"))
                .Select(text => (text, (JsonElement?)null)).ToList(), listStyle: Text("numberFormat"));

        case "composeDocumentWithHeaderText":
            var withBody = await Compose(client, new List<(string, JsonElement?)> { (Text("bodyText"), null) }, listStyle: null);
            return await Commit(client, withBody, new HeaderFooterOp { Header = Text("headerText"), Scope = "default" });

        case "composeDocumentWithTable":
            var rows = operation.GetProperty("tableCellTextRows").EnumerateArray().Select(Strings).ToList();
            var blank = client.CreateBlank("composed.docx");
            return await Commit(client, blank, new InsertTableOp
            {
                Target = new TextSpanAnchor { ParaId = (await client.InspectAsync(Handle(blank))).Paragraphs[0].ParaId, Expect = "" },
                Position = InsertPosition.After,
                Table = new TableData { Headers = rows[0], Rows = rows.Skip(1).ToList() },
                Mode = ChangeMode.Direct,
            });

        case "composeDocumentWithHyperlink":
            throw new UnsupportedException("OfficeAgent plans have no verb that creates a hyperlink");

        case "composeDocumentWithCompatibilityMode":
            throw new UnsupportedException("OfficeAgent plans have no verb that sets w:compat settings");

        case "mergeTableCellsInRow":
            throw new UnsupportedException("OfficeAgent plans have no verb that merges table cells");

        case "applyNumberingToAnchorParagraph":
            throw new UnsupportedException(
                "OfficeAgent numbers a paragraph by list style and creates the numbering definition itself; " +
                "it has no verb that attaches an existing w:num instance by id");

        case var other:
            throw new UnsupportedException($"officeagent adapter does not implement operation '{other}'");
    }
}

static DocumentHandle Handle(byte[] document) => new StreamHandle(new MemoryStream(document, writable: false), "document.docx");

static NodeAnchor Table(int index) => new() { Kind = "table", Path = $"table#{index}" };

static IReadOnlyList<string> Strings(JsonElement array) => array.EnumerateArray().Select(item => item.GetString()!).ToList();

static async Task<TextSpanAnchor> FirstOccurrence(OfficeAgentClient client, byte[] document, string text)
{
    var hits = await client.FindAsync(Handle(document), new FindQuery(text) { Options = new MatchOptions { CaseSensitive = true } });
    return hits.FirstOrDefault()?.Anchor as TextSpanAnchor
        ?? throw new InvalidOperationException($"text not found: '{text}'");
}

static async Task<TextSpanAnchor> WholeParagraph(OfficeAgentClient client, byte[] document, string text)
{
    var anchor = await FirstOccurrence(client, document, text);
    var paragraph = (await client.InspectAsync(Handle(document))).Paragraphs.First(p => p.ParaId == anchor.ParaId);
    return new TextSpanAnchor { ParaId = paragraph.ParaId, Expect = paragraph.Text };
}

static FormatOp WithRunFormatting(TextSpanAnchor target, JsonElement formatting) => new()
{
    Target = target,
    Bold = formatting.TryGetProperty("bold", out var bold) ? bold.GetBoolean() : null,
    Italic = formatting.TryGetProperty("italic", out var italic) ? italic.GetBoolean() : null,
    Underline = formatting.TryGetProperty("underline", out var underline) ? underline.GetBoolean() : null,
    SizeHalfPoints = formatting.TryGetProperty("fontSizeHalfPoints", out var size) ? size.GetInt32() : null,
    Mode = ChangeMode.Direct,
};

// A new document is OfficeAgent's blank package plus one plan per paragraph: the first fills
// the blank document's empty paragraph, each later one is inserted after the previous.
static async Task<byte[]> Compose(OfficeAgentClient client, IReadOnlyList<(string Text, JsonElement? Formatting)> paragraphs, string? listStyle)
{
    var document = client.CreateBlank("composed.docx");
    string? previous = null;
    foreach (var (text, formatting) in paragraphs)
    {
        var anchor = (await client.InspectAsync(Handle(document))).Paragraphs.Last(p => p.In is null);
        document = previous is null
            ? await Commit(client, document, new ChangeTextOp
            {
                Target = new TextSpanAnchor { ParaId = anchor.ParaId, Expect = "" },
                With = text,
                Mode = ChangeMode.Direct,
            })
            : await Commit(client, document, new InsertOp
            {
                Target = new TextSpanAnchor { ParaId = anchor.ParaId, Expect = previous },
                Position = InsertPosition.After,
                Text = text,
                Mode = ChangeMode.Direct,
            });
        previous = text;

        var written = await FirstOccurrence(client, document, text);
        if (formatting is { } runFormatting)
            document = await Commit(client, document, WithRunFormatting(written, runFormatting));
        if (listStyle is not null)
            document = await Commit(client, document, new FormatOp
            {
                Target = written,
                ListStyle = listStyle,
                ListLevel = 0,
                ListId = 1,
                Mode = ChangeMode.Direct,
            });
    }
    return document;
}

static async Task<byte[]> Commit(OfficeAgentClient client, byte[] document, PlanOperation operation)
{
    var plan = new DocumentPlan
    {
        ContractVersion = DocumentPlan.CurrentContractVersion,
        Format = OfficeAgent.Abstractions.DocumentFormat.Word,
        Operations = new[] { operation },
    };
    using var result = await client.CommitAsync(Handle(document), plan);
    if (result.Committed) return result.ToBytes();

    var errors = result.Report.Errors;
    var message = string.Join("; ", errors.Select(error => $"{error.Code}: {error.Message}"));
    if (errors.Count > 0 && errors.All(error => error.Code == ValidationErrorCodes.UnsupportedOperation))
        throw new UnsupportedException($"OfficeAgent declined the plan: {message}");
    throw new InvalidOperationException($"OfficeAgent did not commit the plan: {message}");
}

sealed class UnsupportedException(string message) : Exception(message);

// docx-platform-tests adapter protocol v1 entrypoint for the Open XML SDK.
//
// Capability map (see README.md):
// - replaceFirstTextOccurrence: intra-w:t replacement via the SDK's typed
//   DOM. A match that spans w:t boundaries is declined (exit 2) rather than
//   approximated with adapter-side algorithms.
// - acceptAllTrackedChanges / rejectAllTrackedChanges and everything else:
//   exit 2 unsupported -- the SDK is a typed DOM over the package and
//   provides no accept/reject-revisions API.
using System.Reflection;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

const string ProtocolVersion = "1";

static string? Arg(string[] argv, string name)
{
    var index = Array.IndexOf(argv, name);
    return index >= 0 && index + 1 < argv.Length ? argv[index + 1] : null;
}

if (args.Contains("--print-library-version"))
{
    var assembly = typeof(WordprocessingDocument).Assembly;
    var informational = assembly
        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
        ?.InformationalVersion;
    var version = informational?.Split('+')[0]
        ?? assembly.GetName().Version?.ToString()
        ?? "unknown";
    Console.WriteLine($"DocumentFormat.OpenXml {version}");
    return 0;
}

var protocol = Arg(args, "--protocol-version");
if (protocol != ProtocolVersion)
{
    Console.WriteLine(
        $"openxml-sdk adapter speaks protocol v{ProtocolVersion}, got {protocol}");
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

using var operation = JsonDocument.Parse(File.ReadAllText(operationPath));
var operationName = operation.RootElement.GetProperty("operationName").GetString();

if (operationName is "acceptAllTrackedChanges" or "rejectAllTrackedChanges")
{
    Console.WriteLine(
        "the Open XML SDK is a typed DOM over the package with no " +
        "accept/reject-revisions API; implementing one would be an " +
        "adapter-side algorithm");
    return 2;
}
if (operationName != "replaceFirstTextOccurrence")
{
    Console.WriteLine(
        $"openxml-sdk adapter does not implement operation '{operationName}'");
    return 2;
}

var findText = operation.RootElement.GetProperty("findText").GetString()!;
var replaceText = operation.RootElement.GetProperty("replaceText").GetString()!;

// Mutate a temp copy and move it into place only on success: exit 2 must
// leave no output file (adapter protocol v1).
var tempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
File.Copy(inputPath, tempPath, overwrite: true);
try
{
    int exitCode = Replace(tempPath, findText, replaceText);
    if (exitCode == 0)
    {
        File.Move(tempPath, outputPath, overwrite: true);
    }
    return exitCode;
}
finally
{
    if (File.Exists(tempPath)) File.Delete(tempPath);
}

static int Replace(string packagePath, string findText, string replaceText)
{
    using var package = WordprocessingDocument.Open(packagePath, isEditable: true);
    var document = package.MainDocumentPart?.Document;
    var body = document?.Body;
    if (document is null || body is null)
    {
        Console.Error.WriteLine("input package has no main document body");
        return 1;
    }

    foreach (var paragraph in body.Descendants<Paragraph>())
    {
        // w:t only, matching the operation's body-text semantics
        // (w:delText is pending-deletion content, not paragraph text).
        var texts = paragraph.Descendants<Text>().ToList();
        var paragraphText = string.Concat(texts.Select(t => t.Text ?? string.Empty));
        if (!paragraphText.Contains(findText, StringComparison.Ordinal)) continue;

        var firstIndex = paragraphText.IndexOf(findText, StringComparison.Ordinal);
        var offset = 0;
        foreach (var text in texts)
        {
            var value = text.Text ?? string.Empty;
            if (firstIndex >= offset && firstIndex + findText.Length <= offset + value.Length)
            {
                var local = firstIndex - offset;
                var rewritten = value
                    .Remove(local, findText.Length)
                    .Insert(local, replaceText);
                text.Text = rewritten;
                if (rewritten != rewritten.Trim())
                {
                    text.Space = SpaceProcessingModeValues.Preserve;
                }
                document.Save();
                return 0;
            }
            offset += value.Length;
        }

        Console.WriteLine(
            "first occurrence spans w:t boundaries; the openxml-sdk adapter " +
            "only performs intra-w:t replacement (glue, not algorithms)");
        return 2;
    }

    Console.Error.WriteLine($"findText not present in any paragraph: '{findText}'");
    return 1;
}

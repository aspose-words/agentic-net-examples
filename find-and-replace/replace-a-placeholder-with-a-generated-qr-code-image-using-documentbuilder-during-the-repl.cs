using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a placeholder.
        const string inputPath = "Input.docx";
        const string outputPath = "Output.docx";

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a QR placeholder: {{QR}}.");
        doc.Save(inputPath);

        // Load the document for replacement.
        var loadedDoc = new Document(inputPath);

        // Verify that the placeholder exists at least once without altering the document.
        int initialMatches = loadedDoc.Range.Replace("{{QR}}", "{{QR}}", new FindReplaceOptions());
        if (initialMatches == 0)
            throw new Exception("Placeholder '{{QR}}' was not found in the document.");

        // Perform replacement using a custom callback that inserts a QR code image.
        var replaceOptions = new FindReplaceOptions
        {
            ReplacingCallback = new QrReplacingCallback()
        };
        loadedDoc.Range.Replace("{{QR}}", "", replaceOptions);

        // Save the modified document.
        loadedDoc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file '{outputPath}'.");
    }
}

// Custom callback that replaces the placeholder with a generated QR code image.
public class QrReplacingCallback : IReplacingCallback
{
    // Aspose.Words versions may require the method name to be Replacing.
    public ReplaceAction Replacing(ReplacingArgs e)
    {
        return Replace(e);
    }

    // Core replacement logic.
    private ReplaceAction Replace(ReplacingArgs e)
    {
        // Ensure the match node is a Run.
        var run = e.MatchNode as Run;
        if (run == null)
            return ReplaceAction.Skip;

        // A tiny 1x1 transparent PNG (as a placeholder for a QR code).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);

        // Insert the image at the location of the placeholder.
        var docBuilder = new DocumentBuilder((Document)run.Document);
        docBuilder.MoveTo(run);
        using (var ms = new MemoryStream(pngBytes))
        {
            docBuilder.InsertImage(ms);
        }

        // Remove the original placeholder run.
        var parent = run.ParentNode;
        parent?.RemoveChild(run);

        // Skip the default replacement since we handled it.
        return ReplaceAction.Skip;
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document that contains the QR code placeholder.
        // -----------------------------------------------------------------
        const string inputPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains a QR code placeholder:");
        builder.Writeln("[QR]");
        builder.Writeln("End of document.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document for processing.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Prepare the callback that will replace the placeholder with an image.
        // -----------------------------------------------------------------
        var callback = new QrCodeReplacingCallback();

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback,
            MatchCase = true
        };

        // -----------------------------------------------------------------
        // 4. Perform the replacement using a literal string placeholder.
        // -----------------------------------------------------------------
        int replacedCount = loadedDoc.Range.Replace("[QR]", string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("No placeholder was replaced.");

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // -----------------------------------------------------------------
        // 6. Write a JSON report of the replacements.
        // -----------------------------------------------------------------
        var report = new
        {
            Placeholder = "[QR]",
            Replacements = callback.MatchedPlaceholders,
            ReplacedCount = replacedCount,
            OutputDocument = outputPath
        };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", json);
    }

    // -----------------------------------------------------------------
    // Callback that inserts a QR code image in place of the matched placeholder.
    // -----------------------------------------------------------------
    private class QrCodeReplacingCallback : IReplacingCallback
    {
        public List<string> MatchedPlaceholders { get; } = new List<string>();

        // A simple 1×1 black PNG used as a placeholder QR code image.
        private static readonly byte[] QrCodePng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Record the matched placeholder text.
            MatchedPlaceholders.Add(args.Match.Value);

            // Insert the QR code image at the location of the match.
            DocumentBuilder imageBuilder = new DocumentBuilder((Document)args.MatchNode.Document);
            imageBuilder.MoveTo(args.MatchNode);
            using (MemoryStream imageStream = new MemoryStream(QrCodePng))
            {
                imageBuilder.InsertImage(imageStream);
            }

            // Let the engine perform the default replacement (replace with empty string).
            return ReplaceAction.Replace;
        }
    }
}

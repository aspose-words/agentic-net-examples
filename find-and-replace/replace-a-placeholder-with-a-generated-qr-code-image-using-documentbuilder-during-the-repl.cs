using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text with a placeholder that will be replaced by a QR code.
        builder.Writeln("Please scan the following QR code:");
        builder.Writeln("_QR_"); // Placeholder to be replaced.

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new QrCodeReplacingCallback()
        };

        // Perform the replacement. The placeholder text will be removed and a QR code field inserted.
        int replaced = doc.Range.Replace("_QR_", string.Empty, options);
        if (replaced == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the resulting document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Callback that inserts a QR code field at the location of each match.
    private class QrCodeReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Create a builder positioned at the start of the matched node.
            // args.MatchNode.Document returns DocumentBase, so cast to Document.
            DocumentBuilder cb = new DocumentBuilder((Document)args.MatchNode.Document);
            cb.MoveTo(args.MatchNode);

            // Insert a QR code field. The QR code encodes a sample URL.
            string qrData = "https://example.com";
            cb.InsertField($"DISPLAYBARCODE QR \"{qrData}\"");

            // Remove the placeholder text by replacing it with an empty string.
            args.Replacement = string.Empty;
            return ReplaceAction.Replace;
        }
    }
}

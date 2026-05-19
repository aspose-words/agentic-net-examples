using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a placeholder for the QR code.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Document start.");
        builder.Writeln("Here is the QR code placeholder: {{QR}}");
        builder.Writeln("Document end.");

        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new QrCodeReplacingCallback()
        };

        // Perform the replacement. The replacement string is empty because the callback inserts an image.
        int replacedCount = loadedDoc.Range.Replace("{{QR}}", string.Empty, options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one QR code placeholder replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Write a simple JSON report of the operation.
        var report = new { Replacements = replacedCount };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", json);
    }

    private class QrCodeReplacingCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Generate a simple placeholder QR code image (black squares on white background).
            using (Bitmap bitmap = new Bitmap(100, 100))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.White);
                    using (SolidBrush blackBrush = new SolidBrush(Color.Black))
                    {
                        // Draw a few black squares to simulate a QR pattern.
                        for (int i = 0; i < 10; i++)
                        {
                            int size = 8;
                            int offset = i * 10;
                            graphics.FillRectangle(blackBrush, offset, offset, size, size);
                        }
                    }
                }

                using (MemoryStream imageStream = new MemoryStream())
                {
                    bitmap.Save(imageStream, ImageFormat.Png);
                    imageStream.Position = 0;

                    // Insert the image at the location of the placeholder.
                    DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
                    builder.MoveTo(args.MatchNode);
                    builder.InsertImage(imageStream);
                }
            }

            // Let Aspose.Words perform the default replacement (empty string) so the match is counted.
            return ReplaceAction.Replace;
        }
    }
}

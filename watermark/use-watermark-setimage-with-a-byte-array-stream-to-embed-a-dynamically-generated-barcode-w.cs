using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one paragraph.
        doc.FirstSection.Body.FirstParagraph.AppendChild(new Run(doc, "Sample content for watermark demonstration."));

        // Create a simple PNG image (1x1 pixel) as a placeholder for the barcode.
        // This avoids the need for an external barcode generator and complies with the rule set.
        // The image data is a base64‑encoded PNG.
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK5cAAAAASUVORK5CYII=");

        // Load the image data into a memory stream.
        using (MemoryStream imageStream = new MemoryStream(pngData))
        {
            // Reset the position just in case.
            imageStream.Position = 0;

            // Configure optional watermark appearance.
            ImageWatermarkOptions options = new ImageWatermarkOptions
            {
                Scale = 5,          // Scale factor for the watermark.
                IsWashout = false   // Disable washout effect.
            };

            // Apply the image as a watermark using the stream overload.
            doc.Watermark.SetImage(imageStream, options);
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "BarcodeWatermark.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

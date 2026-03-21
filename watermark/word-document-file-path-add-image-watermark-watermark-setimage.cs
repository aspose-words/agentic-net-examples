using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create temporary file paths.
        string inputDocPath = Path.Combine(Path.GetTempPath(), "SourceDocument.docx");
        string watermarkImagePath = Path.Combine(Path.GetTempPath(), "Watermark.png");
        string outputDocPath = Path.Combine(Path.GetTempPath(), "WatermarkedDocument.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple Word document to work with.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        doc.Save(inputDocPath);

        // -----------------------------------------------------------------
        // 2. Create a simple PNG image that will be used as a watermark.
        // -----------------------------------------------------------------
        // A minimal 1x1 transparent PNG (base64 encoded).
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(watermarkImagePath, pngBytes);

        // -----------------------------------------------------------------
        // 3. Load the document and apply the image watermark.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            Scale = 5,
            IsWashout = false
        };
        loadedDoc.Watermark.SetImage(watermarkImagePath, options);

        // -----------------------------------------------------------------
        // 4. Save the watermarked document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputDocPath);

        Console.WriteLine($"Watermarked document saved to: {outputDocPath}");
    }
}

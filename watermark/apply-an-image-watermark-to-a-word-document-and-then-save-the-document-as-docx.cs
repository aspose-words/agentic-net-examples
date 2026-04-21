using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample image, the source document and the output document.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string imagePath = Path.Combine(artifactsDir, "watermark.png");
        string outputPath = Path.Combine(artifactsDir, "WatermarkedDocument.docx");

        // Create a simple 1x1 pixel PNG image from a base64 string.
        // This avoids using System.Drawing or external image files.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a blank Word document and add a paragraph of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document contains an image watermark.");

        // Apply the image watermark using the Document.Watermark API.
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            // Example: make the watermark opaque (no washout) and scale it.
            IsWashout = false,
            Scale = 1.0
        };
        doc.Watermark.SetImage(imagePath, options);

        // Save the watermarked document as DOCX.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}

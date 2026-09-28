using System;
using System.IO;
using Aspose.Words;
using SkiaSharp;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with barcode watermark.");

        // Base64‑encoded PNG image (1×1 pixel) to simulate a barcode image.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6WcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Decode the image bytes into an SKBitmap (required by Aspose.Words Watermark API).
        using SKBitmap bitmap = SKBitmap.Decode(imageBytes);
        if (bitmap == null)
        {
            Console.WriteLine("Failed to decode the barcode image.");
            return;
        }

        // Apply the image as a watermark.
        doc.Watermark.SetImage(bitmap);

        // Save the document.
        string outputPath = "WatermarkedDoc.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}

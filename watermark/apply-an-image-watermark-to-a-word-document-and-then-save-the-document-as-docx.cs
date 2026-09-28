using System;
using System.IO;
using Aspose.Words;
using SkiaSharp;   // Required for SKBitmap used by Watermark.SetImage

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document content.");

        // Create a sample PNG image to use as a watermark.
        string imagePath = "watermark.png";
        if (!File.Exists(imagePath))
        {
            // A 1x1 pixel transparent PNG encoded in Base64.
            string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ZcAAAAASUVORK5CYII=";
            byte[] imageBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, imageBytes);
        }

        // Load the image into an SKBitmap (required by the current Watermark API).
        using SKBitmap bitmap = SKBitmap.Decode(imagePath);
        if (bitmap == null)
        {
            Console.WriteLine("Failed to load watermark image.");
            return;
        }

        // Apply the image watermark to the document.
        doc.Watermark.SetImage(bitmap);

        // Save the watermarked document as DOCX.
        string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Watermarked document saved successfully.");
        }
        else
        {
            Console.WriteLine("Failed to save watermarked document.");
        }
    }
}

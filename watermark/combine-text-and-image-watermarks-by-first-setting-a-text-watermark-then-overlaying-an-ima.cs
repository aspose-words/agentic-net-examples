using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for the output document and temporary image file.
        const string outputDocPath = "CombinedWatermark.docx";
        const string imagePath = "watermark.png";

        // Create a minimal PNG image (1x1 pixel, transparent) from a Base64 string.
        // This avoids using System.Drawing APIs.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank document.
        Document doc = new Document();

        // Add some sample content so the watermarks are visible.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document demonstrates combined text and image watermarks.");
        builder.Writeln("The text watermark appears behind the page content,");
        builder.Writeln("and the image watermark is overlaid on top.");

        // Apply a text watermark.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Apply an image watermark using the previously created PNG file.
        ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Increase the size of the image watermark.
            IsWashout = false   // Keep the image colors (no washout effect).
        };
        doc.Watermark.SetImage(imagePath, imgOptions);

        // Save the document with both watermarks applied.
        doc.Save(outputDocPath);

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
        {
            File.Delete(imagePath);
        }

        // Optional verification (no console output required by the task).
        // The file existence check ensures the document was saved.
        bool saved = File.Exists(outputDocPath);
    }
}

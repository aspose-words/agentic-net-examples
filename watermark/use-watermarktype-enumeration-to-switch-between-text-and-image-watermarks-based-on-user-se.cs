using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Create a sample PNG image (1x1 pixel) for the image watermark.
        string imagePath = Path.Combine(artifactsDir, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Choose the watermark type to apply.
        WatermarkType selectedWatermark = WatermarkType.Image; // Change to WatermarkType.Text for text watermark.

        // Ensure any existing watermark is removed before applying a new one.
        if (doc.Watermark.Type != WatermarkType.None)
            doc.Watermark.Remove();

        // Apply the chosen watermark.
        if (selectedWatermark == WatermarkType.Text)
        {
            doc.Watermark.SetText("Confidential");
        }
        else if (selectedWatermark == WatermarkType.Image)
        {
            // Use the overload that accepts a file path and optional options.
            doc.Watermark.SetImage(imagePath, null);
        }

        // Save the resulting document.
        string outputPath = Path.Combine(artifactsDir, "output.docx");
        doc.Save(outputPath);

        // Simple verification output.
        Console.WriteLine($"Watermark applied: {doc.Watermark.Type}");
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

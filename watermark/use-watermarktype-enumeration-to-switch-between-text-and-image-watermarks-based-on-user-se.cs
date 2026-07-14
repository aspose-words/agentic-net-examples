using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Simulate user selection: change this value to WatermarkType.Image to test image watermark.
        WatermarkType selectedWatermark = WatermarkType.Text;

        // Path for the output document.
        string outputPath = "WatermarkedDocument.docx";

        if (selectedWatermark == WatermarkType.Text)
        {
            // Apply a text watermark.
            doc.Watermark.SetText("Sample Text Watermark");
        }
        else if (selectedWatermark == WatermarkType.Image)
        {
            // Prepare a deterministic 1x1 PNG image.
            string imagePath = "sample.png";
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
            File.WriteAllBytes(imagePath, pngBytes);

            // Apply an image watermark using the created image file.
            // Use the overload that accepts a file path and ImageWatermarkOptions.
            ImageWatermarkOptions options = new ImageWatermarkOptions();
            doc.Watermark.SetImage(imagePath, options);
        }

        // Save the document.
        doc.Save(outputPath);

        // Simple validation output.
        Console.WriteLine($"Watermark applied: {doc.Watermark.Type}");
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class WatermarkDemo
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a blank document.
        Document doc = new Document();

        // Create a sample 1x1 pixel PNG image for the image watermark.
        string imagePath = Path.Combine(outputDir, "sample.png");
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, imageBytes);

        // Simulate user selection. Change this value to WatermarkType.Image to test image watermark.
        WatermarkType selectedWatermark = WatermarkType.Text;

        // Apply the appropriate watermark based on the selected type.
        if (selectedWatermark == WatermarkType.Text)
        {
            // Simple text watermark.
            doc.Watermark.SetText("Sample Text Watermark");
        }
        else if (selectedWatermark == WatermarkType.Image)
        {
            // Use the overload that accepts a file path and optional options.
            // Passing null for options uses default settings.
            doc.Watermark.SetImage(imagePath, null);
        }

        // Save the document.
        string outputPath = Path.Combine(outputDir, "WatermarkedDocument.docx");
        doc.Save(outputPath);

        // Simple validation: ensure the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully with {(selectedWatermark == WatermarkType.Text ? "text" : "image")} watermark.");
        }
    }
}

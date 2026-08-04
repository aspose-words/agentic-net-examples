using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare a blank document.
        Document doc = new Document();

        // Create a deterministic sample image file (1x1 pixel PNG) if it does not exist.
        string imagePath = "sample.png";
        if (!File.Exists(imagePath))
        {
            // Base64-encoded PNG data for a single transparent pixel.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6V8AAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // Simulate user selection without interactive input.
        // Change this value to WatermarkType.Text to apply a text watermark instead.
        WatermarkType selectedType = WatermarkType.Image;

        // Apply the appropriate watermark based on the selected type.
        if (selectedType == WatermarkType.Text)
        {
            doc.Watermark.SetText("Sample Text Watermark");
        }
        else if (selectedType == WatermarkType.Image)
        {
            var imageOptions = new ImageWatermarkOptions
            {
                Scale = 1,          // No scaling.
                IsWashout = false   // Keep original colors.
            };
            doc.Watermark.SetImage(imagePath, imageOptions);
        }

        // Save the resulting document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}

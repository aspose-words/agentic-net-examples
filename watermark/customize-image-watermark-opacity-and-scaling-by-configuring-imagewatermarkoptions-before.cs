using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple 1x1 PNG image file to use as a watermark.
        string imagePath = "sample.png";
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XG6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank document.
        Document doc = new Document();

        // Configure watermark options: set scale and disable washout (full opacity).
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // Scale to 50% of the original size.
            IsWashout = false    // Do not apply washout effect (full opacity).
        };

        // Insert the image watermark using the configured options.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the resulting document.
        string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);
    }
}

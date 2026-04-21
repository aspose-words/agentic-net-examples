using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample image and the output document.
        string baseDir = Directory.GetCurrentDirectory();
        string imagePath = Path.Combine(baseDir, "watermark.png");
        string outputPath = Path.Combine(baseDir, "Watermarked.docx");

        // Create a simple 1x1 pixel PNG image (transparent) from a Base64 string.
        // This avoids using System.Drawing or external image files.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X0ZcAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank Word document.
        Document doc = new Document();

        // Configure watermark options: set scale and disable washout (full opacity).
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            Scale = 0.5,      // Scale factor (50% of the original image size).
            IsWashout = false // Keep the original opacity of the image.
        };

        // Apply the image watermark using the configured options.
        doc.Watermark.SetImage(imagePath, options);

        // Save the resulting document.
        doc.Save(outputPath);
    }
}

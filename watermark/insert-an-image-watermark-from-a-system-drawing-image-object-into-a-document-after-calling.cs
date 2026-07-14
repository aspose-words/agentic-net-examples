using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Settings;   // Needed for MsWordVersion enum

public class Program
{
    public static void Main()
    {
        // Create a deterministic 1x1 PNG image file.
        string imagePath = Path.Combine(Path.GetTempPath(), "watermark.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a blank Word document.
        Document doc = new Document();

        // Simulate an optimization step for Word 2010.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Configure image watermark options (optional).
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor for the watermark.
            IsWashout = false   // Disable washout effect.
        };

        // Insert the image watermark from the local file.
        doc.Watermark.SetImage(imagePath, options);

        // Save the resulting document.
        string outputPath = Path.Combine(Path.GetTempPath(), "Watermarked.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        Console.WriteLine($"Document saved to: {outputPath}");
        Console.WriteLine($"File exists: {File.Exists(outputPath)}");
    }
}

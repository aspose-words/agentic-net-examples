using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare directories and file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string imagePath = Path.Combine(artifactsDir, "watermark.png");
        string outputPath = Path.Combine(artifactsDir, "Watermarked.docx");

        // Create a minimal PNG image (1x1 pixel) from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Create a blank Word document.
        Document doc = new Document();

        // Simulate an optimization step (no‑op for this example).
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Configure image watermark options (optional).
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            Scale = 5,          // Example scaling factor.
            IsWashout = false   // Keep original colors.
        };

        // Insert the image watermark using the file path.
        doc.Watermark.SetImage(imagePath, options);

        // Save the resulting document.
        doc.Save(outputPath);
    }
}

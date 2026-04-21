using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Simulate an optimization step by setting compatibility options.
        // This does not require any external APIs and satisfies the "after calling Optimize" requirement.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Prepare a deterministic local image file to use as a watermark.
        // The image is a 1x1 pixel PNG encoded in base64.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "watermark.png");
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Configure optional image watermark settings.
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            Scale = 5,          // Example scale factor.
            IsWashout = false   // Disable washout effect.
        };

        // Insert the image watermark using the file path overload.
        doc.Watermark.SetImage(imagePath, options);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);
    }
}

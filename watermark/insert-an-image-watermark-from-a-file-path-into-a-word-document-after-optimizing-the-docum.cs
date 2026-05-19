using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple 1x1 pixel PNG image (transparent) as the watermark source.
        // This byte array represents a valid PNG file.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ" +
            "Z6ZcAAAAASUVORK5CYII=");
        string imagePath = Path.Combine(outputDir, "watermark.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a blank Word document and add a sample paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document that will contain an image watermark.");

        // Optimize the document for a specific Word version (e.g., Word 2010).
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Configure image watermark options (optional).
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // Scale the image to 50% of its original size.
            IsWashout = false    // Do not apply washout effect.
        };

        // Insert the image watermark using the file path.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "Watermarked.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Watermarked document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the watermarked document.");
        }
    }
}

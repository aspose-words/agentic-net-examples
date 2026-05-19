using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OptimizedWithWatermark.docx");

        // Create a simple 1x1 PNG image (transparent) and save it locally
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK9cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a large document by adding many paragraphs
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 1000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}");
        }

        // Optimize the document for a specific Word version to reduce overhead
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Configure image watermark options
        ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // Scale the watermark to 50% of the page width/height
            IsWashout = false    // Keep original colors (no washout effect)
        };

        // Apply the image watermark using the local image file
        doc.Watermark.SetImage(imagePath, imgOptions);

        // Save the optimized document with the watermark
        doc.Save(outputPath);

        // Optional: simple verification that the file was created
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {outputPath}");
        }
    }
}

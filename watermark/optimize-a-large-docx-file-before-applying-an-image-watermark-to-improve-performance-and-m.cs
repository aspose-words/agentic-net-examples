using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for intermediate and final files
        string largeDocPath = Path.Combine(outputDir, "LargeDocument.docx");
        string optimizedDocPath = Path.Combine(outputDir, "OptimizedDocument.docx");
        string finalDocPath = Path.Combine(outputDir, "WatermarkedDocument.docx");
        string imagePath = Path.Combine(outputDir, "watermark.png");

        // 1. Create a large sample document
        Document largeDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(largeDoc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: The quick brown fox jumps over the lazy dog.");
        }
        largeDoc.Save(largeDocPath);

        // 2. Optimize the document by saving with memory optimization enabled
        SaveOptions optOptions = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
        optOptions.MemoryOptimization = true;
        largeDoc.Save(optimizedDocPath, optOptions);

        // 3. Load the optimized document
        Document doc = new Document(optimizedDocPath);

        // 4. Create a simple PNG image for the watermark (1x1 pixel)
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // 5. Apply the image watermark
        ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // Scale to 50% of the page width/height
            IsWashout = false    // Keep original colors
        };
        doc.Watermark.SetImage(imagePath, imgOptions);

        // 6. Save the final document with the watermark
        doc.Save(finalDocPath);

        // Simple validation output
        Console.WriteLine($"Final document saved: {File.Exists(finalDocPath)}");
    }
}

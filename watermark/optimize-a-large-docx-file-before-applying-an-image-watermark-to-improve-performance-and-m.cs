using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample PNG image for the watermark (a 1x1 pixel transparent PNG).
        string imagePath = Path.Combine(outputDir, "watermark.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAoMBgZcKcVQAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a large document (e.g., 1000 paragraphs spread over many pages).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 1000; i++)
        {
            builder.Writeln($"Paragraph {i}");
            // Insert a page break every 50 paragraphs to increase page count.
            if (i % 50 == 0)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Optimize compatibility options for a specific Word version to reduce overhead.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

        // Configure image watermark options.
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // Scale the image to 50% of its original size.
            IsWashout = false    // Keep the image colors unchanged.
        };

        // Apply the image watermark using the local image file.
        doc.Watermark.SetImage(imagePath, watermarkOptions);

        // Save the document with memory optimization enabled to lower memory consumption.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            MemoryOptimization = true
        };
        string outputPath = Path.Combine(outputDir, "OptimizedWatermarked.docx");
        doc.Save(outputPath, saveOptions);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        else
            Console.WriteLine("Failed to save the document.");
    }
}

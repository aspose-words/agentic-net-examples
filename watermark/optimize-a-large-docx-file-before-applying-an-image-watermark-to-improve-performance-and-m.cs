using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a large sample DOCX document.
        // -----------------------------------------------------------------
        string largeDocPath = Path.Combine(outputDir, "LargeDocument.docx");
        Document largeDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(largeDoc);

        // Add many paragraphs to simulate a large file.
        for (int i = 1; i <= 2000; i++)
        {
            builder.Writeln($"This is line {i} of a large document used for performance testing.");
        }

        largeDoc.Save(largeDocPath);

        // -----------------------------------------------------------------
        // 2. Optimize the document by saving it with memory optimization.
        // -----------------------------------------------------------------
        string optimizedDocPath = Path.Combine(outputDir, "OptimizedDocument.docx");
        Document docToOptimize = new Document(largeDocPath);

        // Create save options with memory optimization enabled.
        SaveOptions opt = SaveOptions.CreateSaveOptions(SaveFormat.Docx);
        opt.MemoryOptimization = true;

        docToOptimize.Save(optimizedDocPath, opt);

        // -----------------------------------------------------------------
        // 3. Create a simple PNG image to be used as a watermark.
        // -----------------------------------------------------------------
        string imagePath = Path.Combine(outputDir, "watermark.png");
        // 1x1 pixel transparent PNG (base64 encoded).
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lK6XAAAAAElFTkSuQmCC";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // 4. Load the optimized document and apply the image watermark.
        // -----------------------------------------------------------------
        Document finalDoc = new Document(optimizedDocPath);

        ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
        {
            Scale = 0.5,          // Scale the watermark to 50% of the page width.
            IsWashout = false    // Keep the original colors.
        };

        // Apply the watermark using the image file path.
        finalDoc.Watermark.SetImage(imagePath, imgOptions);

        // -----------------------------------------------------------------
        // 5. Save the final document with the watermark.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(outputDir, "WatermarkedDocument.docx");
        finalDoc.Save(outputPath);

        // Simple validation that the output file exists.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Watermark applied and document saved successfully:");
            Console.WriteLine(outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the watermarked document.");
        }
    }
}

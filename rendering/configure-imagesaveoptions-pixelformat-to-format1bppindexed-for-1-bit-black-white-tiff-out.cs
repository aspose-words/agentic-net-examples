using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered as a 1‑bit black‑white TIFF image.");

        // Configure image save options for 1‑bit TIFF
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format1bppIndexed
        };

        // Define output file path
        string outputPath = Path.Combine(outputDir, "sample_1bpp.tiff");

        // Save the document as TIFF
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        Console.WriteLine($"TIFF image saved successfully to: {outputPath}");
    }
}

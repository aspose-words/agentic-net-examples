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
        string outputPath = Path.Combine(outputDir, "1bpp.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World! This document will be saved as a 1‑bit black‑white TIFF.");

        // Configure image save options for 1‑bit TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // CCITT compression is appropriate for 1‑bit images.
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as TIFF using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Indicate success (no interactive prompts).
        Console.WriteLine("TIFF file created at: " + outputPath);
    }
}

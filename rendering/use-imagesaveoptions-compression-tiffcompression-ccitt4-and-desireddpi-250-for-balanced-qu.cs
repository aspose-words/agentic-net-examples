using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as a black‑and‑white TIFF image.");
        builder.Writeln("Using CCITT4 compression and a DPI of 250 for balanced quality.");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Apply CCITT4 compression (good for B/W images).
            TiffCompression = TiffCompression.Ccitt4,
            // Set both horizontal and vertical resolution (DPI) to 250.
            Resolution = 250,
            // Render as black‑and‑white.
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Save the document as a TIFF file.
        string outputPath = Path.Combine(outputDir, "BalancedBw.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Report success.
        Console.WriteLine($"TIFF saved to: {outputPath}");
    }
}

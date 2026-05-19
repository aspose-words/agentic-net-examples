using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with a heading and a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Balanced Quality Black‑and‑White TIFF");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This document is rendered to a multi‑page TIFF using CCITT4 compression and a DPI of 250.");

        // Prepare the output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "BalancedQuality.tiff");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression for black‑and‑white images.
            TiffCompression = TiffCompression.Ccitt4,
            // Set the desired resolution (dots per inch) for both axes.
            Resolution = 250f,
            // Render the pages as black‑and‑white.
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Save the document as a TIFF file.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Output the file path and size for quick verification.
        Console.WriteLine($"TIFF saved to: {outputPath}");
        Console.WriteLine($"File size: {new FileInfo(outputPath).Length} bytes");
    }
}

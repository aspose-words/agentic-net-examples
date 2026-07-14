using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Build a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Document");
        builder.Writeln("This document will be saved as a binary TIFF image with a high dithering threshold to make it darker.");

        // Configure TIFF save options to use Floyd‑Steinberg dithering with a high threshold.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 150
        };

        // Save the document as a TIFF file.
        string outputPath = Path.Combine(artifactsDir, "DarkenedBinary.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Output basic information about the generated file.
        Console.WriteLine($"TIFF saved to: {outputPath}");
        Console.WriteLine($"File size: {new FileInfo(outputPath).Length} bytes");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Document");
        builder.Writeln("This document will be saved as a TIFF image with Floyd‑Steinberg dithering.");

        // Prepare the output folder.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Sample.tiff");

        // Configure ImageSaveOptions for TIFF with Floyd‑Steinberg dithering.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression (common for B&W TIFFs).
            TiffCompression = TiffCompression.Ccitt3,
            // Apply Floyd‑Steinberg dithering.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            // Set the threshold to 180 for moderately dark images.
            ThresholdForFloydSteinbergDithering = 180
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, report success.
        Console.WriteLine($"TIFF image saved successfully to: {outputPath}");
    }
}

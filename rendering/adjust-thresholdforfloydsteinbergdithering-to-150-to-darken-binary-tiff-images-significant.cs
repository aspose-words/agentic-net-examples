using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content so the rendered TIFF has visible data.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Document");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This document is rendered to a binary (1‑bpp) TIFF image.");
        builder.Writeln("The ThresholdForFloydSteinbergDithering is set to 150 to produce a darker output.");

        // Configure image save options for TIFF with Floyd‑Steinberg dithering.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression (common for binary TIFFs).
            TiffCompression = TiffCompression.Ccitt3,
            // Apply Floyd‑Steinberg error diffusion.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            // Increase the threshold to 150 to darken the binary image.
            ThresholdForFloydSteinbergDithering = 150
        };

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.tiff");

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF output file.");

        Console.WriteLine($"TIFF image saved successfully to: {outputPath}");
    }
}

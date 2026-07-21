using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "DarkerTiffOutput.tiff");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Document");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This document is rendered to a TIFF image with a higher dithering threshold to produce a darker output.");

        // Configure ImageSaveOptions for TIFF with Floyd‑Steinberg dithering.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT3 compression which works with 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt3,
            // Apply Floyd‑Steinberg dithering.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            // Increase the threshold to 200 (default is 128) for a darker result.
            ThresholdForFloydSteinbergDithering = (byte)200
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");

        // Optionally, output the file size for quick confirmation.
        Console.WriteLine($"TIFF file saved successfully. Size: {new FileInfo(outputPath).Length} bytes");
    }
}

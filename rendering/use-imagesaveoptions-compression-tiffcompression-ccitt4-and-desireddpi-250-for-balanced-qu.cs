using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, this is a sample document for TIFF rendering.");

        // Configure TIFF save options for black‑and‑white output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 4 compression for monochrome TIFF.
            TiffCompression = TiffCompression.Ccitt4,
            // Set DPI for balanced quality.
            Resolution = 250
        };

        // Define output file path.
        string outputPath = "SampleOutput.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");
        }

        // Indicate success.
        Console.WriteLine($"TIFF file successfully created at '{Path.GetFullPath(outputPath)}'.");
    }
}

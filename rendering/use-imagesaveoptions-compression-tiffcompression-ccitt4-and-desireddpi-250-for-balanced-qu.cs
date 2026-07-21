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

        // Add sample content.
        builder.Writeln("Aspose.Words rendering example.");
        builder.Writeln("This document will be saved as a black‑and‑white TIFF.");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression for balanced quality and file size.
            TiffCompression = TiffCompression.Ccitt4,
            // Set both horizontal and vertical resolution to 250 DPI.
            Resolution = 250,
            // Render the output as black‑and‑white.
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RenderedDocument.tiff");

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Report success.
        Console.WriteLine($"TIFF file successfully saved to: {outputPath}");
    }
}

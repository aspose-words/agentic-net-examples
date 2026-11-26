using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with a paragraph of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering with CCITT3 compression and 1bpp indexed pixel format.");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression for smallest file size.
            TiffCompression = TiffCompression.Ccitt3,
            // Use black‑and‑white (1‑bit) color mode to achieve 1bpp indexed output.
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image with the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");

        // Output the file size to demonstrate compression effect.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file saved successfully. Size: {fileSize} bytes.");
    }
}

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
        builder.Writeln("Sample text for TIFF rendering with compression and 1bpp indexed pixel format.");

        // Define the output TIFF file path.
        string outputPath = "output.tiff";

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression for smallest size.
            TiffCompression = TiffCompression.Ccitt3,
            // Set color mode to BlackAndWhite which results in a 1‑bit per pixel image.
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the TIFF file.");

        // Output the file size.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file created successfully. Size: {fileSize} bytes.");
    }
}

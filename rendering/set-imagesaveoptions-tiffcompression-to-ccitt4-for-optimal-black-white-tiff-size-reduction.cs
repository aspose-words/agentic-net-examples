using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some black text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for TIFF rendering with CCITT4 compression.");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression for optimal black‑white size reduction.
            TiffCompression = TiffCompression.Ccitt4,
            // Optional: set resolution if desired.
            Resolution = 300
        };

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, report success.
        Console.WriteLine($"TIFF file saved successfully at '{Path.GetFullPath(outputPath)}'.");
    }
}

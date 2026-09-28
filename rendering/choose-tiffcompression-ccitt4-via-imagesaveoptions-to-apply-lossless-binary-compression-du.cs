using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a simple paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering with CCITT4 compression.");

        // Configure image save options for TIFF with CCITT4 (lossless binary) compression.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.TiffCompression = TiffCompression.Ccitt4;

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, indicate success (no interactive input required).
        Console.WriteLine("TIFF file saved successfully with CCITT4 compression.");
    }
}

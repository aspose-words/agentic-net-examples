using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering.");

        // Configure TIFF save options to use CCITT4 compression.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.TiffCompression = TiffCompression.Ccitt4;

        string outputFile = "output.tiff";

        // Save the document as a TIFF image with the specified compression.
        doc.Save(outputFile, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Report the result.
        long fileSize = new FileInfo(outputFile).Length;
        Console.WriteLine($"TIFF saved to '{outputFile}' ({fileSize} bytes).");
    }
}

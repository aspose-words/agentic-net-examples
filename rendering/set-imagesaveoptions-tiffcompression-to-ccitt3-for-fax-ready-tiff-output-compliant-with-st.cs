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
        builder.Writeln("Sample text for TIFF rendering.");

        // Set up TIFF save options to use CCITT3 compression (fax‑ready).
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.TiffCompression = TiffCompression.Ccitt3;

        string outputFile = "fax_ready_output.tiff";

        // Render the document to a TIFF file using the specified options.
        doc.Save(outputFile, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputFile))
        {
            throw new InvalidOperationException("Failed to create the TIFF file.");
        }

        // Optionally, report the file size.
        Console.WriteLine($"TIFF file saved: {outputFile} ({new FileInfo(outputFile).Length} bytes)");
    }
}

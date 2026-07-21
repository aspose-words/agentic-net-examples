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
        builder.Writeln("Hello, fax‑ready TIFF!");

        // Configure image save options for TIFF with CCITT3 compression.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Ccitt3;

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FaxReady.tiff");

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, indicate success (no interactive output required).
        Console.WriteLine("TIFF file created successfully.");
    }
}

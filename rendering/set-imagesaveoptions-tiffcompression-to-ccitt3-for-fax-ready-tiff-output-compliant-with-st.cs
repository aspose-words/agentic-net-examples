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
        builder.Writeln("This is a sample document for fax‑ready TIFF output.");

        // Configure TIFF save options with CCITT3 compression.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.TiffCompression = TiffCompression.Ccitt3;

        // Define output file path.
        string outputPath = "fax_ready_output.tiff";

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, tiffOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine($"TIFF file saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}

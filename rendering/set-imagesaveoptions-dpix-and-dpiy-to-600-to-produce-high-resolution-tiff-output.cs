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
        builder.Writeln("Sample text for high‑resolution TIFF rendering.");

        // Define the output file path.
        string outputPath = "high_res_output.tiff";

        // Configure ImageSaveOptions for TIFF with 600 DPI.
        // In Aspose.Words the DPI is set via the Resolution property (applies to both axes).
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 600
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");
        }

        // Indicate success.
        Console.WriteLine($"TIFF file saved successfully at '{outputPath}'.");
    }
}

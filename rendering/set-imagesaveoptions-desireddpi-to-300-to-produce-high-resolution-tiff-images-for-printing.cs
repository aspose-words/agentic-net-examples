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
        builder.Writeln("Hello, high‑resolution TIFF!");

        // Configure image save options for TIFF with 300 DPI.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // The DPI setting is specified via the Resolution property.
            Resolution = 300
        };

        // Define output file path.
        string outputPath = "high_res_output.tiff";

        // Save the document as a TIFF image using the specified DPI.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Report success (no interactive input required).
        Console.WriteLine($"TIFF image saved successfully with 300 DPI at: {Path.GetFullPath(outputPath)}");
    }
}

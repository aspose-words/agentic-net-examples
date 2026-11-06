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
        builder.Writeln("Hello, low‑resolution TIFF!");

        // Configure TIFF save options with low DPI (72).
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // The Resolution property sets the DPI for the saved image.
            Resolution = 72
        };

        // Define output file path.
        string outputPath = "low_res_output.tiff";

        // Save the document as TIFF using the configured options.
        doc.Save(outputPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");
        }

        // Report success (no interactive input required).
        Console.WriteLine($"TIFF saved successfully to '{outputPath}' with Resolution = 72 DPI.");
    }
}

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
        builder.Writeln("This is a sample black‑and‑white TIFF rendered with CCITT4 compression at 250 DPI.");

        // Configure TIFF save options.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,
            Resolution = 250
        };

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Output a confirmation (non‑interactive).
        Console.WriteLine($"TIFF file successfully saved to: {outputPath}");
    }
}

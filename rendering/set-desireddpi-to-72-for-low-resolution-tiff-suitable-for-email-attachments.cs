using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for low‑resolution TIFF.");

        // Configure TIFF save options to use a low DPI (72).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // The Resolution property sets the DPI for the rendered image.
        saveOptions.Resolution = 72;

        // Define the output file path.
        string outputPath = "LowResolution.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create TIFF file at {outputPath}");
        }

        // Indicate success.
        Console.WriteLine("TIFF saved successfully.");
    }
}

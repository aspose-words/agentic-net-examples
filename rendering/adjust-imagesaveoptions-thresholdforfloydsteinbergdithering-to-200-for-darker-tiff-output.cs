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
        builder.Writeln("Sample text for TIFF rendering with custom dithering threshold.");

        // Configure ImageSaveOptions for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Set the threshold for Floyd‑Steinberg dithering to a higher value (darker output).
        saveOptions.ThresholdForFloydSteinbergDithering = 200;

        // Define output file path.
        string outputPath = "output.tiff";

        // Render the document to a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, report success (no user interaction required).
        Console.WriteLine($"TIFF file saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}

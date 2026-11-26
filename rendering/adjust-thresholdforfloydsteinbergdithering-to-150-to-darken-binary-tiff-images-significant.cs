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
        builder.Writeln("Sample text to demonstrate TIFF dithering threshold adjustment.");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Increase the threshold for Floyd‑Steinberg dithering to darken the binary image.
            ThresholdForFloydSteinbergDithering = 150
        };

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, indicate success (no interactive input required).
        Console.WriteLine($"TIFF image saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}

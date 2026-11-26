using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering with dithering.");

        // Configure ImageSaveOptions for TIFF with a higher dithering threshold.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.ThresholdForFloydSteinbergDithering = 200; // Darker output.

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine("TIFF file created successfully with custom dithering threshold.");
    }
}

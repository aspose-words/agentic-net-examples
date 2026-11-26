using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF conversion with dithering.");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Adjust the dithering threshold to 180 for moderately dark images.
        saveOptions.ThresholdForFloydSteinbergDithering = 180;

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, output a confirmation (no user interaction required).
        Console.WriteLine($"TIFF file saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}

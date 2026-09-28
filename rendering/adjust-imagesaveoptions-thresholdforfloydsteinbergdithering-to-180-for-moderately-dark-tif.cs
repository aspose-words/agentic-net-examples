using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering with Floyd‑Steinberg dithering.");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Adjust the dithering threshold to 180 for moderately dark images.
            ThresholdForFloydSteinbergDithering = 180
        };

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
        }

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine($"TIFF image saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}

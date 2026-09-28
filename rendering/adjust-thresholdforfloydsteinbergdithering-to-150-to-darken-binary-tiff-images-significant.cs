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
        builder.Writeln("Sample text to render as a binary TIFF image.");
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Configure image save options for TIFF with black‑and‑white color mode.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the pages as black‑and‑white (binary) images.
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Set the threshold for Floyd‑Steinberg dithering to 150 to darken the output.
            ThresholdForFloydSteinbergDithering = 150,
            // Save only the first page (single‑page example).
            PageSet = new PageSet(0)
        };

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);

        // Report the file size to indicate that rendering succeeded.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF image saved successfully: {outputPath} ({fileSize} bytes)");
    }
}

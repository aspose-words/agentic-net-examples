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
        builder.Writeln("Sample text for OCR preprocessing. This document will be saved as a binary TIFF.");

        // Configure image save options for binary TIFF with a high dithering threshold.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the image in black-and-white (binary) mode.
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Set the threshold for Floyd‑Steinberg dithering to 90 (lighter binary output).
            ThresholdForFloydSteinbergDithering = 90
        };

        // Define the output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, indicate success (no interactive input required).
        Console.WriteLine($"TIFF file successfully saved to '{Path.GetFullPath(outputPath)}'.");
    }
}

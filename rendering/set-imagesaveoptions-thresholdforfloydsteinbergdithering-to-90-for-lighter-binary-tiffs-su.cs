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

        // Define the output TIFF file path.
        string outputPath = "output.tiff";

        // Configure image save options for TIFF with a high dithering threshold.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set the threshold for Floyd‑Steinberg dithering to 90 (lighter binary image).
            ThresholdForFloydSteinbergDithering = 90,
            // Optional: ensure the image is saved as a black‑and‑white bitmap.
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Optionally, output a confirmation (no user interaction required).
        Console.WriteLine($"TIFF file successfully created at '{Path.GetFullPath(outputPath)}'.");
    }
}

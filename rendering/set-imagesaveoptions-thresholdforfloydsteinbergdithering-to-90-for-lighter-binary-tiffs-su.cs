using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a folder for the output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for OCR preprocessing.");

        // Configure TIFF save options:
        // - Use CCITT4 compression (common for binary TIFFs).
        // - Apply Floyd‑Steinberg dithering.
        // - Set a high threshold (90) to produce lighter binary images.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 90
        };

        // Save the document as a TIFF image.
        string tiffPath = Path.Combine(outputDir, "LighterBinary.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Indicate success.
        Console.WriteLine($"TIFF saved to: {tiffPath}");
    }
}

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for grayscale TIFF conversion.");
        // Insert a placeholder image (a small generated bitmap is not required; the document can be saved without it).

        // Configure TIFF save options:
        // - Grayscale color mode.
        // - CCITT3 compression (binary image compression).
        // - Use Floyd‑Steinberg dithering with a threshold of 100 to obtain a lighter grayscale result.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.Grayscale,
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 100
        };

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(outputDir, "GrayscaleThreshold.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Optionally, inform that the process completed.
        Console.WriteLine($"TIFF saved to: {tiffPath}");
    }
}

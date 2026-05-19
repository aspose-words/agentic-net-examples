using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for grayscale TIFF conversion.");

        // Configure TIFF save options:
        // - Grayscale color mode.
        // - CCITT4 compression (binary).
        // - Floyd‑Steinberg dithering with a custom threshold of 100.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,
            ImageColorMode = ImageColorMode.Grayscale,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 100
        };

        // Save the document as a TIFF image.
        string outPath = Path.Combine(artifactsDir, "GrayscaleThreshold.tiff");
        doc.Save(outPath, options);

        // Verify that the file was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Inform the user (no input required).
        Console.WriteLine($"TIFF saved to: {outPath}");
    }
}

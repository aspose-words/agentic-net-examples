using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a folder for output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Build a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for lighter grayscale TIFF conversion.");

        // Configure TIFF save options with a threshold of 100.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression (suitable for 1‑bpp images).
            TiffCompression = TiffCompression.Ccitt4,
            // Render the page in grayscale.
            ImageColorMode = ImageColorMode.Grayscale,
            // Apply Floyd‑Steinberg dithering with a custom threshold.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 100
        };

        // Save the document as a TIFF image.
        string outPath = Path.Combine(artifactsDir, "GrayscaleThreshold.tiff");
        doc.Save(outPath, options);

        // Verify that the file was created.
        if (!File.Exists(outPath))
            throw new Exception("Failed to create the TIFF file.");

        // Indicate successful completion.
        Console.WriteLine($"TIFF saved to: {outPath}");
    }
}

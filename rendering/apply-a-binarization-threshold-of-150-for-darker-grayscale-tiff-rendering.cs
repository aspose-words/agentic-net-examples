using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Binarized.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample text for TIFF rendering with a binarization threshold.");

        // Configure TIFF save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression (suitable for black‑and‑white images).
            TiffCompression = TiffCompression.Ccitt4,

            // Render the document as grayscale before binarization.
            ImageColorMode = ImageColorMode.Grayscale,

            // Apply Floyd‑Steinberg dithering with a custom threshold of 150.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 150,

            // Optional: set resolution for better quality.
            Resolution = 300
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Indicate success.
        Console.WriteLine($"TIFF file saved successfully to: {outputPath}");
    }
}

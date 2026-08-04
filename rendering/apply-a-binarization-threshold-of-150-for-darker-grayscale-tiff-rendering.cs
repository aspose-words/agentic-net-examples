using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for TIFF rendering with binarization.");

        // Configure TIFF save options with a darker grayscale binarization threshold.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression (suitable for binary images).
            TiffCompression = TiffCompression.Ccitt4,
            // Apply Floyd‑Steinberg dithering with a custom threshold.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 150,
            // Render the page as grayscale before binarization.
            ImageColorMode = ImageColorMode.Grayscale
        };

        // Save the document as a TIFF image.
        string outputPath = Path.Combine(artifactsDir, "Binarized.tiff");
        doc.Save(outputPath, tiffOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Indicate successful completion.
        Console.WriteLine($"TIFF file saved to: {outputPath}");
    }
}

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
        builder.Writeln("This is a sample document for TIFF binarization.");

        // Configure TIFF rendering options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression for black‑and‑white images.
            TiffCompression = TiffCompression.Ccitt4,
            // Render the page as a grayscale image before binarization.
            ImageColorMode = ImageColorMode.Grayscale,
            // Apply Floyd‑Steinberg dithering with a custom threshold.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 150
        };

        // Save the document as a TIFF file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Binarized.tiff");
        doc.Save(outputPath, options);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, indicate success (no console input required).
        Console.WriteLine("TIFF file saved to: " + outputPath);
    }
}

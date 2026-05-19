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
        builder.Writeln("This is a sample document for TIFF rendering.");
        builder.Writeln("The image will be rendered in black and white with a custom binarization threshold.");

        // Configure TIFF save options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression (common for B&W TIFF).
            TiffCompression = TiffCompression.Ccitt4,
            // Render as black and white.
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // Apply Floyd‑Steinberg dithering with a custom threshold of 150.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 150
        };

        // Save the document as a TIFF image.
        string outputPath = Path.Combine(artifactsDir, "Binarized.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Optionally, output the result path.
        Console.WriteLine($"TIFF file saved to: {outputPath}");
    }
}

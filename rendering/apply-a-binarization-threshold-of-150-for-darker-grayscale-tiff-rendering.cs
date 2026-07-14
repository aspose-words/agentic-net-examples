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
        builder.Writeln("This is a sample document for TIFF rendering.");
        builder.Writeln("The following line will be rendered in darker grayscale.");
        builder.Font.Size = 24;
        builder.Writeln("Dark grayscale text.");

        // Configure TIFF save options with binarization threshold 150.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression (suitable for B&W images).
            TiffCompression = TiffCompression.Ccitt3,
            // Apply Floyd‑Steinberg dithering with a custom threshold.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 150,
            // Render as black‑and‑white (1 bpp) image.
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Save the document as a TIFF image.
        string tiffPath = Path.Combine(outputDir, "Rendered.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("TIFF file was not created.", tiffPath);

        // Optionally, report success.
        Console.WriteLine($"TIFF image saved to: {tiffPath}");
    }
}

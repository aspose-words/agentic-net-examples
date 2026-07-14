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
        builder.Writeln("Hello World! This is a sample document for TIFF rendering.");

        // Configure ImageSaveOptions for TIFF with Floyd‑Steinberg dithering and a higher threshold.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 200
        };

        // Save the document as a TIFF image.
        string outPath = Path.Combine(artifactsDir, "Dithered.tiff");
        doc.Save(outPath, options);

        // Validate that the file was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Output the result path and file size.
        Console.WriteLine($"TIFF saved to {outPath} ({new FileInfo(outPath).Length} bytes).");
    }
}

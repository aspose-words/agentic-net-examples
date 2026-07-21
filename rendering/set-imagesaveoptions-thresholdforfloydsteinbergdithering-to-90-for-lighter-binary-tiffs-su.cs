using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for OCR preprocessing.");

        // Configure TIFF save options with Floyd‑Steinberg dithering and a light threshold.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 90
        };

        // Save the document as a TIFF image.
        string outPath = Path.Combine(outputDir, "Sample_OCR.tiff");
        doc.Save(outPath, options);

        // Verify that the file was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        Console.WriteLine($"TIFF saved to: {outPath}");
    }
}

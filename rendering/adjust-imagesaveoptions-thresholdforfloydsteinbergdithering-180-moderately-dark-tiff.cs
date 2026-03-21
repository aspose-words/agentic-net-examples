using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Prepare a temporary image file (1x1 pixel PNG).
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeExample");
        Directory.CreateDirectory(tempDir);
        string imagePath = Path.Combine(tempDir, "Sample.png");

        // Base64-encoded 1x1 transparent PNG.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X3WcAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new document and add the image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(imagePath);

        // Set up ImageSaveOptions for TIFF output with Floyd‑Steinberg dithering.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 180
        };

        // Prepare output directory.
        string outputDir = Path.Combine(tempDir, "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Converted.tiff");

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        Console.WriteLine($"TIFF saved to: {outputPath}");
    }
}

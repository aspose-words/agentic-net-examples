using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create temporary directories.
        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeOutput");
        string imageDir = Path.Combine(Path.GetTempPath(), "AsposeImages");
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(imageDir);

        // Create a small PNG image (1x1 pixel, white) if it doesn't exist.
        string imagePath = Path.Combine(imageDir, "sample.png");
        if (!File.Exists(imagePath))
        {
            // Base64 for a 1x1 white PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK5cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for OCR preprocessing.");
        builder.InsertImage(imagePath);

        // Configure image save options for a binary TIFF suitable for OCR.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 90
        };

        // Save the document as a TIFF image with the specified options.
        string outputPath = Path.Combine(outputDir, "OcrReady.tiff");
        doc.Save(outputPath, options);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

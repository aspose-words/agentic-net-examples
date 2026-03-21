using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample text.
        builder.Writeln("Sample document for TIFF dithering.");

        // Insert a tiny PNG image from an embedded base64 string.
        byte[] imageBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=");
        builder.InsertImage(imageBytes);

        // Ensure the output directory exists.
        Directory.CreateDirectory("Output");

        // Configure ImageSaveOptions for TIFF output with Floyd‑Steinberg dithering.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 150
        };

        // Save the document as a binary TIFF using the configured options.
        doc.Save(Path.Combine("Output", "DitheredImage.tiff"), options);
    }
}

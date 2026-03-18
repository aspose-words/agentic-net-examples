using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Define directories (adjust as needed for your environment).
        string imageDir = Path.Combine(Environment.CurrentDirectory, "Images");
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(imageDir);
        Directory.CreateDirectory(artifactsDir);

        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF conversion.");
        builder.InsertImage(Path.Combine(imageDir, "SampleImage.jpg"));

        // Configure TIFF save options to use Floyd‑Steinberg dithering with a custom threshold of 150.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT4 compression (common for 1‑bpp TIFFs).
            TiffCompression = TiffCompression.Ccitt4,
            // Apply the Floyd‑Steinberg dithering method.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            // Set the binarization error threshold to 150 (darker output).
            ThresholdForFloydSteinbergDithering = 150
        };

        // Save the document as a TIFF image with the specified binarization settings.
        doc.Save(Path.Combine(artifactsDir, "DarkerGrayscale.tiff"), tiffOptions);
    }
}

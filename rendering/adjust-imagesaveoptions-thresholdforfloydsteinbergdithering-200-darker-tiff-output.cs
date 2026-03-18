using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load or create a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content and an image to demonstrate the effect.
        builder.Writeln("Sample text for TIFF rendering with darker dithering.");
        builder.InsertImage("ImageDir\\SampleImage.jpg"); // Replace with actual image path.

        // Create ImageSaveOptions for TIFF output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT3 compression (common for 1‑bpp TIFF).
            TiffCompression = TiffCompression.Ccitt3,

            // Enable Floyd‑Steinberg dithering for binarization.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,

            // Increase the threshold to 200 to produce a darker image.
            ThresholdForFloydSteinbergDithering = 200
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save("ArtifactsDir\\DarkerDithered.tiff", options);
    }
}

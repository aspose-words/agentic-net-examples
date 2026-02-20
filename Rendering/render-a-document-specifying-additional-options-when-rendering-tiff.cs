using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add some content to demonstrate rendering.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");
        builder.InsertImage("SampleImage.png"); // Replace with an actual image path if needed.

        // Configure TIFF rendering options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use LZW compression for a good balance between size and quality.
            TiffCompression = TiffCompression.Lzw,

            // Apply Floyd‑Steinberg dithering when binarizing (useful for CCITT compression).
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,

            // Set a higher threshold for the dithering algorithm (default is 128).
            ThresholdForFloydSteinbergDithering = 200,

            // Set image resolution to 300 DPI for high‑quality output.
            HorizontalResolution = 300,
            VerticalResolution = 300,

            // Enable high‑quality rendering (slower but better visual results).
            UseHighQualityRendering = true
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save("RenderedDocument.tiff", tiffOptions);
    }
}

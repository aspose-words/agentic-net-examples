using System;
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

        // Insert an image that will be converted to a binary TIFF.
        // Replace the path with a valid image file on your system.
        builder.InsertImage(@"Images\SampleImage.jpg");

        // Configure ImageSaveOptions for TIFF output with Floyd‑Steinberg dithering.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression (common for 1‑bpp TIFFs).
            TiffCompression = TiffCompression.Ccitt3,

            // Enable Floyd‑Steinberg dithering for binarization.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,

            // Increase the threshold to 150 to make the resulting binary image darker.
            ThresholdForFloydSteinbergDithering = 150
        };

        // Save the document as a binary TIFF using the configured options.
        doc.Save(@"Output\DitheredImage.tiff", options);
    }
}

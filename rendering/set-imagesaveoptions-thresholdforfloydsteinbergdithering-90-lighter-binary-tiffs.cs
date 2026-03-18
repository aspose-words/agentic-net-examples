using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths – replace with actual locations.
        string artifactsDir = @"C:\Output\";
        string imageDir = @"C:\Images\";

        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for OCR preprocessing.");
        builder.InsertImage(imageDir + "sample.png");

        // Configure image save options for a binary TIFF suitable for OCR.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression (common for OCR‑friendly TIFFs).
            TiffCompression = TiffCompression.Ccitt3,

            // Apply Floyd‑Steinberg dithering when converting to 1‑bpp.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,

            // Set a higher threshold (90) to produce a lighter binary image.
            ThresholdForFloydSteinbergDithering = 90
        };

        // Save the document as a TIFF image with the specified options.
        doc.Save(artifactsDir + "OcrReady.tiff", options);
    }
}

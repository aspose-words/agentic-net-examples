using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph of text.
        builder.Writeln("Sample document for TIFF binarization.");

        // Insert an image (replace with a valid path to an image file).
        builder.InsertImage("Logo.jpg");

        // -------------------------------------------------
        // Save the document as TIFF using the default Threshold binarization method.
        // -------------------------------------------------
        ImageSaveOptions thresholdOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression (suitable for 1‑bpp images).
            TiffCompression = TiffCompression.Ccitt3,
            // Explicitly set the binarization method to Threshold (default).
            TiffBinarizationMethod = ImageBinarizationMethod.Threshold
        };

        doc.Save("Threshold.tiff", thresholdOptions);

        // -------------------------------------------------
        // Save the document as TIFF using Floyd‑Steinberg dithering with a custom threshold.
        // -------------------------------------------------
        ImageSaveOptions ditherOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            // Use Floyd‑Steinberg dithering for binarization.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            // Set a custom threshold (0‑255). Higher values produce darker output.
            ThresholdForFloydSteinbergDithering = 240
        };

        doc.Save("FloydSteinberg.tiff", ditherOptions);
    }
}

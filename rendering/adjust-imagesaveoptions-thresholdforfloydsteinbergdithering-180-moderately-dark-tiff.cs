using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and add an image that will be converted to TIFF.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(@"ImageDir\Sample.jpg"); // Replace with your image path.

        // Set up ImageSaveOptions for TIFF output with Floyd‑Steinberg dithering.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression (typical for 1‑bpp TIFF files).
            TiffCompression = TiffCompression.Ccitt3,
            // Enable Floyd‑Steinberg dithering.
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            // Adjust the dithering threshold to 180 for moderately dark images.
            ThresholdForFloydSteinbergDithering = 180
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save(@"ArtifactsDir\Converted.tiff", saveOptions); // Replace with your output path.
    }
}

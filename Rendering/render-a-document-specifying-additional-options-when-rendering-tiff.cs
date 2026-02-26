using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and add some sample content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page with an image.");
        // Replace with a valid image path as needed.
        builder.InsertImage("ImageDir/Logo.jpg");

        // Configure ImageSaveOptions for TIFF rendering.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set the resolution (dots per inch).
            Resolution = 300,

            // Define the output image size in pixels (optional).
            ImageSize = new Size(2480, 3508), // Approx. A4 at 300 dpi.

            // Choose compression. Here we use CCITT4 with Floyd‑Steinberg dithering.
            TiffCompression = TiffCompression.Ccitt4,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            ThresholdForFloydSteinbergDithering = 200,

            // Enable anti‑aliasing and high‑quality rendering algorithms.
            UseAntiAliasing = true,
            UseHighQualityRendering = true,

            // Render each document page as a separate frame in a multi‑frame TIFF.
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Save the document as a TIFF file using the configured options.
        doc.Save("ArtifactsDir/RenderedDocument.tiff", tiffOptions);
    }
}

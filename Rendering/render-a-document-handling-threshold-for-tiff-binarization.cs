using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading and an image to the document.
        builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
        builder.Writeln("Hello world!");
        builder.InsertImage("ImageDir/Logo.jpg"); // Replace with the actual image path.

        // Configure TIFF save options to use the Threshold binarization method.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression (common for 1‑bpp TIFFs).
            TiffCompression = TiffCompression.Ccitt3,
            // Apply the simple threshold method for converting the image to black and white.
            TiffBinarizationMethod = ImageBinarizationMethod.Threshold
            // No additional threshold value is required for the Threshold method.
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save("ArtifactsDir/ImageSaveOptions.Threshold.tiff", tiffOptions);
    }
}

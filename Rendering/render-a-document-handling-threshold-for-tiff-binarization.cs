using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading and an image to the document.
        builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
        builder.Writeln("Hello world!");
        builder.InsertImage("ImageDir/Logo.jpg"); // Replace with actual image path.

        // Set up TIFF save options to use threshold binarization.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            TiffBinarizationMethod = ImageBinarizationMethod.Threshold
        };

        // Save the document as a TIFF file using the configured options.
        doc.Save("ArtifactsDir/ImageSaveOptions.Threshold.tiff", tiffOptions);
    }
}

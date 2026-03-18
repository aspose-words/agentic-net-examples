using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for asymmetric resolution TIFF.");

        // Initialize ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set horizontal DPI to 72 and vertical DPI to 300.
        options.HorizontalResolution = 72f; // DpiX
        options.VerticalResolution = 300f;  // DpiY

        // Save the document as a TIFF image using the configured options.
        doc.Save("AsymmetricResolution.tiff", options);
    }
}

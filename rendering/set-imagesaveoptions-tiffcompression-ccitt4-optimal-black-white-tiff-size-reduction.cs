using System;
using Aspose.Words;
using Aspose.Words.Saving;

class TiffCompressionExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // (Optional) Add some content so the TIFF is not empty.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for black‑white TIFF.");
        // Insert an image if desired.
        // builder.InsertImage("path/to/image.jpg");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the compression to CCITT4 – optimal for black‑white images.
        options.TiffCompression = TiffCompression.Ccitt4;

        // Save the document as a TIFF file using the specified options.
        doc.Save("OutputCcitt4.tiff", options);
    }
}

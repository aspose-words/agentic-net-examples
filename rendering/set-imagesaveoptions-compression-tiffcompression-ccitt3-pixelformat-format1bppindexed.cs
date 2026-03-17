using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF compression.");

        // Configure ImageSaveOptions for smallest TIFF file size.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT3 compression (monochrome, good for text).
            TiffCompression = TiffCompression.Ccitt3,
            // Render the image as 1‑bit indexed (black & white).
            PixelFormat = ImagePixelFormat.Format1bppIndexed
        };

        // Save the document as a TIFF image using the configured options.
        doc.Save("Output.tiff", options);
    }
}

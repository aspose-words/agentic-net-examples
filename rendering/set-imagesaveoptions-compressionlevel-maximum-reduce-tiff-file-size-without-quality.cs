using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load or create a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a sample image so the TIFF will contain visual content.
        // Replace with a valid path to an image file if needed.
        builder.InsertImage("sample.jpg");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the compression scheme to the strongest loss‑less option (LZW).
        // This reduces file size without degrading image quality.
        tiffOptions.TiffCompression = TiffCompression.Lzw;

        // Save the document as a TIFF image using the configured options.
        doc.Save("output.tiff", tiffOptions);
    }
}

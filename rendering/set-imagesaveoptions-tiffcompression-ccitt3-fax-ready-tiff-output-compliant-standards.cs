using System;
using Aspose.Words;
using Aspose.Words.Saving;

class FaxReadyTiffExample
{
    static void Main()
    {
        // Load or create a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for fax‑ready TIFF.");

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Set the compression to CCITT3, which is required for fax‑ready TIFF files.
        tiffOptions.TiffCompression = TiffCompression.Ccitt3;

        // Optional: set a 1‑bpp binarization method when using CCITT compression.
        // tiffOptions.TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering;

        // Save the document as a TIFF image with the specified compression.
        doc.Save("FaxReadyOutput.tiff", tiffOptions);
    }
}

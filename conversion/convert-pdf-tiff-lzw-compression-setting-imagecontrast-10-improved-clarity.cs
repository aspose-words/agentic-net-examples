using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToTiffConverter
{
    static void Main()
    {
        // Load the source PDF document.
        Document pdfDoc = new Document("Input.pdf");

        // Configure image save options for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Apply LZW compression to the TIFF file.
            TiffCompression = TiffCompression.Lzw,
            // Increase image contrast (value is a float; the API accepts any float, though typical range is 0‑1).
            ImageContrast = 10f
        };

        // Save the document as a TIFF image using the specified options.
        pdfDoc.Save("Output.tiff", tiffOptions);
    }
}

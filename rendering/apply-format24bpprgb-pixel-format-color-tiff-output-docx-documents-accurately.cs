using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure image save options for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use 24 bits per pixel RGB format.
            PixelFormat = ImagePixelFormat.Format24BppRgb,

            // Optional: set compression (LZW is default, can be changed if needed).
            // TiffCompression = TiffCompression.Lzw
        };

        // Save the document pages as a single TIFF file with the specified pixel format.
        doc.Save("Output.tiff", tiffOptions);
    }
}

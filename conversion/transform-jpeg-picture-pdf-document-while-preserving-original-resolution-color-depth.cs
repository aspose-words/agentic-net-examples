using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source JPEG image.
        string jpegPath = "input.jpg";

        // Path where the resulting PDF will be saved.
        string pdfPath = "output.pdf";

        // Create a new empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the JPEG image into the document.
        // The image is embedded with its original resolution and color depth.
        builder.InsertImage(jpegPath);

        // Configure PDF save options to preserve the original image quality.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use JPEG compression so we can control the quality.
            ImageCompression = PdfImageCompression.Jpeg,
            // 100 = best quality, no additional compression (original bytes are kept).
            JpegQuality = 100
        };

        // Disable downsampling to keep the original resolution of the image.
        pdfOptions.DownsampleOptions.DownsampleImages = false;

        // Save the document as a PDF file with the specified options.
        doc.Save(pdfPath, pdfOptions);
    }
}

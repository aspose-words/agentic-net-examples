using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("input.docx");

        // Configure PDF save options to apply high compression to images.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use JPEG compression for all images in the PDF.
            ImageCompression = PdfImageCompression.Jpeg,
            // Lower JPEG quality (0‑100) to increase compression.
            // A value around 30 provides strong compression while keeping acceptable visual quality.
            JpegQuality = 30
        };

        // Save the document as PDF using the configured options.
        doc.Save("output.pdf", pdfOptions);
    }
}

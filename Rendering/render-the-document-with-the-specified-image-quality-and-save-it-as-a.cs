using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderDocumentToPdf
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\Output.pdf";

        // Load the document (lifecycle: create/load).
        Document doc = new Document(inputPath);

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use JPEG compression for all images in the PDF.
            ImageCompression = PdfImageCompression.Jpeg,

            // Set the desired JPEG quality (0‑100). Higher values give better quality.
            JpegQuality = 80,

            // Enable high‑quality (slower) rendering algorithms for the PDF.
            UseHighQualityRendering = true
        };

        // Save the document as PDF with the specified options (lifecycle: save).
        doc.Save(outputPath, pdfOptions);
    }
}

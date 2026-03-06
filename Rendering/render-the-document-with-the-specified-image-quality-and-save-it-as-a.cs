using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderDocumentToPdf
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("InputDocument.docx");

        // Create PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the JPEG quality for images embedded in the PDF (0‑100).
        // Higher values give better image quality but larger file size.
        pdfOptions.JpegQuality = 85;

        // Optional: use high‑quality rendering algorithms (slower but better).
        pdfOptions.UseHighQualityRendering = true;

        // Save the document as PDF using the configured options.
        doc.Save("RenderedDocument.pdf", pdfOptions);
    }
}

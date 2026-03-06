using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfRenderExample
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("InputDocument.docx");

        // Create PdfSaveOptions to customize PDF output.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set image quality for JPEG images embedded in the PDF (0‑100).
            JpegQuality = 90,

            // Enable anti‑aliasing for smoother text rendering.
            UseAntiAliasing = true,

            // Use high‑quality rendering algorithms (slower but better visual quality).
            UseHighQualityRendering = true,

            // Embed all fonts fully into the PDF to preserve appearance.
            EmbedFullFonts = true,

            // Set the PDF compliance level (e.g., PDF/A‑1b for archival).
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF using the configured options.
        doc.Save("RenderedOutput.pdf", pdfOptions);
    }
}

using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Image quality for JPEG images embedded in the PDF (0‑100).
            JpegQuality = 90,

            // Render text with anti‑aliasing and high‑quality algorithms.
            UseAntiAliasing = true,
            UseHighQualityRendering = true,

            // Embed all fonts used in the document into the PDF.
            // In older Aspose.Words versions the FontEmbeddingMode enum may not exist;
            // use the boolean property EmbedFullFonts instead.
            EmbedFullFonts = true,

            // Set PDF/A‑1b compliance (preserves visual appearance).
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

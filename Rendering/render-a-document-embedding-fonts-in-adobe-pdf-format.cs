using System;
using Aspose.Words;
using Aspose.Words.Saving;

class EmbedFontsToPdf
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("InputDocument.docx");

        // Create PdfSaveOptions to control PDF rendering.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Embed all fonts used in the document.
        // EmbedFullFonts = true embeds the complete font files (no subsetting).
        pdfOptions.EmbedFullFonts = true;

        // Ensure that all fonts are embedded, not only non‑standard ones.
        pdfOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;

        // Do not substitute TrueType fonts with core PDF Type 1 fonts.
        pdfOptions.UseCoreFonts = false;

        // Save the document as a PDF with the specified options.
        doc.Save("OutputDocument.pdf", pdfOptions);
    }
}

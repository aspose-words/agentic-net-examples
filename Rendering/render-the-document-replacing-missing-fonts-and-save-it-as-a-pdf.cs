using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderDocumentWithMissingFonts
{
    static void Main()
    {
        // Path to the source document (DOCX, DOC, etc.).
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\RenderedDocument.pdf";

        // Load the source document.
        Document doc = new Document(inputPath);

        // Configure PDF save options to replace missing fonts with core PDF Type 1 fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Substitute Arial, Times New Roman, Courier New and Symbol with core PDF fonts.
            UseCoreFonts = true,

            // Embed only non‑standard fonts; standard fonts are replaced by core fonts.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNonstandard
        };

        // Save the document as PDF using the configured options.
        doc.Save(outputPath, pdfOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class EmbedMissingFontsToPdf
{
    static void Main()
    {
        // Path to the source DOCX file.
        string docxPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting PDF will be saved.
        string pdfPath = @"C:\Docs\ResultDocument.pdf";

        // Load the existing DOCX document.
        Document doc = new Document(docxPath);

        // Configure PDF save options to embed all fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed the full font files (no sub‑setting) to guarantee that missing fonts are present.
            EmbedFullFonts = true,

            // Ensure that every font used in the document is embedded.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF using the configured options.
        doc.Save(pdfPath, pdfOptions);
    }
}

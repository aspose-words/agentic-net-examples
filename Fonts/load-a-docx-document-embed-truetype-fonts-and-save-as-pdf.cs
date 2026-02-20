using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Configure PDF save options to embed all TrueType fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed every font used in the document.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,
            // Ensure that the full font data (all glyphs) is embedded.
            EmbedFullFonts = true
        };

        // Save the document as PDF with the specified options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

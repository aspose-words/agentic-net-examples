using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("MyDir/Document.docx");

        // Configure PDF save options to embed fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed all fonts used in the document (including standard Windows fonts).
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,
            // Embed the complete font files (no subsetting) to preserve all glyphs.
            EmbedFullFonts = true,
            // Do not replace TrueType fonts with core PDF Type 1 fonts.
            UseCoreFonts = false
        };

        // Save the document as a PDF with the specified embedding options.
        doc.Save("ArtifactsDir/Document.EmbeddedFonts.pdf", pdfOptions);
    }
}

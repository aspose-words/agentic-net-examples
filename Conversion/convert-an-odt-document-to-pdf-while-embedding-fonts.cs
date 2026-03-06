using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source ODT document.
        Document doc = new Document("input.odt");

        // Configure PDF save options to embed all fonts fully.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed the complete font files (no subsetting) to ensure they are available in the PDF.
            EmbedFullFonts = true,
            // Explicitly request embedding of all fonts.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF with the specified options.
        doc.Save("output.pdf", pdfOptions);
    }
}

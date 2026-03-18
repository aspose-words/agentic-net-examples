using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure that all fonts used in the document are embedded in the PDF.
        // EmbedAll is the default, but we set it explicitly for clarity.
        pdfOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;

        // Embed the complete font files (no subsetting) so that every glyph is available.
        // This increases file size but guarantees visual fidelity on any platform.
        pdfOptions.EmbedFullFonts = true;

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

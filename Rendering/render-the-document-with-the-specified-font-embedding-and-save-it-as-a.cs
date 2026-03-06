using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Create PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Specify the font embedding mode.
        // Options: EmbedAll, EmbedNonstandard, EmbedNone.
        pdfOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNonstandard;

        // Determine whether to embed full fonts (true) or subset them (false).
        pdfOptions.EmbedFullFonts = true;

        // Save the document as a PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

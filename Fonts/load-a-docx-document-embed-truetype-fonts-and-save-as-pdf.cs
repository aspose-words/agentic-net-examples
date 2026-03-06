using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk.
        Document doc = new Document("input.docx");

        // Enable embedding of TrueType fonts when the document is saved.
        FontInfoCollection fontInfos = doc.FontInfos;
        fontInfos.EmbedTrueTypeFonts = true;   // embed TrueType fonts
        fontInfos.SaveSubsetFonts = false;     // embed the full font (no subsetting)

        // Configure PDF save options to embed fonts into the PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.EmbedFullFonts = true;                     // embed full font files
        pdfOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll; // embed all fonts

        // Save the document as a PDF with the specified options.
        doc.Save("output.pdf", pdfOptions);
    }
}

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

        // Enable embedding of TrueType fonts in the document.
        FontInfoCollection fontInfos = doc.FontInfos;
        fontInfos.EmbedTrueTypeFonts = true;
        fontInfos.EmbedSystemFonts = true;
        fontInfos.SaveSubsetFonts = true;

        // Configure PDF save options to embed full fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.EmbedFullFonts = true;                     // embed complete fonts (no subsetting)
        pdfOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll; // embed all fonts

        // Save the document as PDF with the specified options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

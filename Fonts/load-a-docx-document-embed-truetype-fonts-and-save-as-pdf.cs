using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = "output.pdf";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Enable embedding of TrueType fonts in the document.
        FontInfoCollection fontInfos = doc.FontInfos;
        fontInfos.EmbedTrueTypeFonts = true;   // Embed TrueType fonts.
        fontInfos.SaveSubsetFonts = false;     // Embed full fonts (no subsetting).

        // Configure PDF save options to embed all fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll, // Embed all fonts.
            EmbedFullFonts = true                               // Ensure full font embedding.
        };

        // Save the document as PDF with the specified options.
        doc.Save(outputPath, pdfOptions);
    }
}

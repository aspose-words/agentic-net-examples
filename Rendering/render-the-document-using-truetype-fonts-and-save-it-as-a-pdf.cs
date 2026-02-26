using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderWithTrueTypeFonts
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\Output.pdf";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Do NOT substitute TrueType fonts with core PDF Type 1 fonts.
            UseCoreFonts = false,

            // Embed all fonts (including TrueType) into the PDF.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,

            // Embed the full font files (no subsetting) to preserve every glyph.
            EmbedFullFonts = true
        };

        // Save the document as PDF using the configured options.
        doc.Save(outputPath, pdfOptions);
    }
}

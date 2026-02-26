using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderAndSavePdf
{
    static void Main()
    {
        // Path to the folder that contains the source document.
        string inputFolder = @"C:\Docs\Input\";
        // Path to the folder where the PDF will be saved.
        string outputFolder = @"C:\Docs\Output\";

        // Load the source Word document.
        Document doc = new Document(inputFolder + "SourceDocument.docx");

        // Configure PDF save options to replace missing fonts with core PDF Type 1 fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Substitute standard TrueType fonts (Arial, Times New Roman, Courier New, Symbol)
            // with their core PDF equivalents when the original fonts are unavailable.
            UseCoreFonts = true,

            // Optionally, embed only non‑standard fonts; standard fonts will be replaced.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedNonstandard
        };

        // Save the rendered document as PDF using the configured options.
        doc.Save(outputFolder + "RenderedDocument.pdf", pdfOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = "output.pdf";

        // Load the existing DOCX document.
        Document doc = new Document(inputPath);

        // Configure PDF save options to embed all fonts (including missing ones).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure that every font used in the document is fully embedded.
            EmbedFullFonts = true,

            // Embed all fonts, not just non‑standard ones.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF using the configured options.
        doc.Save(outputPath, pdfOptions);
    }
}

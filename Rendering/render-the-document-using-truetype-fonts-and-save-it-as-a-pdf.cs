using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Create PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure that TrueType fonts are not substituted with core PDF Type 1 fonts.
        // Setting UseCoreFonts to false keeps the original TrueType fonts in the PDF.
        pdfOptions.UseCoreFonts = false;

        // Embed all fonts (including TrueType) into the PDF to guarantee correct rendering.
        pdfOptions.FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

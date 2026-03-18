using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX file.
        Document doc = new Document("Input.docx");

        // Create PDF save options and configure font embedding.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Embed all fonts (including custom fonts) into the PDF.
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the document as PDF with the specified options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

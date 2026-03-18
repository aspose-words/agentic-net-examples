using System;
using Aspose.Words;
using Aspose.Words.Saving;

class EmbedAllFontsExample
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("InputDocument.docx");

        // Create PDF save options and enable full font embedding.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // When true, the complete font files are embedded into the PDF.
            EmbedFullFonts = true
        };

        // Save the document as PDF with all used fonts fully embedded.
        doc.Save("OutputDocument.pdf", pdfOptions);
    }
}

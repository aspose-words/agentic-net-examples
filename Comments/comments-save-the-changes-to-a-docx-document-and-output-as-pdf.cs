using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Create PDF save options (customize if needed).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Example customization:
        // pdfOptions.EmbedFullFonts = true;

        // Save the document as PDF using the specified options.
        doc.Save("output.pdf", pdfOptions);
    }
}

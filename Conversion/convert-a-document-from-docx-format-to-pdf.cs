using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk.
        Document doc = new Document("input.docx");

        // Create PDF save options (optional – can be customized as needed).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as a PDF file using the specified options.
        doc.Save("output.pdf", pdfOptions);
    }
}

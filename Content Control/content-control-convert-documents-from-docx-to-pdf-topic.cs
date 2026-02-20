using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX file from disk.
        Document doc = new Document("input.docx");

        // Create PDF save options (optional – can customize PDF output here).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as a PDF file.
        doc.Save("output.pdf", pdfOptions);
    }
}

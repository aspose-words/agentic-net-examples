using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("input.docx");

        // Create PDF save options and set the desired PDF/A compliance level.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA1b; // Change to PdfA2u, PdfA3a, etc., as needed.

        // Save the document as a PDF file using the configured options.
        doc.Save("output.pdf", pdfOptions);
    }
}

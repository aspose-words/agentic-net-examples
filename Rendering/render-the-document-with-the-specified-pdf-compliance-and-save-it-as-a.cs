using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Specify the PDF compliance level (e.g., PDF/A-1b).
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the document as a PDF using the specified compliance.
        doc.Save("Output.pdf", pdfOptions);
    }
}

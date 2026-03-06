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

        // Enable PDF/UA compliance (PDF/UA-2 in this example).
        pdfOptions.Compliance = PdfCompliance.PdfUa2;

        // Export the document structure (tags) – required for PDF/UA.
        pdfOptions.ExportDocumentStructure = true;

        // Show the document title in the PDF viewer’s title bar – required for PDF/UA.
        pdfOptions.DisplayDocTitle = true;

        // Save the document as a PDF with the specified options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

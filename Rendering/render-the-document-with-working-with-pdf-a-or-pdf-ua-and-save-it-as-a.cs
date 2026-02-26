using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // -------------------------------------------------
        // Save the document as PDF/A-1b (preserves visual appearance).
        // -------------------------------------------------
        PdfSaveOptions pdfAOptions = new PdfSaveOptions();
        pdfAOptions.Compliance = PdfCompliance.PdfA1b; // Set PDF/A-1b compliance.
        doc.Save("Output_PdfA1b.pdf", pdfAOptions);

        // -------------------------------------------------
        // Save the document as PDF/UA-1 (accessibility compliant).
        // -------------------------------------------------
        PdfSaveOptions pdfUaOptions = new PdfSaveOptions();
        pdfUaOptions.Compliance = PdfCompliance.PdfUa1; // Set PDF/UA-1 compliance.
        doc.Save("Output_PdfUa1.pdf", pdfUaOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfComplianceExample
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Data\SampleDocument.docx";

        // Path where the PDF files will be saved.
        string outputPathPdfA2u = @"C:\Output\SampleDocument_PdfA2u.pdf";
        string outputPathPdfUa2 = @"C:\Output\SampleDocument_PdfUa2.pdf";

        // Load the Word document.
        Document doc = new Document(inputPath);

        // -------------------------------------------------
        // Save as PDF/A-2u (preserves visual appearance and
        // allows Unicode text extraction).
        // -------------------------------------------------
        PdfSaveOptions pdfA2uOptions = new PdfSaveOptions
        {
            // Set the compliance level to PDF/A-2u.
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the document using the configured options.
        doc.Save(outputPathPdfA2u, pdfA2uOptions);

        // -------------------------------------------------
        // Save as PDF/UA-2 (accessibility compliant PDF).
        // -------------------------------------------------
        PdfSaveOptions pdfUa2Options = new PdfSaveOptions
        {
            // Set the compliance level to PDF/UA-2.
            Compliance = PdfCompliance.PdfUa2
        };

        // Save the document using the configured options.
        doc.Save(outputPathPdfUa2, pdfUa2Options);
    }
}

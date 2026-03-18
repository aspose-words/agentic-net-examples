using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfA1bExport
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("InputDocument.docx");

        // Configure PDF save options for PDF/A‑1b compliance.
        // PDF/A compliance automatically embeds the required ICC color profile.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF/A‑1b compliant PDF.
        doc.Save("OutputDocument.pdf", pdfOptions);
    }
}

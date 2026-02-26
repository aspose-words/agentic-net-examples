using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderToPdfUa
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Ensure the page layout is up‑to‑date before rendering.
        doc.UpdatePageLayout();

        // Configure PDF/UA compliance options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set PDF/UA‑1 compliance (ISO 14289‑1).
            Compliance = PdfCompliance.PdfUa1,

            // Export document structure (tags) – required for PDF/UA.
            ExportDocumentStructure = true
        };

        // Save the document as a PDF/UA file.
        doc.Save("output.pdf", pdfOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Rebuild the page layout to ensure accurate rendering.
        doc.UpdatePageLayout();

        // Create PDF save options and configure PDF/UA compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // PDF/UA-1 compliance (ISO 14289-1). Use PdfUa2 for PDF/UA-2.
            Compliance = PdfCompliance.PdfUa1,

            // Document structure is required for PDF/UA; the property is ignored for PDF/UA
            // but setting it makes the intent explicit.
            ExportDocumentStructure = true,

            // Optional: improve rendering quality (may increase processing time).
            UseHighQualityRendering = true
        };

        // Save the document as a PDF/UA compliant file.
        doc.Save("Output.pdf", pdfOptions);
    }
}

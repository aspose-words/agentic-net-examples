using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Create and configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Example: apply ZIP (Flate) compression to all text content.
        pdfOptions.TextCompression = PdfTextCompression.Flate;

        // Example: export custom document properties as XMP metadata.
        pdfOptions.CustomPropertiesExport = PdfCustomPropertiesExport.Metadata;

        // Example: set PDF compliance level (e.g., PDF/A-1b).
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the document as PDF using the configured options.
        doc.Save("output.pdf", pdfOptions);
    }
}

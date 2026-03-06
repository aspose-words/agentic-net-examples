using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document from disk.
        // The Document constructor handles the creation and loading lifecycle.
        Document doc = new Document("InputDocument.docx");

        // Create a PdfSaveOptions object to control PDF rendering.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set PDF compliance to PDF/A-1b to ensure visual fidelity.
            Compliance = PdfCompliance.PdfA1b,

            // Export the document structure (tags) for better accessibility.
            ExportDocumentStructure = true,

            // Use high‑quality rendering for drawing objects.
            UseHighQualityRendering = true
        };

        // Save the document as a PDF using the specified options.
        // The Save method with (string, SaveOptions) follows the required save rule.
        doc.Save("OutputDocument.pdf", pdfOptions);
    }
}

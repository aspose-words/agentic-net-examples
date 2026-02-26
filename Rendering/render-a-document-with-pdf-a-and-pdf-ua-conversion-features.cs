using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Prepare PDF save options for PDF/A‑4 combined with PDF/UA‑2 compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // PDF/A‑4 + PDF/UA‑2 ensures long‑term visual fidelity and accessibility.
            Compliance = PdfCompliance.PdfA4Ua2,

            // Exporting the document structure is required for PDF/UA compliance.
            ExportDocumentStructure = true
        };

        // Save the document as a PDF using the configured options.
        doc.Save("Output.pdf", saveOptions);
    }
}

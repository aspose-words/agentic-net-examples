using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Create a PdfSaveOptions object and configure it for PDF/A‑1b compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA1b; // PDF/A‑1b ensures visual fidelity.

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", saveOptions);
    }
}

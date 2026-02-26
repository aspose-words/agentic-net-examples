using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfComplianceExample
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Create a PdfSaveOptions object to configure PDF output.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the desired PDF compliance level.
        // Options include PdfCompliance.PdfA1a, PdfCompliance.PdfA1b, PdfCompliance.PdfA2a,
        // PdfCompliance.PdfA2u, PdfCompliance.PdfA3a, PdfCompliance.PdfA3u,
        // PdfCompliance.PdfA4, PdfCompliance.PdfA4f, PdfCompliance.PdfA4Ua2,
        // PdfCompliance.PdfUa1, PdfCompliance.PdfUa2, etc.
        saveOptions.Compliance = PdfCompliance.PdfA1b; // Example: PDF/A-1b compliance.

        // Save the document as a PDF using the configured compliance level.
        doc.Save("Output.pdf", saveOptions);
    }
}

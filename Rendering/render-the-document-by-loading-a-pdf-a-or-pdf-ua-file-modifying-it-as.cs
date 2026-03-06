using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing PDF/A or PDF/UA document.
        // The Document constructor automatically detects the format.
        Document doc = new Document("Input.pdf");

        // Create a PdfSaveOptions object to control PDF saving behavior.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the desired PDF compliance level.
        // Example: PDF/A-2u (preserves visual appearance and allows text extraction).
        saveOptions.Compliance = PdfCompliance.PdfA2u;

        // Save the document as a PDF using the specified compliance settings.
        doc.Save("Output.pdf", saveOptions);
    }
}

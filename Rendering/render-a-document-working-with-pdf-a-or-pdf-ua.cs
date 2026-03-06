using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document from disk.
        Document doc = new Document("Input.docx");

        // Create a PdfSaveOptions instance to control PDF conversion.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the PDF compliance level.
        // Use PdfA1b for PDF/A‑1b compliance (visual appearance preservation).
        // Alternatively, use PdfUa1 for PDF/UA‑1 compliance (accessibility).
        saveOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the document as a PDF file that conforms to the selected standard.
        doc.Save("Output_PdfA1b.pdf", saveOptions);
    }
}

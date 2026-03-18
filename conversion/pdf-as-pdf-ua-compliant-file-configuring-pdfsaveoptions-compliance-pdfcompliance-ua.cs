using Aspose.Words;
using Aspose.Words.Saving;

class PdfUaExport
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Create PDF save options and set the compliance level to PDF/UA‑1.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfUa1
        };

        // Save the document as a PDF/UA‑compliant file.
        doc.Save("Output.PdfUa.pdf", pdfOptions);
    }
}

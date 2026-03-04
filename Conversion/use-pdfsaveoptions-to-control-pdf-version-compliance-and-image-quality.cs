using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Create a PdfSaveOptions instance to customize PDF output.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the PDF compliance level (e.g., PDF/A-1b for archival quality).
        pdfOptions.Compliance = PdfCompliance.PdfA1b;

        // Control the quality of JPEG images embedded in the PDF (0‑100).
        pdfOptions.JpegQuality = 80;

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

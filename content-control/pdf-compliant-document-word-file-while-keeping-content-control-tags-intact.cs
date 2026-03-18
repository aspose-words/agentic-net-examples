using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("input.docx");

        // Configure PDF/A‑1a compliance and preserve content‑control (SDT) tags.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // PDF/A‑1a requires a tagged PDF (includes document structure).
            Compliance = PdfCompliance.PdfA1a,

            // Export the document structure so that tags are retained.
            ExportDocumentStructure = true,

            // Preserve form fields (including content controls) in the PDF.
            PreserveFormFields = true,

            // Use the SDT tag as the name of the form field in the PDF.
            UseSdtTagAsFormFieldName = true
        };

        // Save the document as a PDF/A‑1a file.
        doc.Save("output.pdf", pdfOptions);
    }
}

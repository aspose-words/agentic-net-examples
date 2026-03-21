using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world!");

        // Configure PDF/A‑1a compliance and preserve content‑control tags.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a,
            ExportDocumentStructure = true,
            PreserveFormFields = true,
            UseSdtTagAsFormFieldName = true,
            UpdateFields = false // Prevent processing of fields that may cause errors.
        };

        // Save the document as a PDF/A‑1a file.
        doc.Save("output.pdf", pdfOptions);
    }
}

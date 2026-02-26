using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello PDF/UA world!");

        // Set the document title – required for PDF/UA compliance.
        doc.BuiltInDocumentProperties.Title = "PDF/UA Sample";

        // Configure PDF save options for PDF/UA compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Use PDF/UA‑2 (ISO 14289‑2:2024) compliance.
            Compliance = PdfCompliance.PdfUa2,
            // Show the document title in the PDF viewer's title bar (required by PDF/UA).
            DisplayDocTitle = true
        };

        // Save the document as a PDF with the specified options.
        doc.Save("OutputPdfUa.pdf", saveOptions);
    }
}

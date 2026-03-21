using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfA1bExport
{
    static void Main()
    {
        // Create a simple in‑memory document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, PDF/A‑1b world!");

        // Configure PDF save options for PDF/A‑1b compliance.
        // PDF/A compliance automatically embeds the required ICC color profile.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF/A‑1b compliant PDF.
        doc.Save("OutputDocument.pdf", pdfOptions);

        Console.WriteLine("PDF/A‑1b file created successfully: OutputDocument.pdf");
    }
}

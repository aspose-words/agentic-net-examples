using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple DOCX document in memory
        Document document = new Document();
        DocumentBuilder builder = new DocumentBuilder(document);
        builder.Writeln("Hello, Aspose.Words! This document will be saved as PDF/A‑2u.");

        // Set PDF save options for PDF/A‑2u compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the document as PDF/A‑2u
        document.Save("ConvertedDocument.pdf", pdfOptions);

        Console.WriteLine("Document successfully converted to PDF/A‑2u: ConvertedDocument.pdf");
    }
}

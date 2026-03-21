using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocToPdfWithCustomPageSize
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a custom page size (in points, 1 point = 1/72 inch).
        builder.PageSetup.PageWidth = 620;   // ~8.61 inches
        builder.PageSetup.PageHeight = 480;  // ~6.67 inches
        builder.PageSetup.PaperSize = PaperSize.Custom;

        // Add sample content.
        builder.Writeln("Hello, this is a test document with a custom page size.");

        // Determine output path in the current directory.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "ResultDocument.pdf");

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PageLayout = PdfPageLayout.SinglePage
        };

        // Save the document as PDF.
        doc.Save(pdfPath, pdfOptions);
        Console.WriteLine($"PDF saved to: {pdfPath}");
    }
}

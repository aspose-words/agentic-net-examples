using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocToPdfWithCustomPageSize
{
    static void Main()
    {
        // Path to the source DOC file.
        string docPath = @"C:\Docs\SourceDocument.docx";

        // Path where the resulting PDF will be saved.
        string pdfPath = @"C:\Docs\ResultDocument.pdf";

        // Load the existing Word document.
        Document doc = new Document(docPath);

        // Use DocumentBuilder to modify the page setup of the first (or current) section.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a custom page size in points (1 point = 1/72 inch).
        // Example: 620 x 480 points (~8.61" x 6.67").
        builder.PageSetup.PageWidth = 620;
        builder.PageSetup.PageHeight = 480;

        // Inform the layout that a custom size is being used.
        builder.PageSetup.PaperSize = PaperSize.Custom;

        // Create PdfSaveOptions – additional PDF‑specific settings can be configured here.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: keep the default single‑page layout.
            PageLayout = PdfPageLayout.SinglePage
        };

        // Save the document as PDF using the configured options.
        doc.Save(pdfPath, pdfOptions);
    }
}

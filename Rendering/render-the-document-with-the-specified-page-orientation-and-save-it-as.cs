using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderWithOrientation
{
    static void Main()
    {
        // Path where the PDF will be saved.
        string outputPath = @"C:\Output\DocumentWithLandscape.pdf";

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and set page orientation.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document is rendered in landscape orientation.");

        // Set the page orientation to Landscape (wide and short).
        builder.PageSetup.Orientation = Orientation.Landscape;

        // Ensure the layout is up‑to‑date before saving.
        doc.UpdatePageLayout();

        // Create PDF save options (default settings are sufficient for this task).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the document as a PDF file.
        doc.Save(outputPath, pdfOptions);
    }
}

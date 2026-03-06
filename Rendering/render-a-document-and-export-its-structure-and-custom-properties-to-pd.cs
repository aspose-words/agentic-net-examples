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
        builder.Writeln("Sample content for PDF export.");

        // Add a custom document property that we want to export to the PDF.
        doc.CustomDocumentProperties.Add("Company", "My Company");

        // Rebuild the page layout to ensure accurate rendering.
        doc.UpdatePageLayout();

        // Configure PDF save options:
        // - ExportDocumentStructure = true makes the PDF contain a tag structure (visible in Acrobat's Tags pane).
        // - CustomPropertiesExport = Standard stores custom properties in the PDF's /Info dictionary.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,
            CustomPropertiesExport = PdfCustomPropertiesExport.Standard
        };

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content so the PDF has visible text.
        builder.Writeln("Hello world!");

        // Add a custom document property that we want to export.
        doc.CustomDocumentProperties.Add("Company", "My value");

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export the document structure (tags) to make the PDF navigable via the Tags pane.
            ExportDocumentStructure = true,

            // Export custom properties as XMP metadata.
            CustomPropertiesExport = PdfCustomPropertiesExport.Metadata
        };

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}

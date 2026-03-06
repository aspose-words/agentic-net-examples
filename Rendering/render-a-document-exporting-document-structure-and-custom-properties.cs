using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExportExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Build simple content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Add a custom document property.
            doc.CustomDocumentProperties.Add("Company", "My value");

            // Configure PDF save options to export document structure and custom properties.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Export the logical structure (tags) so it appears in the PDF navigation pane.
                ExportDocumentStructure = true,

                // Export custom properties as standard entries in the PDF /Info dictionary.
                CustomPropertiesExport = PdfCustomPropertiesExport.Standard
            };

            // Save the document as PDF with the specified options.
            doc.Save("ExportedDocument.pdf", pdfOptions);
        }
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportDocumentStructureAndProperties
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Document Title");

        // Add a normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This document demonstrates exporting document structure and custom properties.");

        // Add a custom document property.
        doc.CustomDocumentProperties.Add("Company", "My value");

        // ---------- Save as PDF with document structure and custom properties ----------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export the logical document structure (tags) to the PDF.
            ExportDocumentStructure = true,

            // Export custom properties as standard entries in the PDF /Info dictionary.
            CustomPropertiesExport = PdfCustomPropertiesExport.Standard
        };

        // Save the document to PDF using the configured options.
        doc.Save("ExportedStructureAndProperties.pdf", pdfOptions);

        // ---------- Save as HTML with document properties ----------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export both built‑in and custom document properties to the HTML output.
            ExportDocumentProperties = true
        };

        // Save the same document to HTML using the configured options.
        doc.Save("ExportedStructureAndProperties.html", htmlOptions);
    }
}

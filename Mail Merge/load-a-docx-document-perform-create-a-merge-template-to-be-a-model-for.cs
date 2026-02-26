using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class MergeTemplateToPdf
{
    static void Main()
    {
        // Path to the source DOCX file that will be turned into a merge template.
        const string sourceDocxPath = @"C:\Docs\SourceDocument.docx";

        // Load the existing DOCX document.
        Document templateDoc = new Document(sourceDocxPath);

        // OPTIONAL: Insert merge fields into the template if they are not already present.
        // This demonstrates how a template can be prepared for future mail‑merge operations.
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.MoveToDocumentEnd();
        builder.Writeln(); // Ensure we are on a new paragraph.
        builder.InsertField("MERGEFIELD CustomerName", "<<[CustomerName]>>");
        builder.InsertField("MERGEFIELD Address", "<<[Address]>>");

        // Save the prepared document as a Word template (.dotx) – this can be reused for new reports.
        const string templateDotxPath = @"C:\Docs\MergeTemplate.dotx";
        templateDoc.Save(templateDotxPath, SaveFormat.Dotx);

        // Convert the same document (now a ready‑to‑use template) to PDF.
        const string outputPdfPath = @"C:\Docs\ResultDocument.pdf";
        templateDoc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}

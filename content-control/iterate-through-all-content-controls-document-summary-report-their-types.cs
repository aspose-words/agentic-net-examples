using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the source document that contains content controls.
        Document sourceDoc = new Document("Input.docx");

        // Build a textual summary of all content controls.
        StringBuilder summary = new StringBuilder();
        summary.AppendLine("Content Controls Summary");
        summary.AppendLine("------------------------");

        // Retrieve the collection of structured document tags (content controls).
        StructuredDocumentTagCollection sdtCollection = sourceDoc.Range.StructuredDocumentTags;

        // Iterate through each content control and record its type and title.
        foreach (IStructuredDocumentTag sdt in sdtCollection)
        {
            SdtType type = sdt.SdtType;
            string title = string.IsNullOrEmpty(sdt.Title) ? "(no title)" : sdt.Title;
            summary.AppendLine($"- Type: {type}, Title: {title}");
        }

        // Create a new document to hold the report.
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);
        builder.Writeln(summary.ToString());

        // Save the report document.
        reportDoc.Save("ContentControlsReport.docx");
    }
}

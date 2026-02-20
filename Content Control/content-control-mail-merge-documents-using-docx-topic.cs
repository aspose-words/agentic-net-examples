using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class ContentControlMailMerge
{
    static void Main()
    {
        // Load a DOCX template that contains content controls (SDTs).
        // The tags of the content controls should match the field names used in the mail merge.
        Document doc = new Document("TemplateWithContentControls.docx");

        // Enable merging into non‑MERGEFIELD elements such as content controls.
        // When this flag is true, the mail merge engine will treat content controls
        // whose Tag property matches a field name as merge targets.
        doc.MailMerge.UseNonMergeFields = true;

        // Create a data source. A DataTable is used here for simplicity,
        // but any supported source (arrays, custom IMailMergeDataSource, etc.) works.
        DataTable data = new DataTable("Employee");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Title");

        data.Rows.Add("John", "Doe", "Software Engineer");
        data.Rows.Add("Jane", "Smith", "Project Manager");

        // Perform the mail merge. Field names must correspond to the Tag values
        // of the content controls in the template.
        doc.MailMerge.Execute(data);

        // Remove any empty paragraphs that may have been left after merging.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;

        // Save the merged document.
        doc.Save("MergedResult.docx");
    }
}

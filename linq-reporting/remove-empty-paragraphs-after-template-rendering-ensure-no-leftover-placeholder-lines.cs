using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class MailMergeCleanupExample
{
    static void Main()
    {
        // Path to the Word template that contains MERGEFIELDs.
        string templatePath = @"C:\Docs\Template.docx";

        // Path where the merged document will be saved.
        string outputPath = @"C:\Docs\MergedResult.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Configure mail merge to remove any paragraphs that become empty after merging.
        // This uses the MailMergeCleanupOptions.RemoveEmptyParagraphs flag.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;

        // Optional: treat paragraphs that contain only punctuation as empty as well.
        doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

        // Prepare a simple data source for demonstration.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Rows.Add(new object[] { "John", "Doe" });
        table.Rows.Add(new object[] { "", "" });               // This row will produce empty paragraphs.
        table.Rows.Add(new object[] { "Jane", "Doe" });

        // Execute the mail merge with regions (if the template uses TableStart/TableEnd fields).
        doc.MailMerge.ExecuteWithRegions(table);

        // Save the cleaned‑up document.
        doc.Save(outputPath);
    }
}

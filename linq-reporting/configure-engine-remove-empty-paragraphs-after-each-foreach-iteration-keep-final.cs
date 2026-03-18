using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a merge field that will be filled inside the loop.
        builder.InsertField(" MERGEFIELD Name ");

        // Configure mail‑merge cleanup to remove empty paragraphs (including those that contain only punctuation).
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

        // Sample data source – a DataTable with some empty rows.
        DataTable table = new DataTable("People");
        table.Columns.Add("Name");
        table.Rows.Add(new object[] { "John Doe" });
        table.Rows.Add(new object[] { null });          // Empty value – will produce an empty paragraph.
        table.Rows.Add(new object[] { "Jane Smith" });

        // Perform mail merge for each row manually to demonstrate cleanup after each iteration.
        foreach (DataRow row in table.Rows)
        {
            // Execute mail merge for a single record.
            doc.MailMerge.Execute(new[] { "Name" }, new object[] { row["Name"] });

            // After the merge, Aspose.Words automatically removes empty paragraphs
            // because CleanupOptions was set to RemoveEmptyParagraphs.
            // No additional code is required here.
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}

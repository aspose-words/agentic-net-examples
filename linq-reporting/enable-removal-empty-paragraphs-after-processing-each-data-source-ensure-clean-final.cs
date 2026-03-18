using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // Create a new blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a mail‑merge region with two fields.
        builder.InsertField(" MERGEFIELD TableStart:Employees");
        builder.InsertField(" MERGEFIELD FirstName ");
        builder.Write(" ");
        builder.InsertField(" MERGEFIELD LastName ");
        builder.InsertField(" MERGEFIELD TableEnd:Employees");

        // First data source – includes an empty row that would generate an empty paragraph.
        DataTable dt1 = new DataTable("Employees");
        dt1.Columns.Add("FirstName");
        dt1.Columns.Add("LastName");
        dt1.Rows.Add(new object[] { "John", "Doe" });
        dt1.Rows.Add(new object[] { "", "" }); // empty values
        dt1.Rows.Add(new object[] { "Jane", "Doe" });

        // Configure mail‑merge to remove empty paragraphs (including those that contain only punctuation).
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

        // Execute the first merge; empty paragraphs are removed automatically.
        doc.MailMerge.ExecuteWithRegions(dt1);

        // Prepare a second data source.
        DataTable dt2 = new DataTable("Employees");
        dt2.Columns.Add("FirstName");
        dt2.Columns.Add("LastName");
        dt2.Rows.Add(new object[] { "Alice", "Smith" });
        dt2.Rows.Add(new object[] { null, null }); // another empty row

        // Re‑apply cleanup options before the next merge (required if options might have changed).
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

        // Insert a new region for the second merge.
        builder.MoveToDocumentEnd();
        builder.InsertParagraph();
        builder.InsertField(" MERGEFIELD TableStart:Employees");
        builder.InsertField(" MERGEFIELD FirstName ");
        builder.Write(" ");
        builder.InsertField(" MERGEFIELD LastName ");
        builder.InsertField(" MERGEFIELD TableEnd:Employees");

        // Execute the second merge; empty paragraphs are again removed.
        doc.MailMerge.ExecuteWithRegions(dt2);

        // Save the final, clean document.
        doc.Save("CleanedMailMerge.docx");
    }
}

// Enable removal of empty paragraphs that contain only merge tags (e.g., {{Field}}) during mail merge.
// The document is created, template tags are inserted, and the mail‑merge cleanup options are set
// so that after the merge any paragraph that became empty (or contains only punctuation) is removed.

using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

class RemoveEmptyParagraphsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert template tags that will be treated as merge fields.
        // These tags are in the "{{FieldName}}" format.
        builder.Writeln("{{FirstName}}");   // Paragraph 1 – will be filled or become empty.
        builder.Writeln("{{LastName}}");    // Paragraph 2 – will be filled or become empty.
        builder.Writeln("{{MiddleInitial}}"); // Paragraph 3 – may stay empty.

        // Configure mail‑merge to recognise the "{{...}}" tags as merge fields.
        doc.MailMerge.UseNonMergeFields = true;          // Enables processing of alternative tags.
        doc.MailMerge.PreserveUnusedTags = false;       // Unused tags become MERGEFIELDs that can be removed.

        // Set cleanup options:
        //   RemoveEmptyParagraphs – deletes paragraphs that are empty after the merge.
        //   CleanupParagraphsWithPunctuationMarks – also treats paragraphs that contain only punctuation as empty.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs;
        doc.MailMerge.CleanupParagraphsWithPunctuationMarks = true;

        // Prepare a data source. In this example only FirstName and LastName have values;
        // MiddleInitial is missing, so its paragraph will be empty after the merge.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        // Note: No column for "MiddleInitial".
        table.Rows.Add(new object[] { "John", "Doe" });

        // Execute the mail merge. The paragraph containing {{MiddleInitial}} will become empty
        // and will be removed because of the cleanup options set above.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}

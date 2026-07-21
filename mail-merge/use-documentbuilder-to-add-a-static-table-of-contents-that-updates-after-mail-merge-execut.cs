using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a static Table of Contents (TOC) field at the beginning of the document.
        // The switches configure the TOC to include heading levels 1‑3, add hyperlinks, hide page numbers for web view, and use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add a heading that will be populated by a mail‑merge field.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.InsertField(" MERGEFIELD Title ");
        builder.Writeln();

        // Add a normal paragraph with another mail‑merge field.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.InsertField(" MERGEFIELD Content ");

        // Prepare mail‑merge data in a DataTable.
        DataTable table = new DataTable("Data");
        table.Columns.Add("Title");
        table.Columns.Add("Content");
        table.Rows.Add("First Document Title", "This is the content of the first document.");
        table.Rows.Add("Second Document Title", "This is the content of the second document.");

        // Execute mail merge for each row; the document will be expanded for each record.
        doc.MailMerge.Execute(table);

        // After mail merge, update all fields (including the TOC) to reflect the new headings.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}

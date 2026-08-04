using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare data for the mail merge.
        DataTable data = new DataTable("Data");
        data.Columns.Add("Title");
        data.Columns.Add("Content");
        data.Rows.Add("Chapter 1", "This is the content of the first chapter.");
        data.Rows.Add("Chapter 2", "This is the content of the second chapter.");

        // Create a blank document and a builder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a static Table of Contents field.
        // Switches: include heading levels 1‑3, make entries hyperlinks, hide page numbers for hidden text, and use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Insert a heading that will be filled by the mail merge.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.InsertField(" MERGEFIELD Title ");
        builder.Writeln();

        // Insert a normal paragraph with a merge field for the body content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.InsertField(" MERGEFIELD Content ");
        builder.Writeln();

        // Execute the mail merge using the prepared data.
        doc.MailMerge.Execute(data);

        // Update all fields in the document, which refreshes the TOC to reflect the merged headings.
        doc.UpdateFields();

        // Save the final document.
        doc.Save("MergedWithToc.docx");
    }
}

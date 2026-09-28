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

        // Insert a static Table of Contents at the beginning of the document.
        // The TOC will include headings of levels 1 to 3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Add a page break after the TOC so the main content starts on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Add headings that will be captured by the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Report Title");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2");

        // Insert a merge field that will be populated by mail merge.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.InsertField("MERGEFIELD CustomerName \\* MERGEFORMAT");
        builder.Writeln();

        // Prepare mail merge data.
        DataTable data = new DataTable("Customers");
        data.Columns.Add("CustomerName");
        data.Rows.Add("John Doe");
        data.Rows.Add("Jane Smith");

        // Execute mail merge.
        doc.MailMerge.Execute(data);

        // Update all fields (including the TOC) after mail merge.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}

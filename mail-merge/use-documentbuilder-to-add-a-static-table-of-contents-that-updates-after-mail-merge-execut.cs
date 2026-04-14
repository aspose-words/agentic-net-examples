using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a static Table of Contents (TOC) field.
        // The switches configure the TOC to include heading levels 1‑3, add hyperlinks, hide page numbers in web view, and use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add content with heading styles that the TOC will reference.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Customer Report");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Customer Details");

        // Insert merge fields that will be populated by mail merge.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Write("Name: ");
        builder.InsertField(" MERGEFIELD CustomerName ");
        builder.Writeln();

        builder.Write("Address: ");
        builder.InsertField(" MERGEFIELD Address ");
        builder.Writeln();

        // Add another heading to demonstrate multiple TOC entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Summary");

        // Prepare a DataTable as the mail merge data source.
        DataTable table = new DataTable("Customers");
        table.Columns.Add("CustomerName");
        table.Columns.Add("Address");
        table.Rows.Add(new object[] { "Thomas Hardy", "120 Hanover Sq., London" });
        table.Rows.Add(new object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

        // Execute the mail merge to fill the fields.
        doc.MailMerge.Execute(table);

        // After mail merge, update all fields (including the TOC) so the TOC reflects the inserted headings.
        doc.UpdateFields();

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MailMergeWithToc.docx");
        doc.Save(outputPath);
    }
}

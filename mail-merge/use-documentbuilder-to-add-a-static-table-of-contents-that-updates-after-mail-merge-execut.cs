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

        // Insert a static Table of Contents (TOC) field.
        // The switches configure the TOC to include heading levels 1‑3, add hyperlinks, and use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add a heading that will be captured by the TOC.
        // The heading contains a MERGEFIELD which will be filled during mail merge.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("Customer: ");
        builder.InsertField("MERGEFIELD CustomerName");

        // Add a normal paragraph with another merge field.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln();
        builder.Write("Address: ");
        builder.InsertField("MERGEFIELD Address");

        // Add a second heading to demonstrate multiple TOC entries.
        builder.Writeln();
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Write("Order Details for ");
        builder.InsertField("MERGEFIELD CustomerName");

        // Prepare mail merge data.
        DataTable table = new DataTable("Data");
        table.Columns.Add("CustomerName");
        table.Columns.Add("Address");
        table.Rows.Add(new object[] { "John Doe", "123 Main St, Anytown" });
        table.Rows.Add(new object[] { "Jane Smith", "456 Oak Ave, Othertown" });

        // Execute mail merge. This will duplicate the document content for each row.
        doc.MailMerge.Execute(table);

        // After mail merge, update all fields (including the TOC) to reflect the merged data.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("MailMergeWithToc.docx");
    }
}

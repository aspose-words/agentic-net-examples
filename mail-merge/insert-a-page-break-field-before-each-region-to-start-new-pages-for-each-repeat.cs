using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define a mail merge region named "MyRegion".
        builder.InsertField(" MERGEFIELD TableStart:MyRegion");

        // Insert a PAGE_BREAK field that will be repeated for each record.
        // This ensures each repeat starts on a new page.
        builder.InsertField(" PAGEBREAK ");

        // Add some content inside the region.
        builder.Write("Name: ");
        builder.InsertField(" MERGEFIELD Name");
        builder.InsertParagraph();

        // End the region.
        builder.InsertField(" MERGEFIELD TableEnd:MyRegion");

        // Prepare a data source with several rows.
        DataTable table = new DataTable("MyRegion");
        table.Columns.Add("Name");
        table.Rows.Add("Alice");
        table.Rows.Add("Bob");
        table.Rows.Add("Charlie");

        // Execute the mail merge with regions. Each row will be placed in the region,
        // and the PAGE_BREAK field ensures each repeat starts on a new page.
        doc.MailMerge.ExecuteWithRegions(table);

        // Save the resulting document.
        doc.Save("MailMergeWithPageBreak.docx");
    }
}

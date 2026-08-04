using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a page break before the mail merge region so each record starts on a new page.
        builder.InsertBreak(BreakType.PageBreak);

        // Begin a mail merge region named "Employees".
        builder.InsertField(" MERGEFIELD TableStart:Employees");

        // Insert merge fields that will be populated from the data source.
        builder.Write("Name: ");
        builder.InsertField(" MERGEFIELD Name");
        builder.Writeln();

        builder.Write("Title: ");
        builder.InsertField(" MERGEFIELD Title");
        builder.Writeln();

        // End the mail merge region.
        builder.InsertField(" MERGEFIELD TableEnd:Employees");

        // Prepare a DataTable that matches the region name and field names.
        DataTable table = new DataTable("Employees");
        table.Columns.Add("Name");
        table.Columns.Add("Title");
        table.Rows.Add(new object[] { "John Doe", "Manager" });
        table.Rows.Add(new object[] { "Jane Smith", "Developer" });
        table.Rows.Add(new object[] { "Bob Johnson", "Designer" });

        // Execute the mail merge with regions. The page break ensures each employee record starts on a new page.
        doc.MailMerge.ExecuteWithRegions(table);

        // Save the resulting document.
        doc.Save("MailMergeWithPageBreakPerRegion.docx");
    }
}

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

        // Start a table that will serve as the mail merge region.
        builder.StartTable();

        // Insert the TableStart field. This marks the beginning of the region named "MyTable".
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableStart:MyTable ");

        // Insert cells that contain the merge fields for each column.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Name ");

        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Quantity ");

        // Insert the TableEnd field. It must be in the same row as TableStart.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableEnd:MyTable ");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Prepare a DataTable that matches the region name and column names.
        DataTable table = new DataTable("MyTable");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Quantity", typeof(int));

        table.Rows.Add("Apples", 10);
        table.Rows.Add("Bananas", 5);
        table.Rows.Add("Cherries", 12);

        // Execute the mail merge with regions. The table rows will be repeated for each record.
        doc.MailMerge.ExecuteWithRegions(table);

        // Save the resulting document.
        doc.Save("MailMergeWithTableRegion.docx");
    }
}

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

        // Start a table that will act as a mail merge region placeholder.
        builder.StartTable();

        // First cell – start of the mail merge region named "Orders".
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableStart:Orders ");

        // Second cell – merge field for the first column.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD OrderID ");

        // Third cell – merge field for the second column.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD Quantity ");

        // Fourth cell – end of the mail merge region.
        builder.InsertCell();
        builder.InsertField(" MERGEFIELD TableEnd:Orders ");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Prepare a DataTable that matches the region name and column names.
        DataTable orders = new DataTable("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("Quantity", typeof(int));

        // Add sample rows.
        orders.Rows.Add(1001, 5);
        orders.Rows.Add(1002, 2);
        orders.Rows.Add(1003, 9);

        // Execute mail merge with regions – the table will be expanded for each row.
        doc.MailMerge.ExecuteWithRegions(orders);

        // Save the result to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MailMergeTableRegion.docx");
        doc.Save(outputPath);
    }
}

using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        DataTable data = new DataTable("ReportData");
        data.Columns.Add("Product", typeof(string));
        data.Columns.Add("Quantity", typeof(int));
        data.Rows.Add("Apples", 120);
        data.Rows.Add("Bananas", 85);
        data.Rows.Add("Oranges", 60);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title of the report.
        builder.Writeln("Sales Report");
        builder.Writeln();

        // Start a table and add a header row.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Insert a row for each data record.
        foreach (DataRow row in data.Rows)
        {
            builder.InsertCell();
            builder.Write(row["Product"].ToString());

            builder.InsertCell();
            builder.Write(row["Quantity"].ToString());

            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Add a summary paragraph with the total quantity.
        int totalQuantity = data.AsEnumerable().Sum(r => Convert.ToInt32(r["Quantity"]));
        builder.Writeln();
        builder.Writeln($"Total quantity: {totalQuantity}");

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
    }
}

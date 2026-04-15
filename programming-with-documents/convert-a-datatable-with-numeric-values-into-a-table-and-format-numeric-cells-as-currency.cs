using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a DataTable with sample numeric data.
        DataTable tableData = new DataTable("Products");
        tableData.Columns.Add("Item", typeof(string));
        tableData.Columns.Add("Price", typeof(decimal));

        tableData.Rows.Add("Apple", 1.25m);
        tableData.Rows.Add("Banana", 0.80m);
        tableData.Rows.Add("Cherry", 2.50m);
        tableData.Rows.Add("Date", 3.10m);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Font.Bold = true;
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Data rows.
        foreach (DataRow row in tableData.Rows)
        {
            // Item cell.
            builder.InsertCell();
            builder.Font.Bold = false;
            builder.Write(row["Item"].ToString());

            // Price cell – format as currency.
            builder.InsertCell();
            decimal price = (decimal)row["Price"];
            builder.Write(price.ToString("C")); // Uses current culture currency format.

            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFromDataTable.docx");
        doc.Save(outputPath);
    }
}

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

        // Sample data source.
        DataTable table = new DataTable("Products");
        table.Columns.Add("Product");
        table.Columns.Add("Quantity", typeof(int));

        table.Rows.Add("Apples", 30);
        table.Rows.Add("Bananas", 45);
        table.Rows.Add("Oranges", 25);

        // Insert a table with a header row.
        builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Font.Bold = true;
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        int totalQuantity = 0;
        foreach (DataRow row in table.Rows)
        {
            builder.InsertCell();
            builder.Font.Bold = false;
            builder.Write(row["Product"].ToString());

            builder.InsertCell();
            int qty = Convert.ToInt32(row["Quantity"]);
            builder.Write(qty.ToString());

            builder.EndRow();

            totalQuantity += qty;
        }

        builder.EndTable();

        // Add a summary paragraph after the table.
        builder.Writeln();
        builder.Font.Bold = true;
        builder.Writeln($"Total quantity: {totalQuantity}");

        // Save the report document.
        doc.Save("Report.docx");
    }
}

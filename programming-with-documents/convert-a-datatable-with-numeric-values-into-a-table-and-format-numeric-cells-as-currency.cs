using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a DataTable with numeric values.
        DataTable table = new DataTable("Products");
        table.Columns.Add("Item", typeof(string));
        table.Columns.Add("Price", typeof(decimal));

        table.Rows.Add("Apple", 1.25m);
        table.Rows.Add("Banana", 0.75m);
        table.Rows.Add("Cherry", 2.50m);
        table.Rows.Add("Date", 3.10m);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with a header row.
        builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Populate the table with DataTable rows.
        foreach (DataRow row in table.Rows)
        {
            // Item cell.
            builder.InsertCell();
            builder.Write(row["Item"].ToString());

            // Price cell formatted as currency.
            builder.InsertCell();
            decimal price = (decimal)row["Price"];
            builder.Write(price.ToString("C")); // e.g., $1.25

            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}

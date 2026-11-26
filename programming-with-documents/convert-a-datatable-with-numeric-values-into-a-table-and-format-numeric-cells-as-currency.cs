using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a DataTable with numeric values.
        DataTable dataTable = new DataTable("Products");
        dataTable.Columns.Add("Item", typeof(string));
        dataTable.Columns.Add("Quantity", typeof(int));
        dataTable.Columns.Add("UnitPrice", typeof(decimal));

        dataTable.Rows.Add("Apple", 10, 0.5m);
        dataTable.Rows.Add("Banana", 5, 0.3m);
        dataTable.Rows.Add("Orange", 8, 0.7m);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table wordTable = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Unit Price");
        builder.EndRow();

        // Populate table rows from the DataTable.
        foreach (DataRow row in dataTable.Rows)
        {
            builder.InsertCell();
            builder.Write(row["Item"].ToString());

            builder.InsertCell();
            builder.Write(row["Quantity"].ToString());

            builder.InsertCell();
            // Format numeric value as currency.
            decimal price = (decimal)row["UnitPrice"];
            builder.Write(price.ToString("C"));
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}

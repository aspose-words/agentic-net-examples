using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a DataTable with sample numeric data.
        DataTable table = new DataTable("Products");
        table.Columns.Add("Item", typeof(string));
        table.Columns.Add("Price", typeof(decimal));

        table.Rows.Add("Apple", 1.25m);
        table.Rows.Add("Banana", 0.75m);
        table.Rows.Add("Carrot", 0.60m);
        table.Rows.Add("Doughnut", 1.50m);

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table wordTable = builder.StartTable();

        // Insert header row.
        builder.InsertCell();
        builder.Font.Bold = true;
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Reset bold for data rows.
        builder.Font.Bold = false;

        // Populate the table with DataTable rows.
        foreach (DataRow row in table.Rows)
        {
            // Item cell.
            builder.InsertCell();
            builder.Write(row["Item"].ToString());

            // Price cell formatted as currency.
            builder.InsertCell();
            decimal price = (decimal)row["Price"];
            builder.Write(string.Format("{0:C}", price));

            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "TableFromDataTable.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}

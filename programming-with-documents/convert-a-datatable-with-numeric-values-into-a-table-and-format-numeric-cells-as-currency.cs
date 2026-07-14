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
        DataTable table = new DataTable("Products");
        table.Columns.Add("Item", typeof(string));
        table.Columns.Add("Price", typeof(decimal));

        table.Rows.Add("Apple", 1.25m);
        table.Rows.Add("Banana", 0.75m);
        table.Rows.Add("Cherry", 2.50m);
        table.Rows.Add("Date", 3.10m);

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

        // Populate the table with data from the DataTable.
        foreach (DataRow row in table.Rows)
        {
            // Item cell.
            builder.InsertCell();
            builder.Write(row["Item"].ToString());

            // Price cell formatted as currency.
            builder.InsertCell();
            if (row["Price"] is decimal price)
            {
                // Use the current culture's currency format.
                builder.Write(price.ToString("C"));
            }
            else
            {
                builder.Write(row["Price"].ToString());
            }

            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "DataTableToTable.docx");
        doc.Save(outputPath, SaveFormat.Docx);
    }
}

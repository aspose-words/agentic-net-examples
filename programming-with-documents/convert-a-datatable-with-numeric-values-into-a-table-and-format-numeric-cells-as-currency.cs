using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a DataTable with sample numeric data.
        DataTable table = new DataTable("Products");
        table.Columns.Add("Product", typeof(string));
        table.Columns.Add("Price", typeof(decimal));
        table.Columns.Add("Quantity", typeof(int));

        table.Rows.Add("Apple", 1.25m, 10);
        table.Rows.Add("Banana", 0.75m, 20);
        table.Rows.Add("Cherry", 2.50m, 15);

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table wordTable = builder.StartTable();

        // Insert header row.
        InsertCell(builder, "Product", true);
        InsertCell(builder, "Price", true);
        InsertCell(builder, "Quantity", true);
        builder.EndRow();

        // Insert data rows.
        foreach (DataRow row in table.Rows)
        {
            InsertCell(builder, row["Product"].ToString(), false);
            // Format the numeric value as currency.
            string price = Convert.ToDecimal(row["Price"]).ToString("C", CultureInfo.CurrentCulture);
            InsertCell(builder, price, false);
            InsertCell(builder, row["Quantity"].ToString(), false);
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DataTableToWord.docx");

        // Save the document.
        doc.Save(outputPath);
    }

    // Helper method to insert a cell with optional bold header formatting.
    private static void InsertCell(DocumentBuilder builder, string text, bool isHeader)
    {
        builder.InsertCell();
        if (isHeader)
        {
            builder.Font.Bold = true;
        }
        else
        {
            builder.Font.Bold = false;
        }
        builder.Write(text);
    }
}

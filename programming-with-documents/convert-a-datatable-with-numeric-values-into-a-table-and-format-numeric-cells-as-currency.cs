using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a DataTable with numeric values.
        DataTable table = new DataTable("Products");
        table.Columns.Add("Item", typeof(string));
        table.Columns.Add("Quantity", typeof(int));
        table.Columns.Add("UnitPrice", typeof(decimal));

        table.Rows.Add("Apple", 10, 0.75m);
        table.Rows.Add("Banana", 5, 0.50m);
        table.Rows.Add("Orange", 8, 0.65m);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table wordTable = builder.StartTable();

        // Insert header row.
        InsertHeaderCell(builder, "Item");
        InsertHeaderCell(builder, "Quantity");
        InsertHeaderCell(builder, "Unit Price");
        builder.EndRow();

        // Insert data rows.
        foreach (DataRow row in table.Rows)
        {
            // Item (string)
            builder.InsertCell();
            builder.Write(row["Item"].ToString());

            // Quantity (int) – keep as plain number.
            builder.InsertCell();
            builder.Write(row["Quantity"].ToString());

            // Unit Price (decimal) – format as currency.
            builder.InsertCell();
            decimal price = (decimal)row["UnitPrice"];
            builder.Write(price.ToString("C")); // Currency format based on current culture.

            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Optionally apply a simple style.
        wordTable.StyleIdentifier = StyleIdentifier.LightListAccent1;
        wordTable.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        doc.Save("OutputTable.docx");
    }

    // Helper method to insert a header cell with bold text.
    private static void InsertHeaderCell(DocumentBuilder builder, string text)
    {
        builder.InsertCell();
        // Make header text bold.
        builder.Font.Bold = true;
        builder.Write(text);
        // Reset bold for subsequent cells.
        builder.Font.Bold = false;
    }
}

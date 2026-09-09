using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Add sample data rows.
        AddDataRow(builder, "Apples", 20);
        AddDataRow(builder, "Bananas", 40);
        AddDataRow(builder, "Carrots", 50);
        builder.EndTable();

        // Define the threshold value.
        const int threshold = 30;

        // Apply conditional formatting: highlight cells with values above the threshold.
        // Skip the header row (row index 0).
        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            // Quantity is in the second cell (index 1).
            Cell quantityCell = row.Cells[1];
            string text = quantityCell.ToString(SaveFormat.Text).Trim();

            if (int.TryParse(text, out int value) && value > threshold)
            {
                // Highlight the cell background.
                quantityCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
            }
        }

        // Save the document to the local file system.
        doc.Save("ConditionalFormattingTable.docx");
    }

    // Helper method to add a data row to the table.
    private static void AddDataRow(DocumentBuilder builder, string item, int quantity)
    {
        builder.InsertCell();
        builder.Write(item);
        builder.InsertCell();
        builder.Write(quantity.ToString());
        builder.EndRow();
    }
}

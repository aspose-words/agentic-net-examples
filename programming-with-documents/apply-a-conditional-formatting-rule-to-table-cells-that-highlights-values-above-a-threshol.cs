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

        // Sample data rows.
        string[,] data = {
            { "Apples",  "20" },
            { "Bananas", "45" },
            { "Carrots", "30" },
            { "Dates",   "55" }
        };

        // Populate the table with the sample data.
        for (int i = 0; i < data.GetLength(0); i++)
        {
            builder.InsertCell();
            builder.Write(data[i, 0]); // Item name.
            builder.InsertCell();
            builder.Write(data[i, 1]); // Quantity as text.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Define the threshold above which cells will be highlighted.
        const int threshold = 30;

        // Iterate over the rows (skip the header row) and apply shading to cells
        // where the numeric value exceeds the threshold.
        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
        {
            Row row = table.Rows[rowIndex];
            Cell quantityCell = row.Cells[1]; // Second column holds the quantity.

            // Try to parse the cell text to an integer.
            if (int.TryParse(quantityCell.ToString(SaveFormat.Text).Trim(), out int value))
            {
                if (value > threshold)
                {
                    // Highlight the cell with a light yellow background.
                    quantityCell.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
                }
            }
        }

        // Save the resulting document.
        doc.Save("ConditionalFormattingTable.docx");
    }
}

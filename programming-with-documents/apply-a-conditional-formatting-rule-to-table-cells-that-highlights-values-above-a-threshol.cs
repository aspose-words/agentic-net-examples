using System;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a header row and some numeric data.
        builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Sample data.
        string[] items = { "Apples", "Bananas", "Carrots", "Dates" };
        double[] quantities = { 20, 45, 10, 60 };

        for (int i = 0; i < items.Length; i++)
        {
            builder.InsertCell();
            builder.Write(items[i]);
            builder.InsertCell();
            builder.Write(quantities[i].ToString());
            builder.EndRow();
        }

        builder.EndTable();

        // Define the threshold for highlighting.
        double threshold = 40;

        // Apply conditional shading to cells in the Quantity column that exceed the threshold.
        Table table = doc.FirstSection.Body.Tables[0];
        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++) // Skip header row.
        {
            Cell qtyCell = table.Rows[rowIndex].Cells[1]; // Quantity column.
            string cellText = qtyCell.ToString(SaveFormat.Text).Trim();

            if (double.TryParse(cellText, out double value) && value > threshold)
            {
                qtyCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
            }
        }

        // Save the document.
        doc.Save("ConditionalFormatting.docx");
    }
}

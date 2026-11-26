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

        // Define a simple table with numeric values.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Value");
        builder.EndRow();

        // Sample data rows.
        string[] items = { "A", "B", "C", "D", "E" };
        double[] values = { 23.5, 67.2, 45.0, 82.1, 38.9 };
        double threshold = 50.0; // Highlight values above this.

        for (int i = 0; i < items.Length; i++)
        {
            builder.InsertCell();
            builder.Write(items[i]);

            builder.InsertCell();
            builder.Write(values[i].ToString());

            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Apply conditional formatting: highlight cells with values above the threshold.
        // Skip the header row (index 0) and start from the first data row.
        for (int i = 1; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];
            Cell valueCell = row.Cells[1]; // Second column holds the numeric value.
            string cellText = valueCell.GetText().Trim(); // Remove end-of-cell marker.

            if (double.TryParse(cellText, out double cellValue) && cellValue > threshold)
            {
                // Set background shading to light yellow.
                valueCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
            }
        }

        // Save the document.
        string outputPath = "ConditionalFormattingTable.docx";
        doc.Save(outputPath);
    }
}

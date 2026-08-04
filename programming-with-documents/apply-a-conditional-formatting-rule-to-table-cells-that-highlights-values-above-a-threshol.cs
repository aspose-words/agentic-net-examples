using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using System.Drawing;

namespace AsposeWordsConditionalFormatting
{
    public class Program
    {
        public static void Main()
        {
            // Define the output file path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ConditionalFormattingTable.docx");

            // Threshold for highlighting.
            const int threshold = 30;

            // Create a new blank document and a DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Item");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.EndRow();

            // Data rows.
            InsertDataRow(builder, "Apples", "20");
            InsertDataRow(builder, "Bananas", "40");
            InsertDataRow(builder, "Carrots", "50");

            // Finish the table.
            builder.EndTable();

            // Apply conditional formatting: highlight quantity cells with values above the threshold.
            // Use index-based loop because Row does not have a RowIndex property.
            for (int i = 1; i < table.Rows.Count; i++) // start from 1 to skip header row
            {
                Row row = table.Rows[i];
                // The quantity is in the second cell (index 1).
                Cell quantityCell = row.Cells[1];

                // Extract the cell text.
                string cellText = quantityCell.ToString(SaveFormat.Text).Trim();

                // Try to parse the numeric value.
                if (int.TryParse(cellText, out int value) && value > threshold)
                {
                    // Highlight the cell background.
                    quantityCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
                }
            }

            // Save the document.
            doc.Save(outputPath);
        }

        // Helper method to insert a data row into the table.
        private static void InsertDataRow(DocumentBuilder builder, string item, string quantity)
        {
            builder.InsertCell();
            builder.Writeln(item);
            builder.InsertCell();
            builder.Writeln(quantity);
            builder.EndRow();
        }
    }
}

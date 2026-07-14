using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsConditionalFormatting
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define a threshold value.
            double threshold = 30.0;

            // Build a simple table with a header row and some numeric values.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Row 1.
            builder.InsertCell();
            builder.Write("Apples");
            builder.InsertCell();
            builder.Write("20");
            builder.EndRow();

            // Row 2.
            builder.InsertCell();
            builder.Write("Bananas");
            builder.InsertCell();
            builder.Write("40");
            builder.EndRow();

            // Row 3.
            builder.InsertCell();
            builder.Write("Carrots");
            builder.InsertCell();
            builder.Write("55");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Apply conditional formatting: highlight cells with values above the threshold.
            // Skip the header row (row index 0).
            for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
            {
                Row row = table.Rows[rowIndex];
                // The quantity is in the second cell (index 1).
                Cell quantityCell = row.Cells[1];
                string cellText = quantityCell.ToString(SaveFormat.Text).Trim();

                if (double.TryParse(cellText, out double value) && value > threshold)
                {
                    // Highlight the cell background with yellow color.
                    quantityCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
                }
            }

            // Save the document to a file.
            doc.Save("ConditionalFormattingTable.docx");
        }
    }
}

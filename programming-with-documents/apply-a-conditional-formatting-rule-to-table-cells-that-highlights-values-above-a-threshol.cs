using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsConditionalFormatting
{
    public class Program
    {
        public static void Main()
        {
            // Define the output folder and ensure it exists.
            string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(artifactsDir);

            // Create a new blank document and a DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with a header row and some numeric values.
            builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Data rows.
            AddDataRow(builder, "Apples", 20);
            AddDataRow(builder, "Bananas", 40);
            AddDataRow(builder, "Carrots", 55);
            AddDataRow(builder, "Dates", 15);

            builder.EndTable();

            // Define the threshold above which cells will be highlighted.
            const int threshold = 30;

            // Iterate over the table rows (skip the header) and apply shading to cells
            // where the numeric value exceeds the threshold.
            Table table = doc.FirstSection.Body.Tables[0];
            for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
            {
                Row row = table.Rows[rowIndex];
                Cell quantityCell = row.Cells[1]; // Second column holds the quantity.

                // Try to parse the cell text as an integer.
                if (int.TryParse(quantityCell.GetText().Trim(), out int value))
                {
                    if (value > threshold)
                    {
                        // Highlight the cell with a light yellow background.
                        quantityCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
                    }
                }
            }

            // Save the document.
            string outputPath = Path.Combine(artifactsDir, "ConditionalFormattingTable.docx");
            doc.Save(outputPath);
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
}

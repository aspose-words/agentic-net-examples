using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ConditionalCellShadingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple table with numeric values.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // Data rows.
            AddDataRow(builder, "Apples", 30);
            AddDataRow(builder, "Bananas", 70);
            AddDataRow(builder, "Cherries", 45);
            AddDataRow(builder, "Dates", 90);
            AddDataRow(builder, "Elderberries", 20);

            builder.EndTable();

            // Apply conditional shading: values > 50 get LightSalmon, otherwise LightGreen.
            foreach (Row row in table.Rows)
            {
                // Skip the header row (the first row of the table).
                if (row == table.FirstRow) continue;

                Cell valueCell = row.Cells[1];
                string text = valueCell.ToString(SaveFormat.Text).Trim();

                if (int.TryParse(text, out int numericValue))
                {
                    if (numericValue > 50)
                        valueCell.CellFormat.Shading.BackgroundPatternColor = Color.LightSalmon;
                    else
                        valueCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
                }
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalCellShading.docx");
            doc.Save(outputPath);
        }

        // Helper method to add a data row with an item name and a numeric quantity.
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

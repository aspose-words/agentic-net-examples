using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableFooterExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and keep a reference to it.
            Table table = builder.StartTable();

            // ---------- Header row ----------
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Quantity");
            builder.EndRow();

            // ---------- Data rows ----------
            string[] items = { "Apples", "Bananas", "Carrots" };
            int[] quantities = { 20, 40, 50 };

            for (int i = 0; i < items.Length; i++)
            {
                builder.InsertCell();
                builder.Write(items[i]);
                builder.InsertCell();
                builder.Write(quantities[i].ToString());
                builder.EndRow();
            }

            // ---------- Calculate column sum ----------
            int totalQuantity = 0;
            // Skip the header row (index 0) when summing.
            for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
            {
                Row row = table.Rows[rowIndex];
                // The quantity is in the second cell (index 1).
                string cellText = row.Cells[1].ToString(SaveFormat.Text).Trim();
                if (int.TryParse(cellText, out int value))
                {
                    totalQuantity += value;
                }
            }

            // ---------- Footer row ----------
            // Apply formatting for the footer row.
            builder.RowFormat.Height = 20;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            builder.Font.Bold = true;

            builder.InsertCell();
            builder.Write("Total");
            builder.InsertCell();
            builder.Write(totalQuantity.ToString());
            builder.EndRow();

            // Reset formatting to defaults for any further content.
            builder.RowFormat.ClearFormatting();
            builder.CellFormat.ClearFormatting();
            builder.Font.Bold = false;

            // End the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithFooter.docx");
            doc.Save(outputPath);
        }
    }
}

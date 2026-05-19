using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableCellShadingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3‑row, 4‑column table.
            Table table = builder.StartTable();

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Iterate through all cells and apply background colors based on column index.
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                Row row = table.Rows[rowIndex];
                for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
                {
                    Cell cell = row.Cells[colIndex];
                    // Example rule: even columns get LightBlue, odd columns get LightGreen.
                    if (colIndex % 2 == 0)
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
                    else
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
                }
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableCellShading.docx");
            doc.Save(outputPath);
        }
    }
}

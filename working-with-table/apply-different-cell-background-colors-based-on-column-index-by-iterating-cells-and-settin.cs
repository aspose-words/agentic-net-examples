using System;
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

            // Start a table and keep a reference to it.
            Table table = builder.StartTable();

            // Build a sample 3x4 table.
            for (int row = 1; row <= 3; row++)
            {
                for (int col = 1; col <= 4; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row}C{col}");
                }
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Iterate through all cells and apply background colors based on column index.
            foreach (Row row in table.Rows)
            {
                for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
                {
                    Cell cell = row.Cells[colIndex];

                    // Example: even columns get LightBlue, odd columns get LightGray.
                    if (colIndex % 2 == 0)
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
                    else
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
            }

            // Save the document.
            doc.Save("ColoredTable.docx");
        }
    }
}

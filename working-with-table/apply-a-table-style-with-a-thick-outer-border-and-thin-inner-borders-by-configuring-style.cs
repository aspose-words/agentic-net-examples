using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleBordersExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a 3x3 table.
            Table table = builder.StartTable();

            // Fill the table with sample text.
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Remove any existing borders.
            table.ClearBorders();

            // Apply thick outer borders (2 points).
            table.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);

            // Apply thin inner borders (0.5 points) to each cell.
            for (int i = 0; i < table.Rows.Count; i++)
            {
                Row row = table.Rows[i];
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    Cell cell = row.Cells[j];
                    // Right border for all but the last column.
                    if (j < row.Cells.Count - 1)
                    {
                        cell.CellFormat.Borders[BorderType.Right].LineStyle = LineStyle.Single;
                        cell.CellFormat.Borders[BorderType.Right].LineWidth = 0.5;
                        cell.CellFormat.Borders[BorderType.Right].Color = Color.Black;
                    }
                    // Bottom border for all but the last row.
                    if (i < table.Rows.Count - 1)
                    {
                        cell.CellFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Single;
                        cell.CellFormat.Borders[BorderType.Bottom].LineWidth = 0.5;
                        cell.CellFormat.Borders[BorderType.Bottom].Color = Color.Black;
                    }
                }
            }

            // Save the document to the local file system.
            doc.Save("TableStyleBorders.docx");
        }
    }
}

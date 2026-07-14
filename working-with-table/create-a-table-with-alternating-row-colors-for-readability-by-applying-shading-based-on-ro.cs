using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            int rows = 6;    // Number of rows to create.
            int columns = 3; // Number of columns per row.

            // Build the table rows and cells.
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {r + 1}, Cell {c + 1}");
                }
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Apply alternating shading based on row index parity.
            for (int i = 0; i < table.Rows.Count; i++)
            {
                Row row = table.Rows[i];
                // Choose a background color: LightGray for even rows, White for odd rows.
                Color bgColor = (i % 2 == 0) ? Color.LightGray : Color.White;

                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = bgColor;
                }
            }

            // Save the document to the local file system.
            string outputPath = "AlternatingRows.docx";
            doc.Save(outputPath);
        }
    }
}

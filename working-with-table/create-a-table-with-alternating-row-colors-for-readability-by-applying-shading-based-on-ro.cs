using System;
using System.IO;
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

            // Start building the table.
            Table table = builder.StartTable();

            int rowCount = 6;   // Number of rows to create.
            int colCount = 3;   // Number of columns per row.

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                // Populate cells for the current row.
                for (int colIndex = 0; colIndex < colCount; colIndex++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {rowIndex + 1}, Col {colIndex + 1}");
                }

                // Finish the current row.
                builder.EndRow();

                // Apply shading based on row index parity.
                // Even rows (0‑based) get LightGray, odd rows get White.
                Color background = (rowIndex % 2 == 0) ? Color.LightGray : Color.White;
                Row currentRow = table.LastRow; // The row just added.
                foreach (Cell cell in currentRow.Cells)
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = background;
                }
            }

            // Complete the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRowsTable.docx");
            doc.Save(outputPath);

            // Optional verification (no interactive output required).
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");
        }
    }
}

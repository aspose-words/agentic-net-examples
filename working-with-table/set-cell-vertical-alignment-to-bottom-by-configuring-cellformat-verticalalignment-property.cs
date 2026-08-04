using System;
using System.IO;
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

            // Build a 2x2 table where each cell's text is aligned to the bottom.
            for (int row = 0; row < 2; row++)
            {
                for (int col = 0; col < 2; col++)
                {
                    // Insert a new cell.
                    builder.InsertCell();

                    // Set vertical alignment for the current cell.
                    builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;

                    // Write some sample text.
                    builder.Write($"Row {row + 1}, Cell {col + 1}");
                }

                // End the current row.
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Verify that every cell has the Bottom vertical alignment.
            foreach (Row r in table.Rows)
            {
                foreach (Cell c in r.Cells)
                {
                    if (c.CellFormat.VerticalAlignment != CellVerticalAlignment.Bottom)
                        throw new InvalidOperationException("A cell does not have Bottom vertical alignment.");
                }
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellVerticalAlignmentBottom.docx");
            doc.Save(outputPath);
        }
    }
}

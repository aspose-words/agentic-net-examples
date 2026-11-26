using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace MergeVerticalCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3‑row, 2‑column table.
            builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("Row 1, Col 1");
            builder.InsertCell();
            builder.Write("Row 1, Col 2");
            builder.EndRow();

            // Row 2
            builder.InsertCell();
            builder.Write("Row 2, Col 1");
            builder.InsertCell();
            builder.Write("Row 2, Col 2");
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("Row 3, Col 1");
            builder.InsertCell();
            builder.Write("Row 3, Col 2");
            builder.EndRow();

            // End the table construction.
            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Merge cells vertically in the first column.
            // The first cell gets the "First" flag, subsequent cells get "Previous".
            if (table.Rows.Count > 0)
            {
                // First cell in the column.
                Cell firstCell = table.Rows[0].Cells[0];
                firstCell.CellFormat.VerticalMerge = CellMerge.First;

                // Remaining cells in the same column.
                for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
                {
                    Cell cell = table.Rows[rowIndex].Cells[0];
                    cell.CellFormat.VerticalMerge = CellMerge.Previous;
                }
            }

            // Save the document.
            string outputPath = "MergedCells.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");
        }
    }
}

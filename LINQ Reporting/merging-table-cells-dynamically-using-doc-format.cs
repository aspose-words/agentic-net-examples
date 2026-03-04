using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCellMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with three rows and three columns.
            Table table = builder.StartTable();

            // -----------------------------------------------------------------
            // Row 1 – Demonstrate horizontal merging of the first two cells.
            // -----------------------------------------------------------------
            // First cell – mark it as the start of a horizontally merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Horizontally merged cells");

            // Second cell – merge it with the previous cell.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Third cell – normal, not merged.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Normal cell");
            builder.EndRow();

            // -----------------------------------------------------------------
            // Row 2 – Demonstrate vertical merging of the first column.
            // -----------------------------------------------------------------
            // First cell – start of a vertically merged range.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Vertically merged cells");

            // Second cell – normal.
            builder.InsertCell();
            builder.Write("Normal cell");

            // Third cell – normal.
            builder.InsertCell();
            builder.Write("Normal cell");
            builder.EndRow();

            // -----------------------------------------------------------------
            // Row 3 – Continue the vertical merge started in Row 2.
            // -----------------------------------------------------------------
            // First cell – merge with the cell above.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            // Second cell – normal.
            builder.InsertCell();
            builder.Write("Normal cell");

            // Third cell – normal.
            builder.InsertCell();
            builder.Write("Normal cell");
            builder.EndRow();

            // End the table construction.
            builder.EndTable();

            // Save the document in DOC format.
            doc.Save("MergedCellsDemo.doc");

            // -----------------------------------------------------------------
            // Load an existing document and programmatically merge cells based on data.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document("MergedCellsDemo.doc");
            Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];

            // Example: merge cells in the second row if the cell text contains "Normal".
            Row secondRow = loadedTable.Rows[1];
            for (int i = 0; i < secondRow.Cells.Count; i++)
            {
                Cell cell = secondRow.Cells[i];
                string text = cell.ToString(SaveFormat.Text).Trim();

                if (text.Equals("Normal cell", StringComparison.OrdinalIgnoreCase))
                {
                    // Merge this cell with the previous one horizontally.
                    cell.CellFormat.HorizontalMerge = CellMerge.First;
                    if (i + 1 < secondRow.Cells.Count)
                        secondRow.Cells[i + 1].CellFormat.HorizontalMerge = CellMerge.Previous;
                    break; // Merge only the first pair for demonstration.
                }
            }

            // Save the modified document.
            loadedDoc.Save("MergedCellsDemo_Modified.doc");
        }
    }
}

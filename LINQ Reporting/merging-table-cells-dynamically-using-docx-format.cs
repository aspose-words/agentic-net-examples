using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMergeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Parameters that control merging.
            int columnsToMergeHorizontally = 3; // Merge first 3 cells of the first row.
            int rowsToMergeVertically = 2;     // Merge first 2 cells of the first column.

            // ---------- First Row (Horizontal Merge) ----------
            // Insert the first cell – it will be the start of the merged range.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Horizontally merged cells");

            // Insert the remaining cells that belong to the horizontal merge.
            for (int i = 1; i < columnsToMergeHorizontally; i++)
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                // No text needed for merged cells.
            }

            // Add an extra unmerged cell to the same row for demonstration.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Unmerged cell");

            // End the first row.
            builder.EndRow();

            // ---------- Subsequent Rows ----------
            // Build the remaining rows. The first column will be vertically merged.
            for (int rowIndex = 1; rowIndex < rowsToMergeVertically; rowIndex++)
            {
                // First cell of the row – part of vertical merge.
                builder.InsertCell();
                builder.CellFormat.VerticalMerge = CellMerge.Previous;
                // No text needed for merged cells.

                // Add the same number of cells as the first row.
                for (int col = 1; col <= columnsToMergeHorizontally; col++)
                {
                    builder.InsertCell();
                    builder.CellFormat.VerticalMerge = CellMerge.None;
                    builder.Write($"R{rowIndex + 1}, C{col + 1}");
                }

                // Add the extra unmerged cell.
                builder.InsertCell();
                builder.CellFormat.VerticalMerge = CellMerge.None;
                builder.Write($"R{rowIndex + 1}, C{columnsToMergeHorizontally + 2}");

                builder.EndRow();
            }

            // Add one more row that is not part of the vertical merge.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("New row, first cell");

            for (int col = 1; col <= columnsToMergeHorizontally; col++)
            {
                builder.InsertCell();
                builder.Write($"R{rowsToMergeVertically + 2}, C{col + 1}");
            }

            builder.InsertCell();
            builder.Write($"R{rowsToMergeVertically + 2}, C{columnsToMergeHorizontally + 2}");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document in DOCX format.
            doc.Save("MergedTable.docx");
        }
    }
}

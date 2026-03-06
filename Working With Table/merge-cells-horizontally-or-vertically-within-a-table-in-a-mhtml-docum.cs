using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableMerge
{
    class Program
    {
        static void Main()
        {
            // Load an existing MHTML document.
            // The path should point to a valid MHTML file that contains at least one table.
            Document doc = new Document("InputDocument.mhtml");

            // Get the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // -------------------------------------------------
            // Horizontal merge: merge the first two cells of the first row.
            // -------------------------------------------------
            // Set the first cell as the start of the merged range.
            Cell firstCell = table.Rows[0].Cells[0];
            firstCell.CellFormat.HorizontalMerge = CellMerge.First;
            firstCell.CellFormat.VerticalMerge = CellMerge.None; // Ensure vertical merge is not set.

            // Set the second cell to merge with the previous cell.
            Cell secondCell = table.Rows[0].Cells[1];
            secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;
            secondCell.CellFormat.VerticalMerge = CellMerge.None;

            // -------------------------------------------------
            // Vertical merge: merge the first cell of the first row with the cell directly below it.
            // -------------------------------------------------
            // Set the top cell as the start of the vertical merge.
            firstCell.CellFormat.VerticalMerge = CellMerge.First;
            // The cell below (row 1, column 0) will merge with the previous cell vertically.
            Cell belowCell = table.Rows[1].Cells[0];
            belowCell.CellFormat.VerticalMerge = CellMerge.Previous;
            // Ensure horizontal merge flags are cleared for these cells.
            belowCell.CellFormat.HorizontalMerge = CellMerge.None;

            // -------------------------------------------------
            // Save the modified document back to MHTML format.
            // -------------------------------------------------
            doc.Save("MergedOutput.mhtml", SaveFormat.Mhtml);
        }
    }
}

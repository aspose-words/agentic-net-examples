using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeTableCellsInMhtml
{
    static void Main()
    {
        // Load the MHTML document.
        Document doc = new Document("InputDocument.mhtml");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // ---------- Horizontal merge ----------
        // Merge the first two cells of the first row.
        // The leftmost cell becomes the first cell in the merged range.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;

        // The cell to the right merges with the previous cell.
        Cell secondCell = table.Rows[0].Cells[1];
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // ---------- Vertical merge ----------
        // Merge the first two cells of the first column (rows 0 and 1).
        // The top cell becomes the first cell in the merged range.
        Cell topCell = table.Rows[0].Cells[0];
        topCell.CellFormat.VerticalMerge = CellMerge.First;

        // The cell below merges with the previous cell vertically.
        Cell bottomCell = table.Rows[1].Cells[0];
        bottomCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the modified document back to MHTML format.
        doc.Save("MergedOutput.mhtml");
    }
}

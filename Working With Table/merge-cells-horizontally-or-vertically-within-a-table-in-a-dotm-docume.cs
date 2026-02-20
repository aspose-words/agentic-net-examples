using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeTableCells
{
    static void Main()
    {
        // Load the DOTM template.
        Document doc = new Document("Template.dotm");

        // Assume the document contains at least one table.
        Table table = doc.FirstSection.Body.Tables[0];

        // ---------- Horizontal merge ----------
        // Merge the first two cells of the first row.
        Cell firstCell = table.Rows[0].Cells[0];
        Cell secondCell = table.Rows[0].Cells[1];

        // Mark the first cell as the start of a horizontal merge range.
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        // Mark the second cell as merged to the previous cell.
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // ---------- Vertical merge ----------
        // Merge the first two cells of the first column (rows 0 and 1).
        Cell topCell = table.Rows[0].Cells[0];
        Cell bottomCell = table.Rows[1].Cells[0];

        // Mark the top cell as the start of a vertical merge range.
        topCell.CellFormat.VerticalMerge = CellMerge.First;
        // Mark the bottom cell as merged to the previous cell vertically.
        bottomCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the modified document.
        doc.Save("MergedCells.dotm");
    }
}

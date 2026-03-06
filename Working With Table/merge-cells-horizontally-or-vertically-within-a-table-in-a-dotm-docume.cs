using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeTableCells
{
    static void Main()
    {
        // Load the existing DOTM template.
        // Replace "Template.dotm" with the actual path to your macro-enabled document.
        Document doc = new Document("Template.dotm");

        // Assume the document contains at least one table.
        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // -------------------------------------------------
        // Horizontal merge: merge the first two cells of the first row.
        // -------------------------------------------------
        // Set the first cell as the start of the merged range.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;

        // Set the second cell to merge with the previous cell.
        Cell secondCell = table.Rows[0].Cells[1];
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // -------------------------------------------------
        // Vertical merge: merge the first cell of the first row with the cell directly below it.
        // -------------------------------------------------
        // Set the top cell as the start of the vertical merge.
        firstCell.CellFormat.VerticalMerge = CellMerge.First;

        // Set the cell in the second row, same column, to merge with the previous (above) cell.
        Cell belowCell = table.Rows[1].Cells[0];
        belowCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // -------------------------------------------------
        // Save the modified document.
        // The output will also be a DOTM to preserve macros.
        // -------------------------------------------------
        doc.Save("MergedOutput.dotm");
    }
}

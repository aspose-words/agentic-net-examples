using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableCellMergeExample
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("Input.docm");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // ---------- Horizontal merge ----------
        // Merge the first two cells of the first row.
        // The leftmost cell becomes the start of the merged range.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;

        // The cell to the right merges into the previous cell.
        Cell secondCell = table.Rows[0].Cells[1];
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // ---------- Vertical merge ----------
        // Merge the first cell of the first three rows vertically.
        // The top cell starts the merged range.
        Cell topCell = table.Rows[0].Cells[0];
        topCell.CellFormat.VerticalMerge = CellMerge.First;

        // The cells below merge into the previous (top) cell.
        for (int rowIndex = 1; rowIndex <= 2; rowIndex++)
        {
            Cell cell = table.Rows[rowIndex].Cells[0];
            cell.CellFormat.VerticalMerge = CellMerge.Previous;
        }

        // Save the modified document as a DOCM.
        doc.Save("Output.docm");
    }
}

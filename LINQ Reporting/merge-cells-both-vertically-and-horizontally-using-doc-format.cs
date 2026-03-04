using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeCellsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // ---- First row, first cell (top‑left) ----
        // This cell will be the first cell of a 2×2 merged block.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // Start horizontal merge.
        builder.CellFormat.VerticalMerge   = CellMerge.First;   // Start vertical merge.
        builder.Write("Merged 2×2");

        // ---- First row, second cell (top‑right) ----
        // Merge horizontally with the cell to the left.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue horizontal merge.
        builder.CellFormat.VerticalMerge   = CellMerge.None;     // No vertical merge for this cell.
        // No text needed; this cell is merged.

        // End the first row.
        builder.EndRow();

        // ---- Second row, first cell (bottom‑left) ----
        // Merge vertically with the cell above.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;     // No horizontal merge for this cell.
        builder.CellFormat.VerticalMerge   = CellMerge.Previous; // Continue vertical merge.
        // No text needed; this cell is merged.

        // ---- Second row, second cell (bottom‑right) ----
        // This cell is independent (not merged).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Independent cell");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document in DOC format.
        doc.Save("MergedCells.doc", SaveFormat.Doc);
    }
}

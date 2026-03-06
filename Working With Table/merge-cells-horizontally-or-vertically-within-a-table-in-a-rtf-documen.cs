using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeTableCellsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        builder.StartTable();

        // ---------- First Row ----------
        // Insert first cell and mark it as the first cell of a horizontal merge range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // Horizontal merge start
        builder.CellFormat.VerticalMerge = CellMerge.First;     // Vertical merge start
        builder.Write("Top‑Left merged cell");

        // Insert second cell and merge it horizontally with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Horizontal merge continuation
        builder.CellFormat.VerticalMerge = CellMerge.None;       // No vertical merge for this cell
        // No text needed – the cell is merged.

        // Insert third cell – normal, not merged.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Insert first cell – this cell will be merged vertically with the cell above.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.Previous; // Vertical merge continuation
        // No text needed – the cell is merged.

        // Insert second cell – normal cell (not part of any merge).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Second row, second cell");

        // Insert third cell – normal cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Second row, third cell");

        // End the second row and the table.
        builder.EndTable();

        // Save the document as RTF.
        doc.Save("MergedCells.rtf");
    }
}

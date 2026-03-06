using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeTableCells
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // ---------- First row ----------
        // Insert first cell and mark it as the start of a horizontal merge range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged horizontally (col 1‑2)");

        // Insert second cell and merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged cell.

        // Insert third cell (not merged).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Cell 3");

        // Finish the first row.
        builder.EndRow();

        // ---------- Second row ----------
        // First cell will be the start of a vertical merge range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged vertically (row 2‑3)");

        // Second cell merges vertically with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for the merged cell.

        // Third cell (not merged).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Cell 3");

        // Finish the second row.
        builder.EndRow();

        // ---------- Third row ----------
        // Continue the vertical merge for the first column.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for the merged cell.

        // Second cell (not merged).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Cell 2");

        // Third cell (not merged).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Cell 3");

        // Finish the third row.
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Save the document to a DOCX file.
        doc.Save("MergedCells.docx");
    }
}

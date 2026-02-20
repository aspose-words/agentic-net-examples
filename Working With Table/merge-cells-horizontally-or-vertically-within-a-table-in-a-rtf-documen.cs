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
        Table table = builder.StartTable();

        // ---------- Horizontal merge (first row) ----------
        // First cell – start of a horizontally merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged cells");

        // Second cell – merged with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged cell.
        builder.EndRow();

        // Reset merge settings for subsequent rows.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // ---------- Vertical merge (first column) ----------
        // First row, first cell – start of a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged cells");

        // First row, second cell – normal cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row, first cell – merged with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for the merged cell.
        // Second row, second cell – normal cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Reset vertical merge for any further rows.
        builder.CellFormat.VerticalMerge = CellMerge.None;

        // End the table.
        builder.EndTable();

        // Save the document as an RTF file.
        doc.Save("MergedCells.rtf");
    }
}

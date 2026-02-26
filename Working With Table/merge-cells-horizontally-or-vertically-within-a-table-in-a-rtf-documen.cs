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
        builder.StartTable();

        // -------------------------------------------------
        // Horizontal merge in the first row.
        // -------------------------------------------------
        // First cell – marks the start of a horizontally merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged cells");

        // Second cell – merges with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Third cell – normal (not merged).
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Normal cell");

        // End the first row.
        builder.EndRow();

        // -------------------------------------------------
        // Vertical merge in the first column.
        // -------------------------------------------------
        // First cell of the second row – merges vertically with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text is needed for a vertically merged cell.

        // Second cell of the second row – regular cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Second row, second column");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as an RTF file.
        doc.Save("MergedCells.rtf");
    }
}

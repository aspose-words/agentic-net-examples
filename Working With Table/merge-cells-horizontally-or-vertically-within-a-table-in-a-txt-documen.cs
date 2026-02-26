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

        // ---------- First row ----------
        // Cell (0,0) – start of a horizontally merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        // Also start a vertically merged range in this cell.
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged horizontally and vertically");

        // Cell (0,1) – merge with the previous cell horizontally.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No vertical merge for this cell.
        builder.CellFormat.VerticalMerge = CellMerge.None;
        // No text needed for merged cell.
        builder.EndRow();

        // ---------- Second row ----------
        // Cell (1,0) – merge with the cell above vertically.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for merged cell.

        // Cell (1,1) – independent cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Regular cell");

        // Finish the table.
        builder.EndTable();

        // Save the document as a plain text file.
        doc.Save("MergedTable.txt");
    }
}

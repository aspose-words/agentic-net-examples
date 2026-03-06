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

        // Start a table with three columns.
        builder.StartTable();

        // ---------- First Row ----------
        // Insert first cell and mark it as the start of a horizontal merge range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged");

        // Insert second cell and merge it with the previous cell horizontally.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged cell.

        // Insert third cell – not merged.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Insert first cell and start a vertical merge range (will span two rows).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged");

        // Insert second cell and merge it vertically with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for the merged cell.

        // Insert third cell – not merged.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        doc.Save("MergedCells.docx");
    }
}

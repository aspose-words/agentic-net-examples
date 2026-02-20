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

        // Start a new table.
        Table table = builder.StartTable();

        // ---------- First row: horizontal merge ----------
        // First cell – mark as the first cell in a horizontally merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged cells");

        // Second cell – merge with the previous cell horizontally.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged cell.
        builder.EndRow();

        // ---------- Second row: start vertical merge ----------
        // First cell – mark as the first cell in a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged cells");

        // Second cell – normal, not merged.
        builder.InsertCell();
        builder.Write("Normal cell");
        builder.EndRow();

        // ---------- Third row: continue vertical merge ----------
        // First cell – merge with the cell above vertically.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for the merged cell.
        // Second cell – normal, not merged.
        builder.InsertCell();
        builder.Write("Another normal cell");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document as a plain‑text file.
        doc.Save("MergedTable.txt");
    }
}

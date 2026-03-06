using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // -------------------------------------------------
        // Row 1 – Horizontal merge of the first two cells.
        // -------------------------------------------------
        builder.InsertCell();                                   // First cell.
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // Mark as the first merged cell.
        builder.Write("Horizontally merged cells.");

        builder.InsertCell();                                   // Second cell (merged with the first).
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged cell.
        builder.EndRow();

        // -------------------------------------------------
        // Row 2 – Begin vertical merge of the first column.
        // -------------------------------------------------
        builder.InsertCell();                                   // First cell of column 1.
        builder.CellFormat.VerticalMerge = CellMerge.First;     // First cell in a vertical merge range.
        builder.Write("Vertically merged cells.");

        builder.InsertCell();                                   // Second cell of column 2 (normal).
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");
        builder.EndRow();

        // -------------------------------------------------
        // Row 3 – Continue vertical merge of the first column.
        // -------------------------------------------------
        builder.InsertCell();                                   // Continuation of the vertical merge.
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for the merged cell.

        builder.InsertCell();                                   // Normal cell in column 2.
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Another normal cell");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document in the legacy DOC format.
        doc.Save("MergedCells.doc", SaveFormat.Doc);
    }
}

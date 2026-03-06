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

        // Start a new table.
        Table table = builder.StartTable();

        // -------------------------------------------------
        // Horizontal merge example (first row).
        // -------------------------------------------------
        bool mergeFirstRow = true; // This could be any runtime condition.

        // First cell – start of the merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = mergeFirstRow ? CellMerge.First : CellMerge.None;
        builder.Write("Header spanning two columns");

        // Second cell – merges with the previous cell if the condition is true.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = mergeFirstRow ? CellMerge.Previous : CellMerge.None;
        // No text needed for a merged‑into cell.
        builder.Write(string.Empty);
        builder.EndRow();

        // -------------------------------------------------
        // Normal cells (second row) – no merging.
        // -------------------------------------------------
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Row 2, Col 1");
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.EndRow();

        // -------------------------------------------------
        // Vertical merge example (first column of rows 3‑4).
        // -------------------------------------------------
        bool mergeVertically = true; // This could be any runtime condition.

        // Row 3, first cell – start of the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = mergeVertically ? CellMerge.First : CellMerge.None;
        builder.Write("Vertically merged cell");
        builder.InsertCell();
        builder.Write("Row 3, Col 2");
        builder.EndRow();

        // Row 4, first cell – merges with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = mergeVertically ? CellMerge.Previous : CellMerge.None;
        // No text needed for a merged‑into cell.
        builder.Write(string.Empty);
        builder.InsertCell();
        builder.Write("Row 4, Col 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document in DOC format.
        doc.Save("MergedTable.doc");
    }
}

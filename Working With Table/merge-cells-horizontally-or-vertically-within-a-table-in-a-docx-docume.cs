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

        // -------------------------------------------------
        // Horizontal merge: first row, first two cells merged.
        // -------------------------------------------------
        // First cell – mark as the first cell in a horizontally merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged cells");

        // Second cell – merge with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Third cell – normal (not merged).
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Normal cell");

        // End the first row.
        builder.EndRow();

        // -------------------------------------------------
        // Vertical merge: first column of the next two rows merged.
        // -------------------------------------------------
        // Row 2, first cell – start of a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged cells");

        // Row 2, second cell – normal.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // End the second row.
        builder.EndRow();

        // Row 3, first cell – merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Row 3, second cell – normal.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // End the third row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a DOCX file.
        doc.Save("MergedCells.docx");
    }
}

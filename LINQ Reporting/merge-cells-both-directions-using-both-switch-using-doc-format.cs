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

        // Start a new table.
        Table table = builder.StartTable();

        // -------------------------------------------------
        // First row - first cell will be the start of a
        // horizontally and vertically merged range.
        // -------------------------------------------------
        builder.InsertCell();
        // Mark this cell as the first cell in a horizontal merge.
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        // Mark this cell as the first cell in a vertical merge.
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged horizontally and vertically");

        // Second cell in the first row – part of the horizontal merge.
        builder.InsertCell();
        // This cell merges with the previous cell horizontally.
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No vertical merge for this cell.
        builder.CellFormat.VerticalMerge = CellMerge.None;
        // No text needed; the content is taken from the first cell.
        builder.EndRow();

        // -------------------------------------------------
        // Second row – continue the vertical merge for the
        // first column, and add normal cells for the rest.
        // -------------------------------------------------
        builder.InsertCell();
        // This cell merges vertically with the cell above.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for merged cells.
        // Insert a normal cell in the second column.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document in DOC format.
        doc.Save("MergedCells.doc", SaveFormat.Doc);
    }
}

using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeCellsExample
{
    static void Main()
    {
        // Create a new blank document and a DocumentBuilder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // -------------------------------------------------
        // First row – horizontally merged cells (2 columns)
        // -------------------------------------------------
        // Insert the first cell. Mark it as the first cell in a horizontal merge
        // and also as the first cell in a vertical merge (we will merge this column later).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge   = CellMerge.First;
        builder.Write("Merged horizontally and vertically");

        // Insert the second cell. Merge it with the previous cell horizontally.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for a merged cell.
        // End the first row.
        builder.EndRow();

        // -------------------------------------------------
        // Second row – vertically merged cell in first column
        // -------------------------------------------------
        // Reset merge flags that are not needed for this row.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;

        // Insert the first cell of the second row. This cell continues the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for the continuation cell.

        // Insert the second cell of the second row – a normal, unmerged cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Regular cell");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("MergedCells.docx");
    }
}

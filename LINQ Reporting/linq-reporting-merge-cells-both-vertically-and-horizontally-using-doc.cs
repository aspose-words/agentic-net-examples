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

        // Start a table.
        builder.StartTable();

        // ---------- First Row ----------
        // Cell (0,0) – first cell of a 2x2 merged block.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First; // start horizontal merge
        builder.CellFormat.VerticalMerge   = CellMerge.First; // start vertical merge
        builder.Write("Merged 2x2");

        // Cell (0,1) – horizontally merged to the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write(string.Empty); // content is ignored for merged cells

        // Cell (0,2) – normal, unmerged cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Cell 0,2");

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Cell (1,0) – vertically merged to the cell above.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.Previous;
        builder.Write(string.Empty);

        // Cell (1,1) – both horizontally and vertically merged to the top‑left cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge   = CellMerge.Previous;
        builder.Write(string.Empty);

        // Cell (1,2) – normal, unmerged cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Cell 1,2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document in DOC format.
        doc.Save("MergedCells.doc");
    }
}

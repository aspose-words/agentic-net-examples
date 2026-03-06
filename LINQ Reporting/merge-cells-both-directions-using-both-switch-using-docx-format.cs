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
        builder.StartTable();

        // ----- First row -----
        // Cell (0,0) – first cell of a 2x2 merged block.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // Start horizontal merge.
        builder.CellFormat.VerticalMerge   = CellMerge.First;   // Start vertical merge.
        builder.Write("Merged cell (2x2)");

        // Cell (0,1) – continues horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge with left cell.
        builder.CellFormat.VerticalMerge   = CellMerge.None;     // No vertical merge.
        builder.Write(string.Empty); // Content is ignored for merged cells.

        // End first row.
        builder.EndRow();

        // ----- Second row -----
        // Cell (1,0) – continues vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;      // No horizontal merge.
        builder.CellFormat.VerticalMerge   = CellMerge.Previous; // Merge with cell above.
        builder.Write(string.Empty);

        // Cell (1,1) – continues both horizontal and vertical merges.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge with left cell.
        builder.CellFormat.VerticalMerge   = CellMerge.Previous; // Merge with cell above.
        builder.Write(string.Empty);

        // End second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document in DOCX format.
        doc.Save("MergedCellsBothDirections.docx");
    }
}

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

        // Start building a table.
        builder.StartTable();

        // ---- First Row ----
        // Cell (0,0) – first cell of a region that will be merged horizontally and vertically.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // start horizontal merge
        builder.CellFormat.VerticalMerge   = CellMerge.First;   // start vertical merge
        builder.Write("Merged Cell");

        // Cell (0,1) – horizontally merged with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // continue horizontal merge
        builder.CellFormat.VerticalMerge   = CellMerge.None;     // no vertical merge
        builder.Write(" "); // placeholder text (will not be displayed)

        // End the first row.
        builder.EndRow();

        // ---- Second Row ----
        // Cell (1,0) – vertically merged with the cell above.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;    // no horizontal merge
        builder.CellFormat.VerticalMerge   = CellMerge.Previous; // continue vertical merge
        builder.Write(" "); // placeholder text

        // Cell (1,1) – regular cell, not merged.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;
        builder.Write("Regular Cell");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Reset merge settings for any further content.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge   = CellMerge.None;

        // Save the document in DOC format.
        doc.Save("MergedCells.doc", SaveFormat.Doc);
    }
}

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

        // Start a table.
        Table table = builder.StartTable();

        // ---------- First row ----------
        // First cell: start of a horizontal merge and start of a vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // horizontal merge start
        builder.CellFormat.VerticalMerge = CellMerge.First;     // vertical merge start
        builder.Write("Merged horizontally and vertically");

        // Second cell: continue the horizontal merge with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // merge with left cell
        builder.CellFormat.VerticalMerge = CellMerge.None;       // no vertical merge
        builder.Write(""); // content not needed for merged cell

        // Third cell: regular, not merged.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // End the first row.
        builder.EndRow();

        // ---------- Second row ----------
        // First cell: continue the vertical merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.Previous; // merge with cell above
        builder.Write(""); // content not needed for merged cell

        // Second cell: regular cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Second row, second cell");

        // Third cell: regular cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Second row, third cell");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document as a PDF file.
        doc.Save("MergedTable.pdf", SaveFormat.Pdf);
    }
}

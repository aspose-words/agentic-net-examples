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
        Table table = builder.StartTable();

        // -------------------------------------------------
        // Horizontal merge example (first row).
        // -------------------------------------------------
        // First cell – marks the start of a horizontally merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged cells");

        // Second cell – merges with the cell to its left.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Third cell – normal, not merged.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // End the first row.
        builder.EndRow();

        // Reset horizontal merge flags so they don't affect subsequent rows.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // -------------------------------------------------
        // Vertical merge example (second column, rows 2‑3).
        // -------------------------------------------------
        // Row 2, first column.
        builder.InsertCell();
        builder.Write("Row2 Col1");

        // Row 2, second column – start of a vertically merged range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged cells");

        // End row 2.
        builder.EndRow();

        // Row 3, first column.
        builder.InsertCell();
        builder.Write("Row3 Col1");

        // Row 3, second column – merges with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text needed for merged cell.

        // End row 3.
        builder.EndRow();

        // Reset vertical merge flags for any further cells.
        builder.CellFormat.VerticalMerge = CellMerge.None;

        // Finish the table.
        builder.EndTable();

        // Save the document as a PDF file.
        doc.Save("MergedTable.pdf", SaveFormat.Pdf);
    }
}

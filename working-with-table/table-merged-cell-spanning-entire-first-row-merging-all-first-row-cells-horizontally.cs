using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeFirstRowExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        builder.StartTable();

        // ---------- First row (merged cell) ----------
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header spanning all columns");

        // Insert the remaining cells of the first row and merge them with the first cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // ---------- Second row (regular cells) ----------
        // Reset merge flags to None for normal cells.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        builder.InsertCell();
        builder.Write("Cell 1");

        builder.InsertCell();
        builder.Write("Cell 2");

        builder.InsertCell();
        builder.Write("Cell 3");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("MergedFirstRow.docx");
    }
}

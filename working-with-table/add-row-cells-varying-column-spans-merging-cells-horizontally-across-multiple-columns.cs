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
        // First row – a cell that spans two columns, then a normal cell.
        // -------------------------------------------------
        // Insert the first cell and mark it as the start of a horizontal merge range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Span 2 columns");

        // Insert the second cell and merge it with the previous one.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged cell.

        // Insert a third cell that is not merged.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // Finish the first row.
        builder.EndRow();

        // -------------------------------------------------
        // Second row – a cell that spans three columns, then a normal cell.
        // -------------------------------------------------
        // Insert the first cell of the second row and start a merge of three columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Span 3 columns");

        // Merge the next two cells with the first one.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Insert a fourth cell that is not merged.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Another cell");

        // Finish the second row.
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("MergedCellsTable.docx");
    }
}

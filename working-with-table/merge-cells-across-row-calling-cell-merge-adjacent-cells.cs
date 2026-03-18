using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document and a DocumentBuilder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a new table.
        builder.StartTable();

        // ---------- First row: merge three cells horizontally ----------
        // Insert the first cell of the merged range and mark it as the start.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("This text spans three merged cells.");

        // Insert the second cell and merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Insert the third cell and also merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // ---------- Second row: normal (unmerged) cells ----------
        // Reset merge flag to ensure cells are independent.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("MergedCells.docx");
    }
}

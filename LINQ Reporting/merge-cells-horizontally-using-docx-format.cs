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

        // Start a new table.
        Table table = builder.StartTable();

        // ---- First row: horizontally merged cells ----
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("This text spans two merged cells.");

        // Insert the second cell and merge it with the previous one.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // ---- Second row: normal (unmerged) cells ----
        // Reset merge setting to None for subsequent cells.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document in DOCX format.
        doc.Save("MergedCells.docx");
    }
}

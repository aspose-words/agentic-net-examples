using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeFirstRowCells
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("This cell spans the entire first row.");

        // Insert additional cells and merge them with the first one.
        // Each subsequent cell must be marked as 'Previous'.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Finish the first row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save("MergedFirstRowTable.docx");
    }
}

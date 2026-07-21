using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // ----- First row: a single cell that spans the full width -----
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header spanning all columns");

        // Insert additional cells and merge them with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first (merged) row.
        builder.EndRow();

        // Reset merge settings for subsequent rows.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // ----- Second row: regular separate cells -----
        builder.InsertCell();
        builder.Write("Cell 1");

        builder.InsertCell();
        builder.Write("Cell 2");

        builder.InsertCell();
        builder.Write("Cell 3");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("MergedCellTable.docx");
    }
}

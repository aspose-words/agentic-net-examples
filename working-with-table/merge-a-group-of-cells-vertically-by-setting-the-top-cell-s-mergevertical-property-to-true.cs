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

        // Start a table.
        Table table = builder.StartTable();

        // ---- First row ----
        // First cell: start a vertical merge range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged vertically");

        // Second cell: regular (not merged) cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Unmerged cell");
        builder.EndRow();

        // ---- Second row ----
        // First cell: merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text is written for merged cells.

        // Second cell: regular cell.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Unmerged cell");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("MergedCells.docx");
    }
}

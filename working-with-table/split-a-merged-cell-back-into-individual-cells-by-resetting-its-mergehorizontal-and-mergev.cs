using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to construct its content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // ---- First row with a horizontally merged cell spanning two columns ----
        // Insert the first cell and mark it as the start of a merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell");

        // Insert the second cell and mark it as merged to the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text is needed for the merged part.

        // End the first row.
        builder.EndRow();

        // ---- Second row with normal (unmerged) cells ----
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Cell 1");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Cell 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document that contains the merged cell.
        string mergedPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCell.docx");
        doc.Save(mergedPath);

        // -------- Split the merged cell back into individual cells --------
        // Access the cells of the first row (the merged ones).
        Row firstRow = table.FirstRow;
        Cell firstCell = firstRow.FirstCell;
        Cell secondCell = firstRow.LastCell;

        // Reset both horizontal and vertical merge flags to None.
        firstCell.CellFormat.HorizontalMerge = CellMerge.None;
        secondCell.CellFormat.HorizontalMerge = CellMerge.None;
        firstCell.CellFormat.VerticalMerge = CellMerge.None;
        secondCell.CellFormat.VerticalMerge = CellMerge.None;

        // Save the document after splitting the cell.
        string splitPath = Path.Combine(Directory.GetCurrentDirectory(), "SplitCell.docx");
        doc.Save(splitPath);
    }
}

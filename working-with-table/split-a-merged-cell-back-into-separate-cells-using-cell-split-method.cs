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

        // Build a table where the first row has a horizontally merged cell.
        Table table = builder.StartTable();

        // First cell – start of the merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell");

        // Second cell – merged with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.EndRow();

        // Add a second row with normal cells to keep the table valid.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Locate the merged cell (the first cell of the first row).
        Cell mergedCell = table.Rows[0].Cells[0];

        // ---- Split the merged cell ----
        // 1. Remove the merge flag from the original cell.
        mergedCell.CellFormat.HorizontalMerge = CellMerge.None;

        // 2. Remove the cell that was previously merged with it (the one having HorizontalMerge.Previous).
        //    This cell is now at index 1 in the row.
        Cell cellToRemove = table.Rows[0].Cells[1];
        cellToRemove.Remove();

        // 3. Insert a new cell after the original one.
        Cell newCell = new Cell(doc);
        table.Rows[0].InsertAfter(newCell, mergedCell);
        // Ensure the new cell contains at least one paragraph (required by Aspose.Words).
        newCell.EnsureMinimum();

        // Verify that the split produced two independent cells without merge flags.
        if (table.Rows[0].Cells.Count != 2 ||
            table.Rows[0].Cells[0].CellFormat.HorizontalMerge != CellMerge.None ||
            table.Rows[0].Cells[1].CellFormat.HorizontalMerge != CellMerge.None)
        {
            throw new InvalidOperationException("Cell split did not produce the expected separate cells.");
        }

        // Save the resulting document.
        const string outputPath = "SplitMergedCell.docx";
        doc.Save(outputPath);
    }
}

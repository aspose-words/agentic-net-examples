using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeVerticalCellsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        builder.StartTable();

        // ---------- First Row ----------
        // Insert the first cell of the first column.
        builder.InsertCell();
        // Mark this cell as the first cell in a vertically merged range.
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged vertically across three rows.");

        // Insert the second cell of the first row (independent column).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None; // Ensure no vertical merging here.
        builder.Write("Row 1, Column 2.");

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Insert the cell that continues the vertical merge.
        builder.InsertCell();
        // This cell merges with the cell directly above.
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // No text is needed for merged cells; they remain empty.
        // Insert the second column cell for this row.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 2, Column 2.");

        // End the second row.
        builder.EndRow();

        // ---------- Third Row ----------
        // Insert the third cell that continues the vertical merge.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // Insert the second column cell for this row.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Row 3, Column 2.");

        // End the third row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("MergedVerticalCells.docx");
    }
}

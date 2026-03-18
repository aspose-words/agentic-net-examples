using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableMergeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ---------- First Row ----------
        // Insert the first cell that will be the top‑left corner of the merged area.
        builder.InsertCell();
        // Mark this cell as the first cell in a horizontal and vertical merge range.
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge   = CellMerge.First;
        builder.Write("Merged 2×3 cell");

        // Insert the next two cells in the same row and merge them horizontally with the first cell.
        for (int i = 0; i < 2; i++)
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.CellFormat.VerticalMerge   = CellMerge.None; // No vertical merge for these helper cells.
            builder.Write(string.Empty);
        }

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Insert the cell that continues the vertical merge of the first cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;   // No horizontal merge needed here.
        builder.CellFormat.VerticalMerge   = CellMerge.Previous; // Merge with the cell above.
        builder.Write(string.Empty);

        // No additional cells are needed because the merged cell already spans three columns.
        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save("TableWithMergedCell.docx");
    }
}

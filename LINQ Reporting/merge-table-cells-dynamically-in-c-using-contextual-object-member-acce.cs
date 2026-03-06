using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeCellsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // ---------- First row ----------
        // Dynamically merge the first three cells horizontally.
        for (int col = 0; col < 5; col++)
        {
            builder.InsertCell();

            // Apply horizontal merge flags based on column index.
            if (col == 0)
                builder.CellFormat.HorizontalMerge = CellMerge.First;      // First cell in the merged range.
            else if (col < 3)
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;   // Merge with the previous cell.
            else
                builder.CellFormat.HorizontalMerge = CellMerge.None;       // No merge for remaining cells.

            builder.Write($"R1C{col + 1}");
        }
        builder.EndRow();

        // ---------- Second row ----------
        // Merge the first two cells vertically with the cells above.
        for (int col = 0; col < 5; col++)
        {
            builder.InsertCell();

            // Apply vertical merge flags based on column index.
            if (col == 0)
                builder.CellFormat.VerticalMerge = CellMerge.First;        // First cell in the vertical merge.
            else if (col == 1)
                builder.CellFormat.VerticalMerge = CellMerge.Previous;     // Merge with the cell above.
            else
                builder.CellFormat.VerticalMerge = CellMerge.None;         // No vertical merge for other cells.

            builder.Write($"R2C{col + 1}");
        }
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("MergedCells.docx");
    }
}

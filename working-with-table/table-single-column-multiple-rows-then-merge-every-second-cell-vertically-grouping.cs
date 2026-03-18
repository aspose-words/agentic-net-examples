using System;
using Aspose.Words;
using Aspose.Words.Tables;

class MergeEverySecondCellVertically
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single column.
        Table table = builder.StartTable();

        // Define the total number of rows (must be even to merge every second cell).
        int totalRows = 6;

        for (int rowIndex = 0; rowIndex < totalRows; rowIndex++)
        {
            // Insert a new cell (the only column in the row).
            builder.InsertCell();

            if (rowIndex % 2 == 0) // First cell of the pair – start of a vertical merge.
            {
                builder.CellFormat.VerticalMerge = CellMerge.First;
                builder.Write($"Group {(rowIndex / 2) + 1}");
            }
            else // Second cell of the pair – merge with the cell above.
            {
                builder.CellFormat.VerticalMerge = CellMerge.Previous;
                // No text is written to merged cells; they inherit the content of the first cell.
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("MergedTable.docx");
    }
}

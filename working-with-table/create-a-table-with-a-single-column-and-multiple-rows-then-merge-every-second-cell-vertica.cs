using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table that will have a single column.
        Table table = builder.StartTable();

        // Define the number of rows you want in the table.
        int totalRows = 7; // Example: 7 rows (adjust as needed).

        for (int rowIndex = 0; rowIndex < totalRows; rowIndex++)
        {
            // Insert a cell for the current row.
            builder.InsertCell();

            // Every second cell (rowIndex % 2 == 1) should merge vertically with the previous one.
            if (rowIndex % 2 == 0) // First cell of a pair (or the last odd row).
            {
                // If there is a following row, start a vertical merge range.
                if (rowIndex + 1 < totalRows)
                {
                    builder.CellFormat.VerticalMerge = CellMerge.First;
                }
                else
                {
                    // No merge needed for the final unpaired row.
                    builder.CellFormat.VerticalMerge = CellMerge.None;
                }

                // Write some content for the first cell of the merge group.
                builder.Write($"Row {rowIndex + 1}");
            }
            else // Second cell of a pair – merge with the previous cell.
            {
                builder.CellFormat.VerticalMerge = CellMerge.Previous;
                // Merged cells must be empty.
                builder.Write(string.Empty);
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableWithVerticalMerges.docx");
        doc.Save(outputPath);
    }
}

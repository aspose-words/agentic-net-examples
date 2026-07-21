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

        // Start a new table.
        Table table = builder.StartTable();

        // ---- First row (merged cell) ----
        // Insert the first cell and mark it as the start of a horizontal merge range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("This cell spans the entire first row.");

        // Insert additional cells in the same row and merge them with the previous cell.
        // The number of cells added determines the total column count of the table.
        for (int i = 0; i < 2; i++) // total 3 columns in this example
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        }

        // End the first row.
        builder.EndRow();

        // ---- Second row (regular cells) ----
        // Reset the merge flag before adding normal cells.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row 2, Cell {i + 1}");
        }

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("MergedTable.docx");
    }
}

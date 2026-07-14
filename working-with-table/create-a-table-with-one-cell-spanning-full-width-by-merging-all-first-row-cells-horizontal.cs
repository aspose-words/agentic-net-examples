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

        // Start a new table.
        Table table = builder.StartTable();

        // Insert the first cell of the first row and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header spanning all columns");

        // Insert additional cells in the same row and merge them with the first cell.
        // The number of cells determines how many columns the table will have.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        }

        // End the first (merged) row.
        builder.EndRow();

        // Add a second row with regular, unmerged cells to illustrate the table layout.
        for (int i = 0; i < 4; i++)
        {
            builder.InsertCell();
            builder.Write($"Cell {i + 1}");
        }
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
        doc.Save(outputPath);
    }
}

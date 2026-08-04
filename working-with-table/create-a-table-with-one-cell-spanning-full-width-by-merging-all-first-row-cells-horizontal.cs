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

        // ---- First row: a single cell that spans the full width ----
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("This cell spans the entire first row.");

        // Insert additional cells in the same row and merge them with the first cell.
        // The number of cells determines how many columns the table will have.
        // Here we add three more cells, but any number works.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            // No text is needed for merged cells.
        }

        // End the first row.
        builder.EndRow();

        // ---- Second row: normal unmerged cells ----
        // Reset merge setting to None for regular cells.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCellTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file at {outputPath}");
        }

        // Inform the user (optional, no interaction required).
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}

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

        // ---- First merged cell (spans two columns) ----
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged across 2 columns");

        // Insert the second cell and merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for merged cells.

        // ---- Normal cell (single column) ----
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Normal cell");

        // ---- Second merged cell (spans two columns) ----
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Another merged cell");

        // Merge the next cell with the previous one.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for merged cells.

        // End the row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Define the output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedCells.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file at {outputPath}");

        // Inform the user (no interactive prompts required).
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

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

        // Start a table.
        Table table = builder.StartTable();

        // -------------------------------------------------
        // Row with cells that have different horizontal spans.
        // -------------------------------------------------

        // First cell – start a merge that will span 2 columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged 2 columns");

        // Second cell – part of the previous merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for merged cells.

        // Third cell – normal, not merged.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // Fourth cell – start a merge that will span 3 columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged 3 columns");

        // Fifth cell – part of the previous merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Sixth cell – part of the previous merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedTable.docx");
        doc.Save(outputPath);

        // Simple validation – ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");

        // Inform the user (no interactive pause required).
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}

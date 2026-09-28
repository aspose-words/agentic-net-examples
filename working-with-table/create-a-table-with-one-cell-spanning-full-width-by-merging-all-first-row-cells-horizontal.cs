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

        // Add a paragraph before the table.
        builder.Writeln("Table with a merged first‑row cell:");

        // Start the table.
        builder.StartTable();

        // ---------- First row (merged cell) ----------
        // Insert the first cell and mark it as the start of a horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Writeln("This cell spans all columns");

        // Insert additional cells for the remaining columns and mark them as continuation of the merge.
        // Number of additional cells determines total column count. Here we create a 3‑column table.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // Reset the merge setting for subsequent rows (optional but keeps formatting clean).
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // ---------- Second row (regular cells) ----------
        builder.InsertCell();
        builder.Writeln("Cell 1");

        builder.InsertCell();
        builder.Writeln("Cell 2");

        builder.InsertCell();
        builder.Writeln("Cell 3");

        // End the second row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}

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

        // Build a simple 3‑column table with two rows.
        builder.StartTable();

        // First row – cells that will be merged.
        builder.InsertCell();
        builder.Writeln("Merged Cell Part 1");
        builder.InsertCell();
        builder.Writeln("Merged Cell Part 2");
        builder.InsertCell();
        builder.Writeln("Merged Cell Part 3");
        builder.EndRow();

        // Second row – regular cells.
        builder.InsertCell();
        builder.Writeln("Row2 Col1");
        builder.InsertCell();
        builder.Writeln("Row2 Col2");
        builder.InsertCell();
        builder.Writeln("Row2 Col3");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];
        Row firstRow = table.Rows[0];

        // Merge the first three cells horizontally using CellFormat.HorizontalMerge.
        // The first cell starts the merge, the following cells continue it.
        firstRow.Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
        firstRow.Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;
        firstRow.Cells[2].CellFormat.HorizontalMerge = CellMerge.Previous;

        // Save the document.
        string outputPath = "MergedCells.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        // Inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}

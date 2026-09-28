using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        builder.StartTable();

        // ---- First row: merged header spanning three columns ----
        // First cell - start of the merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Writeln("Header");

        // Second cell - continues the merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.Writeln(string.Empty);

        // Third cell - continues the merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.Writeln(string.Empty);

        // End the first row.
        builder.EndRow();

        // ---- Second row: regular three columns ----
        builder.InsertCell();
        builder.Writeln("Column 1");

        builder.InsertCell();
        builder.Writeln("Column 2");

        builder.InsertCell();
        builder.Writeln("Column 3");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "MergedHeaderTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}

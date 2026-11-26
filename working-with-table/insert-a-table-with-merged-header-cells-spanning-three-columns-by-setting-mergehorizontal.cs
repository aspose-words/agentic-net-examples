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

        // Start building the table.
        builder.StartTable();

        // ---------- Header row with merged cells ----------
        // First cell – start of the horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Writeln("Header spanning three columns");

        // Second cell – part of the merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Third cell – part of the merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the header row.
        builder.EndRow();

        // Reset merge setting for subsequent rows.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // ---------- Data row ----------
        builder.InsertCell();
        builder.Writeln("Cell 1");

        builder.InsertCell();
        builder.Writeln("Cell 2");

        builder.InsertCell();
        builder.Writeln("Cell 3");

        // End the data row.
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

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}

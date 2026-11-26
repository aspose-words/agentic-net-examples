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

        // Start a table with four logical columns.
        builder.StartTable();

        // ---- First cell (will span two columns) ----
        builder.InsertCell();
        builder.Writeln("Cell 1 (colspan 2)");
        // Mark this cell as the first part of a horizontal merge.
        Cell firstCell = (Cell)builder.CurrentParagraph.ParentNode;
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;

        // ---- Second cell (merged with the first) ----
        builder.InsertCell();
        // This cell is the continuation of the merge.
        Cell mergedCell = (Cell)builder.CurrentParagraph.ParentNode;
        mergedCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // ---- Third cell (normal) ----
        builder.InsertCell();
        builder.Writeln("Cell 3");

        // ---- Fourth cell (normal) ----
        builder.InsertCell();
        builder.Writeln("Cell 4");

        // End the row and the table.
        builder.EndRow();
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
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}

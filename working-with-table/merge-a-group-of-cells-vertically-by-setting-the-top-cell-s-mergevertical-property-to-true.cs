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
        builder.StartTable();

        // ---------- First Row ----------
        // Insert the first cell that will be merged vertically.
        builder.InsertCell();
        // Get the cell just created.
        Cell topCell = (Cell)builder.CurrentParagraph.ParentNode;
        // Enable vertical merge for the top cell (first cell in the merge group).
        topCell.CellFormat.VerticalMerge = CellMerge.First;
        builder.Writeln("Top merged cell");

        // Insert a regular cell in the same row.
        builder.InsertCell();
        builder.Writeln("Cell 1,2");

        // End the first row.
        builder.EndRow();

        // ---------- Second Row ----------
        // Insert the cell that continues the vertical merge.
        builder.InsertCell();
        Cell mergedContinuation = (Cell)builder.CurrentParagraph.ParentNode;
        // Continue the vertical merge for this cell (previous cell in the merge group).
        mergedContinuation.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.Writeln("Continued merged cell");

        // Insert another regular cell.
        builder.InsertCell();
        builder.Writeln("Cell 2,2");

        // End the second row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Indicate success.
        Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
    }
}

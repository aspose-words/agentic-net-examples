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

        // Start a table.
        builder.StartTable();

        // First row – simple three cells.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Second row – first cell spans two columns (horizontal merge).
        // Insert first cell and mark it as the start of a merge.
        builder.InsertCell();
        Cell firstCell = (Cell)builder.CurrentParagraph.ParentNode;
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Writeln("Spans 2 columns");

        // Insert second cell and mark it as a continuation of the previous merge.
        builder.InsertCell();
        Cell secondCell = (Cell)builder.CurrentParagraph.ParentNode;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged part.

        // Insert third cell – normal cell.
        builder.InsertCell();
        builder.Writeln("Normal cell");

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "MergedCellsTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Optionally, inform that the process completed.
        Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
    }
}

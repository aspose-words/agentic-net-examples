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

        // Start a table with 5 columns (implicitly created by inserting cells).
        builder.StartTable();

        // First row – simple header for illustration.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.InsertCell();
        builder.Writeln("Header 4");
        builder.InsertCell();
        builder.Writeln("Header 5");
        builder.EndRow();

        // Second row – cells with varying column spans.

        // Cell that spans the first two columns.
        builder.InsertCell();
        builder.Writeln("Span 2 columns");
        // Get the cell that was just created and start horizontal merge.
        Cell cell = (Cell)builder.CurrentParagraph.ParentNode;
        cell.CellFormat.HorizontalMerge = CellMerge.First;

        // Insert the merged cell (second column) and mark it as continuation.
        builder.InsertCell();
        cell = (Cell)builder.CurrentParagraph.ParentNode;
        cell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Normal single‑column cell.
        builder.InsertCell();
        builder.Writeln("Normal cell");

        // Cell that spans the last two columns.
        builder.InsertCell();
        builder.Writeln("Span 2 columns");
        cell = (Cell)builder.CurrentParagraph.ParentNode;
        cell.CellFormat.HorizontalMerge = CellMerge.First;

        // Insert the merged cell (fifth column) and mark it as continuation.
        builder.InsertCell();
        cell = (Cell)builder.CurrentParagraph.ParentNode;
        cell.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "MergedCellsTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Inform that the process completed.
        Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
    }
}

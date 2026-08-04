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

        // -----------------------------------------------------------------
        // First row – simple header cells (no merging).
        // -----------------------------------------------------------------
        builder.InsertCell();
        builder.Write("Header 1");

        builder.InsertCell();
        builder.Write("Header 2");

        builder.InsertCell();
        builder.Write("Header 3");

        builder.EndRow();

        // -----------------------------------------------------------------
        // Second row – cells with varying horizontal spans.
        // -----------------------------------------------------------------

        // Cell that spans two columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First; // start of merge range
        builder.Write("Span 2 columns");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // merged with previous cell
        // No text needed for merged cell.

        // Normal (unmerged) cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // Cell that spans three columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Span 3 columns");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text for the merged cells.

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a local file.
        string outputPath = "MergedCells.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Inform the user (optional, no interaction required).
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}

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

        // ---------- First row ----------
        // Insert the first cell and mark it as the start of a merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged across two columns");

        // Insert the second cell and merge it with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Insert a third cell that is not merged.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Third column");

        // End the first row.
        builder.EndRow();

        // ---------- Second row (regular cells) ----------
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Row 2, Col 1");
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.InsertCell();
        builder.Write("Row 2, Col 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "MergedCells.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Document was not saved.");

        // Optional: inform the user (no interactive input required).
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
